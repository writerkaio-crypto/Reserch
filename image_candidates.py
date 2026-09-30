"""画像 2 枚以下の eBay 出品について、差し替え用の画像候補を集める。

  python image_candidates.py <販売管理シートCSV> <low_image_listings CSV> [SKU ...]

SKU を指定するとその SKU だけを処理する (取得済みでも取り直す)。

販売管理シートの D 列 = SKU。対象 SKU ごとに:
  1) AM〜AQ 列の仕入れ先ページを順に開き、画像が 3 枚以上ある最初のページを採用
  2) 無ければ AS 列のメルカリ検索を「売り切れ」に切り替え、画像 3 枚以上の商品を候補として集める
     (付属品が似ているかの最終判断は人が行う)
結果は output/image_candidates.json に保存し、再実行時は取得済みの SKU を飛ばす。
"""

import asyncio
import csv
import json
import os
import re
import sys
import urllib.parse
from pathlib import Path

from playwright.async_api import async_playwright

OUT_DIR = Path(__file__).parent / "output"
RESULT_PATH = Path(os.environ.get("RESULT_PATH", OUT_DIR / "image_candidates.json"))
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"

MIN_IMAGES = 3
SUPPLIER_COLS = range(38, 43)  # AM〜AQ
MERCARI_SEARCH_COL = 44  # AS
SKU_COL = 3  # D
NAME_COL = 4  # E
MAX_SOLD_CHECK = 12  # 売り切れ検索で詳細を確認する最大件数
MAX_SOLD_CANDIDATES = 5
CONCURRENCY = 4

# 接続が許可されていない / 画像を取得できないドメイン
UNSUPPORTED = ("item.fril.jp", "shop.lashinbang.com", "www.amiami.jp",
               "ec.treasure-f.com", "netmall.hardoff.co.jp")


def ld_json_images(html):
    for m in re.finditer(r'<script[^>]*type="application/ld\+json"[^>]*>(.*?)</script>',
                         html, re.S):
        try:
            data = json.loads(m.group(1))
        except ValueError:
            continue
        for x in data if isinstance(data, list) else [data]:
            if isinstance(x, dict) and x.get("@type") == "Product":
                im = x.get("image") or []
                return [im] if isinstance(im, str) else list(im), x.get("name", "")
    return [], ""


def next_data(html):
    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', html, re.S)
    return json.loads(m.group(1)) if m else {}


async def open_page(ctx, url, capture, expect=False):
    """ページを開き、capture(url) が真になる API レスポンスの JSON を集める。

    expect=True なら、該当レスポンスが届くまで最大 30 秒待つ。
    """
    page = await ctx.new_page()
    captured = []

    async def on_response(r):
        if r.request.resource_type in ("xhr", "fetch") and capture(r.url):
            try:
                captured.append((r.url, await r.json()))
            except Exception:
                pass

    page.on("response", lambda r: asyncio.ensure_future(on_response(r)))
    try:
        resp = await page.goto(url, wait_until="domcontentloaded", timeout=60000)
        await page.wait_for_timeout(4000)
        for _ in range(26 if expect else 0):
            if captured:
                break
            await page.wait_for_timeout(1000)
        await page.wait_for_timeout(500)
        html = await page.content()
        return resp.status if resp else 0, page.url, html, captured, page
    except Exception:
        await page.close()
        raise


async def fetch_images(ctx, url):
    """{'images': [...], 'title': str, 'status': str} を返す。"""
    host = urllib.parse.urlparse(url).netloc
    if host in UNSUPPORTED:
        return {"error": "接続不可のサイト"}
    mid = re.search(r"mercari\.com/item/(m\d+)", url)
    shop = re.search(r"mercari\.com/shops/product/(\w+)", url)
    capture = (lambda u: "items/get" in u) if mid else \
              (lambda u: "shops/products/" in u) if shop else (lambda u: False)
    status, final_url, html, captured, page = await open_page(
        ctx, url, capture, expect=bool(mid or shop))
    await page.close()
    if status >= 400:
        return {"error": f"HTTP {status}"}

    if mid:
        for _, d in captured:
            data = d.get("data") or {}
            if data.get("id") == mid.group(1):
                return {"images": data.get("photos") or [], "title": data.get("name", ""),
                        "status": data.get("status", "")}
        photos = sorted(set(re.findall(
            rf"https://static\.mercdn\.net/item/detail/orig/photos/{mid.group(1)}_\d+\.jpg",
            html)), key=lambda u: int(re.search(r"_(\d+)\.jpg", u).group(1)))
        title = re.sub(r" - メルカリ$", "", re.search(r"<title>(.*?)</title>", html).group(1))
        return {"images": photos, "title": title, "status": ""}
    if shop:
        for _, d in captured:
            if d.get("name") == shop.group(1):
                return {"images": (d.get("productDetail") or {}).get("photos") or [],
                        "title": d.get("displayName", ""), "status": "shops"}
        return {"error": "商品データを取得できず"}
    if "paypayfleamarket" in host:
        item = (next_data(html).get("props", {}).get("initialState", {})
                .get("itemsState", {}).get("items", {}).get("item", {}))
        return {"images": [i["url"] for i in item.get("images") or [] if i.get("url")],
                "title": item.get("title", ""), "status": item.get("status", "")}
    images, title = ld_json_images(html)
    if images:
        return {"images": images, "title": title, "status": ""}
    return {"error": "画像を特定できず"}


def sold_search_url(final_url):
    """メルカリ検索 URL の status を売り切れに差し替える。"""
    parts = urllib.parse.urlparse(final_url)
    q = [(k, v) for k, v in urllib.parse.parse_qsl(parts.query) if k != "status"]
    q.append(("status", "sold_out|trading"))
    return parts._replace(query=urllib.parse.urlencode(q)).geturl()


async def sold_candidates(ctx, search_url):
    # 保存済み検索条件 (search_condition_id) を開くと条件付きの URL にリダイレクトされる
    _, final_url, _, _, page = await open_page(ctx, search_url, lambda u: False)
    await page.close()
    url = sold_search_url(final_url)
    _, _, _, captured, page = await open_page(
        ctx, url, lambda u: "entities:search" in u, expect=True)
    await page.close()
    items = []
    for _, d in captured:
        items = d.get("items") or items
    items = [i for i in items if i.get("status") in (
        "ITEM_STATUS_SOLD_OUT", "ITEM_STATUS_TRADING", "sold_out", "trading")]

    cands, checked = [], 0
    for it in items[:MAX_SOLD_CHECK]:
        item_url = (f"https://jp.mercari.com/shops/product/{it['id']}"
                    if it.get("itemType") == "ITEM_TYPE_BEYOND"
                    else f"https://jp.mercari.com/item/{it['id']}")
        checked += 1
        try:
            info = await fetch_images(ctx, item_url)
        except Exception as e:
            continue
        if len(info.get("images") or []) >= MIN_IMAGES:
            cands.append({"url": item_url, "price": it.get("price"), **info})
            if len(cands) >= MAX_SOLD_CANDIDATES:
                break
    return {"search_url": url, "sold_found": len(items), "checked": checked,
            "candidates": cands}


async def process_sku(ctx, row):
    suppliers = []
    for col in SUPPLIER_COLS:
        url = row[col].strip()
        if not url.startswith("http"):
            continue
        try:
            info = await fetch_images(ctx, url)
        except Exception as e:
            info = {"error": f"読み込み失敗: {type(e).__name__}"}
        suppliers.append({"col": col, "url": url, **info})
        if len(info.get("images") or []) >= MIN_IMAGES:
            return {"suppliers": suppliers, "selected": "supplier"}

    result = {"suppliers": suppliers, "selected": None}
    search = row[MERCARI_SEARCH_COL].strip()
    if search.startswith("http"):
        try:
            result["sold"] = await sold_candidates(ctx, search)
        except Exception as e:
            result["sold"] = {"error": f"読み込み失敗: {type(e).__name__}", "candidates": []}
        if result["sold"]["candidates"]:
            result["selected"] = "mercari_sold"
    return result


async def main():
    sheet_csv, low_csv = sys.argv[1], sys.argv[2]
    rows = list(csv.reader(open(sheet_csv, encoding="utf-8")))
    by_sku = {r[SKU_COL].strip(): r for r in rows[2:] if r[SKU_COL].strip()}
    targets = list(csv.DictReader(open(low_csv, encoding="utf-8-sig")))
    only = set(sys.argv[3:])
    if only:
        targets = [t for t in targets if t["sku"] in only]
        missing = only - {t["sku"] for t in targets}
        if missing:
            print(f"画像2枚以下の一覧に無い SKU: {' '.join(sorted(missing))}")

    results = json.loads(RESULT_PATH.read_text()) if RESULT_PATH.exists() else {}
    todo = [t for t in targets if only or t["sku"] not in results]
    print(f"対象 {len(targets)} SKU / 取得済み {len(targets) - len(todo)} / 残り {len(todo)}")

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            executable_path=CHROME, args=["--no-sandbox"],
            proxy={"server": os.environ["HTTPS_PROXY"]})
        ctx = await browser.new_context(locale="ja-JP")
        await ctx.route(re.compile(r".*\.(png|jpe?g|webp|gif|woff2?|mp4)(\?.*)?$"),
                        lambda route: route.abort())
        sem = asyncio.Semaphore(CONCURRENCY)
        done = 0

        async def run(t):
            nonlocal done
            row = by_sku.get(str(int(t["sku"])) if t["sku"].isdigit() else t["sku"])
            async with sem:
                if row is None:
                    res = {"error": "販売管理シートに SKU が無い"}
                else:
                    res = await process_sku(ctx, row)
                    res["sheet_name"] = row[NAME_COL]
            res.update({"item_id": t["item_id"], "ebay_title": t["title"],
                        "ebay_images": t["image_count"]})
            results[t["sku"]] = res
            done += 1
            RESULT_PATH.write_text(json.dumps(results, ensure_ascii=False, indent=1))
            print(f"[{done}/{len(todo)}] {t['sku']}: {res.get('selected')}", flush=True)

        await asyncio.gather(*(run(t) for t in todo))
        await browser.close()

    sel = [r.get("selected") for r in results.values()]
    print(f"\n仕入れ先で見つかった: {sel.count('supplier')} / メルカリ売り切れ候補あり: "
          f"{sel.count('mercari_sold')} / 候補なし: {sel.count(None)}")


if __name__ == "__main__":
    asyncio.run(main())
