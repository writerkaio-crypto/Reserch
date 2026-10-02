"""中古品の画像差し替え用に、仕入れ先と同じ / 近い状態のメルカリ出品を集める。

  python mercari_condition_search.py supplier <URL> [...]          仕入れ先ページの状態・説明・画像を取得
  python mercari_condition_search.py search <SKU> <キーワード> [...]  メルカリ検索 (販売中+売り切れ) の各商品を取得

結果は output/mercari_condition/<SKU or supplier>.json に保存する。
採否 (同じ商品か・付属品・状態が仕入れ先と同じか1段階差で傷の内容が近いか) は人が判定する。
"""

import asyncio
import json
import os
import re
import sys
import urllib.parse
from pathlib import Path

from playwright.async_api import async_playwright

OUT_DIR = Path(__file__).parent / "output" / "mercari_condition"
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"
MAX_ITEMS = 40  # 1 キーワードあたり詳細を開く最大件数
CONCURRENCY = 4

# メルカリの「商品の状態」(良い順)。段階差の計算に使う
CONDITIONS = ["新品、未使用", "未使用に近い", "目立った傷や汚れなし",
              "やや傷や汚れあり", "傷や汚れあり", "全体的に状態が悪い"]


def rank(cond):
    for i, c in enumerate(CONDITIONS):
        if cond and cond.startswith(c):
            return i
    return None


async def open_page(ctx, url, capture=lambda u: False, wait_capture=False):
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
        await page.wait_for_timeout(3000)
        for _ in range(25 if wait_capture else 0):
            if captured:
                break
            await page.wait_for_timeout(1000)
        await page.wait_for_timeout(500)
        html = await page.content()
        text = await page.inner_text("body")
        return {"status": resp.status if resp else 0, "url": page.url, "html": html,
                "text": text, "captured": captured}
    finally:
        await page.close()


def mercari_item(data):
    return {
        "title": data.get("name", ""),
        "price": data.get("price"),
        "status": data.get("status", ""),
        "condition": (data.get("item_condition") or {}).get("name", ""),
        "description": data.get("description") or "",
        "images": data.get("photos") or [],
    }


def mercari_shop(d):
    detail = d.get("productDetail") or {}
    cond = (detail.get("condition") or {}).get("displayName", "")
    return {
        "title": d.get("displayName", ""),
        "price": d.get("price"),
        "status": "shops",
        "condition": cond,
        "description": detail.get("description") or "",
        "images": detail.get("photos") or [],
    }


async def fetch(ctx, url):
    mid = re.search(r"mercari\.com/item/(m\d+)", url)
    shop = re.search(r"mercari\.com/shops/product/(\w+)", url)
    if mid:
        p = await open_page(ctx, url, lambda u: "items/get" in u, True)
        for _, d in p["captured"]:
            if (d.get("data") or {}).get("id") == mid.group(1):
                return {"url": url, **mercari_item(d["data"])}
        return {"url": url, "error": f"取得失敗 HTTP {p['status']}"}
    if shop:
        p = await open_page(ctx, url, lambda u: "shops/products/" in u, True)
        for _, d in p["captured"]:
            if d.get("name") == shop.group(1):
                return {"url": url, **mercari_shop(d)}
        return {"url": url, "error": f"取得失敗 HTTP {p['status']}"}
    # その他のサイトは本文テキストと画像 URL をそのまま保存して人が読む
    p = await open_page(ctx, url)
    imgs = sorted(set(re.findall(r'https?://[^"\'\s>]+\.(?:jpe?g|png|webp)', p["html"])))
    title = re.search(r"<title>(.*?)</title>", p["html"], re.S)
    return {"url": url, "http": p["status"], "title": title.group(1).strip() if title else "",
            "text": re.sub(r"\n\s*\n+", "\n", p["text"])[:6000], "images": imgs[:80]}


async def search(ctx, keyword, status):
    q = urllib.parse.urlencode({"keyword": keyword, "status": status})
    p = await open_page(ctx, f"https://jp.mercari.com/search?{q}",
                        lambda u: "entities:search" in u, True)
    items = []
    for _, d in p["captured"]:
        items += d.get("items") or []
    return [(f"https://jp.mercari.com/shops/product/{i['id']}"
             if i.get("itemType") == "ITEM_TYPE_BEYOND"
             else f"https://jp.mercari.com/item/{i['id']}") for i in items]


async def main():
    mode, args = sys.argv[1], sys.argv[2:]
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            executable_path=CHROME, args=["--no-sandbox"],
            proxy={"server": os.environ["HTTPS_PROXY"]})
        ctx = await browser.new_context(locale="ja-JP")
        await ctx.route(re.compile(r".*\.(png|jpe?g|webp|gif|woff2?|mp4)(\?.*)?$"),
                        lambda route: route.abort())
        sem = asyncio.Semaphore(CONCURRENCY)

        async def guarded(url):
            async with sem:
                try:
                    return await fetch(ctx, url)
                except Exception as e:
                    return {"url": url, "error": f"{type(e).__name__}: {e}"[:200]}

        if mode == "supplier":
            res = await asyncio.gather(*(guarded(u) for u in args))
            for r in res:
                name = re.sub(r"\W+", "_", r["url"])[-60:]
                (OUT_DIR / f"supplier_{name}.json").write_text(
                    json.dumps(r, ensure_ascii=False, indent=1))
                print(json.dumps({k: (v if k != "text" else v[:1500]) for k, v in r.items()
                                  if k not in ("images", "html")}, ensure_ascii=False)[:2500])
                print("images:", len(r.get("images") or []))
        else:
            sku, keywords = args[0], args[1:]
            urls = []
            for kw in keywords:
                for st in ("on_sale", "sold_out|trading"):
                    found = await search(ctx, kw, st)
                    print(f"{kw} [{st}]: {len(found)}", file=sys.stderr)
                    urls += [u for u in found[:MAX_ITEMS] if u not in urls]
            path = OUT_DIR / f"{sku}.json"
            old = {r["url"]: r for r in json.loads(path.read_text())} if path.exists() else {}
            todo = [u for u in urls if u not in old or "error" in old[u]]
            res = await asyncio.gather(*(guarded(u) for u in todo))
            old.update({r["url"]: r for r in res})
            path.write_text(json.dumps(list(old.values()), ensure_ascii=False, indent=1))
            print(f"{sku}: {len(old)} 件保存 ({len(todo)} 件取得)", file=sys.stderr)
        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
