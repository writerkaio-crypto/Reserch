"""自分の出品中リストから、登録画像が N 枚以下の出品を SKU で洗い出す。

SKU は出品者本人しか取得できないため、次のどちらかで実行する。

  1) ユーザートークン方式 (全自動)
       環境変数 EBAY_USER_TOKEN に自分のアカウントのトークンを設定して実行。
       Trading API GetSellerList で出品中の全件の SKU と画像 URL を取得する。
         python low_image_listings.py

  2) Seller Hub CSV 方式 (トークン不要)
       Seller Hub > Listings > Active から出品中リストを CSV でダウンロードし、
       そのファイルを指定して実行。CSV の「Item number」「Custom label (SKU)」を使い、
       画像枚数は Browse API (EBAY_CLIENT_ID / EBAY_CLIENT_SECRET) で取得する。
         python low_image_listings.py --csv active_listings.csv

出力: output/low_image_listings_YYYY-MM-DD.csv
"""

import argparse
import csv
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from ebay_research import MARKETPLACE, OUT_DIR, get_token

TRADING_URL = "https://api.ebay.com/ws/api.dll"
BROWSE_ITEM_URL = "https://api.ebay.com/buy/browse/v1/item"
NS = {"e": "urn:ebay:apis:eBLBaseComponents"}
PAGE_SIZE = 200
CACHE_PATH = OUT_DIR / ".image_cache.json"


# ---------- 1) ユーザートークン方式 (Trading API) ----------

def trading_call(token, call, body):
    headers = {
        "X-EBAY-API-CALL-NAME": call,
        "X-EBAY-API-SITEID": "0",
        "X-EBAY-API-COMPATIBILITY-LEVEL": "1349",
        "Content-Type": "text/xml",
    }
    creds = ""
    if token.startswith("v^"):  # OAuth ユーザートークン
        headers["X-EBAY-API-IAF-TOKEN"] = token
    else:  # Auth'n'Auth トークン
        creds = (
            "<RequesterCredentials><eBayAuthToken>"
            f"{token}</eBayAuthToken></RequesterCredentials>"
        )
    xml = (
        '<?xml version="1.0" encoding="utf-8"?>'
        f'<{call}Request xmlns="urn:ebay:apis:eBLBaseComponents">'
        f"{creds}{body}</{call}Request>"
    )
    req = urllib.request.Request(TRADING_URL, data=xml.encode(), headers=headers)
    with urllib.request.urlopen(req, timeout=60) as r:
        root = ET.fromstring(r.read())
    if root.findtext("e:Ack", namespaces=NS) not in ("Success", "Warning"):
        msgs = [e.findtext("e:LongMessage", namespaces=NS)
                for e in root.findall("e:Errors", NS)]
        raise RuntimeError(f"{call} failed: {msgs}")
    return root


def fetch_via_trading(token):
    # GetSellerList の EndTime 範囲は最大 120 日。GTC 出品は 30 日ごとに更新されるので
    # 「今〜120日後に終了予定」で出品中の全件をカバーできる。
    now = datetime.now(timezone.utc)
    fmt = "%Y-%m-%dT%H:%M:%S.000Z"
    selectors = "".join(
        f"<OutputSelector>{s}</OutputSelector>" for s in (
            "ItemArray.Item.ItemID",
            "ItemArray.Item.SKU",
            "ItemArray.Item.Title",
            "ItemArray.Item.PictureDetails.PictureURL",
            "ItemArray.Item.SellingStatus.ListingStatus",
            "ItemArray.Item.ListingDetails.ViewItemURL",
            "ItemArray.Item.Variations.Variation.SKU",
            "PaginationResult",
            "HasMoreItems",
        )
    )
    items, page = [], 1
    while True:
        body = (
            f"<EndTimeFrom>{now.strftime(fmt)}</EndTimeFrom>"
            f"<EndTimeTo>{(now + timedelta(days=119)).strftime(fmt)}</EndTimeTo>"
            "<DetailLevel>ReturnAll</DetailLevel>"
            f"<Pagination><EntriesPerPage>{PAGE_SIZE}</EntriesPerPage>"
            f"<PageNumber>{page}</PageNumber></Pagination>{selectors}"
        )
        root = trading_call(token, "GetSellerList", body)
        for it in root.iterfind("e:ItemArray/e:Item", NS):
            status = it.findtext("e:SellingStatus/e:ListingStatus", namespaces=NS)
            if status and status != "Active":
                continue
            var_skus = [v.text for v in it.iterfind(
                "e:Variations/e:Variation/e:SKU", NS) if v.text]
            items.append({
                "item_id": it.findtext("e:ItemID", namespaces=NS),
                "sku": it.findtext("e:SKU", "", namespaces=NS),
                "variation_skus": " / ".join(var_skus),
                "title": it.findtext("e:Title", "", namespaces=NS),
                "image_count": len(it.findall("e:PictureDetails/e:PictureURL", NS)),
                "url": it.findtext("e:ListingDetails/e:ViewItemURL", "",
                                   namespaces=NS),
            })
        total_pages = int(root.findtext(
            "e:PaginationResult/e:TotalNumberOfPages", "1", namespaces=NS))
        print(f"page {page}/{total_pages}: {len(items)} items")
        if page >= total_pages:
            return items
        page += 1


# ---------- 2) Seller Hub CSV 方式 (Browse API) ----------

def read_seller_hub_csv(path):
    """1 行 = 1 出品として読む。先頭のメタ行はスキップする。"""
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    for i, row in enumerate(rows):
        cols = [c.strip().lower() for c in row]
        if "item number" in cols:
            break
    else:
        sys.exit("CSV に 'Item number' 列が見つかりません")
    idx = {c: n for n, c in enumerate(cols)}
    sku_col = next((n for c, n in idx.items() if c.startswith("custom label")), None)

    def get(row, col):
        n = idx.get(col) if isinstance(col, str) else col
        return row[n].strip() if n is not None and n < len(row) else ""

    listings, seen = [], set()
    for row in rows[i + 1:]:
        item_id = get(row, "item number")
        if not item_id.isdigit() or item_id in seen:
            continue
        seen.add(item_id)
        listings.append({
            "item_id": item_id,
            "sku": get(row, sku_col),
            "title": get(row, "title"),
            "site": get(row, "listing site"),
            "quantity": get(row, "available quantity"),
        })
    return listings


def browse_get(token, path, params):
    url = f"{BROWSE_ITEM_URL}/{path}?{urllib.parse.urlencode(params)}"
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {token}",
        "X-EBAY-C-MARKETPLACE-ID": MARKETPLACE,
    })
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.load(r)


def image_count(item):
    return (1 if item.get("image") else 0) + len(item.get("additionalImages") or [])


def fetch_image_count(token, item_id):
    """(画像枚数, URL) を返す。バリエーション出品は全バリエーション中の最大枚数。"""
    try:
        it = browse_get(token, "get_item_by_legacy_id", {"legacy_item_id": item_id})
        return image_count(it), it.get("itemWebUrl", "")
    except urllib.error.HTTPError as e:
        if e.code != 400:  # バリエーション出品以外のエラーはそのまま上げる
            raise
    group = browse_get(token, "get_items_by_item_group", {"item_group_id": item_id})
    items = group.get("items") or []
    url = items[0].get("itemWebUrl", "") if items else ""
    return max((image_count(i) for i in items), default=0), url


def fetch_with_retry(state, item_id, attempts=5):
    """通信エラーは待って再試行、トークン期限切れ (401) は再取得して再試行する。"""
    for n in range(attempts):
        try:
            return fetch_image_count(state["token"], item_id)
        except urllib.error.HTTPError as e:
            if e.code != 401 or n == attempts - 1:
                raise
            state["token"] = get_token()
        except (urllib.error.URLError, TimeoutError, ConnectionError, OSError):
            if n == attempts - 1:
                raise
            time.sleep(2 ** (n + 1))


def fetch_via_csv(path):
    """Browse API で各出品の画像枚数を取得する。

    Browse API は 1 日 5000 回までなので、結果を CACHE_PATH に保存し、
    上限 (429) に達したら中断 → 翌日再実行で続きから取得する。
    SKU ごとの代表 (US 優先) を先に取得し、残りのサイトは後回しにする。
    """
    listings = read_seller_hub_csv(path)
    cache = json.loads(CACHE_PATH.read_text()) if CACHE_PATH.exists() else {}

    first, seen_sku = [], set()
    for li in sorted(listings, key=lambda li: li["site"] != "US"):
        if li["sku"] not in seen_sku:
            seen_sku.add(li["sku"])
            first.append(li)
    first_ids = {li["item_id"] for li in first}
    todo = [li for li in first + [li for li in listings if li["item_id"] not in first_ids]
            if li["item_id"] not in cache]
    print(f"{len(listings)} 件中 取得済み {len(listings) - len(todo)} 件 / 残り {len(todo)} 件")

    state = {"token": get_token()}
    try:
        for n, li in enumerate(todo, 1):
            try:
                count, _ = fetch_with_retry(state, li["item_id"])
                cache[li["item_id"]] = count
            except urllib.error.HTTPError as e:
                if e.code == 429:
                    print("Browse API の 1 日の上限に達しました。リセット後に再実行してください。")
                    break
                if e.code == 404:  # 在庫切れ等で表示されていない出品
                    cache[li["item_id"]] = None
                else:
                    print(f"[ERR] {li['item_id']}: {e.code} {e.read()[:200]!r}",
                          file=sys.stderr)
            except (urllib.error.URLError, TimeoutError, ConnectionError, OSError) as e:
                print(f"[ERR] {li['item_id']}: {e}", file=sys.stderr)  # 再実行で再取得
            if n % 100 == 0:
                CACHE_PATH.write_text(json.dumps(cache))
                print(f"{n}/{len(todo)} checked")
            time.sleep(0.05)
    finally:
        CACHE_PATH.write_text(json.dumps(cache))

    for li in listings:
        li["image_count"] = cache.get(li["item_id"], "未取得")
        li["url"] = f"https://www.ebay.com/itm/{li['item_id']}"
    return listings


def summarize_by_sku(listings, max_images):
    """SKU ごとにまとめ、どこかのサイトで画像が max_images 枚以下の SKU を返す。"""
    by_sku = {}
    for li in listings:
        by_sku.setdefault(li["sku"], []).append(li)
    out = []
    for sku, lis in by_sku.items():
        counts = [li for li in lis if isinstance(li["image_count"], int)]
        low = [li for li in counts if li["image_count"] <= max_images]
        if not low:
            continue
        rep = next((li for li in lis if li["site"] == "US"), lis[0])
        out.append({
            "sku": sku,
            "min_images": min(li["image_count"] for li in low),
            "low_sites": " ".join(sorted(f"{li['site']}:{li['image_count']}" for li in low)),
            "all_sites_low": "YES" if len(low) == len(lis) else "",
            "site_count": len(lis),
            "unchecked_sites": sum(1 for li in lis if li["image_count"] == "未取得"),
            "title": rep["title"],
            "low_item_ids": " ".join(li["item_id"] for li in low),
        })
    return sorted(out, key=lambda r: (r["min_images"], r["sku"]))


def write_csv(path, rows, fields):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    print(f"wrote {path} ({len(rows)} rows)")


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--csv", help="Seller Hub からダウンロードした出品中リスト CSV")
    ap.add_argument("--max-images", type=int, default=2,
                    help="この枚数以下を抽出 (既定: 2)")
    args = ap.parse_args()
    OUT_DIR.mkdir(exist_ok=True)
    stamp = date.today().isoformat()

    if args.csv:
        listings = fetch_via_csv(args.csv)
        skus = summarize_by_sku(listings, args.max_images)
        write_csv(OUT_DIR / f"low_image_skus_{stamp}.csv", skus, list(skus[0]) if skus else ["sku"])
        detail = sorted(
            (li for li in listings
             if isinstance(li["image_count"], int) and li["image_count"] <= args.max_images),
            key=lambda li: (li["image_count"], li["sku"], li["site"]))
        write_csv(OUT_DIR / f"low_image_listings_{stamp}.csv", detail,
                  ["sku", "site", "item_id", "image_count", "quantity", "title", "url"])
        unchecked = sum(1 for li in listings if li["image_count"] == "未取得")
        hidden = sum(1 for li in listings if li["image_count"] is None)
        print(f"\n出品 {len(listings)} 件 / SKU {len({li['sku'] for li in listings})} 種")
        print(f"画像 {args.max_images} 枚以下: {len(detail)} 出品 / {len(skus)} SKU")
        print(f"  (うち全サイトで {args.max_images} 枚以下の SKU: "
              f"{sum(1 for s in skus if s['all_sites_low'])})")
        if hidden:
            print(f"  非表示で確認できない出品 (在庫切れ等): {hidden} 件")
        if unchecked:
            print(f"  未取得: {unchecked} 件 → API 上限リセット後に同じコマンドを再実行")
        return

    if not os.environ.get("EBAY_USER_TOKEN"):
        sys.exit("EBAY_USER_TOKEN を設定するか、--csv で Seller Hub の CSV を指定してください")
    items = fetch_via_trading(os.environ["EBAY_USER_TOKEN"])
    hits = sorted((i for i in items if i["image_count"] <= args.max_images),
                  key=lambda i: (i["image_count"], i["sku"] or i["variation_skus"]))
    write_csv(OUT_DIR / f"low_image_listings_{stamp}.csv", hits,
              ["sku", "variation_skus", "item_id", "image_count", "title", "url"])
    no_sku = sum(1 for i in hits if not i["sku"] and not i["variation_skus"])
    print(f"\n出品中 {len(items)} 件中、画像 {args.max_images} 枚以下: {len(hits)} 件")
    if no_sku:
        print(f"  うち SKU 未設定: {no_sku} 件 (item_id で確認してください)")


if __name__ == "__main__":
    main()
