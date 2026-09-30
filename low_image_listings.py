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
    """Item number ごとに SKU・タイトルをまとめる。先頭のメタ行はスキップする。"""
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
    title_col = idx.get("title")

    items = {}
    for row in rows[i + 1:]:
        if len(row) <= idx["item number"]:
            continue
        item_id = row[idx["item number"]].strip()
        if not item_id.isdigit():
            continue
        it = items.setdefault(item_id, {"item_id": item_id, "skus": [], "title": ""})
        if sku_col is not None and row[sku_col].strip():
            it["skus"].append(row[sku_col].strip())
        if title_col is not None and not it["title"]:
            it["title"] = row[title_col].strip()
    return list(items.values())


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


def fetch_via_csv(path):
    token = get_token()
    rows = read_seller_hub_csv(path)
    items = []
    for n, r in enumerate(rows, 1):
        skus = list(dict.fromkeys(r["skus"]))
        try:
            count, url = fetch_image_count(token, r["item_id"])
        except urllib.error.HTTPError as e:
            print(f"[ERR] {r['item_id']}: {e.code} {e.read()[:200]!r}",
                  file=sys.stderr)
            continue
        items.append({
            "item_id": r["item_id"],
            "sku": skus[0] if len(skus) == 1 else "",
            "variation_skus": " / ".join(skus) if len(skus) > 1 else "",
            "title": r["title"],
            "image_count": count,
            "url": url,
        })
        if n % 50 == 0 or n == len(rows):
            print(f"{n}/{len(rows)} checked")
        time.sleep(0.05)
    return items


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--csv", help="Seller Hub からダウンロードした出品中リスト CSV")
    ap.add_argument("--max-images", type=int, default=2,
                    help="この枚数以下を抽出 (既定: 2)")
    args = ap.parse_args()

    if args.csv:
        items = fetch_via_csv(args.csv)
    elif os.environ.get("EBAY_USER_TOKEN"):
        items = fetch_via_trading(os.environ["EBAY_USER_TOKEN"])
    else:
        sys.exit("EBAY_USER_TOKEN を設定するか、--csv で Seller Hub の CSV を指定してください")

    hits = sorted((i for i in items if i["image_count"] <= args.max_images),
                  key=lambda i: (i["image_count"], i["sku"] or i["variation_skus"]))

    OUT_DIR.mkdir(exist_ok=True)
    path = OUT_DIR / f"low_image_listings_{date.today().isoformat()}.csv"
    fields = ["sku", "variation_skus", "item_id", "image_count", "title", "url"]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        w.writerows(hits)

    no_sku = sum(1 for i in hits if not i["sku"] and not i["variation_skus"])
    print(f"\n出品中 {len(items)} 件中、画像 {args.max_images} 枚以下: {len(hits)} 件")
    if no_sku:
        print(f"  うち SKU 未設定: {no_sku} 件 (item_id で確認してください)")
    print(f"wrote {path}")


if __name__ == "__main__":
    main()
