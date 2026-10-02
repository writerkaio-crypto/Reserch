"""自分の eBay 出品中商品のうち「中古 × 画像2枚以下」を SKU で洗い出す。

Trading API GetSellerList を使用。環境変数 EBAY_USER_TOKEN (OAuth ユーザートークン) が必要。
出力: output/used_low_images_<日付>.csv (該当分), output/active_listings_<日付>.csv (全件)
"""

import csv
import os
import sys
import urllib.request
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

API_URL = "https://api.ebay.com/ws/api.dll"
NS = {"e": "urn:ebay:apis:eBLBaseComponents"}
MAX_IMAGES = 2
OUT_DIR = Path(__file__).parent / "output"

# 新品扱いの ConditionID。これ以外(2500 以上)を中古とみなす。
# 1000=New, 1500=New other, 1750=New with defects, 2000/2010=Certified refurbished 等
NEW_CONDITION_IDS = {"1000", "1500", "1750"}


def call(name, body):
    xml = (
        '<?xml version="1.0" encoding="utf-8"?>'
        f'<{name}Request xmlns="urn:ebay:apis:eBLBaseComponents">{body}</{name}Request>'
    )
    req = urllib.request.Request(
        API_URL,
        data=xml.encode(),
        headers={
            "X-EBAY-API-CALL-NAME": name,
            "X-EBAY-API-SITEID": "0",
            "X-EBAY-API-COMPATIBILITY-LEVEL": "1271",
            "X-EBAY-API-IAF-TOKEN": os.environ["EBAY_USER_TOKEN"],
            "Content-Type": "text/xml",
        },
    )
    with urllib.request.urlopen(req, timeout=60) as r:
        root = ET.fromstring(r.read())
    if root.findtext("e:Ack", namespaces=NS) not in ("Success", "Warning"):
        errs = [e.findtext("e:LongMessage", namespaces=NS) for e in root.findall("e:Errors", NS)]
        sys.exit(f"[ERR] {name}: {errs}")
    return root


def fetch_active_items():
    # GetSellerList は終了日時で範囲指定 (最大120日)。出品中 = 終了日時が今以降。
    now = datetime.now(timezone.utc)
    end_to = now + timedelta(days=119)
    fmt = "%Y-%m-%dT%H:%M:%S.000Z"
    page, items = 1, []
    while True:
        root = call("GetSellerList", f"""
            <DetailLevel>ReturnAll</DetailLevel>
            <EndTimeFrom>{now.strftime(fmt)}</EndTimeFrom>
            <EndTimeTo>{end_to.strftime(fmt)}</EndTimeTo>
            <IncludeVariations>true</IncludeVariations>
            <Pagination><EntriesPerPage>200</EntriesPerPage><PageNumber>{page}</PageNumber></Pagination>""")
        items += root.findall("e:ItemArray/e:Item", NS)
        total_pages = int(root.findtext("e:PaginationResult/e:TotalNumberOfPages", "1", NS))
        print(f"page {page}/{total_pages}: {len(items)} items", file=sys.stderr)
        if page >= total_pages:
            return items
        page += 1


def parse(item):
    t = lambda path: item.findtext(path, "", NS)
    pics = [p.text for p in item.findall("e:PictureDetails/e:PictureURL", NS)]
    pics += [p.text for p in item.findall("e:PictureDetails/e:ExternalPictureURL", NS)]
    var_skus = [v.findtext("e:SKU", "", NS) for v in item.findall("e:Variations/e:Variation", NS)]
    var_pics = item.findall("e:Variations/e:Pictures/e:VariationSpecificPictureSet/e:PictureURL", NS)
    return {
        "sku": t("e:SKU"),
        "item_id": t("e:ItemID"),
        "title": t("e:Title"),
        "site": t("e:Site"),
        "category": t("e:PrimaryCategory/e:CategoryName"),
        "condition_id": t("e:ConditionID"),
        "condition": t("e:ConditionDisplayName"),
        "image_count": len(pics),
        "variation_image_count": len(var_pics),
        "variation_skus": "/".join(s for s in var_skus if s),
        "listing_status": t("e:SellingStatus/e:ListingStatus"),
        "listing_type": t("e:ListingType"),
        "price": t("e:SellingStatus/e:CurrentPrice"),
        "quantity_available": int(t("e:Quantity") or 0) - int(t("e:SellingStatus/e:QuantitySold") or 0),
        "start_time": t("e:ListingDetails/e:StartTime"),
        "url": t("e:ListingDetails/e:ViewItemURL"),
    }


def is_used(row):
    return row["condition_id"] != "" and row["condition_id"] not in NEW_CONDITION_IDS


def write_csv(path, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    print(f"wrote {path} ({len(rows)} rows)", file=sys.stderr)


def main():
    OUT_DIR.mkdir(exist_ok=True)
    rows = [parse(i) for i in fetch_active_items()]
    rows = [r for r in rows if r["listing_status"] == "Active"]
    hits = [r for r in rows if is_used(r) and r["image_count"] <= MAX_IMAGES]
    hits.sort(key=lambda r: (r["image_count"], r["sku"]))

    print("\n[状態の内訳]", file=sys.stderr)
    for (cid, name), n in Counter((r["condition_id"], r["condition"]) for r in rows).most_common():
        print(f"  {cid or '-':>5} {name or '(なし)'}: {n}", file=sys.stderr)
    print(f"\n出品中 {len(rows)} 件 / 中古 {sum(map(is_used, rows))} 件 / "
          f"中古かつ画像{MAX_IMAGES}枚以下 {len(hits)} 件", file=sys.stderr)

    # 同一 SKU が複数サイトに出品されているので SKU 単位にまとめる
    by_sku = {}
    for r in hits:
        by_sku.setdefault(r["sku"] or f"(SKU未設定:{r['item_id']})", []).append(r)
    sku_rows = [{
        "sku": sku,
        "title": rs[0]["title"],
        "condition": rs[0]["condition"],
        "min_image_count": min(r["image_count"] for r in rs),
        "listing_count": len(rs),
        "sites": "/".join(sorted({r["site"] for r in rs})),
        "item_ids": "/".join(r["item_id"] for r in rs),
    } for sku, rs in sorted(by_sku.items(), key=lambda kv: kv[0].zfill(10))]
    no_cond = [r for r in rows if not r["condition_id"] and r["image_count"] <= MAX_IMAGES]
    print(f"SKU 単位 {len(sku_rows)} 件 (状態未設定で画像{MAX_IMAGES}枚以下: {len(no_cond)} 出品は対象外)",
          file=sys.stderr)

    stamp = date.today().isoformat()
    if rows:
        write_csv(OUT_DIR / f"active_listings_{stamp}.csv", rows)
    if hits:
        write_csv(OUT_DIR / f"used_low_images_{stamp}.csv", hits)
        write_csv(OUT_DIR / f"used_low_images_sku_{stamp}.csv", sku_rows)


if __name__ == "__main__":
    main()
