"""画像少なめの出品一覧に、eBay に登録されている商品の状態 (Condition) を付ける。

  python listing_conditions.py <low_image_listings CSV>

環境変数 EBAY_USER_TOKEN (出品者本人のトークン) を使用。
GetItem で出品ごとに ConditionID / 状態名 / 現在の画像枚数 / 出品状況を取り直す。
ConditionID 1000 = New (新品)。1500 (New other) などは新品扱いにしない。
結果: output/low_image_conditions_YYYY-MM-DD.csv
"""

import csv
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from pathlib import Path

from low_image_listings import NS, trading_call

OUT_DIR = Path(__file__).parent / "output"

NEW_CONDITION_ID = "1000"


def get_condition(token, item_id, attempts=4):
    for i in range(attempts):
        try:
            root = trading_call(token, "GetItem", (
                f"<ItemID>{item_id}</ItemID>"
                "<OutputSelector>Item.ConditionID</OutputSelector>"
                "<OutputSelector>Item.ConditionDisplayName</OutputSelector>"
                "<OutputSelector>Item.SellingStatus.ListingStatus</OutputSelector>"
                "<OutputSelector>Item.PictureDetails</OutputSelector>"))
            break
        except Exception:
            if i == attempts - 1:
                raise
            time.sleep(2 ** i)
    it = root.find("e:Item", NS)
    return {
        "condition_id": it.findtext("e:ConditionID", "", namespaces=NS),
        "condition": it.findtext("e:ConditionDisplayName", "", namespaces=NS),
        "status": it.findtext("e:SellingStatus/e:ListingStatus", "", namespaces=NS),
        "image_count_now": len(it.findall("e:PictureDetails/e:PictureURL", NS)),
    }


def main():
    token = os.environ.get("EBAY_USER_TOKEN", "")
    if len(token) < 100:
        sys.exit("EBAY_USER_TOKEN が未設定か途中で切れています")
    with open(sys.argv[1], encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))

    def work(row):
        try:
            return {**row, **get_condition(token, row["item_id"])}
        except Exception as e:
            return {**row, "condition_id": "", "condition": f"エラー: {e}",
                    "status": "", "image_count_now": ""}

    with ThreadPoolExecutor(4) as ex:
        out = list(ex.map(work, rows))
    for r in out:
        r["is_new"] = "YES" if r["condition_id"] == NEW_CONDITION_ID else ""

    path = OUT_DIR / f"low_image_conditions_{date.today().isoformat()}.csv"
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(out[0].keys()))
        w.writeheader()
        w.writerows(out)
    print(f"wrote {path} ({len(out)} rows, new={sum(r['is_new'] == 'YES' for r in out)})")


if __name__ == "__main__":
    main()
