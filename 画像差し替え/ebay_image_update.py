"""「差し替え予定」シートの反映=OK 行について、eBay 出品画像を差し替える。

画像1〜15 列のハイパーリンクを順に使う。
- i.ebayimg.com の URL は現在の出品にある画像をそのまま使う
- それ以外は UploadSiteHostedPictures で eBay (EPS) にアップロードしてから使う
画像リンクが空の OK 行はスキップする（採用ページだけ入っていて画像が未取得の行）。

環境変数 EBAY_USER_TOKEN を使用。
使い方: python ebay_image_update.py <Excel> [--dry-run] [SKU ...]
"""

import html
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from datetime import date

import openpyxl

API_URL = "https://api.ebay.com/ws/api.dll"
NS = 'xmlns="urn:ebay:apis:eBLBaseComponents"'
IMAGE_COLS = range(10, 25)  # K〜Y 列（0始まり）


def call(verb, inner):
    body = f'<?xml version="1.0" encoding="utf-8"?><{verb}Request {NS}>{inner}</{verb}Request>'
    req = urllib.request.Request(
        API_URL,
        data=body.encode(),
        headers={
            "X-EBAY-API-CALL-NAME": verb,
            "X-EBAY-API-SITEID": "0",
            "X-EBAY-API-COMPATIBILITY-LEVEL": "1225",
            "X-EBAY-API-IAF-TOKEN": os.environ["EBAY_USER_TOKEN"],
            "Content-Type": "text/xml",
        },
    )
    for attempt in range(4):
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                return r.read().decode()
        except (urllib.error.URLError, TimeoutError):
            if attempt == 3:
                raise
            time.sleep(2 ** (attempt + 1))  # 一時的な通信切れは 2, 4, 8 秒待って再試行


def current_pictures(item_id):
    x = call("GetItem", f"<ItemID>{item_id}</ItemID><DetailLevel>ReturnAll</DetailLevel>")
    return [html.unescape(u) for u in re.findall("<PictureURL>(.*?)</PictureURL>", x)]


def upload(url):
    x = call(
        "UploadSiteHostedPictures",
        f"<ExternalPictureURL>{html.escape(url)}</ExternalPictureURL><PictureSet>Supersize</PictureSet>",
    )
    found = re.findall("<FullURL>(.*?)</FullURL>", x)
    if not found:
        raise RuntimeError(f"upload failed: {url}\n{x[:600]}")
    return html.unescape(found[0])


def revise(item_id, urls):
    pics = "".join(f"<PictureURL>{html.escape(u)}</PictureURL>" for u in urls)
    x = call(
        "ReviseFixedPriceItem",
        f"<Item><ItemID>{item_id}</ItemID><PictureDetails>{pics}</PictureDetails></Item>",
    )
    ack = re.findall("<Ack>(.*?)</Ack>", x)[0]
    errors = [
        m
        for m in re.findall("<LongMessage>(.*?)</LongMessage>", x)
        if "business policies" not in m and "Best Offer" not in m  # 毎回出る定型の警告
    ]
    return ack, errors


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    dry_run = "--dry-run" in sys.argv
    path, only = args[0], {int(s) for s in args[1:]}
    ws = openpyxl.load_workbook(path)["差し替え予定"]
    log = []

    for row in ws.iter_rows(min_row=2):
        sku, item_id = row[1].value, str(row[2].value)
        if row[0].value != "OK" or (only and sku not in only):
            continue
        links = [row[i].hyperlink.target for i in IMAGE_COLS if row[i].hyperlink]
        if not links:
            print(f"{sku}: 画像リンクなし → スキップ")
            continue
        before = current_pictures(item_id)
        if dry_run:
            print(f"{sku}: {len(before)} → {len(links)} 枚（dry-run）")
            continue

        new = []
        for url in links:
            if "ebayimg.com" in url:
                key = re.search(r"/z/([^/]+)/", url).group(1)
                match = [u for u in before if key in u]
                if not match:
                    raise RuntimeError(f"{sku}: 現在の出品に無い eBay 画像 {url}")
                new.append(match[0])
            else:
                new.append(upload(url))
        ack, errors = revise(item_id, new)
        after = current_pictures(item_id)
        print(f"{sku}: {ack} {len(before)} → {len(after)} 枚 {errors or ''}")
        log.append({"sku": sku, "item": item_id, "ack": ack, "errors": errors, "before": before, "after": after})

    if log:
        out = f"update_log_{date.today()}.json"
        if os.path.exists(out):  # 同じ日の複数回実行は追記する
            with open(out, encoding="utf-8") as f:
                log = json.load(f) + log
        with open(out, "w", encoding="utf-8") as f:
            json.dump(log, f, ensure_ascii=False, indent=1)
        print(f"ログ: {out}")


if __name__ == "__main__":
    main()
