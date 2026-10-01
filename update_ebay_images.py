"""画像差し替えシートの「OK」行について、eBay 出品の画像を差し替える。

入力: 差し替えシート(xlsx) の「差し替え予定」シート
  A列 反映(OK/NG) が "OK" で始まる行だけを対象にし、
  K〜Y列(画像1〜15)のハイパーリンク先を画像の並び順として使う。

処理(1出品ずつ):
  1. GetItem で現在の画像URLをバックアップ(output/picture_backup_<日付>.json)
  2. eBay 外の画像は UploadSiteHostedPictures で eBay(EPS) に取り込む
     (eBay 現行画像 i.ebayimg.com/images/g/<id>/ は GetItem の EPS URL に置き換えて再利用)
  3. ReviseFixedPriceItem で PictureDetails を丸ごと差し替え
  結果は output/image_update_result_<日付>.csv に出力。

認証(Trading API。いずれか1つ):
  EBAY_USER_TOKEN      OAuth ユーザーアクセストークン
  EBAY_REFRESH_TOKEN   OAuth リフレッシュトークン(+ EBAY_CLIENT_ID / EBAY_CLIENT_SECRET)
  EBAY_AUTH_TOKEN      Auth'n'Auth トークン

使い方:
  python update_ebay_images.py 差し替え.xlsx --dry-run   # URL確認のみ(eBayは変更しない)
  python update_ebay_images.py 差し替え.xlsx             # 実行
  python update_ebay_images.py 差し替え.xlsx --sku 943 977
"""

import argparse
import base64
import csv
import io
import json
import os
import re
import sys
import urllib.parse
import urllib.request
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

import openpyxl
from PIL import Image

API_URL = "https://api.ebay.com/ws/api.dll"
TOKEN_URL = "https://api.ebay.com/identity/v1/oauth2/token"
OAUTH_SCOPES = " ".join([
    "https://api.ebay.com/oauth/api_scope",
    "https://api.ebay.com/oauth/api_scope/sell.inventory",
])
COMPAT_LEVEL = "1349"
NS = {"e": "urn:ebay:apis:eBLBaseComponents"}
SHEET = "差し替え予定"
IMG_COL_START = 10  # K列(0始まり)
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124 Safari/537.36")
OUT_DIR = Path(__file__).parent / "output"


def read_targets(xlsx, skus=None):
    ws = openpyxl.load_workbook(xlsx)[SHEET]
    targets = []
    for row in ws.iter_rows(min_row=2):
        flag = str(row[0].value or "").strip()
        if not flag.upper().startswith("OK"):
            continue
        sku, item_id = str(row[1].value), str(row[2].value)
        if skus and sku not in skus:
            continue
        urls = [c.hyperlink.target for c in row[IMG_COL_START:] if c.hyperlink]
        targets.append({"sku": sku, "item_id": item_id, "urls": urls,
                        "expected": row[7].value})
    return targets


def check_image(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=30) as r:
        im = Image.open(io.BytesIO(r.read()))
    if im.format not in ("JPEG", "PNG"):
        return f"形式 {im.format}"
    if max(im.size) < 500:
        return f"解像度不足 {im.size}"
    return ""


class Trading:
    def __init__(self):
        self.iaf, self.authn = None, None
        if os.environ.get("EBAY_USER_TOKEN"):
            self.iaf = os.environ["EBAY_USER_TOKEN"]
        elif os.environ.get("EBAY_REFRESH_TOKEN"):
            self.iaf = self._refresh()
        elif os.environ.get("EBAY_AUTH_TOKEN"):
            self.authn = os.environ["EBAY_AUTH_TOKEN"]
        else:
            sys.exit("eBay ユーザートークンが未設定です "
                     "(EBAY_USER_TOKEN / EBAY_REFRESH_TOKEN / EBAY_AUTH_TOKEN)")

    def _refresh(self):
        cid = os.environ["EBAY_CLIENT_ID"]
        secret = os.environ["EBAY_CLIENT_SECRET"]
        auth = base64.b64encode(f"{cid}:{secret}".encode()).decode()
        body = urllib.parse.urlencode({
            "grant_type": "refresh_token",
            "refresh_token": os.environ["EBAY_REFRESH_TOKEN"],
            "scope": OAUTH_SCOPES,
        }).encode()
        req = urllib.request.Request(TOKEN_URL, data=body, headers={
            "Authorization": f"Basic {auth}",
            "Content-Type": "application/x-www-form-urlencoded",
        })
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.load(r)["access_token"]

    def call(self, name, inner):
        creds = ""
        if self.authn:
            creds = (f"<RequesterCredentials><eBayAuthToken>{escape(self.authn)}"
                     f"</eBayAuthToken></RequesterCredentials>")
        xml = (f'<?xml version="1.0" encoding="utf-8"?>'
               f'<{name}Request xmlns="urn:ebay:apis:eBLBaseComponents">'
               f"{creds}<ErrorLanguage>en_US</ErrorLanguage>"
               f"<WarningLevel>High</WarningLevel>{inner}</{name}Request>")
        headers = {
            "X-EBAY-API-CALL-NAME": name,
            "X-EBAY-API-SITEID": "0",
            "X-EBAY-API-COMPATIBILITY-LEVEL": COMPAT_LEVEL,
            "Content-Type": "text/xml",
        }
        if self.iaf:
            headers["X-EBAY-API-IAF-TOKEN"] = self.iaf
        req = urllib.request.Request(API_URL, data=xml.encode(), headers=headers)
        with urllib.request.urlopen(req, timeout=60) as r:
            root = ET.fromstring(r.read())
        ack = root.findtext("e:Ack", "", NS)
        errors = [
            f"[{e.findtext('e:SeverityCode', '', NS)}] "
            f"{e.findtext('e:ErrorCode', '', NS)} "
            f"{e.findtext('e:LongMessage', '', NS)}"
            for e in root.findall("e:Errors", NS)
        ]
        if ack not in ("Success", "Warning"):
            raise RuntimeError(f"{name} {ack}: {' / '.join(errors)}")
        return root, errors

    def get_item(self, item_id):
        root, _ = self.call("GetItem", f"<ItemID>{item_id}</ItemID>"
                            "<DetailLevel>ReturnAll</DetailLevel>")
        item = root.find("e:Item", NS)
        return {
            "title": item.findtext("e:Title", "", NS),
            "status": item.findtext("e:SellingStatus/e:ListingStatus", "", NS),
            "pictures": [p.text for p in
                         item.findall("e:PictureDetails/e:PictureURL", NS)],
        }

    def upload_external(self, url):
        root, _ = self.call(
            "UploadSiteHostedPictures",
            f"<ExternalPictureURL>{escape(url)}</ExternalPictureURL>"
            "<PictureSet>Supersize</PictureSet>")
        return root.findtext("e:SiteHostedPictureDetails/e:FullURL", "", NS)

    def revise_pictures(self, item_id, eps_urls):
        pics = "".join(f"<PictureURL>{escape(u)}</PictureURL>" for u in eps_urls)
        _, warnings = self.call(
            "ReviseFixedPriceItem",
            f"<Item><ItemID>{item_id}</ItemID>"
            f"<PictureDetails>{pics}</PictureDetails></Item>")
        return warnings


def ebay_gallery_id(url):
    """i.ebayimg.com の画像ID(images/g/<id>/ または /z/<id>/)を返す。"""
    m = re.search(r"i\.ebayimg\.com/(?:images/g|.*?/z)/([^/]+)/", url)
    return m.group(1) if m else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--sku", nargs="*")
    args = ap.parse_args()

    targets = read_targets(args.xlsx, args.sku)
    print(f"対象 {len(targets)} 件 / 画像 {sum(len(t['urls']) for t in targets)} 枚")

    problems = 0
    for t in targets:
        if t["expected"] and len(t["urls"]) != t["expected"]:
            print(f"[WARN] {t['sku']}: 画像リンク {len(t['urls'])} 枚 ≠ "
                  f"差し替え後枚数 {t['expected']}")
        for u in t["urls"]:
            if ebay_gallery_id(u):
                continue  # eBay 現行画像は EPS の元画像をそのまま再利用するので検査不要
            try:
                msg = check_image(u)
            except Exception as e:
                msg = f"取得失敗 {e}"
            if msg:
                problems += 1
                print(f"[NG] {t['sku']} {u}: {msg}")
    if problems:
        sys.exit(f"画像の問題 {problems} 件。修正してから再実行してください")
    if args.dry_run:
        print("dry-run: 画像URLはすべて取得可能(JPEG/PNG・長辺500px以上)")
        return

    api = Trading()
    OUT_DIR.mkdir(exist_ok=True)
    stamp = date.today().isoformat()
    backup_path = OUT_DIR / f"picture_backup_{stamp}.json"
    backup = json.loads(backup_path.read_text()) if backup_path.exists() else {}
    results = []
    for t in targets:
        sku, item_id = t["sku"], t["item_id"]
        res = {"sku": sku, "item_id": item_id, "before": "", "after": "",
               "result": "", "message": ""}
        try:
            cur = api.get_item(item_id)
            backup.setdefault(item_id, {"sku": sku, **cur})
            backup_path.write_text(json.dumps(backup, ensure_ascii=False, indent=1))
            res["before"] = len(cur["pictures"])
            if cur["status"] != "Active":
                raise RuntimeError(f"出品状態が {cur['status']}")
            current = {ebay_gallery_id(p): p for p in cur["pictures"]}
            eps = []
            for u in t["urls"]:
                gid = ebay_gallery_id(u)
                eps.append(current.get(gid) or api.upload_external(u))
            warnings = api.revise_pictures(item_id, eps)
            after = api.get_item(item_id)["pictures"]
            res["after"] = len(after)
            res["result"] = "OK" if len(after) == len(t["urls"]) else "要確認"
            res["message"] = " / ".join(warnings)
        except Exception as e:
            res["result"], res["message"] = "失敗", str(e)
        results.append(res)
        print(f"{sku} {item_id}: {res['result']} "
              f"{res['before']}→{res['after']}枚 {res['message']}")

    path = OUT_DIR / f"image_update_result_{stamp}.csv"
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(results[0].keys()))
        w.writeheader()
        w.writerows(results)
    print(f"wrote {path}")


if __name__ == "__main__":
    main()
