"""eBay Browse API で横展開キーワードの現行出品をリサーチする。

環境変数 EBAY_CLIENT_ID / EBAY_CLIENT_SECRET を使用。
出力: output/listings.csv (全件), output/summary.csv (キーワード別集計)
"""

import base64
import csv
import json
import os
import statistics
import sys
import urllib.parse
import urllib.request
from datetime import date
from pathlib import Path

from keywords import KEYWORDS

TOKEN_URL = "https://api.ebay.com/identity/v1/oauth2/token"
SEARCH_URL = "https://api.ebay.com/buy/browse/v1/item_summary/search"
SCOPE = "https://api.ebay.com/oauth/api_scope"
MARKETPLACE = "EBAY_US"
PRICE_MIN, PRICE_MAX = 300, 2000
LIMIT = 20
OUT_DIR = Path(__file__).parent / "output"


def get_token():
    cid = os.environ["EBAY_CLIENT_ID"]
    secret = os.environ["EBAY_CLIENT_SECRET"]
    auth = base64.b64encode(f"{cid}:{secret}".encode()).decode()
    body = urllib.parse.urlencode(
        {"grant_type": "client_credentials", "scope": SCOPE}
    ).encode()
    req = urllib.request.Request(
        TOKEN_URL,
        data=body,
        headers={
            "Authorization": f"Basic {auth}",
            "Content-Type": "application/x-www-form-urlencoded",
        },
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.load(r)["access_token"]


def search(token, query):
    params = {
        "q": query,
        "limit": LIMIT,
        "filter": f"price:[{PRICE_MIN}..{PRICE_MAX}],priceCurrency:USD",
    }
    url = f"{SEARCH_URL}?{urllib.parse.urlencode(params)}"
    req = urllib.request.Request(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "X-EBAY-C-MARKETPLACE-ID": MARKETPLACE,
        },
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.load(r)


def shipping_cost(item):
    opts = item.get("shippingOptions") or []
    if not opts:
        return ""
    cost = opts[0].get("shippingCost") or {}
    return cost.get("value", "")


def main():
    OUT_DIR.mkdir(exist_ok=True)
    token = get_token()
    rows, summary = [], []

    for axis, label, query in KEYWORDS:
        try:
            data = search(token, query)
        except urllib.error.HTTPError as e:
            print(f"[ERR] {query}: {e.code} {e.read()[:200]!r}", file=sys.stderr)
            continue
        items = data.get("itemSummaries") or []
        prices = []
        for rank, it in enumerate(items, 1):
            price = float(it["price"]["value"])
            prices.append(price)
            seller = it.get("seller") or {}
            loc = it.get("itemLocation") or {}
            rows.append({
                "axis": axis,
                "keyword_ja": label,
                "query": query,
                "rank": rank,
                "title": it.get("title", ""),
                "price_usd": price,
                "shipping_usd": shipping_cost(it),
                "condition": it.get("condition", ""),
                "buying_options": "/".join(it.get("buyingOptions") or []),
                "seller": seller.get("username", ""),
                "seller_feedback": seller.get("feedbackScore", ""),
                "seller_pct": seller.get("feedbackPercentage", ""),
                "location": loc.get("country", ""),
                "listed_at": it.get("itemCreationDate", ""),
                "category": "/".join(
                    c.get("categoryName", "") for c in it.get("categories") or []
                ),
                "item_id": it.get("itemId", ""),
                "url": it.get("itemWebUrl", ""),
            })
        summary.append({
            "axis": axis,
            "keyword_ja": label,
            "query": query,
            "total_hits_in_range": data.get("total", 0),
            "fetched": len(items),
            "jp_sellers": sum(
                1 for it in items
                if (it.get("itemLocation") or {}).get("country") == "JP"
            ),
            "min_usd": min(prices) if prices else "",
            "median_usd": round(statistics.median(prices), 2) if prices else "",
            "max_usd": max(prices) if prices else "",
        })
        print(f"{query}: total={data.get('total', 0)} fetched={len(items)}")

    stamp = date.today().isoformat()
    for name, data in (("listings", rows), ("summary", summary)):
        if not data:
            continue
        path = OUT_DIR / f"{name}_{stamp}.csv"
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(data[0].keys()))
            w.writeheader()
            w.writerows(data)
        print(f"wrote {path} ({len(data)} rows)")


if __name__ == "__main__":
    main()
