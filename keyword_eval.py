"""横展開キーワード候補を eBay（Browse API）とメルカリで評価する。

使い方: python3 keyword_eval.py candidates.json result.json
  candidates.json: [["日本語の語句", "eBay検索語", [メルカリ題名の必須語...], [eBay題名の必須語...]], ...]

指標:
  - eBay: $70-700 の出品数、上位60件のうち販売実績あり件数・販売個数・日本からの出品数
    （Browse API は出品中の商品の販売数しか返さないため、1点物は少なめに出る）
  - 仕入れ上限(円) = eBay価格中央値 × RATE × (1 - FEE) - SHIP
  - メルカリ: 販売中・1万〜10万円のうち、仕入れ上限以下の件数
点数 = 販売実績あり件数 × (1 - 日本比率)。仕入れ上限以下の商品がある語の中から点数順に採用する。
環境変数 EBAY_CLIENT_ID / EBAY_CLIENT_SECRET が必要。依存: pip install mercapi
"""
import asyncio, base64, json, os, statistics, sys, urllib.parse, urllib.request
from concurrent.futures import ThreadPoolExecutor
from mercapi import Mercapi
from mercapi.requests import SearchRequestData
RATE, FEE, SHIP = 145, 0.20, 6000
cid,sec=os.environ["EBAY_CLIENT_ID"],os.environ["EBAY_CLIENT_SECRET"]
auth=base64.b64encode(f"{cid}:{sec}".encode()).decode()
T=json.load(urllib.request.urlopen(urllib.request.Request("https://api.ebay.com/identity/v1/oauth2/token",data=urllib.parse.urlencode({"grant_type":"client_credentials","scope":"https://api.ebay.com/oauth/api_scope"}).encode(),headers={"Authorization":"Basic "+auth,"Content-Type":"application/x-www-form-urlencoded"}),timeout=30))["access_token"]
def get(url):
    import time
    r=urllib.request.Request(url,headers={"Authorization":"Bearer "+T,"X-EBAY-C-MARKETPLACE-ID":"EBAY_US"})
    for n in range(4):
        try: return json.load(urllib.request.urlopen(r,timeout=30))
        except urllib.error.HTTPError: raise
        except Exception:
            if n==3: raise
            time.sleep(2**(n+1))
def ebay(q, keys):
    d=get("https://api.ebay.com/buy/browse/v1/item_summary/search?"+urllib.parse.urlencode({"q":q,"limit":200,"filter":"price:[70..700],priceCurrency:USD"}))
    its=[i for i in d.get("itemSummaries",[]) if any(k.lower() in i["title"].lower() for k in keys)]
    def one(i):
        try: return get("https://api.ebay.com/buy/browse/v1/item/"+urllib.parse.quote(i["itemId"]))
        except Exception: return {}
    with ThreadPoolExecutor(8) as ex: got=list(ex.map(one,its[:60]))
    sold_l=sum(1 for g in got if sum(a.get("estimatedSoldQuantity",0) for a in g.get("estimatedAvailabilities",[]))>0)
    sold_q=sum(sum(a.get("estimatedSoldQuantity",0) for a in g.get("estimatedAvailabilities",[])) for g in got)
    jp=sum(1 for g in got if g.get("itemLocation",{}).get("country")=="JP")
    pr=[float(i["price"]["value"]) for i in its]
    return dict(total=d.get("total",0), rel=len(its), checked=len(got), sold_listings=sold_l, sold_qty=sold_q, jp=jp,
                jp_ratio=round(jp/len(got),2) if got else None, ebay_med=statistics.median(pr) if pr else None)
async def merc(m, q, keys):
    items=[]; r=await m.search(q, price_min=10000, price_max=100000, status=[SearchRequestData.Status.STATUS_ON_SALE])
    while True:
        items+=r.items
        if not r.meta.next_page_token or len(items)>=600: break
        r=await r.next_page()
    rel=[i for i in items if 10000<=i.price<=100000 and any(k.lower() in i.name.lower() for k in keys)]
    return rel
async def main(cands):
    m=Mercapi(); out=[]
    for ja,en,jk,ek in cands:
        e=ebay(en,ek); rel=await merc(m,ja,jk)
        cap=round(e["ebay_med"]*RATE*(1-FEE)-SHIP) if e["ebay_med"] else 0
        row={"ja":ja,"en":en,**e,"m_count":len(rel),"m_med":statistics.median([i.price for i in rel]) if rel else None,"cap":cap,"m_under_cap":sum(1 for i in rel if i.price<=cap)}
        out.append(row); print(json.dumps(row,ensure_ascii=False),flush=True)
    return out
if __name__=="__main__":
    cands=json.load(open(sys.argv[1])); res=asyncio.run(main(cands))
    json.dump(res,open(sys.argv[2],"w"),ensure_ascii=False,indent=1)
