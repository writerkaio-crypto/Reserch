"""low_image_listings.py --csv の取得結果 (画像キャッシュ) から Excel レポートを作る。

  python make_excel_report.py active_listings.csv

出力: output/eBay画像枚数チェック_YYYY-MM-DD.xlsx
  - サマリー   : 件数の集計 (数式)
  - 要対応SKU  : いずれかのサイトで画像が 2 枚以下の SKU
  - 要対応出品 : 上記に該当する出品 (サイト別・商品ページへのリンク付き)
  - 全SKU一覧  : 全 SKU × サイトの画像枚数
"""

import json
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from low_image_listings import CACHE_PATH, OUT_DIR, read_seller_hub_csv

MAX_IMAGES = 2
SITES = ["US", "UK", "DE", "FR", "IT", "ES", "CA", "AU"]
HIDDEN, UNCHECKED = "非表示", "未取得"

FONT = "Meiryo UI"
F_BASE = Font(name=FONT, size=10)
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True)
F_NOTE = Font(name=FONT, size=9, color="666666")
F_INPUT = Font(name=FONT, size=10, bold=True, color="0000FF")
F_LINK = Font(name=FONT, size=10, color="0563C1", underline="single")
FILL_HEAD = PatternFill("solid", fgColor="305496")
FILL_RED = PatternFill("solid", fgColor="F8CBAD")
FILL_YELLOW = PatternFill("solid", fgColor="FFF2CC")
FILL_GRAY = PatternFill("solid", fgColor="EDEDED")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")
HEAD_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
THRESHOLD = "サマリー!$C$4"


def site_value(lis):
    """同じサイトに複数出品がある場合は最少枚数。数値が無ければ状態文字列。"""
    nums = [li["image_count"] for li in lis if isinstance(li["image_count"], int)]
    if nums:
        return min(nums)
    if any(li["image_count"] is None for li in lis):
        return HIDDEN
    return UNCHECKED


def load():
    listings = read_seller_hub_csv(sys.argv[1])
    cache = json.loads(CACHE_PATH.read_text()) if CACHE_PATH.exists() else {}
    for li in listings:
        li["image_count"] = cache.get(li["item_id"], UNCHECKED)
    by_sku = {}
    for li in listings:
        by_sku.setdefault(li["sku"], []).append(li)
    skus = []
    for sku, lis in by_sku.items():
        rep = next((li for li in lis if li["site"] == "US"), lis[0])
        per_site = {s: site_value([li for li in lis if li["site"] == s])
                    for s in SITES if any(li["site"] == s for li in lis)}
        nums = [v for v in per_site.values() if isinstance(v, int)]
        skus.append({"sku": sku, "title": rep["title"], "sites": per_site,
                     "min": min(nums) if nums else None})
    skus.sort(key=lambda r: (r["min"] is None, r["min"] or 0,
                             int(r["sku"]) if r["sku"].isdigit() else 10**9, r["sku"]))
    return listings, skus


def header(ws, row, names, widths):
    for c, (name, w) in enumerate(zip(names, widths), 1):
        cell = ws.cell(row=row, column=c, value=name)
        cell.font, cell.fill, cell.alignment, cell.border = F_HEAD, FILL_HEAD, HEAD_ALIGN, BORDER
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[row].height = 32
    ws.freeze_panes = ws.cell(row=row + 1, column=3)


def style_body(ws, first, last, ncols, center_cols=()):
    for row in ws.iter_rows(min_row=first, max_row=last, max_col=ncols):
        for cell in row:
            cell.border = BORDER
            if cell.font != F_LINK:
                cell.font = F_BASE
            if cell.column in center_cols:
                cell.alignment = CENTER


def sku_matrix_sheet(ws, skus):
    """SKU × サイトの画像枚数表。最少枚数・該当サイト数は数式。"""
    names = ["SKU", "商品名 (US)", "最少枚数", "2枚以下の\nサイト数", "出品サイト数",
             "全サイト\n該当"] + SITES
    header(ws, 1, names, [9, 60, 9, 11, 11, 9] + [6] * len(SITES))
    first_site = get_column_letter(7)
    last_site = get_column_letter(6 + len(SITES))
    for r, s in enumerate(skus, 2):
        rng = f"{first_site}{r}:{last_site}{r}"
        ws.cell(row=r, column=1, value=int(s["sku"]) if s["sku"].isdigit() else s["sku"])
        ws.cell(row=r, column=2, value=s["title"])
        ws.cell(row=r, column=3, value=f'=IF(COUNT({rng})=0,"-",MIN({rng}))')
        ws.cell(row=r, column=4, value=f'=COUNTIF({rng},"<="&{THRESHOLD})')
        ws.cell(row=r, column=5, value=f"=COUNTA({rng})")
        ws.cell(row=r, column=6, value=f'=IF(AND(D{r}>0,D{r}=E{r}),"全て","")')
        for c, site in enumerate(SITES, 7):
            if site in s["sites"]:
                ws.cell(row=r, column=c, value=s["sites"][site])
    last = len(skus) + 1
    style_body(ws, 2, last, len(names), center_cols=set(range(3, len(names) + 1)))
    grid = f"{first_site}2:{last_site}{last}"
    ws.conditional_formatting.add(grid, FormulaRule(
        formula=[f"AND(ISNUMBER(G2),G2<={THRESHOLD})"], fill=FILL_RED))
    ws.conditional_formatting.add(grid, FormulaRule(
        formula=['OR(G2="非表示",G2="未取得")'], fill=FILL_GRAY))
    ws.conditional_formatting.add(f"C2:C{last}", FormulaRule(
        formula=[f"AND(ISNUMBER(C2),C2<={THRESHOLD})"], fill=FILL_RED,
        font=Font(name=FONT, bold=True, color="C00000")))
    ws.conditional_formatting.add(f"F2:F{last}", CellIsRule(
        operator="equal", formula=['"全て"'], fill=FILL_YELLOW,
        font=Font(name=FONT, bold=True)))
    ws.auto_filter.ref = f"A1:{last_site}{last}"
    return last


def listing_sheet(ws, listings):
    names = ["SKU", "サイト", "商品番号", "画像枚数", "在庫数", "タイトル", "商品ページ"]
    header(ws, 1, names, [9, 8, 15, 9, 8, 70, 12])
    ws.freeze_panes = "A2"
    for r, li in enumerate(listings, 2):
        ws.cell(row=r, column=1, value=int(li["sku"]) if li["sku"].isdigit() else li["sku"])
        ws.cell(row=r, column=2, value=li["site"])
        ws.cell(row=r, column=3, value=li["item_id"])
        ws.cell(row=r, column=4, value=li["image_count"])
        ws.cell(row=r, column=5, value=int(li["quantity"]) if li["quantity"].isdigit()
                else li["quantity"])
        ws.cell(row=r, column=6, value=li["title"])
        link = ws.cell(row=r, column=7, value="開く")
        link.hyperlink = f"https://www.ebay.com/itm/{li['item_id']}"
        link.font = F_LINK
    last = len(listings) + 1
    style_body(ws, 2, last, len(names), center_cols={1, 2, 3, 4, 5, 7})
    ws.conditional_formatting.add(f"D2:D{last}", CellIsRule(
        operator="lessThanOrEqual", formula=["1"], fill=FILL_RED,
        font=Font(name=FONT, bold=True, color="C00000")))
    ws.auto_filter.ref = f"A1:G{last}"


def summary_sheet(ws, n_listings, stamp):
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 70
    ws["B2"] = "eBay 出品画像チェック"
    ws["B2"].font = F_TITLE
    ws["D2"] = f"作成日: {stamp}"
    ws["D2"].font = F_NOTE

    rows = [
        ("抽出条件: 画像枚数がこの枚数以下", MAX_IMAGES,
         "青字を変えると件数と色分けが再計算されます（要対応シートは2枚で抽出済み）"),
        None,
        ("出品数（全サイト合計）", n_listings,
         "Seller Hub の出品中レポート（CSV）の行数"),
        ("SKU数", "=COUNTA(全SKU一覧!A:A)-1", ""),
        ("画像が条件以下のSKU（いずれかのサイト）", '=COUNTIF(全SKU一覧!D:D,">0")',
         "1サイトでも該当すれば対象"),
        ("　うち 全サイトで条件以下", '=COUNTIF(全SKU一覧!F:F,"全て")',
         "元の画像自体が少ない → 画像追加の優先度 高"),
        ("　うち 一部サイトのみ条件以下", "=C8-C9",
         "そのサイトだけ画像の同期漏れの可能性"),
        ("画像が条件以下の出品数", f'=COUNTIF(全SKU一覧!G:N,"<="&{THRESHOLD})',
         "SKU×サイト単位（同一サイトに複数出品がある場合は1件として集計）"),
        ("画像を確認できなかった出品", '=COUNTIF(全SKU一覧!G:N,"非表示")',
         "在庫0などで eBay 上に表示されていない出品"),
        ("未取得の出品", '=COUNTIF(全SKU一覧!G:N,"未取得")',
         "API 上限のため未取得（0 になれば全件確認済み）"),
    ]
    for i, row in enumerate(rows, 4):
        if row is None:
            continue
        label, value, note = row
        ws.cell(row=i, column=2, value=label).font = F_BASE
        c = ws.cell(row=i, column=3, value=value)
        c.font = F_INPUT if i == 4 else Font(name=FONT, size=11, bold=True)
        c.alignment = CENTER
        c.number_format = "#,##0"
        ws.cell(row=i, column=4, value=note).font = F_NOTE
        for col in (2, 3, 4):
            ws.cell(row=i, column=col).border = BORDER
    ws["C4"].fill = FILL_YELLOW

    notes = [
        "シートの見方",
        "・要対応SKU：いずれかのサイトで画像が2枚以下の SKU。サイト列は各サイトの画像枚数（赤＝2枚以下）",
        "・要対応出品：該当する出品を1行ずつ（サイト別）。「開く」で商品ページへ",
        "・全SKU一覧：全 SKU の画像枚数（フィルターで絞り込み可）",
        "・「全サイト該当」が「全て」の SKU は、どのサイトでも画像が少ない商品です",
        "・灰色の「非表示」は在庫0などで表示されず確認できなかった出品です",
        f"・データ元：Seller Hub 出品中レポート（{n_listings:,} 出品）＋ eBay Browse API で取得した画像枚数",
    ]
    for i, t in enumerate(notes, 16):
        ws.cell(row=i, column=2, value=t).font = (
            Font(name=FONT, size=10, bold=True) if i == 16 else F_BASE)
    ws.sheet_view.showGridLines = False


def main():
    listings, skus = load()
    stamp = date.today().isoformat()
    wb = Workbook()
    summary_sheet(wb.active, len(listings), stamp)
    wb.active.title = "サマリー"

    def is_low(v):
        return isinstance(v, int) and v <= MAX_IMAGES

    low_skus = [s for s in skus if any(is_low(v) for v in s["sites"].values())]
    low_listings = sorted(
        (li for li in listings if is_low(li["image_count"])),
        key=lambda li: (li["image_count"],
                        int(li["sku"]) if li["sku"].isdigit() else 10**9,
                        SITES.index(li["site"]) if li["site"] in SITES else 99))
    sku_matrix_sheet(wb.create_sheet("要対応SKU"), low_skus)
    listing_sheet(wb.create_sheet("要対応出品"), low_listings)
    sku_matrix_sheet(wb.create_sheet("全SKU一覧"), skus)

    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        if ws.title != "サマリー":
            ws.print_title_rows = "1:1"

    OUT_DIR.mkdir(exist_ok=True)
    path = OUT_DIR / f"eBay画像枚数チェック_{stamp}.xlsx"
    wb.save(path)
    print(f"wrote {path}: 要対応 {len(low_skus)} SKU / {len(low_listings)} 出品")


if __name__ == "__main__":
    main()
