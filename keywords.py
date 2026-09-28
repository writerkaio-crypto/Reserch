"""横展開キーワード定義。

元キーワード: 「五十嵐大介 画集 海獣とタマシイ」
  要素分解:  [作家]=五十嵐大介 / [形態]=画集 / [作品]=海獣とタマシイ(海獣の子供)
  展開ルール:
    A. 作家据え置き × 形態/作品入れ替え   (同じ作家の別の高額アイテム)
    B. 作家入れ替え × 形態据え置き(画集)  (絵画的・作家性の強い漫画家の画集)
    C. 作家入れ替え × 形態入れ替え(原画・サイン色紙) (高単価帯を狙う)
"""

# (axis, 日本語ラベル, eBay検索クエリ)
KEYWORDS = [
    # A. 作家据え置き
    ("A", "五十嵐大介 画集(海獣とタマシイ)", "Daisuke Igarashi art book"),
    ("A", "五十嵐大介 海獣の子供", "Children of the Sea Igarashi"),
    ("A", "五十嵐大介 サイン/原画", "Daisuke Igarashi signed"),
    # B. 作家入れ替え × 画集
    ("B", "松本大洋 画集", "Taiyo Matsumoto art book"),
    ("B", "寺田克也 画集", "Katsuya Terada art book"),
    ("B", "弐瓶勉 画集", "Tsutomu Nihei art book"),
    ("B", "鶴田謙二 画集", "Kenji Tsuruta art book"),
    ("B", "大友克洋 画集", "Katsuhiro Otomo art book"),
    ("B", "士郎正宗 画集", "Masamune Shirow Intron Depot"),
    ("B", "諸星大二郎 画集", "Daijiro Morohoshi"),
    ("B", "丸尾末広 画集", "Suehiro Maruo art book"),
    ("B", "市川春子 画集(宝石の国)", "Haruko Ichikawa art book"),
    ("B", "森薫 画集", "Kaoru Mori art book"),
    ("B", "安倍吉俊 画集", "Yoshitoshi ABe art book"),
    ("B", "村田蓮爾 画集", "Range Murata art book"),
    # C. 作家入れ替え × 原画・サイン
    ("C", "松本大洋 サイン/原画", "Taiyo Matsumoto signed"),
    ("C", "弐瓶勉 サイン/原画", "Tsutomu Nihei signed"),
    ("C", "鶴田謙二 サイン/原画", "Kenji Tsuruta signed"),
    ("C", "寺田克也 原画/ドローイング", "Katsuya Terada original drawing"),
    ("C", "大友克洋 サイン/原画", "Katsuhiro Otomo signed"),
]
