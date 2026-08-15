"""月次予算割振ブックを生成する。"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FONT = "Meiryo UI"
BLUE = "0000FF"       # 入力値
BLACK = "000000"      # 数式
YELLOW = "FFF2CC"     # 入力セル
GRAY = "F2F2F2"       # 見出し
HEAD = "D9E2F3"       # 表ヘッダ
TOTALFILL = "E2EFDA"  # 合計行

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

YEN = '#,##0;[Red]-#,##0;"-"'
PCT = '0.0%;[Red]-0.0%;"-"'
NUM = '#,##0.##;[Red]-#,##0.##;"-"'

wb = openpyxl.Workbook()


def style(ws, ref, *, bold=False, size=11, color=BLACK, fill=None, fmt=None,
          align=None, border=False, wrap=False):
    rng = ws[ref]
    rows = rng if isinstance(rng, tuple) else ((rng,),)
    if rows and not isinstance(rows[0], tuple):
        rows = (rows,)
    for row in rows:
        for c in row:
            c.font = Font(name=FONT, size=size, bold=bold, color=color)
            if fill:
                c.fill = PatternFill("solid", fgColor=fill)
            if fmt:
                c.number_format = fmt
            if align or wrap:
                c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
            if border:
                c.border = BORDER


def put(ws, ref, value, **kw):
    ws[ref] = value
    style(ws, ref, **kw)


def title(ws, text, sub=None):
    put(ws, "A1", text, bold=True, size=14)
    if sub:
        put(ws, "A2", sub, size=9, color="808080")


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


# =====================================================================
# 1. 設定
# =====================================================================
s = wb.active
s.title = "設定"
title(s, "月次予算割振", "対象月を変えるだけで、日付・週の区切り・WEEK番号がすべて入れ替わります。")

put(s, "A4", "■ 対象月", bold=True, fill=GRAY)
for row, (label, val) in enumerate(
        [("年", 2026), ("月", 4)], start=5):
    put(s, f"A{row}", label, bold=True)
    put(s, f"B{row}", val, color=BLUE, fill=YELLOW, border=True, align="center")
put(s, "A7", "開始WEEK", bold=True)
put(s, "B7", "=B16", color=BLUE, fill=YELLOW, border=True, align="center")
put(s, "C7", "← 自動計算値が入っています。別の番号にしたいときは直接上書きしてください。",
    size=9, color="808080")

put(s, "A9", "■ 自動計算（変更しないでください）", bold=True, fill=GRAY)
auto = [
    ("月初", "=DATE(B5,B6,1)", "yyyy/m/d"),
    ("月末", "=EOMONTH(B11,0)", "yyyy/m/d"),
    ("第1週の木曜", "=B11+MOD(4-WEEKDAY(B11,2)+7,7)", "yyyy/m/d"),
    ("3/1(当年)の週の木曜", "=DATE(B5,3,1)+MOD(4-WEEKDAY(DATE(B5,3,1),2)+7,7)", "yyyy/m/d"),
    ("3/1(前年)の週の木曜", "=DATE(B5-1,3,1)+MOD(4-WEEKDAY(DATE(B5-1,3,1),2)+7,7)", "yyyy/m/d"),
    ("開始WEEK（自動）", "=IF(B13>=B14,(B13-B14)/7+1,(B13-B15)/7+1)", "0"),
    ("週数", "=IF(B12<=B13,1,1+CEILING((B12-B13)/7,1))", "0"),
]
for i, (label, f, fmt) in enumerate(auto, start=11):
    put(s, f"A{i}", label)
    put(s, f"B{i}", f, fmt=fmt, align="center", border=True)

put(s, "A19", "■ 週の区切り（金曜〜木曜。月初・月末は端数の週になります）", bold=True, fill=GRAY)
for col, label in zip("ABCDE", ["週", "WEEK", "開始", "終了", "日数"]):
    put(s, f"{col}20", label, bold=True, fill=HEAD, align="center", border=True)
for i in range(1, 7):
    r = 20 + i
    put(s, f"A{r}", i, align="center", border=True)
    put(s, f"B{r}", f'=IF($A{r}<=$B$17,"W"&($B$7+$A{r}-1),"")', align="center", border=True)
    put(s, f"C{r}", f'=IF($A{r}<=$B$17,IF($A{r}=1,$B$11,$B$13+7*($A{r}-2)+1),"")',
        fmt="m/d", align="center", border=True)
    put(s, f"D{r}", f'=IF($A{r}<=$B$17,MIN($B$12,$B$13+7*($A{r}-1)),"")',
        fmt="m/d", align="center", border=True)
    put(s, f"E{r}", f'=IF($A{r}<=$B$17,$D{r}-$C{r}+1,"")', align="center", border=True)

put(s, "A28", "■ 部門構成比（接客部門シートで使用）", bold=True, fill=GRAY)
for i, (name, ratio) in enumerate([
        ("FLOOR", 0.442736342367074),
        ("CONCESSION", 0.4321965545697439),
        ("STORE", 0.03629238541840829),
        ("OFFICE", 0.08877471764477365)], start=29):
    put(s, f"A{i}", name, border=True)
    put(s, f"B{i}", ratio, color=BLUE, fill=YELLOW, fmt="0.00%", align="center", border=True)
put(s, "A33", "合計", bold=True, border=True)
put(s, "B33", "=SUM(B29:B32)", bold=True, fmt="0.00%", align="center", border=True)
put(s, "C33", "← 100.00% になるようにしてください", size=9, color="808080")

put(s, "A35", "■ 凡例", bold=True, fill=GRAY)
put(s, "A36", "入力するセル", fill=YELLOW, color=BLUE, border=True)
put(s, "B36", "黄色 ＝ 手で入力するところ（青字）", size=9, color="808080")
put(s, "A37", "自動計算", border=True)
put(s, "B37", "白 ＝ 数式が入っています（触らないでください）", size=9, color="808080")
put(s, "A39", "※ 元の値は 2026年4月シートから写しています。実態に合わせて書き換えてください。",
    size=9, color="808080")
widths(s, {"A": 24, "B": 14, "C": 52, "D": 12, "E": 8})

# =====================================================================
# 2. 動員
# =====================================================================
a = wb.create_sheet("動員")
title(a, "日別動員", "D列に日別動員を入力すると、週別の構成比まで自動で出ます（未入力でも他シートは使えます）。")
for col, label in zip("ABCDE", ["日付", "曜日", "WEEK", "動員", "構成比"]):
    put(a, f"{col}3", label, bold=True, fill=HEAD, align="center", border=True)
for i in range(31):
    r = 4 + i
    put(a, f"A{r}", f'=IF({i + 1}<=DAY(設定!$B$12),設定!$B$11+{i},"")',
        fmt="m/d", align="center", border=True)
    put(a, f"B{r}", f'=IF($A{r}="","",CHOOSE(WEEKDAY($A{r},2),"月","火","水","木","金","土","日"))',
        align="center", border=True)
    put(a, f"C{r}", f'=IF($A{r}="","",IF($A{r}<=設定!$B$13,1,1+CEILING(($A{r}-設定!$B$13)/7,1)))',
        align="center", border=True)
    put(a, f"D{r}", None, color=BLUE, fill=YELLOW, fmt=YEN, border=True)
    put(a, f"E{r}", f'=IF($A{r}="","",IF($D$35=0,0,$D{r}/$D$35))', fmt=PCT, border=True)
put(a, "A35", "合計", bold=True, fill=TOTALFILL, border=True)
for col in "BC":
    put(a, f"{col}35", None, fill=TOTALFILL, border=True)
put(a, "D35", "=SUM(D4:D34)", bold=True, fill=TOTALFILL, fmt=YEN, border=True)
put(a, "E35", '=IF($D$35=0,"",SUM(E4:E34))', bold=True, fill=TOTALFILL, fmt=PCT, border=True)

for col, label in zip("GHIJK", ["週", "WEEK", "期間", "動員", "構成比"]):
    put(a, f"{col}3", label, bold=True, fill=HEAD, align="center", border=True)
for i in range(1, 7):
    r = 3 + i
    sr = 20 + i  # 設定シートの対応行
    put(a, f"G{r}", i, align="center", border=True)
    put(a, f"H{r}", f'=設定!$B${sr}', align="center", border=True)
    put(a, f"I{r}", f'=IF(設定!$B${sr}="","",TEXT(設定!$C${sr},"m/d")&"〜"&TEXT(設定!$D${sr},"m/d"))',
        align="center", border=True)
    put(a, f"J{r}", f'=IF(設定!$B${sr}="","",SUMIF($C$4:$C$34,$G{r},$D$4:$D$34))',
        fmt=YEN, border=True)
    put(a, f"K{r}", f'=IF(設定!$B${sr}="","",IF($J$10=0,0,$J{r}/$J$10))', fmt=PCT, border=True)
put(a, "I10", "合計", bold=True, fill=TOTALFILL, align="center", border=True)
put(a, "J10", "=SUM(J4:J9)", bold=True, fill=TOTALFILL, fmt=YEN, border=True)
put(a, "K10", '=IF($J$10=0,"",SUM(K4:K9))', bold=True, fill=TOTALFILL, fmt=PCT, border=True)
for cell in ("G3", "H3", "I3", "J3", "K3"):
    pass
widths(a, {"A": 9, "B": 7, "C": 7, "D": 12, "E": 10, "F": 3,
           "G": 6, "H": 8, "I": 16, "J": 12, "K": 10})
a.freeze_panes = "A4"

# =====================================================================
# 3. 割振（メイン）
# =====================================================================
p = wb.create_sheet("割振")
title(p, "項目別の実施予定",
      "黄色のセルに入力します。金額 ＝ 単価 × 時間 × 人数 × 回数（回数＝その週に何回・何日実施するか）。行は自由に増やせます。")

p.merge_cells("F3:K3")
put(p, "F3", "週ごとの回数（日数）を入力", bold=True, fill=HEAD, align="center", border=True)
p.merge_cells("N3:S3")
put(p, "N3", "週別金額（自動計算）", bold=True, fill=GRAY, align="center", border=True)

heads = {"A": "項目", "B": "内容", "C": "単価(時給)", "D": "時間", "E": "人数",
         "L": "回数計", "M": "金額計"}
for col, label in heads.items():
    put(p, f"{col}4", label, bold=True, fill=HEAD, align="center", border=True, wrap=True)
for i in range(6):
    wk = get_column_letter(6 + i)     # F..K
    amt = get_column_letter(14 + i)   # N..S
    ref = f"設定!$B${21 + i}"
    put(p, f"{wk}4", f"={ref}", bold=True, fill=HEAD, align="center", border=True)
    put(p, f"{amt}4", f"={ref}", bold=True, fill=GRAY, align="center", border=True)

ITEMS = [
    ("障がい者雇用", "駒澤さん", 1236, 5, 1, [3, 3, 3, 3, 3, None]),
    ("新規採用研修（トレーニー）", "週2〜3名", 1226, 5.5, 2, [None] * 6),
    ("新規採用研修（トレーニー）", "カンポリ参加", 1226, 2.15, 3, [None] * 6),
    ("新規採用研修（トレーナー）", "週2〜3名", 1251, 5.5, 1, [None] * 6),
    ("新規採用研修（トレーナー）", "通常", 1251, 5, 4, [None] * 6),
    ("既存スタッフ研修", "週1", 1251, 1, 1, [1, None, None, None, None, None]),
    ("既存スタッフ研修", "週2", 1251, 1.5, 3, [None, 1, None, None, None, None]),
    ("既存スタッフ研修", "週3", 1251, 1.75, 6, [None, None, 1, None, None, None]),
    ("既存スタッフ研修", "週4", 1251, 1.5, 6, [None, None, None, 1, None, None]),
    ("防災訓練", "週1", 1251, 0.5, 10, [1, None, None, None, None, None]),
    ("防災訓練", "週2", 1251, 0.5, 15, [None, 1, None, None, None, None]),
    ("防災訓練", "週3", 1251, 0.75, 8, [None, None, 1, None, None, None]),
    ("防災訓練", "週4", 1251, 0.75, 2, [None, None, None, 1, None, None]),
    ("防災訓練", "総合訓練", 1251, 1, 8, [None] * 6),
    ("棚卸（予算分）", "棚卸", 1251, 5, 5, [None] * 6),
    ("ストア準備・返品（予算分）", "定例（単価以外は変更不可）", 1251, 2, 1, [1, 1, 1, 1, 1, None]),
    ("ストア準備・返品（追加分）", "準備", 1251, 2, 2, [None] * 6),
    ("ストア準備・返品（追加分）", "返品", 1251, 6, 2, [None] * 6),
    ("ストア準備・返品（追加分）", "陳列", 1251, 6, 4, [None] * 6),
    ("その他", "", 1251, None, None, [None] * 6),
    ("リーダー手当", "手当（金額を単価欄に入れ、時間・人数・回数を1にする）", 93000, 1, 1, [None] * 6),
]

FIRST, LAST = 5, 104
for r in range(FIRST, LAST + 1):
    idx = r - FIRST
    item = ITEMS[idx] if idx < len(ITEMS) else None
    put(p, f"A{r}", item[0] if item else None, color=BLUE, fill=YELLOW, border=True)
    put(p, f"B{r}", item[1] if item else None, color=BLUE, fill=YELLOW, border=True)
    put(p, f"C{r}", item[2] if item else None, color=BLUE, fill=YELLOW, fmt=YEN, border=True)
    put(p, f"D{r}", item[3] if item else None, color=BLUE, fill=YELLOW, fmt=NUM, border=True)
    put(p, f"E{r}", item[4] if item else None, color=BLUE, fill=YELLOW, fmt=NUM, border=True)
    for i in range(6):
        wk = get_column_letter(6 + i)
        amt = get_column_letter(14 + i)
        put(p, f"{wk}{r}", item[5][i] if item else None,
            color=BLUE, fill=YELLOW, fmt=NUM, align="center", border=True)
        put(p, f"{amt}{r}",
            f'=IF(OR($A{r}="",{amt}$4=""),"",$C{r}*$D{r}*$E{r}*{wk}{r})',
            fmt=YEN, border=True)
    put(p, f"L{r}", f'=IF($A{r}="","",SUMPRODUCT(($F$4:$K$4<>"")*$F{r}:$K{r}))',
        fmt=NUM, align="center", border=True)
    put(p, f"M{r}", f'=IF($A{r}="","",SUM($N{r}:$S{r}))', fmt=YEN, border=True)

put(p, f"A{LAST + 2}", "合計", bold=True, fill=TOTALFILL, border=True)
for col in "BCDE":
    put(p, f"{col}{LAST + 2}", None, fill=TOTALFILL, border=True)
for i in range(6):
    wk = get_column_letter(6 + i)
    amt = get_column_letter(14 + i)
    put(p, f"{wk}{LAST + 2}", f"=SUM({wk}{FIRST}:{wk}{LAST})",
        bold=True, fill=TOTALFILL, fmt=NUM, align="center", border=True)
    put(p, f"{amt}{LAST + 2}", f"=SUM({amt}{FIRST}:{amt}{LAST})",
        bold=True, fill=TOTALFILL, fmt=YEN, border=True)
put(p, f"L{LAST + 2}", f"=SUM(L{FIRST}:L{LAST})", bold=True, fill=TOTALFILL, fmt=NUM,
    align="center", border=True)
put(p, f"M{LAST + 2}", f"=SUM(M{FIRST}:M{LAST})", bold=True, fill=TOTALFILL, fmt=YEN, border=True)
put(p, f"A{LAST + 4}",
    "※ 項目名は「集計」シートの項目名と同じ文字にしてください（違うと集計に載りません。集計シートの下にチェック欄があります）。",
    size=9, color="808080")

widths(p, {"A": 26, "B": 30, "C": 11, "D": 8, "E": 8,
           **{get_column_letter(6 + i): 7 for i in range(6)},
           "L": 8, "M": 12, "N": 3,
           **{get_column_letter(14 + i): 12 for i in range(6)}})
p.column_dimensions["N"].width = 12
p.freeze_panes = "F5"

# =====================================================================
# 4. 集計
# =====================================================================
g = wb.create_sheet("集計")
title(g, "項目別 集計", "「割振」シートの明細を項目ごとに集計します。予算（I列）だけ入力してください。")

for col, label in zip("A", ["項目"]):
    put(g, "A3", "項目", bold=True, fill=HEAD, align="center", border=True)
for i in range(6):
    col = get_column_letter(2 + i)  # B..G
    put(g, f"{col}3", f"=設定!$B${21 + i}", bold=True, fill=HEAD, align="center", border=True)
for col, label in zip("HIJK", ["月計", "予算", "予算比", "差額"]):
    put(g, f"{col}3", label, bold=True, fill=HEAD, align="center", border=True)

CATS = ["障がい者雇用", "新規採用研修（トレーニー）", "新規採用研修（トレーナー）",
        "既存スタッフ研修", "防災訓練", "棚卸（予算分）", "ストア準備・返品（予算分）",
        "ストア準備・返品（追加分）", "その他", "リーダー手当"]
CAT_FIRST = 4
CAT_LAST = CAT_FIRST + 14  # 予備行を含めて15行

for i in range(15):
    r = CAT_FIRST + i
    name = CATS[i] if i < len(CATS) else None
    put(g, f"A{r}", name, color=BLUE, fill=YELLOW, border=True)
    for j in range(6):
        col = get_column_letter(2 + j)
        src = get_column_letter(14 + j)
        put(g, f"{col}{r}",
            f'=IF(OR($A{r}="",{col}$3=""),"",SUMIF(割振!$A$5:$A$104,$A{r},割振!{src}$5:{src}$104))',
            fmt=YEN, border=True)
    put(g, f"H{r}", f'=IF($A{r}="","",SUM($B{r}:$G{r}))', bold=True, fmt=YEN, border=True)
    put(g, f"I{r}", None, color=BLUE, fill=YELLOW, fmt=YEN, border=True)
    put(g, f"J{r}", f'=IF(OR($A{r}="",$I{r}=0),"",$H{r}/$I{r})', fmt=PCT, border=True)
    put(g, f"K{r}", f'=IF(OR($A{r}="",$I{r}=0),"",$I{r}-$H{r})', fmt=YEN, border=True)

LEAVE_ROW = CAT_LAST + 2   # 有給休暇
CHG_ROW = LEAVE_ROW + 1    # 更衣時間
put(g, f"A{LEAVE_ROW}", "有給休暇", border=True)
put(g, f"A{CHG_ROW}", "更衣時間", border=True)
for j in range(6):
    col = get_column_letter(2 + j)
    put(g, f"{col}{LEAVE_ROW}", f"='有休・更衣'!$E${8 + j}", fmt=YEN, border=True)
    put(g, f"{col}{CHG_ROW}", f"='有休・更衣'!$D${27 + j}", fmt=YEN, border=True)
for r, budget in ((LEAVE_ROW, "='有休・更衣'!$B$5"), (CHG_ROW, "='有休・更衣'!$B$34")):
    put(g, f"H{r}", f"=SUM($B{r}:$G{r})", bold=True, fmt=YEN, border=True)
    put(g, f"I{r}", budget, fmt=YEN, border=True)
    put(g, f"J{r}", f'=IF($I{r}=0,"",$H{r}/$I{r})', fmt=PCT, border=True)
    put(g, f"K{r}", f'=IF($I{r}=0,"",$I{r}-$H{r})', fmt=YEN, border=True)

TOT_ROW = CHG_ROW + 2
put(g, f"A{TOT_ROW}", "合計（有休を除く）", bold=True, fill=TOTALFILL, border=True)
for j in range(6):
    col = get_column_letter(2 + j)
    put(g, f"{col}{TOT_ROW}",
        f'=IF({col}$3="","",SUM({col}${CAT_FIRST}:{col}${CAT_LAST})+{col}{CHG_ROW})',
        bold=True, fill=TOTALFILL, fmt=YEN, border=True)
put(g, f"H{TOT_ROW}", f"=SUM($B{TOT_ROW}:$G{TOT_ROW})", bold=True, fill=TOTALFILL, fmt=YEN, border=True)
put(g, f"I{TOT_ROW}", f"=SUM($I${CAT_FIRST}:$I${CAT_LAST})+$I{CHG_ROW}", bold=True,
    fill=TOTALFILL, fmt=YEN, border=True)
put(g, f"J{TOT_ROW}", f'=IF($I{TOT_ROW}=0,"",$H{TOT_ROW}/$I{TOT_ROW})', bold=True,
    fill=TOTALFILL, fmt=PCT, border=True)
put(g, f"K{TOT_ROW}", f'=IF($I{TOT_ROW}=0,"",$I{TOT_ROW}-$H{TOT_ROW})', bold=True,
    fill=TOTALFILL, fmt=YEN, border=True)

CHK = TOT_ROW + 2
put(g, f"A{CHK}", "■ チェック", bold=True, fill=GRAY)
put(g, f"A{CHK + 1}", "割振シートの金額計")
put(g, f"B{CHK + 1}", "=割振!$M$106", fmt=YEN, border=True)
put(g, f"A{CHK + 2}", "上の表の項目合計")
put(g, f"B{CHK + 2}", f"=SUM($H${CAT_FIRST}:$H${CAT_LAST})", fmt=YEN, border=True)
put(g, f"A{CHK + 3}", "差（集計に載っていない金額）", bold=True)
put(g, f"B{CHK + 3}", f"=$B${CHK + 1}-$B${CHK + 2}", bold=True, fmt=YEN, border=True)
put(g, f"C{CHK + 3}", "← 0 以外なら、割振シートの項目名がこの表の項目名と一致していません",
    size=9, color="808080")

widths(g, {"A": 28, **{get_column_letter(2 + i): 13 for i in range(6)},
           "H": 14, "I": 14, "J": 10, "K": 13})
g.freeze_panes = "B4"

# =====================================================================
# 5. 接客部門
# =====================================================================
h = wb.create_sheet("接客部門")
title(h, "接客部門 週別・部門別割振", "月間予算を、週別動員の構成比と部門構成比で自動的に割り振ります。")
put(h, "A3", "月間予算", bold=True)
put(h, "B3", 22210643, color=BLUE, fill=YELLOW, fmt=YEN, border=True)
put(h, "C3", "← 接客部門の月間予算を入力", size=9, color="808080")

put(h, "A5", "項目", bold=True, fill=HEAD, align="center", border=True)
for i in range(6):
    col = get_column_letter(2 + i)
    put(h, f"{col}5", f"=設定!$B${21 + i}", bold=True, fill=HEAD, align="center", border=True)
put(h, "H5", "月計", bold=True, fill=HEAD, align="center", border=True)

put(h, "A6", "週予算", bold=True, border=True)
for i in range(6):
    col = get_column_letter(2 + i)
    put(h, f"{col}6", f'=IF({col}$5="","",$B$3*動員!$K${4 + i})', bold=True, fmt=YEN, border=True)
put(h, "H6", "=SUM(B6:G6)", bold=True, fmt=YEN, border=True)

DEPTS = ["FLOOR", "CONCESSION", "STORE", "OFFICE"]
put(h, "A8", "■ 予算配分（週予算 × 部門構成比）", bold=True, fill=GRAY)
for k, d in enumerate(DEPTS):
    r = 9 + k
    put(h, f"A{r}", d, border=True)
    for i in range(6):
        col = get_column_letter(2 + i)
        put(h, f"{col}{r}", f'=IF({col}$5="","",{col}$6*設定!$B${29 + k})', fmt=YEN, border=True)
    put(h, f"H{r}", f"=SUM(B{r}:G{r})", fmt=YEN, border=True)

put(h, "A14", "■ 予算外（ミーティング・棚卸・返品など）", bold=True, fill=GRAY)
for k, d in enumerate(DEPTS):
    r = 15 + k
    put(h, f"A{r}", d, border=True)
    for i in range(6):
        col = get_column_letter(2 + i)
        put(h, f"{col}{r}", None, color=BLUE, fill=YELLOW, fmt=YEN, border=True)
    put(h, f"H{r}", f"=SUM(B{r}:G{r})", fmt=YEN, border=True)

put(h, "A20", "■ 合計（予算配分＋予算外）", bold=True, fill=GRAY)
for k, d in enumerate(DEPTS):
    r = 21 + k
    put(h, f"A{r}", d, bold=True, border=True)
    for i in range(6):
        col = get_column_letter(2 + i)
        put(h, f"{col}{r}", f'=IF({col}$5="","",{col}{9 + k}+{col}{15 + k})', fmt=YEN, border=True)
    put(h, f"H{r}", f"=SUM(B{r}:G{r})", fmt=YEN, border=True)
put(h, "A25", "週合計", bold=True, fill=TOTALFILL, border=True)
for i in range(6):
    col = get_column_letter(2 + i)
    put(h, f"{col}25", f'=IF({col}$5="","",SUM({col}21:{col}24))', bold=True,
        fill=TOTALFILL, fmt=YEN, border=True)
put(h, "H25", "=SUM(B25:G25)", bold=True, fill=TOTALFILL, fmt=YEN, border=True)

put(h, "A27", "予算比", bold=True)
put(h, "B27", '=IF($B$3=0,"",$H$25/$B$3)', bold=True, fmt=PCT, border=True)
put(h, "A28", "差額（予算 − 合計）", bold=True)
put(h, "B28", '=IF($B$3=0,"",$B$3-$H$25)', bold=True, fmt=YEN, border=True)
widths(h, {"A": 30, **{get_column_letter(2 + i): 13 for i in range(6)}, "H": 14})
h.freeze_panes = "B6"

# =====================================================================
# 6. 有休・更衣
# =====================================================================
lv = wb.create_sheet("有休・更衣")
title(lv, "有休・更衣時間", "人数を入力すると金額が出ます。有休は予算との差額を各週へ均等に振り分けます。")

put(lv, "A3", "有休単価", bold=True)
put(lv, "B3", 979, color=BLUE, fill=YELLOW, fmt=YEN, border=True)
put(lv, "A4", "時間", bold=True)
put(lv, "B4", 5, color=BLUE, fill=YELLOW, fmt=NUM, border=True)
put(lv, "A5", "予算", bold=True)
put(lv, "B5", 325789, color=BLUE, fill=YELLOW, fmt=YEN, border=True)

for col, label in zip("ABCDEF", ["週", "WEEK", "人数", "金額", "補正後", "1日あたり"]):
    put(lv, f"{col}7", label, bold=True, fill=HEAD, align="center", border=True)
for i in range(6):
    r = 8 + i
    put(lv, f"A{r}", i + 1, align="center", border=True)
    put(lv, f"B{r}", f"=設定!$B${21 + i}", align="center", border=True)
    put(lv, f"C{r}", None, color=BLUE, fill=YELLOW, fmt=NUM, border=True)
    put(lv, f"D{r}", f'=IF($B{r}="","",$B$3*$B$4*$C{r})', fmt=YEN, border=True)
    put(lv, f"E{r}", f'=IF($B{r}="","",$D{r}+$B$16)', fmt=YEN, border=True)
    put(lv, f"F{r}", f'=IF(OR($B{r}="",$C{r}=0),"",$E{r}/$C{r})', fmt=YEN, border=True)
put(lv, "A14", "合計", bold=True, fill=TOTALFILL, border=True)
for col in "BC":
    put(lv, f"{col}14", None, fill=TOTALFILL, border=True)
put(lv, "C14", "=SUM(C8:C13)", bold=True, fill=TOTALFILL, fmt=NUM, border=True)
put(lv, "D14", "=SUM(D8:D13)", bold=True, fill=TOTALFILL, fmt=YEN, border=True)
put(lv, "E14", "=SUM(E8:E13)", bold=True, fill=TOTALFILL, fmt=YEN, border=True)
put(lv, "A15", "予算 − 単純計算額")
put(lv, "B15", "=$B$5-$D$14", fmt=YEN, border=True)
put(lv, "A16", "1週あたりの調整額")
put(lv, "B16", '=IF(設定!$B$17=0,0,$B$15/設定!$B$17)', fmt=YEN, border=True)
put(lv, "C16", "← 各週へ均等に振り分け、補正後の合計が予算と一致します", size=9, color="808080")

put(lv, "A20", "■ 更衣時間", bold=True, fill=GRAY)
put(lv, "A22", "単価", bold=True)
put(lv, "B22", 1132, color=BLUE, fill=YELLOW, fmt=YEN, border=True)
put(lv, "A23", "時間（分）", bold=True)
put(lv, "B23", 9, color=BLUE, fill=YELLOW, fmt=NUM, border=True)
put(lv, "A24", "予算基準額", bold=True)
put(lv, "B24", 81178, color=BLUE, fill=YELLOW, fmt=YEN, border=True)
put(lv, "C24", "← 予算 ＝ 予算基準額 × 接客部門の予算比", size=9, color="808080")
for col, label in zip("ABCD", ["週", "WEEK", "人数", "金額"]):
    put(lv, f"{col}26", label, bold=True, fill=HEAD, align="center", border=True)
for i in range(6):
    r = 27 + i
    put(lv, f"A{r}", i + 1, align="center", border=True)
    put(lv, f"B{r}", f"=設定!$B${21 + i}", align="center", border=True)
    put(lv, f"C{r}", None, color=BLUE, fill=YELLOW, fmt=NUM, border=True)
    put(lv, f"D{r}", f'=IF($B{r}="","",$B$22*($B$23/60)*$C{r})', fmt=YEN, border=True)
put(lv, "A33", "合計", bold=True, fill=TOTALFILL, border=True)
put(lv, "C33", "=SUM(C27:C32)", bold=True, fill=TOTALFILL, fmt=NUM, border=True)
put(lv, "D33", "=SUM(D27:D32)", bold=True, fill=TOTALFILL, fmt=YEN, border=True)
put(lv, "A34", "予算")
put(lv, "B34", '=$B$24*IF(接客部門!$B$3=0,0,接客部門!$H$25/接客部門!$B$3)', fmt=YEN, border=True)
put(lv, "A35", "差（合計 − 予算）")
put(lv, "B35", "=$D$33-$B$34", fmt=YEN, border=True)
widths(lv, {"A": 20, "B": 14, "C": 10, "D": 14, "E": 14, "F": 14})

for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

wb.calculation.fullCalcOnLoad = True  # 開いたときに必ず再計算させる

out = "/home/user/test/excel/月次予算割振.xlsx"
wb.save(out)
print("saved", out)
