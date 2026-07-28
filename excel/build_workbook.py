"""
売店発注ツール（全劇場対応版）Excelワークブックを生成する。

    python3 excel/build_workbook.py [出力パス] [初期マスタ用の在庫CSV]

設計方針
--------
* 1ファイルで全劇場分の在庫CSVを保持し、「設定」シートで選んだ劇場について
  商品別の発注計算を行う。劇場横断の状況は「全劇場サマリ」で見る。
* 劇場ごとの定数は 規模区分（大/中/小）の標準値を基本とし、
  劇場マスタの個別欄に値が入っていればそちらを優先する。
* PI値（来場者100人あたり消費数）は「規模区分_ドリンク提供方式」の
  6プロファイル別に商品マスタで持つ。1杯売り店とドリンクバー店で
  カップ構成が違っても同じマスタで扱える。
* 取扱商品は在庫CSVの実データから導出する。当日CSVと前回CSVの和集合を取るため、
  期中に入れ替わった商品（新商品・終売品）も取りこぼさない。

数式はすべて INDEX/MATCH・SUMIFS など旧来関数で記述している
（XLOOKUP等はLibreOfficeでの検算が通らないため）。
"""

import csv
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------------------
# 共通定義
# ---------------------------------------------------------------------------

FONT = 'Meiryo UI'

# 配色は題材から取っている。客席の暗がり、スクリーンに落ちる光、
# ポップコーンのバター色。バター色は入力欄など「人が触るところ」にだけ使い、
# 状態を表す色（余裕／要発注／今日が期限）とは層を分けている。
C_INK = 'FF1B1D24'
C_TEXT2 = 'FF4E5163'
C_TEXT3 = 'FF7A7D90'
C_PAPER = 'FFFAF7F0'
C_LINE = 'FFE2DCCE'
C_BUTTER = 'FFE8B44A'
C_BUTTER_SOFT = 'FFF9EAC6'
C_OK = 'FF00795A'
C_WARN = 'FFC9761A'
C_CRIT = 'FFA32A3C'
C_OK_SOFT = 'FFDCEBE5'
C_WARN_SOFT = 'FFF7E7D2'
C_CRIT_SOFT = 'FFF3DDE1'

F_BASE = Font(name=FONT, size=11, color=C_INK)
F_BOLD = Font(name=FONT, size=11, bold=True, color=C_INK)
F_TITLE = Font(name=FONT, size=15, bold=True, color=C_INK)
F_NOTE = Font(name=FONT, size=9, color=C_TEXT3)
F_HEAD = Font(name=FONT, size=10, bold=True, color=C_TEXT2)
F_INPUT = Font(name=FONT, size=11, bold=True, color='FF7A5200')
F_HERO = Font(name=FONT, size=36, bold=True, color=C_CRIT)
F_HERO_UNIT = Font(name=FONT, size=13, bold=True, color=C_INK)
F_DAY = Font(name=FONT, size=8, color=C_TEXT3)

FILL_HEAD = PatternFill('solid', fgColor='FFF1ECE1')
FILL_INPUT = PatternFill('solid', fgColor='FFFDF3D8')
FILL_AUTO = PatternFill('solid', fgColor='FFFBFAF7')
FILL_ACCENT = PatternFill('solid', fgColor=C_PAPER)
FILL_ALERT = PatternFill('solid', fgColor=C_CRIT_SOFT)
FILL_TRACK = PatternFill('solid', fgColor='FFF4F1EA')

THIN = Side(style='thin', color=C_LINE)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FMT_INT = '#,##0_);[Red](#,##0)'
FMT_DEC = '#,##0.0;[Red]\\-#,##0.0'
FMT_DAYS = '#,##0.0"日";[Red]\\-#,##0.0"日"'
FMT_YEN = '[$¥-411]#,##0;[Red][$¥-411]\\-#,##0'
FMT_DATE = 'yyyy/m/d\\(aaa\\)'

# 在庫CSVの列（システム出力そのままの並び）
CSV_HEADERS = ['対象日付', '劇場コード', '劇場名', '支払先コード', '支払先名',
               '大分類コード', '小分類コード', '商品分類名', '商品コード', '商品名',
               '入数', '在庫数ケース', '在庫数バラ', '総数', '規定数', '資産廃棄',
               '税抜単価', '仕入税区分', '仕入税率']
CSV_FIRST = 2          # データ開始行
CSV_LAST = 10001       # データ終了行（75劇場×約130商品を想定）
CSV_HELPER = 'U'       # 連番ヘルパー列

SIZES = ['大規模', '中規模', '小規模']
DRINK_STYLES = ['1杯売り', 'ドリンクバー']
# PIプロファイル = 規模区分_提供方式（商品マスタのPI値列と対応する）
PROFILES = [f'{s}_{d}' for d in DRINK_STYLES for s in SIZES]

BASIS_OPTIONS = ['実績消費', 'PI値予測', '規定数', '最大値']
ORDER_STATUS = ['発注済', '入荷済み', '取消']

MASTER_FIRST, MASTER_LAST = 6, 1005      # 商品マスタ
THEATER_FIRST, THEATER_LAST = 6, 105     # 劇場マスタ
CALC_FIRST, CALC_LAST = 6, 405           # 発注計算
PO_FIRST, PO_LAST = 6, 1005              # 発注管理

SPAN_DAYS = 14                           # タイムテーブルに表示する日数
TT_FIRST = 10                            # タイムテーブルのデータ開始行
TT_LAST = TT_FIRST + (CALC_LAST - CALC_FIRST)
TT_DAY_COL = 12                          # L列から14日分のグリッド

S_TT = 'タイムテーブル'
S_SET = '設定'
S_THEATER = '劇場マスタ'
S_SIZE = '規模マスタ'
S_SEASON = '季節係数マスタ'
S_ITEM = '商品マスタ'
S_CUR = '在庫CSV_当日'
S_PRV = '在庫CSV_前回'
S_PO = '発注管理'
S_CALC = '発注計算'
S_DASH = 'ダッシュボード'
S_ALL = '全劇場サマリ'


def col(sheet, letter, first=None, last=None):
    """絶対参照の列範囲を返す。"""
    first = CSV_FIRST if first is None else first
    last = CSV_LAST if last is None else last
    return f'{sheet}!${letter}${first}:${letter}${last}'


def csv_col(sheet, letter):
    return col(sheet, letter, CSV_FIRST, CSV_LAST)


def item_col(letter):
    return col(S_ITEM, letter, MASTER_FIRST, MASTER_LAST)


def theater_col(letter):
    return col(S_THEATER, letter, THEATER_FIRST, THEATER_LAST)


def po_col(letter):
    return col(S_PO, letter, PO_FIRST, PO_LAST)


def style_row(ws, row, cols, font=F_BASE, fill=None, fmt=None, align=None, border=True):
    for c in cols:
        cell = ws[f'{c}{row}']
        cell.font = font
        if fill:
            cell.fill = fill
        if fmt:
            cell.number_format = fmt
        if align:
            cell.alignment = Alignment(horizontal=align, vertical='center')
        if border:
            cell.border = BORDER


def put_headers(ws, row, headers, kinds=None, start_col=2):
    """見出し行（と、その上の 入力/自動/選択 の区分行）を書く。"""
    for i, name in enumerate(headers):
        letter = get_column_letter(start_col + i)
        if kinds:
            k = ws[f'{letter}{row - 1}']
            k.value = kinds[i]
            k.font = F_NOTE
            k.alignment = Alignment(horizontal='center')
        cell = ws[f'{letter}{row}']
        cell.value = name
        cell.font = F_HEAD
        cell.fill = FILL_HEAD
        cell.border = BORDER
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


def sheet_title(ws, text, note=None):
    ws['B2'] = text
    ws['B2'].font = F_TITLE
    if note:
        ws['B3'] = note
        ws['B3'].font = F_NOTE
    ws.sheet_view.showGridLines = False


def add_validation(ws, formula, cells):
    dv = DataValidation(type='list', formula1=formula, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(cells)
    return dv


# ---------------------------------------------------------------------------
# 各シート
# ---------------------------------------------------------------------------

def build_intro(wb):
    ws = wb.create_sheet('はじめに')
    sheet_title(ws, '売店発注ツール（全劇場対応版）')
    ws.column_dimensions['B'].width = 4
    ws.column_dimensions['C'].width = 100

    lines = [
        ('h', 'このツールでできること'),
        ('t', '在庫システムから出力した在庫一覧CSVを貼り付けるだけで、劇場ごとの理論値在庫から'),
        ('t', '発注が必要な商品と推奨発注数を自動算出します。'),
        ('t', '全劇場分のCSVをまとめて貼り付けても、設定シートで選んだ劇場について計算します。'),
        ('', ''),
        ('h', 'ご利用の流れ'),
        ('s', '① 劇場マスタ・規模マスタを登録する'),
        ('t', '　 劇場コード・劇場名・規模区分（大/中/小）・ドリンク提供方式を登録します。'),
        ('t', '　 安全在庫日数や発注サイクルは規模区分ごとの標準値が適用されます。'),
        ('t', '　 特定の劇場だけ変えたい場合は、劇場マスタの個別欄に値を入れると優先されます。'),
        ('s', '② 商品マスタを整える'),
        ('t', '　 リードタイム・最低ロット・PI値を登録します。'),
        ('t', '　 PI値は「規模区分_ドリンク提供方式」の6プロファイル別に持てます。'),
        ('t', '　 1杯売り店とドリンクバー店でカップ構成が違っても同じマスタで管理できます。'),
        ('s', '③ 在庫CSVを貼り付ける'),
        ('t', '　 「在庫CSV_当日」に最新の在庫一覧を、「在庫CSV_前回」に2ヶ月前などの在庫一覧を'),
        ('t', '　 A2セルから貼り付けます（見出し行は貼らずに、データ行だけ貼り付けてください）。'),
        ('t', '　 2時点の在庫を持つことで、期中に入れ替わった商品も取りこぼしません。'),
        ('s', '④ 発注管理シートに発注・入荷を記録する'),
        ('t', '　 実発注数と入荷日を記録すると、期間中の納品数として消費量の計算に使われ、'),
        ('t', '　 未入荷分は「発注残」として推奨発注数から自動で差し引かれます（二重発注の防止）。'),
        ('s', '⑤ 設定シートで対象劇場を選び、発注計算シートを確認する'),
        ('t', '　 推奨発注数（ケース）を確認し、発注管理シートに記録します。'),
        ('', ''),
        ('h', '発注数の考え方'),
        ('t', '　 理論値在庫　　＝ 在庫CSVの「総数」（＝在庫数ケース×入数＋在庫数バラ）'),
        ('t', '　 期間消費数　　＝ 前回在庫 ＋ 期間納品数 － 当日在庫'),
        ('t', '　 実績消費/日　 ＝ 期間消費数 ÷ 期間日数'),
        ('t', '　 PI予測消費/日 ＝ 想定来場者数/日 × 季節係数 × PI値 ÷ 100'),
        ('t', '　 安全在庫数量　＝ 採用消費/日 × 安全在庫日数'),
        ('t', '　 発注点　　　　＝ 採用消費/日 ×（リードタイム＋発注サイクル）＋ 安全在庫数量'),
        ('t', '　 推奨発注数　　＝ 発注点 － 当日在庫 － 発注残 → 入数で割り、最低ロット単位に切り上げ'),
        ('', ''),
        ('t', '「採用消費/日」に実績・PI予測・規定数のどれを使うかは、設定シートで切り替えられます。'),
        ('t', '通常週は実績消費、大作公開週はPI値予測、といった使い分けを想定しています。'),
        ('', ''),
        ('h', '季節変動（繁忙期・閑散期）の調整'),
        ('t', '　 季節係数マスタで月別の係数を設定できます（例：12月＝120%、6月＝85%）。'),
        ('t', '　 特定期間だけ変えたい場合は「特別期間」に開始日・終了日・係数を登録すると、'),
        ('t', '　 その期間は月別係数より優先されます（大型作品の公開週など）。'),
        ('', ''),
        ('h', 'ご利用時のポイント'),
        ('t', '　・黄色のセルのみ入力してください。それ以外は数式が入っています。'),
        ('t', '　・在庫CSVの劇場コードは「0761」のような先頭ゼロ付きの文字列です。'),
        ('t', '　　貼り付け後、設定シートの「CSV該当行数」が0のままなら、コードの型がずれています。'),
        ('t', '　・理論値在庫がマイナスの商品は、納品の未計上など元データ側のずれを示しています。'),
        ('t', '　　全劇場サマリの「マイナス在庫件数」で確認できます。'),
    ]
    r = 5
    for kind, text in lines:
        cell = ws[f'C{r}']
        cell.value = text
        if kind == 'h':
            cell.font = F_BOLD
            cell.fill = FILL_ACCENT
        elif kind == 's':
            cell.font = F_BOLD
        else:
            cell.font = F_BASE
        r += 1
    return ws


def build_size_master(wb):
    ws = wb.create_sheet(S_SIZE)
    sheet_title(ws, '規模マスタ', '規模区分ごとの標準定数です。劇場マスタで個別に上書きできます。')
    headers = ['規模区分', '想定来場者数/日', '安全在庫日数', '発注サイクル日数', '備考']
    put_headers(ws, 5, headers, ['固定', '入力', '入力', '入力', '入力'])
    rows = [
        ('大規模', 3000, 4, 7, 'スクリーン数が多く、来場者数の多い劇場'),
        ('中規模', 1500, 3, 7, '標準的な規模の劇場'),
        ('小規模', 700, 3, 14, '来場者数が少なく、発注頻度を落とす劇場'),
    ]
    for i, row in enumerate(rows):
        r = 6 + i
        for j, v in enumerate(row):
            ws.cell(r, 2 + j, v)
        style_row(ws, r, 'BCDEF', fill=FILL_INPUT)
        ws[f'B{r}'].fill = FILL_AUTO
        ws[f'C{r}'].number_format = FMT_INT
    for c, w in zip('BCDEF', [12, 18, 14, 18, 44]):
        ws.column_dimensions[c].width = w
    return ws


def build_season_master(wb):
    ws = wb.create_sheet(S_SEASON)
    sheet_title(ws, '季節係数マスタ',
                '当日基準日が特別期間に含まれる場合はその係数を、含まれない場合は月別係数を使います。')
    put_headers(ws, 5, ['月', '係数'], ['固定', '入力'])
    for i in range(12):
        r = 6 + i
        ws.cell(r, 2, i + 1)
        ws.cell(r, 3, 1.0)
        style_row(ws, r, 'BC', fill=FILL_INPUT)
        ws[f'B{r}'].fill = FILL_AUTO
        ws[f'C{r}'].number_format = '0%'

    put_headers(ws, 5, ['特別期間名', '開始日', '終了日', '係数'],
                ['入力', '入力', '入力', '入力'], start_col=5)
    samples = [('年末年始', None, None, 1.3), ('大型作品公開週', None, None, 1.5),
               ('閑散期', None, None, 0.85)]
    for i in range(20):
        r = 6 + i
        if i < len(samples):
            ws.cell(r, 5, samples[i][0])
            ws.cell(r, 8, samples[i][3])
        style_row(ws, r, 'EFGH', fill=FILL_INPUT)
        ws[f'F{r}'].number_format = FMT_DATE
        ws[f'G{r}'].number_format = FMT_DATE
        ws[f'H{r}'].number_format = '0%'
    ws['E27'] = '※ 開始日・終了日が空欄の行は無視されます。期間が重なる場合は係数が最大の行が使われます。'
    ws['E27'].font = F_NOTE
    for c, w in zip('BCDEFGH', [8, 10, 3, 22, 14, 14, 10]):
        ws.column_dimensions[c].width = w
    return ws


def build_theater_master(wb):
    ws = wb.create_sheet(S_THEATER)
    sheet_title(ws, '劇場マスタ',
                'PIプロファイルは規模区分と提供方式から自動で決まり、商品マスタのPI値列と対応します。')
    headers = ['劇場コード', '劇場名', '規模区分', 'ドリンク提供方式', 'PIプロファイル',
               '想定来場者数/日(個別)', '安全在庫日数(個別)', '発注サイクル(個別)', '備考']
    kinds = ['入力', '入力', '選択', '選択', '自動', '任意入力', '任意入力', '任意入力', '入力']
    put_headers(ws, 5, headers, kinds)

    for r in range(THEATER_FIRST, THEATER_LAST + 1):
        ws[f'F{r}'] = f'=IF(OR($D{r}="",$E{r}=""),"",$D{r}&"_"&$E{r})'
        style_row(ws, r, 'BCDEGHIJ', fill=FILL_INPUT)
        style_row(ws, r, 'F', fill=FILL_AUTO)
        ws[f'B{r}'].number_format = '@'
        ws[f'G{r}'].number_format = FMT_INT
    ws[f'B{THEATER_FIRST}'] = '0761'
    ws[f'C{THEATER_FIRST}'] = '新宿'
    ws[f'D{THEATER_FIRST}'] = '大規模'
    ws[f'E{THEATER_FIRST}'] = '1杯売り'
    ws[f'J{THEATER_FIRST}'] = 'サンプル。実際の規模区分・提供方式に置き換えてください。'

    add_validation(ws, '"' + ','.join(SIZES) + '"', f'D{THEATER_FIRST}:D{THEATER_LAST}')
    add_validation(ws, '"' + ','.join(DRINK_STYLES) + '"', f'E{THEATER_FIRST}:E{THEATER_LAST}')
    for c, w in zip('BCDEFGHIJ', [12, 22, 12, 18, 20, 20, 18, 18, 40]):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = 'B6'
    return ws


def build_item_master(wb, items):
    ws = wb.create_sheet(S_ITEM)
    sheet_title(ws, '商品マスタ',
                'PI値は来場者100人あたりの消費数です。取り扱いのないプロファイルは空欄のままで構いません。')
    headers = ['商品コード', '商品名', '商品分類名', '支払先名', '入数', '税抜単価',
               'L/T日数', '最低ロット', '発注単位'] + [f'PI値 {p}' for p in PROFILES] + ['備考']
    kinds = ['入力', '入力', '入力', '入力', '入力', '入力', '入力', '入力', '入力'] + \
            ['入力'] * len(PROFILES) + ['入力']
    put_headers(ws, 5, headers, kinds)

    last_col_idx = 2 + len(headers) - 1
    letters = [get_column_letter(i) for i in range(2, last_col_idx + 1)]
    for r in range(MASTER_FIRST, MASTER_LAST + 1):
        style_row(ws, r, letters, fill=FILL_INPUT)
        ws[f'B{r}'].number_format = '@'
        ws[f'F{r}'].number_format = FMT_INT
        ws[f'G{r}'].number_format = FMT_YEN
        for i in range(len(PROFILES)):
            ws.cell(r, 11 + i).number_format = FMT_DEC

    for i, it in enumerate(items):
        r = MASTER_FIRST + i
        if r > MASTER_LAST:
            break
        ws[f'B{r}'] = it['商品コード']
        ws[f'C{r}'] = it['商品名']
        ws[f'D{r}'] = it['商品分類名']
        ws[f'E{r}'] = it['支払先名']
        ws[f'F{r}'] = int(it['入数'])
        ws[f'G{r}'] = float(it['税抜単価'])
        ws[f'H{r}'] = 3
        ws[f'I{r}'] = 1
        ws[f'J{r}'] = 'ケース'

    widths = [16, 30, 16, 22, 8, 12, 9, 11, 10] + [14] * len(PROFILES) + [24]
    for c, w in zip(letters, widths):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = 'D6'
    return ws


def build_csv_sheet(wb, name, other, label):
    ws = wb.create_sheet(name)
    sheet_title(ws, label, '在庫一覧CSVのデータ行を A2 セルから貼り付けてください（見出し行は不要です）。')
    ws['B3'] = ''
    for i, h in enumerate(CSV_HEADERS):
        cell = ws.cell(1, 1 + i, h)
        cell.font = F_HEAD
        cell.fill = FILL_HEAD
        cell.border = BORDER
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
    helper = ws.cell(1, 21, '連番(自動)')
    helper.font = F_HEAD
    helper.fill = FILL_ACCENT
    helper.border = BORDER

    for r in range(CSV_FIRST, CSV_LAST + 1):
        if other is None:
            # 当日CSV: 対象劇場の行に通し番号を振る
            ws[f'{CSV_HELPER}{r}'] = f'=IF($B{r}={S_SET}!$C$4,N({CSV_HELPER}{r-1})+1,N({CSV_HELPER}{r-1}))'
        else:
            # 前回CSV: 対象劇場の行のうち、当日CSVに存在しない商品だけに番号を振る
            ws[f'{CSV_HELPER}{r}'] = (
                f'=IF($B{r}<>{S_SET}!$C$4,N({CSV_HELPER}{r-1}),'
                f'IF(COUNTIFS({csv_col(other, "B")},$B{r},{csv_col(other, "I")},$I{r})>0,'
                f'N({CSV_HELPER}{r-1}),N({CSV_HELPER}{r-1})+1))')
        ws[f'{CSV_HELPER}{r}'].font = F_NOTE
        ws[f'B{r}'].number_format = '@'
        ws[f'I{r}'].number_format = '@'
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 11
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['E'].width = 26
    ws.column_dimensions['H'].width = 16
    ws.column_dimensions['I'].width = 16
    ws.column_dimensions['J'].width = 30
    ws.column_dimensions['U'].width = 11
    ws.freeze_panes = 'A2'
    return ws


def build_settings(wb):
    ws = wb.create_sheet(S_SET)
    sheet_title(ws, '設定', '黄色のセルを入力・選択してください。')
    ws.column_dimensions['B'].width = 28
    ws.column_dimensions['C'].width = 26
    ws.column_dimensions['D'].width = 62

    cur_date = (f'IFERROR(DATE(LEFT(TEXT(INDEX({csv_col(S_CUR, "A")},'
                f'MATCH($C$4,{csv_col(S_CUR, "B")},0)),"00000000"),4),'
                f'MID(TEXT(INDEX({csv_col(S_CUR, "A")},MATCH($C$4,{csv_col(S_CUR, "B")},0)),"00000000"),5,2),'
                f'RIGHT(TEXT(INDEX({csv_col(S_CUR, "A")},MATCH($C$4,{csv_col(S_CUR, "B")},0)),"00000000"),2)),"")')
    prv_date = cur_date.replace(S_CUR, S_PRV)

    def const(individual, size_col):
        """劇場個別の値があればそれを、無ければ規模マスタの標準値を返す。"""
        return (f'=IF($C$4="","",IF(IFERROR(INDEX({theater_col(individual)},'
                f'MATCH($C$4,{theater_col("B")},0)),"")<>"",'
                f'INDEX({theater_col(individual)},MATCH($C$4,{theater_col("B")},0)),'
                f'IFERROR(INDEX({col(S_SIZE, size_col, 6, 8)},MATCH($C$10,{col(S_SIZE, "B", 6, 8)},0)),"")))')

    special = (f'SUMPRODUCT(MAX(({col(S_SEASON, "F", 6, 25)}<=$C$6)*'
               f'({col(S_SEASON, "G", 6, 25)}>=$C$6)*{col(S_SEASON, "H", 6, 25)}))')

    rows = [
        ('対象劇場コード', None, '発注計算・ダッシュボードの対象になる劇場です。劇場マスタから選択します。', FILL_INPUT, '@'),
        ('劇場名', f'=IFERROR(INDEX({theater_col("C")},MATCH($C$4,{theater_col("B")},0)),"")',
         '劇場マスタから自動取得します。', FILL_AUTO, None),
        ('当日基準日', f'={cur_date}', '在庫CSV_当日の対象日付から自動取得します。', FILL_AUTO, FMT_DATE),
        ('前回基準日', f'={prv_date}', '在庫CSV_前回の対象日付から自動取得します。', FILL_AUTO, FMT_DATE),
        ('期間日数', '=IF(OR($C$6="",$C$7=""),"",$C$6-$C$7)', '実績消費/日の算出に使う日数です。', FILL_AUTO, FMT_INT),
        ('算出基準', None,
         '実績消費＝在庫2時点の差、PI値予測＝来場者予測、規定数＝CSVの規定数、最大値＝実績とPI予測の大きい方。',
         FILL_INPUT, None),
        ('規模区分', f'=IFERROR(INDEX({theater_col("D")},MATCH($C$4,{theater_col("B")},0)),"")',
         '劇場マスタから自動取得します。', FILL_AUTO, None),
        ('PIプロファイル', f'=IFERROR(INDEX({theater_col("F")},MATCH($C$4,{theater_col("B")},0)),"")',
         '商品マスタのどのPI値列を使うかを決めます。', FILL_AUTO, None),
        ('安全在庫日数', const('H', 'D'), '劇場個別の設定があればそちらが優先されます。', FILL_AUTO, FMT_DEC),
        ('発注サイクル日数', const('I', 'E'), '次回発注までの日数。発注点の算出に使います。', FILL_AUTO, FMT_INT),
        ('想定来場者数/日', const('G', 'C'), 'PI値予測の基礎になる来場者数です。', FILL_AUTO, FMT_INT),
        ('季節係数', f'=IF($C$6="",1,IF({special}=0,'
                     f'IFERROR(INDEX({col(S_SEASON, "C", 6, 17)},MONTH($C$6)),1),{special}))',
         '当日基準日が特別期間に該当すればその係数、なければ月別係数です。', FILL_AUTO, '0%'),
        ('発注担当者', None, '発注書・発注管理に記録する担当者名です。', FILL_INPUT, None),
        ('当日CSV該当行数', f'=COUNTIF({csv_col(S_CUR, "B")},$C$4)',
         '0のままなら劇場コードが一致していません（先頭ゼロが落ちていないか確認）。', FILL_AUTO, FMT_INT),
        ('前回CSV該当行数', f'=COUNTIF({csv_col(S_PRV, "B")},$C$4)',
         '前回CSVを貼っていない場合は0です（実績消費は算出できません）。', FILL_AUTO, FMT_INT),
        ('取扱商品数', f'=MAX({csv_col(S_CUR, CSV_HELPER)})+MAX({csv_col(S_PRV, CSV_HELPER)})',
         '当日CSVと前回CSVの和集合です。発注計算シートの行数と一致します。', FILL_AUTO, FMT_INT),
    ]

    for i, (label, formula, note, fill, fmt) in enumerate(rows):
        r = 4 + i
        ws[f'B{r}'] = label
        ws[f'B{r}'].font = F_BOLD
        ws[f'B{r}'].fill = FILL_HEAD
        ws[f'B{r}'].border = BORDER
        cell = ws[f'C{r}']
        if formula:
            cell.value = formula
        cell.fill = fill
        cell.border = BORDER
        cell.font = F_INPUT if fill is FILL_INPUT else F_BASE
        if fmt:
            cell.number_format = fmt
        ws[f'D{r}'] = note
        ws[f'D{r}'].font = F_NOTE

    ws['C4'] = '0761'
    ws['C9'] = '実績消費'
    add_validation(ws, f'={theater_col("B")}', 'C4')
    add_validation(ws, '"' + ','.join(BASIS_OPTIONS) + '"', 'C9')

    ws['B22'] = '※ 在庫CSVを貼り替えたら、この画面の該当行数と基準日が想定どおりか必ず確認してください。'
    ws['B22'].font = F_NOTE
    return ws


def build_calc(wb):
    ws = wb.create_sheet(S_CALC)
    sheet_title(ws, '発注計算',
                '設定シートで選んだ劇場の商品別計算です。行は在庫CSV（当日＋前回）から自動生成されます。')
    headers = ['No', '商品コード', '商品名', '商品分類名', '支払先名', '入数', '税抜単価',
               'L/T日数', '最低ロット', '当日理論在庫', '前回理論在庫', '期間納品数', '期間消費数',
               '実績消費/日', 'PI値', 'PI予測消費/日', '規定数', '採用消費/日', '在庫日数',
               '安全在庫数量', '発注点', '発注残', '推奨発注数(バラ)', '推奨発注数(ケース)',
               '発注要否', '在庫金額', '状態', '推奨発注金額']
    kinds = ['自動'] * len(headers)
    put_headers(ws, 5, headers, kinds)

    n_cur = f'MAX({csv_col(S_CUR, CSV_HELPER)})'
    letters = [get_column_letter(i) for i in range(2, 2 + len(headers))]

    for r in range(CALC_FIRST, CALC_LAST + 1):
        no = r - CALC_FIRST + 1
        ws[f'B{r}'] = no
        # 商品コード: 当日CSVの対象劇場行 → 足りない分を前回CSVの未登場商品で埋める
        ws[f'C{r}'] = (
            f'=IFERROR(IF($B{r}<={n_cur},'
            f'INDEX({csv_col(S_CUR, "I")},MATCH($B{r},{csv_col(S_CUR, CSV_HELPER)},0)),'
            f'INDEX({csv_col(S_PRV, "I")},MATCH($B{r}-{n_cur},{csv_col(S_PRV, CSV_HELPER)},0))),"")')

        def from_master_or_csv(master_col, csv_letter):
            return (f'=IF($C{r}="","",IFERROR(INDEX({item_col(master_col)},'
                    f'MATCH($C{r},{item_col("B")},0)),'
                    f'IFERROR(INDEX({csv_col(S_CUR, csv_letter)},'
                    f'MATCH($C{r},{csv_col(S_CUR, "I")},0)),'
                    f'IFERROR(INDEX({csv_col(S_PRV, csv_letter)},'
                    f'MATCH($C{r},{csv_col(S_PRV, "I")},0)),""))))')

        ws[f'D{r}'] = from_master_or_csv('C', 'J')     # 商品名
        ws[f'E{r}'] = from_master_or_csv('D', 'H')     # 商品分類名
        ws[f'F{r}'] = from_master_or_csv('E', 'E')     # 支払先名
        ws[f'G{r}'] = from_master_or_csv('F', 'K')     # 入数
        ws[f'H{r}'] = from_master_or_csv('G', 'Q')     # 税抜単価
        ws[f'I{r}'] = (f'=IF($C{r}="","",IFERROR(INDEX({item_col("H")},'
                       f'MATCH($C{r},{item_col("B")},0)),3))')
        ws[f'J{r}'] = (f'=IF($C{r}="","",MAX(1,IFERROR(INDEX({item_col("I")},'
                       f'MATCH($C{r},{item_col("B")},0)),1)))')

        ws[f'K{r}'] = (f'=IF($C{r}="","",SUMIFS({csv_col(S_CUR, "N")},'
                       f'{csv_col(S_CUR, "B")},{S_SET}!$C$4,{csv_col(S_CUR, "I")},$C{r}))')
        ws[f'L{r}'] = (f'=IF($C{r}="","",SUMIFS({csv_col(S_PRV, "N")},'
                       f'{csv_col(S_PRV, "B")},{S_SET}!$C$4,{csv_col(S_PRV, "I")},$C{r}))')
        # 期間納品数: 発注管理の入荷実績（前回基準日〜当日基準日）
        ws[f'M{r}'] = (f'=IF($C{r}="","",IF(OR({S_SET}!$C$6="",{S_SET}!$C$7=""),0,'
                       f'SUMIFS({po_col("S")},{po_col("C")},{S_SET}!$C$4,{po_col("F")},$C{r},'
                       f'{po_col("R")},">="&{S_SET}!$C$7,{po_col("R")},"<="&{S_SET}!$C$6)))')
        ws[f'N{r}'] = (f'=IF(OR($C{r}="",{S_SET}!$C$15=0),"",$L{r}+$M{r}-$K{r})')
        ws[f'O{r}'] = (f'=IF(OR($N{r}="",{S_SET}!$C$8="",{S_SET}!$C$8<=0),"",'
                       f'MAX(0,$N{r})/{S_SET}!$C$8)')
        # PI値: 商品を縦、PIプロファイルを横に検索する。
        # 列見出しは「PI値 <プロファイル名>」なので、検索値も同じ形に組み立てる。
        ws[f'P{r}'] = (f'=IF($C{r}="","",IFERROR(INDEX({col(S_ITEM, "K", MASTER_FIRST, MASTER_LAST)}:'
                       f'${get_column_letter(10 + len(PROFILES))}${MASTER_LAST},'
                       f'MATCH($C{r},{item_col("B")},0),'
                       f'MATCH("PI値 "&{S_SET}!$C$11,'
                       f'{S_ITEM}!$K$5:${get_column_letter(10 + len(PROFILES))}$5,0)),""))')
        ws[f'Q{r}'] = (f'=IF(OR($C{r}="",$P{r}="",$P{r}=0),"",'
                       f'{S_SET}!$C$14*{S_SET}!$C$15*$P{r}/100)')
        ws[f'R{r}'] = (f'=IF($C{r}="","",SUMIFS({csv_col(S_CUR, "O")},'
                       f'{csv_col(S_CUR, "B")},{S_SET}!$C$4,{csv_col(S_CUR, "I")},$C{r}))')
        ws[f'S{r}'] = (f'=IF($C{r}="","",IF({S_SET}!$C$9="実績消費",N($O{r}),'
                       f'IF({S_SET}!$C$9="PI値予測",N($Q{r}),MAX(N($O{r}),N($Q{r})))))')
        ws[f'T{r}'] = f'=IF(OR($S{r}="",$S{r}<=0),"",$K{r}/$S{r})'
        ws[f'U{r}'] = f'=IF($S{r}="","",ROUNDUP(N($S{r})*{S_SET}!$C$12,0))'
        ws[f'V{r}'] = (f'=IF($C{r}="","",IF({S_SET}!$C$9="規定数",N($R{r}),'
                       f'ROUNDUP(N($S{r})*(N($I{r})+N({S_SET}!$C$13))+N($U{r}),0)))')
        ws[f'W{r}'] = (f'=IF($C{r}="","",SUMIFS({po_col("L")},{po_col("C")},{S_SET}!$C$4,'
                       f'{po_col("F")},$C{r},{po_col("Q")},"発注済"))')
        ws[f'X{r}'] = f'=IF($C{r}="","",MAX(0,N($V{r})-N($K{r})-N($W{r})))'
        ws[f'Y{r}'] = (f'=IF($C{r}="","",IF(N($X{r})<=0,0,'
                       f'CEILING($X{r}/MAX(1,N($G{r})),MAX(1,N($J{r})))))')
        ws[f'Z{r}'] = f'=IF($C{r}="","",IF(N($Y{r})>0,"要発注","正常"))'
        ws[f'AA{r}'] = f'=IF($C{r}="","",N($K{r})*N($H{r}))'
        ws[f'AB{r}'] = (
            f'=IF($C{r}="","",_xlfn.TEXTJOIN(" / ",TRUE,'
            f'IF(COUNTIF({item_col("B")},$C{r})=0,"マスタ未登録",""),'
            f'IF(COUNTIFS({csv_col(S_PRV, "B")},{S_SET}!$C$4,{csv_col(S_PRV, "I")},$C{r})=0,"新商品",""),'
            f'IF(COUNTIFS({csv_col(S_CUR, "B")},{S_SET}!$C$4,{csv_col(S_CUR, "I")},$C{r})=0,"終売候補",""),'
            f'IF(N($K{r})<0,"在庫マイナス",""),'
            f'IF(AND({S_SET}!$C$9="PI値予測",$P{r}=""),"PI値未設定","")))')
        # 集計用（空文字を含まない数値列にしておき、ダッシュボードから合計する）
        ws[f'AC{r}'] = f'=N($Y{r})*N($H{r})'
        ws[f'AC{r}'].number_format = FMT_YEN

        style_row(ws, r, letters, fill=FILL_AUTO)
        for c in 'KLMNRUVWX':
            ws[f'{c}{r}'].number_format = FMT_INT
        for c in 'OPQS':
            ws[f'{c}{r}'].number_format = FMT_DEC
        ws[f'T{r}'].number_format = FMT_DAYS
        ws[f'G{r}'].number_format = FMT_INT
        ws[f'H{r}'].number_format = FMT_YEN
        ws[f'AA{r}'].number_format = FMT_YEN
        ws[f'Y{r}'].number_format = FMT_INT
        ws[f'Y{r}'].font = F_BOLD

    ws.conditional_formatting.add(
        f'Z{CALC_FIRST}:Z{CALC_LAST}',
        CellIsRule(operator='equal', formula=['"要発注"'], fill=FILL_ALERT, font=F_BOLD))
    ws.conditional_formatting.add(
        f'K{CALC_FIRST}:K{CALC_LAST}',
        CellIsRule(operator='lessThan', formula=['0'], fill=FILL_ALERT))

    widths = [5, 16, 30, 16, 22, 8, 11, 9, 10, 13, 13, 12, 12, 12, 9, 14, 10, 12,
              11, 13, 11, 10, 15, 17, 10, 13, 26, 15]
    for c, w in zip(letters, widths):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = 'D6'
    return ws


def build_timetable(wb):
    """
    在庫を上映スケジュール表の形で見せる画面。
    横軸は今日から14日。各商品の帯は在庫が尽きるまでの日数で、
    色が状態（余裕／要発注／今日が期限）、太い縦罫線が発注デッドラインを示す。
    発注数を入れると、伸びた分がバター色で帯の先に足される。
    """
    ws = wb.create_sheet(S_TT)
    ws.sheet_view.showGridLines = False
    ws['B2'] = 'タイムテーブル'
    ws['B2'].font = F_TITLE

    state_col = col(S_TT, 'G', TT_FIRST, TT_LAST)
    amount_col = col(S_TT, 'K', TT_FIRST, TT_LAST)

    # ---- 結論ブロック ----
    ws['B4'] = '今日の結論'
    ws['B4'].font = F_NOTE
    ws.merge_cells('B5:C5')          # 36ptの数字が「#」に潰れないよう幅を確保する
    ws.row_dimensions[5].height = 46
    ws['B5'] = f'=COUNTIF({state_col},"今日が期限")+COUNTIF({state_col},"要発注")'
    ws['B5'].font = F_HERO
    ws['B5'].alignment = Alignment(horizontal='left', vertical='center')
    ws['D5'] = '品目に発注が必要です'
    ws['D5'].font = F_HERO_UNIT
    ws['D5'].alignment = Alignment(horizontal='left', vertical='center')
    ws['B7'] = (f'=IF($B$5=0,"すべての商品が発注サイクルを越える在庫を持っています。",'
                f'"うち "&COUNTIF({state_col},"今日が期限")&'
                f'"品目は今日発注しないとリードタイムに間に合いません。 "&'
                f'COUNTIF({state_col},"要発注")&"品目は次の発注サイクル（"&{S_SET}!$C$13&'
                f'"日）内に在庫が切れます。")')
    ws['B7'].font = Font(name=FONT, size=10, color=C_TEXT2)

    ws['I4'] = '発注予定額（税抜）'
    ws['I4'].font = F_NOTE
    ws['I4'].alignment = Alignment(horizontal='right')
    ws['I5'] = f'=SUM({amount_col})'
    ws['I5'].font = Font(name=FONT, size=16, bold=True, color=C_INK)
    ws['I5'].number_format = FMT_YEN
    ws['I5'].alignment = Alignment(horizontal='right', vertical='center')

    # ---- 見出し ----
    headers = ['No', '商品名', '分類', '残り日数', 'デッドライン', '状態', '発注後日数',
               '推奨(ケース)', '発注数', '金額']
    kinds = ['自動', '自動', '自動', '自動', '自動', '自動', '自動', '自動', '入力', '自動']
    put_headers(ws, 9, headers, kinds)

    # 日付の目盛り。行9に日付、行8に日番号（条件付き書式が参照する）
    for d in range(SPAN_DAYS):
        letter = get_column_letter(TT_DAY_COL + d)
        head = ws[f'{letter}9']
        head.value = f'=IF({S_SET}!$C$6="","",{S_SET}!$C$6+{d})'
        head.number_format = 'm/d'
        head.font = F_DAY
        head.alignment = Alignment(horizontal='center')
        head.fill = FILL_HEAD
        head.border = BORDER
        n = ws[f'{letter}8']
        n.value = d + 1
        n.font = Font(name=FONT, size=8, color='FFD5D0C4')
        n.alignment = Alignment(horizontal='center')
        ws.column_dimensions[letter].width = 3.4

    # 行8は「入力/自動」の区分行と日番号が入るため、軸の説明は凡例側に置く

    # ---- データ行 ----
    for r in range(TT_FIRST, TT_LAST + 1):
        c = CALC_FIRST + (r - TT_FIRST)
        ws[f'B{r}'] = r - TT_FIRST + 1
        ws[f'C{r}'] = f'=IF({S_CALC}!$C{c}="","",{S_CALC}!$D{c})'
        ws[f'D{r}'] = f'=IF({S_CALC}!$C{c}="","",{S_CALC}!$E{c})'
        ws[f'E{r}'] = f'=IF({S_CALC}!$C{c}="","",{S_CALC}!$T{c})'
        ws[f'F{r}'] = f'=IF(ISNUMBER($E{r}),$E{r}-N({S_CALC}!$I{c}),"")'
        ws[f'G{r}'] = (f'=IF({S_CALC}!$C{c}="","",IF(NOT(ISNUMBER($F{r})),"データなし",'
                       f'IF($F{r}<=0,"今日が期限",IF($F{r}<={S_SET}!$C$13,"要発注","余裕"))))')
        ws[f'H{r}'] = (f'=IF(OR({S_CALC}!$C{c}="",NOT(ISNUMBER({S_CALC}!$S{c})),{S_CALC}!$S{c}<=0),"",'
                       f'({S_CALC}!$K{c}+N($J{r})*N({S_CALC}!$G{c}))/{S_CALC}!$S{c})')
        ws[f'I{r}'] = f'=IF({S_CALC}!$C{c}="","",{S_CALC}!$Y{c})'
        ws[f'K{r}'] = f'=IF($J{r}="","",N($J{r})*N({S_CALC}!$H{c}))'

        style_row(ws, r, 'BCDEFGHIJK', fill=FILL_AUTO)
        ws[f'J{r}'].fill = FILL_INPUT
        ws[f'J{r}'].font = F_INPUT
        ws[f'E{r}'].number_format = FMT_DAYS
        ws[f'F{r}'].number_format = FMT_DAYS
        ws[f'H{r}'].number_format = FMT_DAYS
        ws[f'I{r}'].number_format = FMT_INT
        ws[f'J{r}'].number_format = FMT_INT
        ws[f'K{r}'].number_format = FMT_YEN
        ws[f'G{r}'].alignment = Alignment(horizontal='center')
        for d in range(SPAN_DAYS):
            cell = ws.cell(r, TT_DAY_COL + d)
            cell.fill = FILL_TRACK
            cell.border = Border(right=Side(style='hair', color='FFEDE8DC'))
        ws.row_dimensions[r].height = 17

    # ---- 帯の描画（条件付き書式） ----
    first_day = get_column_letter(TT_DAY_COL)
    last_day = get_column_letter(TT_DAY_COL + SPAN_DAYS - 1)
    band = f'{first_day}{TT_FIRST}:{last_day}{TT_LAST}'
    # 数式は範囲の左上セル基準で書く。列は相対、日番号の行8と各項目の列は固定。
    day_ref = f'{first_day}$8'

    def rule(formula, color):
        return FormulaRule(formula=[formula],
                           fill=PatternFill('solid', bgColor=color), stopIfTrue=False)

    # 発注デッドライン（この日を過ぎるとリードタイム的に間に合わない）
    ws.conditional_formatting.add(band, FormulaRule(
        formula=[f'AND(ISNUMBER($F{TT_FIRST}),$F{TT_FIRST}>0,'
                 f'{day_ref}=ROUNDUP($F{TT_FIRST},0))'],
        border=Border(left=Side(style='medium', color=C_INK)), stopIfTrue=False))
    ws.conditional_formatting.add(band, rule(
        f'AND(ISNUMBER($E{TT_FIRST}),{day_ref}<=$E{TT_FIRST},$G{TT_FIRST}="今日が期限")', C_CRIT))
    ws.conditional_formatting.add(band, rule(
        f'AND(ISNUMBER($E{TT_FIRST}),{day_ref}<=$E{TT_FIRST},$G{TT_FIRST}="要発注")', C_WARN))
    ws.conditional_formatting.add(band, rule(
        f'AND(ISNUMBER($E{TT_FIRST}),{day_ref}<=$E{TT_FIRST},$G{TT_FIRST}="余裕")', C_OK))
    # 発注すると伸びる分
    ws.conditional_formatting.add(band, rule(
        f'AND(ISNUMBER($H{TT_FIRST}),{day_ref}>N($E{TT_FIRST}),{day_ref}<=$H{TT_FIRST})', C_BUTTER))

    # 状態列は色だけに頼らず、文字でも読めるようにする
    for text, fill, font_color in (('今日が期限', C_CRIT_SOFT, C_CRIT),
                                   ('要発注', C_WARN_SOFT, C_WARN),
                                   ('余裕', C_OK_SOFT, C_OK)):
        ws.conditional_formatting.add(
            f'G{TT_FIRST}:G{TT_LAST}',
            CellIsRule(operator='equal', formula=[f'"{text}"'],
                       fill=PatternFill('solid', bgColor=fill),
                       font=Font(name=FONT, size=10, bold=True, color=font_color)))

    # ---- 凡例 ----
    legend_row = TT_LAST + 2
    legend = [('今日が期限（今日発注しないと間に合わない）', C_CRIT),
              ('要発注（発注サイクル内に在庫が切れる）', C_WARN),
              ('余裕', C_OK),
              ('発注すると伸びる分', C_BUTTER)]
    for i, (text, color) in enumerate(legend):
        cell = ws.cell(legend_row + i, 12)
        cell.fill = PatternFill('solid', fgColor=color)
        label = ws.cell(legend_row + i, 13, text)
        label.font = F_NOTE
    ws.cell(legend_row + len(legend), 12, '│').font = Font(name=FONT, size=9, bold=True, color=C_INK)
    ws.cell(legend_row + len(legend), 13, '発注デッドライン（在庫が尽きる日 − リードタイム）').font = F_NOTE
    axis = ws.cell(legend_row + len(legend) + 2, 12,
                   '横軸は当日基準日からの14日間です。帯の長さがその商品の在庫が持つ日数を表します。')
    axis.font = F_NOTE
    ws.cell(legend_row + len(legend) + 3, 12,
            '発注数を入れると、伸びた分がバター色で帯の先に足されます。').font = F_NOTE

    for c, w in zip('BCDEFGHIJK', [5, 30, 15, 11, 13, 12, 12, 12, 11, 13]):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = f'C{TT_FIRST}'
    return ws


def build_po(wb):
    ws = wb.create_sheet(S_PO)
    sheet_title(ws, '発注管理',
                '発注と入荷を記録します。入荷済みの数量は期間納品数として、未入荷分は発注残として使われます。')
    headers = ['発注No', '劇場コード', '劇場名', '発注日', '商品コード', '商品名', '支払先名',
               '入数', '推奨発注数(ケース)', '実発注数(ケース)', '発注数量(バラ)', '税抜単価',
               '発注金額', 'L/T日数', '納品予定日', '発注状況', '入荷日', '入荷数量(バラ)',
               '担当者', '備考']
    kinds = ['自動', '入力', '自動', '入力', '入力', '自動', '自動', '自動', '自動', '入力',
             '自動', '自動', '自動', '自動', '自動', '選択', '入力', '自動', '入力', '入力']
    put_headers(ws, 5, headers, kinds)
    letters = [get_column_letter(i) for i in range(2, 2 + len(headers))]
    input_cols = 'CEFKQRTU'

    for r in range(PO_FIRST, PO_LAST + 1):
        ws[f'B{r}'] = (f'=IF(OR($E{r}="",$F{r}=""),"","PO-"&TEXT($E{r},"yyyymmdd")&"-"&'
                       f'TEXT(COUNTIFS($E${PO_FIRST}:$E{r},$E{r},$C${PO_FIRST}:$C{r},$C{r}),"000"))')
        ws[f'D{r}'] = f'=IFERROR(INDEX({theater_col("C")},MATCH($C{r},{theater_col("B")},0)),"")'
        ws[f'G{r}'] = f'=IFERROR(INDEX({item_col("C")},MATCH($F{r},{item_col("B")},0)),"")'
        ws[f'H{r}'] = f'=IFERROR(INDEX({item_col("E")},MATCH($F{r},{item_col("B")},0)),"")'
        ws[f'I{r}'] = f'=IFERROR(INDEX({item_col("F")},MATCH($F{r},{item_col("B")},0)),"")'
        ws[f'J{r}'] = (f'=IF(OR($F{r}="",$C{r}<>{S_SET}!$C$4),"",'
                       f'IFERROR(INDEX({col(S_CALC, "Y", CALC_FIRST, CALC_LAST)},'
                       f'MATCH($F{r},{col(S_CALC, "C", CALC_FIRST, CALC_LAST)},0)),""))')
        ws[f'L{r}'] = f'=IF($K{r}="","",N($K{r})*N($I{r}))'
        ws[f'M{r}'] = f'=IFERROR(INDEX({item_col("G")},MATCH($F{r},{item_col("B")},0)),"")'
        ws[f'N{r}'] = f'=IF($K{r}="","",N($K{r})*N($M{r}))'
        ws[f'O{r}'] = f'=IFERROR(INDEX({item_col("H")},MATCH($F{r},{item_col("B")},0)),"")'
        ws[f'P{r}'] = f'=IF(OR($E{r}="",$O{r}=""),"",WORKDAY($E{r},$O{r}))'
        ws[f'S{r}'] = f'=IF($Q{r}="入荷済み",N($L{r}),0)'

        style_row(ws, r, letters, fill=FILL_AUTO)
        style_row(ws, r, input_cols, fill=FILL_INPUT, font=F_INPUT)
        ws[f'C{r}'].number_format = '@'
        ws[f'F{r}'].number_format = '@'
        for c in 'EPR':
            ws[f'{c}{r}'].number_format = FMT_DATE
        for c in 'IJKLS':
            ws[f'{c}{r}'].number_format = FMT_INT
        for c in 'MN':
            ws[f'{c}{r}'].number_format = FMT_YEN

    add_validation(ws, f'={theater_col("B")}', f'C{PO_FIRST}:C{PO_LAST}')
    add_validation(ws, '"' + ','.join(ORDER_STATUS) + '"', f'Q{PO_FIRST}:Q{PO_LAST}')
    widths = [18, 12, 14, 14, 16, 28, 22, 8, 17, 17, 15, 11, 13, 9, 14, 12, 14, 15, 12, 24]
    for c, w in zip(letters, widths):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = 'D6'
    return ws


def build_all_theaters(wb):
    ws = wb.create_sheet(S_ALL)
    sheet_title(ws, '全劇場サマリ', '劇場マスタに登録された劇場について、当日CSVと発注管理から集計します。')
    headers = ['劇場コード', '劇場名', '規模区分', '提供方式', '取扱商品数', '当日在庫金額',
               'マイナス在庫件数', '未入荷発注件数', '当月発注金額']
    put_headers(ws, 5, headers, ['自動'] * len(headers))
    letters = [get_column_letter(i) for i in range(2, 2 + len(headers))]

    for r in range(THEATER_FIRST, THEATER_LAST + 1):
        src = r  # 劇場マスタと同じ行を参照する
        ws[f'B{r}'] = f'=IF({S_THEATER}!$B{src}="","",{S_THEATER}!$B{src})'
        ws[f'C{r}'] = f'=IF($B{r}="","",{S_THEATER}!$C{src})'
        ws[f'D{r}'] = f'=IF($B{r}="","",{S_THEATER}!$D{src})'
        ws[f'E{r}'] = f'=IF($B{r}="","",{S_THEATER}!$E{src})'
        ws[f'F{r}'] = (f'=IF($B{r}="","",COUNTIFS({csv_col(S_CUR, "B")},$B{r},'
                       f'{csv_col(S_CUR, "I")},"<>"))')
        ws[f'G{r}'] = (f'=IF($B{r}="","",SUMPRODUCT(({csv_col(S_CUR, "B")}=$B{r})*'
                       f'{csv_col(S_CUR, "N")}*{csv_col(S_CUR, "Q")}))')
        ws[f'H{r}'] = (f'=IF($B{r}="","",COUNTIFS({csv_col(S_CUR, "B")},$B{r},'
                       f'{csv_col(S_CUR, "N")},"<0"))')
        ws[f'I{r}'] = (f'=IF($B{r}="","",COUNTIFS({po_col("C")},$B{r},{po_col("Q")},"発注済"))')
        ws[f'J{r}'] = (f'=IF(OR($B{r}="",{S_SET}!$C$6=""),"",'
                       f'SUMIFS({po_col("N")},{po_col("C")},$B{r},'
                       f'{po_col("E")},">="&DATE(YEAR({S_SET}!$C$6),MONTH({S_SET}!$C$6),1),'
                       f'{po_col("E")},"<="&DATE(YEAR({S_SET}!$C$6),MONTH({S_SET}!$C$6)+1,0)))')
        style_row(ws, r, letters, fill=FILL_AUTO)
        ws[f'B{r}'].number_format = '@'
        for c in 'FHI':
            ws[f'{c}{r}'].number_format = FMT_INT
        for c in 'GJ':
            ws[f'{c}{r}'].number_format = FMT_YEN

    ws.conditional_formatting.add(
        f'H{THEATER_FIRST}:H{THEATER_LAST}',
        CellIsRule(operator='greaterThan', formula=['0'], fill=FILL_ALERT))
    for c, w in zip(letters, [12, 22, 12, 16, 13, 16, 18, 16, 16]):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = 'D6'
    return ws


def build_dashboard(wb, categories):
    ws = wb.create_sheet(S_DASH)
    sheet_title(ws, 'ダッシュボード', '設定シートで選んだ劇場の状況です。')
    ws.column_dimensions['B'].width = 24
    for c in 'CDEFGH':
        ws.column_dimensions[c].width = 16

    head = [('対象劇場', f'=IF({S_SET}!$C$4="","",{S_SET}!$C$4&"　"&{S_SET}!$C$5)', None),
            ('当日基準日', f'={S_SET}!$C$6', FMT_DATE),
            ('前回基準日', f'={S_SET}!$C$7', FMT_DATE),
            ('算出基準', f'={S_SET}!$C$9', None),
            ('季節係数', f'={S_SET}!$C$15', '0%')]
    for i, (label, f, fmt) in enumerate(head):
        r = 5 + i
        ws[f'B{r}'] = label
        ws[f'B{r}'].font = F_BOLD
        ws[f'B{r}'].fill = FILL_HEAD
        ws[f'B{r}'].border = BORDER
        ws[f'C{r}'] = f
        ws[f'C{r}'].font = F_BASE
        ws[f'C{r}'].border = BORDER
        if fmt:
            ws[f'C{r}'].number_format = fmt

    calc_z = col(S_CALC, 'Z', CALC_FIRST, CALC_LAST)
    calc_c = col(S_CALC, 'C', CALC_FIRST, CALC_LAST)
    calc_aa = col(S_CALC, 'AA', CALC_FIRST, CALC_LAST)
    calc_e = col(S_CALC, 'E', CALC_FIRST, CALC_LAST)
    calc_k = col(S_CALC, 'K', CALC_FIRST, CALC_LAST)
    calc_ab = col(S_CALC, 'AB', CALC_FIRST, CALC_LAST)
    calc_ac = col(S_CALC, 'AC', CALC_FIRST, CALC_LAST)
    order_amount = f'SUM({calc_ac})'

    kpis = [
        # 空行の数式結果は "" になるため、COUNTIF("<>") では数えられない
        ('取扱商品数', f'=SUMPRODUCT(--({calc_c}<>""))', FMT_INT),
        ('要発注 品目数', f'=COUNTIF({calc_z},"要発注")', FMT_INT),
        ('推奨発注金額（税抜）', f'={order_amount}', FMT_YEN),
        ('当日在庫金額（税抜）', f'=SUM({calc_aa})', FMT_YEN),
        ('在庫マイナス 品目数', f'=COUNTIF({calc_k},"<0")', FMT_INT),
        ('マスタ未登録 品目数', f'=COUNTIF({calc_ab},"*マスタ未登録*")', FMT_INT),
        ('新商品 品目数', f'=COUNTIF({calc_ab},"*新商品*")', FMT_INT),
        ('終売候補 品目数', f'=COUNTIF({calc_ab},"*終売候補*")', FMT_INT),
    ]
    for i, (label, f, fmt) in enumerate(kpis):
        r = 12 + i
        ws[f'B{r}'] = label
        ws[f'B{r}'].font = F_BOLD
        ws[f'B{r}'].fill = FILL_HEAD
        ws[f'B{r}'].border = BORDER
        ws[f'C{r}'] = f
        ws[f'C{r}'].font = F_BOLD
        ws[f'C{r}'].number_format = fmt
        ws[f'C{r}'].border = BORDER

    ws['B22'] = '商品分類別'
    ws['B22'].font = F_BOLD
    put_headers(ws, 23, ['商品分類名', '在庫金額', '推奨発注金額', '要発注品目数'],
                ['入力', '自動', '自動', '自動'])
    for i, name in enumerate(categories):
        r = 24 + i
        ws[f'B{r}'] = name
        ws[f'C{r}'] = f'=SUMIF({calc_e},$B{r},{calc_aa})'
        ws[f'D{r}'] = f'=SUMIF({calc_e},$B{r},{calc_ac})'
        ws[f'E{r}'] = f'=COUNTIFS({calc_e},$B{r},{calc_z},"要発注")'
        style_row(ws, r, 'BCDE', fill=FILL_AUTO)
        ws[f'B{r}'].fill = FILL_INPUT
        ws[f'C{r}'].number_format = FMT_YEN
        ws[f'D{r}'].number_format = FMT_YEN
        ws[f'E{r}'].number_format = FMT_INT
    last = 24 + len(categories) - 1

    chart = BarChart()
    chart.type = 'bar'
    chart.title = '商品分類別 在庫金額'
    chart.height, chart.width = 10, 14
    data = Reference(ws, min_col=3, min_row=23, max_row=last)
    cats = Reference(ws, min_col=2, min_row=24, max_row=last)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.series[0].graphicalProperties.solidFill = C_INK[2:]
    ws.add_chart(chart, 'G22')

    chart2 = BarChart()
    chart2.type = 'bar'
    chart2.title = '商品分類別 推奨発注金額'
    chart2.height, chart2.width = 10, 14
    data2 = Reference(ws, min_col=4, min_row=23, max_row=last)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats)
    chart2.series[0].graphicalProperties.solidFill = C_BUTTER[2:]
    ws.add_chart(chart2, 'O22')

    ws[f'B{last + 3}'] = '※ 商品分類名は在庫CSVの値です。分類が増えたらこの表に追記してください。'
    ws[f'B{last + 3}'].font = F_NOTE
    return ws


# ---------------------------------------------------------------------------

def load_items(csv_path):
    with open(csv_path, encoding='cp932') as fh:
        rows = list(csv.DictReader(fh))
    seen, items = set(), []
    for r in rows:
        code = r['商品コード'].strip()
        if code and code not in seen:
            seen.add(code)
            items.append(r)
    categories = []
    for r in rows:
        if r['商品分類名'] not in categories:
            categories.append(r['商品分類名'])
    return items, categories


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path('売店発注ツール.xlsx')
    src = Path(sys.argv[2]) if len(sys.argv) > 2 else None
    items, categories = load_items(src) if src else ([], ['コンセ包材'])

    wb = Workbook()
    wb.remove(wb.active)
    build_intro(wb)
    build_settings(wb)
    build_calc(wb)
    build_timetable(wb)
    build_dashboard(wb, categories)
    build_po(wb)
    build_csv_sheet(wb, S_CUR, None, '在庫CSV（当日）')
    build_csv_sheet(wb, S_PRV, S_CUR, '在庫CSV（前回）')
    build_item_master(wb, items)
    build_theater_master(wb)
    build_size_master(wb)
    build_season_master(wb)
    build_all_theaters(wb)

    order = ['はじめに', S_SET, S_TT, S_DASH, S_CALC, S_PO, S_CUR, S_PRV,
             S_ITEM, S_THEATER, S_SIZE, S_SEASON, S_ALL]
    wb._sheets = [wb[n] for n in order]
    for ws in wb.worksheets:
        if ws.title in (S_TT, S_SET):
            ws.sheet_properties.tabColor = C_BUTTER[2:]
        elif ws.title in (S_DASH, S_CALC):
            ws.sheet_properties.tabColor = C_INK[2:]

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f'saved: {out}  sheets={len(wb.worksheets)}  items={len(items)}')


if __name__ == '__main__':
    main()
