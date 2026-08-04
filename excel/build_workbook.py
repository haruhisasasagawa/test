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
from openpyxl.chart.marker import DataPoint
from openpyxl.formatting.rule import (CellIsRule, ColorScaleRule, DataBarRule,
                                      FormulaRule)
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
# 曜日書式(aaa)は環境によって日付そのものを壊すことがあるため使わない。
# 曜日が要る場所は CHOOSE(WEEKDAY(...)) で別に組み立てる。
FMT_DATE = 'yyyy/m/d'
FMT_PI = '#,##0.00'          # 実測PI値。小さい値が多いので小数2桁で見る

# 在庫CSVの列（システム出力そのままの並び）
CSV_HEADERS = ['対象日付', '劇場コード', '劇場名', '支払先コード', '支払先名',
               '大分類コード', '小分類コード', '商品分類名', '商品コード', '商品名',
               '入数', '在庫数ケース', '在庫数バラ', '総数', '規定数', '資産廃棄',
               '税抜単価', '仕入税区分', '仕入税率']
CSV_HEAD = 5           # 見出し行。1〜4行目は案内と取込判定に使う
CSV_FIRST = 6          # データ開始行
CSV_LAST = 10005       # データ終了行（75劇場×約130商品を想定）
CSV_HELPER = 'U'       # 連番ヘルパー列

SIZES = ['大規模', '中規模', '小規模']
DRINK_STYLES = ['1杯売り', 'ドリンクバー']
# PIプロファイル＝ドリンクの提供方式。PI値は「来場者100人あたりの消費数」なので、
# 劇場の規模が変わっても値そのものは変わらない（規模は来場者数の側で吸収される）。
# 提供方式だけが商品構成を変えるため、プロファイルはこの2種類で足りる。
PROFILES = list(DRINK_STYLES)

# 採用消費/日の決め方。
#   実績消費 … 在庫2時点の差を期間日数で割っただけ。動員とは連動しない
#   動員連動 … 実績消費を「これからの動員 ÷ 期間の平均動員」で伸縮させる。
#              購買率を人が入れなくても、実績と動員の比から自動で効く（既定）
#   PI値予測 … 商品マスタに手入力したPI値を使う
#   最大値   … 上の3つのうち最大。欠品を避けたいとき
BASIS_OPTIONS = ['動員連動', '実績消費', 'PI値予測', '最大値']
ORDER_STATUS = ['発注済', '入荷済み', '取消']
# 実績消費をどの期間で見るか。2ヶ月はならして安定、1ヶ月は最近の動きに追随する
SPAN_OPTIONS = ['2ヶ月', '直近1ヶ月']
# 発注量は定数から決める。消費量は使わない。
#   発注数 ＝ 定数（最大保管数） − 今の在庫 − 発注済み
# 消費量（＝前回在庫＋納品−当日在庫）は、期間中の納品が発注管理にすべて
# 記録されていることが前提で、記録が抜けると過小に出る。実運用でその前提は
# 成り立たないため、発注量の計算からは外した。
# 消費量は「定数をいくつにするか」の目安と、残り日数の表示にだけ使う。

MASTER_FIRST, MASTER_LAST = 6, 1005      # 商品マスタ
THEATER_FIRST, THEATER_LAST = 6, 105     # 劇場マスタ
CALC_FIRST, CALC_LAST = 6, 405           # 発注計算
PO_FIRST, PO_LAST = 6, 1005              # 発注管理

SPAN_DAYS = 14                           # タイムテーブルに表示する日数
TT_FIRST = 10                            # タイムテーブルのデータ開始行
TT_LAST = TT_FIRST + (CALC_LAST - CALC_FIRST)
TT_DAY_COL = 14                          # N列から14日分のグリッド
# ②発注数を決める の列。定数管理が主線なので、定数と在庫を明細に出す
TT_COL = {'no': 'B', 'name': 'C', 'group': 'D', 'par': 'E', 'stock': 'F',
          'days': 'G', 'dead': 'H', 'state': 'I', 'after': 'J',
          'rec': 'K', 'qty': 'L', 'amt': 'M', 'pick': 'AB'}

# シート名。日々の作業順に番号を振り、上から順にたどれば発注が終わるようにしている。
# シート名は「何をするか」で付ける。毎日さわる3枚に①②③を振り、
# たまにしか触らないシートは頻度を名前に書いておく。
S_INTRO = 'はじめに'
S_MECH = '発注のしくみ'
S_COUNT = '実棚を入れる'
S_SET = '設定（最初に1回）'
S_CUR = '①今日の在庫を貼る'
S_TT = '②発注数を決める'
S_SHEET = '③発注書'
S_MID = '月1回_1ヶ月前の在庫を貼る'
S_PRV = '月1回_2ヶ月前の在庫を貼る'
S_PLAN = '随時_動員を入れる'
S_DASH = 'ダッシュボード'
S_PO = '発注管理'
S_ITEM = '商品マスタ'
S_CONST = '定数マスタ'
S_CALC = '発注計算'
S_ALL = '全劇場サマリ'

# 数式の中では常にシングルクォートで囲む。丸数字などを含む名前でも安全に参照できる。
Q_SET = f"'{S_SET}'"
Q_PRV = f"'{S_PRV}'"
Q_MID = f"'{S_MID}'"
Q_CUR = f"'{S_CUR}'"
Q_PLAN = f"'{S_PLAN}'"
Q_DASH = f"'{S_DASH}'"
Q_TT = f"'{S_TT}'"
Q_SHEET = f"'{S_SHEET}'"
Q_PO = f"'{S_PO}'"
Q_CALC = f"'{S_CALC}'"
Q_CONST = f"'{S_CONST}'"
Q_ITEM = f"'{S_ITEM}'"
Q_THEATER = Q_CONST
Q_ALL = f"'{S_ALL}'"
Q_COUNT = f"'{S_COUNT}'"

# 実棚シートは発注計算と行が1対1で対応する（No が同じ）。
# 検索せずに同じ順番で並べるので、400商品でも数式が重くならない。
COUNT_FIRST = 10          # 4〜6行はまとめ、8行が入力/自動、9行が見出し
COUNT_LAST = COUNT_FIRST + (CALC_LAST - CALC_FIRST)

PLAN_FIRST, PLAN_LAST = 16, 235          # 動員（前回基準日から220日分）
PLAN_DOW_ROW = 6                         # 曜日別実績平均ブロックの先頭行

# 定数マスタは1枚に4ブロックを横に並べる。マスタが増えるほど管理が重くなるため、
# 劇場・規模・月別係数・特別期間をこの1枚に集約している。
SIZE_COL = {'B': 'N', 'C': 'O', 'D': 'P', 'E': 'Q', 'F': 'R'}      # 規模区分ブロック
MONTH_COL = {'B': 'T', 'C': 'U'}                                    # 月別係数ブロック
SPECIAL_COL = {'E': 'W', 'F': 'X', 'G': 'Y', 'H': 'Z'}              # 特別期間ブロック
EQUIP_FIRST, EQUIP_LAST = 110, 309       # ⑤保管設備（劇場ブロックの下）
PI_FIRST_COL = 12                        # 商品マスタのPI値開始列（L列）
# ⑥発注サイクルは AB列から。列幅はシート単位でしか決められないため、
# 曜日の○（幅3）を置く列は他のブロックと重ならない位置に逃がしている。
CYCLE_FIRST, CYCLE_LAST = 6, 15
CYCLE_NAME, CYCLE_CNT, CYCLE_DAYS, CYCLE_TODAY, CYCLE_NOTE = 'AB', 'AJ', 'AK', 'AL', 'AM'
CYCLE_DAY1 = 29                          # AC列＝月曜。WEEKDAY(日付,2) の 1〜7 と並びが揃う
WEEKDAY_COLS = [get_column_letter(CYCLE_DAY1 + i) for i in range(7)]
# 「次にこのグループを発注できる日」を求める補助列。AN〜AT が0〜6日先の判定
CYCLE_OFF1 = 40                          # AN列
CYCLE_WAIT, CYCLE_NEXT = 'AU', 'AV'      # 次の発注日までの日数／次の発注日
CYCLE_LBL_NOW, CYCLE_LBL_WAIT = 'AW', 'AX'   # 帯に出す文言

# 商品グループごとの発注サイクル。TOHOシネマズの実際の頻度を初期値にしてある
ORDER_GROUPS = [
    ('冷凍品', '月水金', ('月', '水', '金'), '週3回。曜日は融通が利く'),
    ('ドリンクシロップ', '月水金', ('月', '水', '金'), '週3回'),
    ('常温包材', '木', ('木',), '毎週木曜の1回のみ'),
    ('その他常温', '木', ('木',), '包材と同じ便で受ける'),
]

# 保管区分。商品マスタと保管設備の両方で使う共通の区分
STORAGE_KINDS = ['常温', '冷蔵', '冷凍']

# ---------------------------------------------------------------------------
# 劇場マスタの初期データ
# ---------------------------------------------------------------------------
# 公式サイト（tohotheater.jp）の各劇場ページURLに出てくる3桁コードと劇場名。
# URLで実際に確認できたものだけを載せている（推測では入れていない）。
#   例: /theater/076/institution.html → 新宿 = 076
#
# 自社の在庫CSVの劇場コードは4桁（新宿＝0761）で、公式3桁＋"1" の形になっている。
# 1劇場ぶんしか実データが無いため、この規則は「暫定」として展開し、
# 劇場マスタのM列（CSV照合）で全劇場ぶんまとめて検証できるようにしてある。
# 貼ったCSVに当たらないコードは即座に「✖」で分かる。
#
# スクリーン数・座席数は公式ページの記載を拾えたものだけ。規模区分の判断材料になる。
# (公式3桁, 劇場名, 都道府県, スクリーン数, 座席数, 補足)
THEATER_SEED = [
    ('089', 'すすきの', '北海道', 10, 1732, ''),
    ('049', 'おいらせ下田', '青森', None, None, ''),
    ('050', '秋田', '秋田', None, None, ''),
    ('078', '仙台', '宮城', 9, 1700, ''),
    ('024', 'ひたちなか', '茨城', 10, 1700, ''),
    ('015', '宇都宮', '栃木', None, None, ''),
    ('075', 'ららぽーと富士見', '埼玉', None, None, ''),
    ('003', '市川コルトンプラザ', '千葉', 9, 2150, ''),
    ('018', 'ららぽーと船橋', '千葉', 10, 1867, ''),
    ('028', '八千代緑が丘', '千葉', None, None, ''),
    ('035', '流山おおたかの森', '千葉', 11, 1900, ''),
    ('071', '市原', '千葉', None, None, ''),
    ('077', '柏', '千葉', None, None, ''),
    ('006', '南大沢', '東京', 9, None, ''),
    ('009', '六本木ヒルズ', '東京', None, None, '旗艦劇場'),
    ('012', '府中', '東京', None, None, ''),
    ('029', '錦糸町（楽天地・オリナス）', '東京', None, None, '2館体制'),
    ('040', '西新井', '東京', None, None, ''),
    ('044', 'お台場（シネマメディアージュ）', '東京', 13, 3034, ''),
    ('073', '日本橋', '東京', None, None, ''),
    ('076', '新宿', '東京', None, None, 'IMAXレーザー'),
    ('080', '上野', '東京', None, None, ''),
    ('081', '日比谷／シャンテ', '東京', 13, 2700, '日比谷11＋宝塚ビル2の一体運営'),
    ('084', '池袋', '東京', None, None, ''),
    ('085', '立川立飛', '東京', None, None, 'IMAX／轟音'),
    ('007', '海老名', '神奈川', None, None, ''),
    ('008', '小田原', '神奈川', 9, None, ''),
    ('010', '川崎', '神奈川', 9, None, ''),
    ('036', 'ららぽーと横浜', '神奈川', None, None, ''),
    ('067', '甲府', '山梨', None, None, ''),
    ('068', '上田', '長野', None, None, ''),
    ('053', 'ファボーレ富山', '富山', None, None, ''),
    ('054', '高岡', '富山', None, None, ''),
    ('020', '岐阜', '岐阜', None, None, ''),
    ('030', 'モレラ岐阜', '岐阜', None, None, ''),
    ('004', '浜松', '静岡', None, None, ''),
    ('016', '木曽川', '愛知', 10, 1828, ''),
    ('021', '東浦', '愛知', None, None, ''),
    ('026', '津島', '愛知', None, None, ''),
    ('079', '赤池', '愛知', None, None, ''),
    ('023', '二条', '京都', None, None, 'IMAXレーザー'),
    ('005', '泉北', '大阪', None, None, ''),
    ('032', 'なんば', '大阪', None, None, 'IMAX'),
    ('037', '梅田', '大阪', None, None, '本館・別館'),
    ('045', '鳳', '大阪', 10, 1950, ''),
    ('072', 'くずはモール', '大阪', None, None, ''),
    ('088', 'ららぽーと門真', '大阪', None, None, ''),
    ('038', '伊丹', '兵庫', None, None, ''),
    ('064', '西宮OS', '兵庫', None, None, ''),
    ('013', '橿原', '奈良', None, None, ''),
    ('031', '岡南', '岡山', 10, 1629, ''),
    ('019', '緑井', '広島', None, None, ''),
    ('017', '高知', '高知', 9, None, '四国最大級'),
    ('022', '直方', '福岡', None, None, ''),
    ('056', '天神（ソラリア館）', '福岡', None, None, ''),
    ('087', 'ららぽーと福岡', '福岡', 9, 1322, ''),
    ('046', '長崎', '長崎', None, None, ''),
    ('014', '光の森', '熊本', 9, None, ''),
    ('027', 'はません', '熊本', None, None, ''),
    ('057', '宇城', '熊本', None, None, ''),
    ('083', '熊本サクラマチ', '熊本', 9, 1578, ''),
    ('055', '大分わさだ', '大分', None, None, ''),
    ('074', 'アミュプラザおおいた', '大分', 10, 1764, ''),
    ('033', '与次郎', '鹿児島', 10, 1984, '2026/8/6 閉館予定'),
]


def seed_size(screens, seats):
    """規模区分の暫定値。スクリーン数・座席数から機械的に置くだけで、
    最終的には現場の実感（動員・売上）で直してもらう前提の初期値。"""
    if (seats or 0) >= 2000 or (screens or 0) >= 11:
        return '大規模'
    if (screens or 0) and screens <= 7:
        return '小規模'
    return '中規模'


def weekday_jp(date_expr):
    """日付式から曜日の文字を組み立てる。書式コード aaa は環境依存で使えない。"""
    return f'CHOOSE(WEEKDAY({date_expr}),"日","月","火","水","木","金","土")'


def size_col(letter, first=6, last=8):
    return col(S_CONST, SIZE_COL[letter], first, last)


def month_col(letter, first=6, last=17):
    return col(S_CONST, MONTH_COL[letter], first, last)


def special_col(letter, first=6, last=25):
    return col(S_CONST, SPECIAL_COL[letter], first, last)


def equip_col(letter):
    return col(S_CONST, letter, EQUIP_FIRST, EQUIP_LAST)


def cycle_col(letter):
    return col(S_CONST, letter, CYCLE_FIRST, CYCLE_LAST)


def col(sheet, letter, first=None, last=None):
    """絶対参照の列範囲を返す。"""
    first = CSV_FIRST if first is None else first
    last = CSV_LAST if last is None else last
    return f"'{sheet}'!${letter}${first}:${letter}${last}"


def csv_col(sheet, letter):
    return col(sheet, letter, CSV_FIRST, CSV_LAST)


def item_col(letter):
    return col(S_ITEM, letter, MASTER_FIRST, MASTER_LAST)


def theater_col(letter):
    return col(S_CONST, letter, THEATER_FIRST, THEATER_LAST)


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
    """
    はじめに。役割ごとに「あなたがやること」を分けて書く。
    毎日の担当者が読むのは最初のブロックだけで済むようにしている。
    """
    ws = wb.create_sheet(S_INTRO)
    sheet_title(ws, '売店発注ツール')
    ws.column_dimensions['B'].width = 4
    ws.column_dimensions['C'].width = 22
    ws.column_dimensions['D'].width = 78

    rows = [
        ('h', '', '毎日やること　これだけです'),
        ('t', '', 'シートの①②③を左から順にやれば、その日の発注は終わります。'),
        ('t', '', '黄色いセルだけ入力します。それ以外のセルは触らないでください。'),
        ('', '', ''),
        ('s', S_CUR, '在庫システムから出した「今日の在庫一覧CSV」を貼る'),
        ('t', '', '　CSVを開いて、見出し行を除いたデータ部分をコピーし、A2セルに貼り付けます。'),
        ('t', '', '　貼り付けたら、シート上部の「該当行数」が0以外になっているか見てください。'),
        ('t', '', '　0のままなら貼り付けに失敗しています。もう一度やり直してください。'),
        ('s', S_TT, '帯を見ながら、黄色の「発注数」に数字を入れる'),
        ('t', '', '　いちばん上に「今日どのグループを発注できるか」が出ます。まずそこを見てください。'),
        ('t', '', '　「定数」と「在庫」が並んでいます。発注数＝定数−在庫−発注済み です。'),
        ('t', '', '　定数が赤く「未設定」の商品は、暫定の数字です。社員に伝えてください。'),
        ('t', '', '　発注グループが薄いグレーの行は、今日が発注日ではない商品です。'),
        ('t', '', '　帯が赤い商品は、今日発注しないと在庫が切れます。'),
        ('t', '', '　迷ったら「推奨(ケース)」に出ている数字をそのまま入れて構いません。'),
        ('t', '', '　数字を入れると帯が伸びます。伸びた先まで在庫が持つ、という意味です。'),
        ('t', '', '　右上に「保管スペースの埋まり具合」が出ます。100%を超えると入りきりません。'),
        ('s', S_COUNT, '★の付いた商品だけ、現物を数えて入れる'),
        ('t', '', '　CSVの在庫は理論値です。ロスや納品の未計上で現物とずれます。'),
        ('t', '', '　全部数える必要はありません。★が付いた商品だけで十分です。'),
        ('t', '', '　数えた数とカウント日（今日の日付）を入れると、そちらで発注数を計算します。'),
        ('s', S_SHEET, '発注先を選んで、印刷する'),
        ('t', '', '　②で入れた数量がここに集まります。発注先を選ぶとその仕入先の分だけになります。'),
        ('t', '', '　そのまま印刷して、いつもの方法で発注してください。'),
        ('', '', ''),
        ('t', '', '発注したら「発注管理」シートに記録します。入荷したら入荷日も入れてください。'),
        ('', '', ''),
        ('s', S_MECH, '「何を計算しているのか」を図で見る（最初に1回だけ読めば十分です）'),
        ('t', '', '　定数がどこにあるか、そこから発注数がどう出るかを1枚にまとめてあります。'),
        ('', '', ''),
        ('h', '', '色の意味'),
        ('t', '', '　黄色のセル … 入力するところ'),
        ('t', '', '　白・灰色のセル … 自動で計算されるところ（触らない）'),
        ('t', '', '　赤い帯 … 今日発注しないと間に合わない'),
        ('t', '', '　オレンジの帯 … 次の発注までに在庫が切れる'),
        ('t', '', '　緑の帯 … 在庫に余裕がある'),
        ('', '', ''),
        ('h', '', '社員の方がやること'),
        ('s', S_MID, '月に1回、1ヶ月前の在庫一覧CSVを貼り替える'),
        ('s', S_PRV, '月に1回、2ヶ月前の在庫一覧CSVを貼り替える'),
        ('t', '', '　取扱商品が変わっていないかを確認するために貼ります。'),
        ('t', '', '　商品は半年ごとに入れ替わります。当日のCSVだけ見ていると、'),
        ('t', '', '　先月まで扱っていた商品が発注リストから静かに消えます。'),
        ('t', '', '　3枚を突き合わせれば、1ヶ月だけ扱った商品も漏れません。'),
        ('t', '', '　貼ったシートの上に、新商品・終売・復活の件数が出ます。'),
        ('s', S_PLAN, '動員の実績を入れる／公開週などの読みを入れる'),
        ('t', '', '　「実績来場者数」は毎日でも、月1回まとめてでもかまいません。'),
        ('t', '', '　実績を入れると、曜日ごとの平均が自動で出ます。'),
        ('t', '', '　予測を入れていない日は、その曜日の実績平均が使われます。'),
        ('t', '', '　「予測来場者数」は大型作品の公開週など、読みが変わる日だけで十分です。'),
        ('t', '', '　ここに入れた動員は、そのまま発注量に効きます。'),
        ('t', '', '　動員が平常の1.3倍なら、消費も1.3倍で見積もって発注数が増えます。'),
        ('s', S_DASH, '劇場全体の状況を確認する'),
        ('s', S_ITEM, 'リードタイム・最低ロット・最大保管数を登録する'),
        ('t', '', '　「最大保管数(ケース)」＝その商品を棚に何ケース置けるか。'),
        ('t', '', '　隣の「目安(自動)」に必要なケース数が出るので、それ以上を入れてください。'),
        ('t', '', '　空欄でも動きますが、入れると発注数がその数を超えなくなります。'),
        ('s', S_CONST, '劇場・規模区分・季節の係数・保管設備を決める'),
        ('t', '', '　人が決める数字はこの1枚に集めてあります。'),
        ('t', '', '　⑤保管設備に冷蔵庫やストッカーを登録すると、収容ケース数が自動で集計され、'),
        ('t', '', '　②の「保管スペースの埋まり具合」に反映されます。'),
        ('s', S_SET, '対象の劇場と、発注数の算出基準を選ぶ'),
        ('', '', ''),
        ('h', '', '発注数はどう決まっているか'),
        ('t', '', '　この商品は1日に何個くらい減るか'),
        ('t', '', '　　（前回と今日の在庫の差 ＝ 実績。これに動員の増減を掛ける）'),
        ('t', '', '　　　　　　↓'),
        ('t', '', '　届くまでの日数と、次に発注するまでの日数のあいだに何個必要か'),
        ('t', '', '　　　　　　↓'),
        ('t', '', '　その必要数から、今ある在庫と発注済みの分を引く'),
        ('t', '', '　　　　　　↓'),
        ('t', '', '　入数で割って、ケース単位に切り上げる ＝ 推奨発注数'),
        ('', '', ''),
        ('t', '', '式で書くと次のとおりです。'),
        ('t', '', '　1日あたり消費 ＝ 実績消費 ×（これからの動員 ÷ 期間の平均動員）'),
        ('t', '', '　安全在庫　＝ 1日あたり消費 × 安全在庫日数'),
        ('t', '', '　発注点　　＝ 1日あたり消費 ×（リードタイム＋発注サイクル）＋ 安全在庫'),
        ('t', '', '　推奨発注数 ＝ 発注点 − 今の在庫 − 発注済み未入荷'),
        ('', '', ''),
        ('h', '', '困ったとき'),
        ('t', '', '　・数字が出ない → 設定シートの「該当行数」が0でないか確認'),
        ('t', '', '　・在庫がマイナス → 納品がシステムに未計上です。'),
        ('t', '', '　　　現物を数えて「実棚を入れる」に入れてください。放置すると大量に発注してしまいます'),
        ('t', '', '　・「マスタ未登録」と出る → 新商品です。商品マスタに追加してください'),
    ]
    r = 5
    for kind, sheet, text in rows:
        if sheet:
            c = ws[f'C{r}']
            c.value = sheet
            c.font = Font(name=FONT, size=10, bold=True, color=C_INK)
            c.fill = FILL_INPUT if sheet in (S_CUR, S_TT, S_SHEET) else FILL_HEAD
            c.border = BORDER
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell = ws[f'D{r}']
        cell.value = text
        if kind == 'h':
            cell.font = Font(name=FONT, size=12, bold=True, color=C_INK)
            cell.fill = FILL_ACCENT
        elif kind == 's':
            cell.font = F_BOLD
        else:
            cell.font = Font(name=FONT, size=10, color=C_TEXT2)
        r += 1
    return ws


def build_mechanism(wb):
    """
    発注のしくみ。よく聞かれる2点だけを図で答えるシート。
      1. 定数はどこにあるのか  → 置き場所マップ（今の値を実際に引いて表示する）
      2. 定数から発注数がどう出るのか → 1商品を選んで数字を順に追う

    説明文を並べるだけだと読まれないため、「今の値」は必ず生きた数式で引いて、
    自分の劇場の実際の数字で読めるようにしている。
    """
    ws = wb.create_sheet(S_MECH)
    ws.sheet_view.showGridLines = False
    for i in range(2, 39):                    # B〜AL列を方眼にする
        ws.column_dimensions[get_column_letter(i)].width = 3.3

    calc_name = col(S_CALC, 'D', CALC_FIRST, CALC_LAST)
    calc_qty = col(S_CALC, 'AH', CALC_FIRST, CALC_LAST)   # 既定表示を選ぶための列

    def mrg(rng, value, font=F_BASE, fill=None, fmt=None, align='left', border=False):
        ws.merge_cells(rng)
        c = ws[rng.split(':')[0]]
        c.value = value
        c.font = font
        if fill:
            c.fill = fill
        if fmt:
            c.number_format = fmt
        c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=True)
        if border:
            for row in ws[rng]:
                for cell in row:
                    cell.border = BORDER
        return c

    def section(row, text, note=None):
        ws[f'B{row}'] = text
        ws[f'B{row}'].font = Font(name=FONT, size=12, bold=True, color=C_INK)
        if note:
            mrg(f'B{row + 1}:AL{row + 1}', note, font=F_NOTE)

    mrg('B2:N2', '発注のしくみ', font=F_TITLE)
    mrg('B3:AL3', 'このツールが何をしているかの説明です。'
                  '最初に1回読めば、あとは②発注数を決めるシートだけ見れば発注できます。',
        font=F_NOTE)

    # ================= 3ステップの図 =================
    section(5, '発注は3つのことしかしていません')
    steps = [
        ('B', 'L', '① 決めておく', C_BUTTER, FILL_INPUT,
         '「この商品は棚に何ケース置けるか」を先に決めておきます。'
         'これが定数（最大保管数）です。毎日は変わりません。',
         '置き場所： 定数マスタ ／ 商品マスタ',
         '触るのは社員だけ。ふだんは開きません。'),
        ('O', 'Y', '② 今の在庫を読む', C_TEXT2, FILL_AUTO,
         'CSVを貼るだけです。ただしCSVの数は理論値なので、'
         'ずれていそうな商品は現物を数えて上書きします。',
         '置き場所： ①今日の在庫を貼る ／ 実棚を入れる',
         '毎日やるのはここだけです。'),
        ('AB', 'AL', '③ 足りない分だけ発注', C_CRIT, FILL_ALERT,
         '定数から、今の在庫と発注済みの分を引きます。'
         '残った数が、今日発注する数です。消費量の計算は要りません。',
         '置き場所： ②発注数を決める → ③発注書',
         '数字は自動で出ます。そのまま入れて構いません。'),
    ]
    for left, right, title, color, fill, body, where, who in steps:
        mrg(f'{left}6:{right}6', title,
            font=Font(name=FONT, size=12, bold=True, color=color), fill=fill,
            align='center', border=True)
        mrg(f'{left}7:{right}9', body,
            font=Font(name=FONT, size=10, color=C_TEXT2), border=True)
        mrg(f'{left}10:{right}10', where, font=F_BOLD, border=True)
        mrg(f'{left}11:{right}11', who, font=F_NOTE, border=True)
    for left in ('M', 'Z'):
        mrg(f'{left}6:{get_column_letter(ws[left + "6"].column + 1)}11', '➡',
            font=Font(name=FONT, size=20, bold=True, color=C_TEXT3), align='center')
    ws.row_dimensions[6].height = 22
    for r in (7, 8, 9):
        ws.row_dimensions[r].height = 15

    # ================= ① 定数の置き場所マップ =================
    section(13, '1. 定数はどこにあるか',
            '人が決める数字は、この2枚にしかありません。'
            '「今の値」はあなたの劇場の現在値です（設定シートで選んだ劇場）。')
    heads = [('B', 'I', '決めておくこと'), ('J', 'O', '置き場所'),
             ('P', 'U', 'ブロック'), ('V', 'AA', '今の値'),
             ('AB', 'AG', '見直す頻度'), ('AH', 'AL', '誰が決める')]
    for left, right, text in heads:
        mrg(f'{left}15:{right}15', text, font=F_HEAD, fill=FILL_HEAD,
            align='center', border=True)

    freeze = f'IFERROR(INDEX({cycle_col(CYCLE_DAYS)},MATCH("冷凍品",{cycle_col(CYCLE_NAME)},0)),"")'
    pack = f'IFERROR(INDEX({cycle_col(CYCLE_DAYS)},MATCH("常温包材",{cycle_col(CYCLE_NAME)},0)),"")'
    const_rows = [
        ('劇場の規模と、ドリンクの提供方式', S_CONST, '① 劇場マスタ（B列）',
         f'=IF({Q_SET}!$C$10="","未設定",{Q_SET}!$C$10&" ／ "&{Q_SET}!$C$11)',
         None, '開店時・方式の変更時', '本社・支配人'),
        ('規模ごとの標準来場者数', S_CONST, '② 規模区分（N列）',
         f'={Q_SET}!$C$14', FMT_INT, '年1回', '本社'),
        ('安全在庫日数（何日ぶん余裕を持つか）', S_CONST, '② 規模区分（N列）',
         f'={Q_SET}!$C$12', FMT_DAYS, '年1回', '本社'),
        ('月ごとの繁閑（季節係数）', S_CONST, '③ 月別係数（T列）',
         f'={Q_SET}!$C$15', '0%', '年1回', '本社'),
        ('GW・大型作品公開などの特別期間', S_CONST, '④ 特別期間（W列）',
         f'=COUNTIF({special_col("E")},"?*")&" 件 登録済み"', None, '都度', '支配人・副支配人'),
        ('冷蔵庫・ストッカーの容量（常温／冷蔵／冷凍）', S_CONST, '⑤ 保管設備（110行目）',
         f'=IF({Q_SET}!$C$4="","",'
         f'TEXT(SUMIFS({equip_col("G")},{equip_col("B")},{Q_SET}!$C$4,'
         f'{equip_col("C")},"常温"),"#,##0")&" ／ "&'
         f'TEXT(SUMIFS({equip_col("G")},{equip_col("B")},{Q_SET}!$C$4,'
         f'{equip_col("C")},"冷蔵"),"#,##0")&" ／ "&'
         f'TEXT(SUMIFS({equip_col("G")},{equip_col("B")},{Q_SET}!$C$4,'
         f'{equip_col("C")},"冷凍"),"#,##0")&" ケース")',
         None, '設備の入れ替え時', '支配人・副支配人'),
        ('発注できる曜日（グループごと）', S_CONST, '⑥ 発注サイクル（AB列）',
         f'="冷凍 "&{freeze}&"日 ／ 包材 "&{pack}&"日"',
         None, '便が変わったとき', '本社'),
        ('入数・単価・リードタイム・最低ロット', S_ITEM, '（商品ごと）',
         f'=SUMPRODUCT(--({item_col("B")}<>""))&" 商品 登録済み"', None,
         '商品の入れ替え時', '社員'),
        ('商品の保管区分と発注グループ', S_ITEM, '（商品ごと）',
         f'=SUMPRODUCT(--({item_col("K")}<>""))&" 商品 設定済み"', None,
         '商品の入れ替え時', '社員'),
        ('期間の平均動員／日', S_PLAN, '実績から自動',
         f'={Q_SET}!$C$24', FMT_INT, '実績を入れれば自動', '副支配人'),
        ('曜日ごとの動員の差', S_PLAN, '実績から自動',
         f'=IF({Q_SET}!$C$21=0,"実績が未入力",{Q_SET}!$C$21&" 日ぶんの実績から算出")',
         None, '実績を入れれば自動', '—'),
        ('商品ごとの最大保管数（棚に置ける上限）', S_ITEM, '（商品ごと）',
         f'=SUMPRODUCT(--({col(S_ITEM, get_column_letter(PI_FIRST_COL + len(PROFILES)), MASTER_FIRST, MASTER_LAST)}<>""))'
         f'&" 商品 設定済み"', None, '棚割りを変えたとき', '副支配人'),
        ('PI値（購買率）', '入力不要', '自動で測定',
         f'=IF({Q_SET}!$C$9="PI値予測","商品マスタの手入力値を使用中","実測値を自動計算")',
         None, '入れなくてよい', '—'),
    ]
    for i, (what, where, block, value, fmt, freq, who) in enumerate(const_rows):
        r = 16 + i
        mrg(f'B{r}:I{r}', what, font=Font(name=FONT, size=10, color=C_INK), border=True)
        mrg(f'J{r}:O{r}', where, font=Font(name=FONT, size=10, bold=True, color=C_INK),
            fill=FILL_INPUT, align='center', border=True)
        mrg(f'P{r}:U{r}', block, font=Font(name=FONT, size=10, color=C_TEXT2),
            align='center', border=True)
        mrg(f'V{r}:AA{r}', value, font=F_BOLD, fmt=fmt, fill=FILL_AUTO,
            align='center', border=True)
        mrg(f'AB{r}:AG{r}', freq, font=F_NOTE, align='center', border=True)
        mrg(f'AH{r}:AL{r}', who, font=F_NOTE, align='center', border=True)
        ws.row_dimensions[r].height = 22
    mrg('B27:AL27', '※ 黄色いシート名が置き場所です。定数はこの2枚以外にはありません。'
                    '②発注数を決めるシートや発注計算シートに、直接数字を書き込む必要はありません。',
        font=F_NOTE)

    # ================= ② 定数 → 在庫 → 発注数 =================
    section(29, '2. その定数から、発注数が出るまで',
            '1商品を選ぶと、その商品の実際の数字で順番に追えます。'
            '空欄のままなら、いま一番多く発注が必要な商品を自動で表示します。')
    # 既定の定数方式では発注点を使わない。どの行が効いているかを先に言っておかないと、
    # 表を上から読んだ人が「発注点から出している」と誤解する。
    mrg('B32:AL32',
        '発注数は「定数 − 今の在庫 − 発注済み」だけで決まります。'
        '下の動員・消費・発注点の行は、定数をいくつにするか決めるときの目安です'
        '（発注数の計算には使っていません）。',
        font=Font(name=FONT, size=10, bold=True, color=C_WARN))
    mrg('B31:E31', '追いかける商品', font=F_BOLD, fill=FILL_HEAD, align='center', border=True)
    pick = mrg('F31:Q31', None, font=F_INPUT, fill=FILL_INPUT, align='left', border=True)
    add_validation(ws, f'={calc_name}', 'F31')
    mrg('R31:AL31', '▼から選べます。選び直すと下の数字がすべて入れ替わります。', font=F_NOTE)

    # 参照用のヘルパー。右に逃がしておく（印刷範囲の外）
    ws['AN31'] = (f'=IF($F$31="",IFERROR(INDEX({calc_name},'
                  f'MATCH(MAX({calc_qty}),{calc_qty},0)),""),$F$31)')
    ws['AN32'] = f'=IFERROR(MATCH($AN$31,{calc_name},0),0)'
    for r in (31, 32):
        ws[f'AN{r}'].font = F_NOTE

    def val(letter, fmt=FMT_INT):
        return (f'=IF($AN$32=0,"",INDEX({col(S_CALC, letter, CALC_FIRST, CALC_LAST)},$AN$32))',
                fmt)

    group = f'INDEX({col(S_CALC, "AF", CALC_FIRST, CALC_LAST)},$AN$32)'
    basis = (f'=IF({Q_SET}!$C$9="実績消費",'
             f'"①今日の在庫 と 月1回_前回の在庫 の差だけ。動員とは連動しません",'
             f'IF({Q_SET}!$C$9="PI値予測","商品マスタ の PI値 × 動員予測",'
             f'"実測PI値 × これからの動員 ÷ 100'
             f'（＝実績消費 ×「これからの動員 ÷ 期間の平均動員」）"))')
    heads2 = [('B', 'C', ''), ('D', 'N', '見ている数字'), ('O', 'S', '値'),
              ('T', 'AL', 'どこから来た数字か')]
    for left, right, text in heads2:
        mrg(f'{left}33:{right}33', text, font=F_HEAD, fill=FILL_HEAD,
            align='center', border=True)

    trace = [
        ('動員', '期間の平均動員／日（実績）', (f'={Q_SET}!$C$24', FMT_INT),
         f'=IF({Q_SET}!$C$21>0,"随時_動員を入れる の実績 "&{Q_SET}!$C$21&"日ぶんの平均",'
         f'"実績が未入力。設定シートの手入力値か標準来場者数を使用中")', False),
        ('動員', 'これからの動員／日（次の発注日まで）', val('AK', FMT_INT),
         '"随時_動員を入れる の予測 × 季節係数。予測がない日はその曜日の実績平均"', False),
        ('定数', '実測PI値（来場者100人あたり何個減ったか）', val('AI', FMT_PI),
         '"期間消費数 ÷（期間の平均動員 ÷ 100）。自動で測るので入力は不要です"', False),
        ('定数', '1日にどれだけ減るか', val('S', FMT_DEC), basis, False),
        ('定数', '届くまでの日数（リードタイム）', val('I', FMT_DAYS),
         '"商品マスタ の L/T日数"', False),
        ('定数', '次の発注までの日数（発注サイクル）', val('AG', FMT_DAYS),
         f'=IF($AN$32=0,"","定数マスタ ⑥発注サイクル　グループ："&{group})', False),
        ('定数', '切らさないための余裕（安全在庫）', val('U'),
         f'="定数マスタ ②規模区分 の安全在庫日数 "&TEXT({Q_SET}!$C$12,"0.0")&"日ぶん"', False),
        ('＝', '持っておくべき数（発注点）', val('V'),
         '"1日に減る数 ×（届くまで＋次の発注まで）＋ 安全在庫"', True),
        ('定数', '棚に置ける上限（最大保管数）', val('AT'),
         f'=IF($AN$32=0,"",IF(INDEX({col(S_CALC, "AT", CALC_FIRST, CALC_LAST)},$AN$32)="",'
         f'"商品マスタに未設定。上限なしで計算しています",'
         f'"商品マスタ の 最大保管数(ケース) × 入数。目安は "&'
         f'INDEX({col(S_CALC, "AU", CALC_FIRST, CALC_LAST)},$AN$32)&" ケース"))', False),
        ('在庫', 'いま在庫にある数', val('K'),
         '=IF($AN$32=0,"",IF(INDEX(' + col(S_CALC, "AM", CALC_FIRST, CALC_LAST) +
         ',$AN$32)<>"","実棚を入れる で数えた数（CSVの理論値 "'
         '&TEXT(INDEX(' + col(S_CALC, "AL", CALC_FIRST, CALC_LAST) +
         ',$AN$32),"#,##0")&" を上書き）",'
         '"①今日の在庫を貼る（CSVの理論値）"))', False),
        ('在庫', '発注済みで、まだ届いていない数', val('W'),
         '"発注管理 の 発注済み"', False),
        ('＝', '足りない数', val('X'),
         f'=IF($AN$32=0,"",IF(INDEX({col(S_CALC, "AT", CALC_FIRST, CALC_LAST)},$AN$32)="",'
         f'"定数が未設定なので、暫定で発注点から計算しています",'
         f'"定数 − いまの在庫 − 発注済み"))', True),
        ('発注', '1ケースに何個入っているか（入数）', val('G'),
         '"商品マスタ の 入数"', False),
        ('＝', '発注する数（ケース）', val('Y'),
         '"足りない数 ÷ 入数 を切り上げ。最低ロットの単位に丸めます"', True),
    ]
    trace_top = 34
    for i, (tag, label, (formula, fmt), src, total) in enumerate(trace):
        r = trace_top + i
        color = {'定数': C_BUTTER, '在庫': C_TEXT2, '発注': C_CRIT,
                 '動員': C_WARN, '＝': C_OK}[tag]
        mrg(f'B{r}:C{r}', tag, font=Font(name=FONT, size=9, bold=True, color=color),
            align='center', border=True)
        mrg(f'D{r}:N{r}', label,
            font=F_BOLD if total else Font(name=FONT, size=10, color=C_INK),
            fill=FILL_ACCENT if total else None, border=True)
        v = mrg(f'O{r}:S{r}', formula, fmt=fmt,
                font=Font(name=FONT, size=13, bold=True, color=C_INK if total else C_TEXT2),
                fill=FILL_ACCENT if total else FILL_AUTO, align='center', border=True)
        v.number_format = fmt
        mrg(f'T{r}:AL{r}', src if src.startswith('=') else f'={src}',
            font=F_NOTE, border=True)
        ws.row_dimensions[r].height = 21

    # ================= 帯の図 =================
    # trace の行数が変わっても崩れないよう、以降の行位置はここから逆算する
    bar_sec = trace_top + len(trace) + 1
    bar_row = bar_sec + 3
    legend_row = bar_row + 1
    faq_sec = legend_row + 2
    section(bar_sec, '同じことを図にすると',
            '帯全体が「持っておくべき数（発注点）」です。'
            '緑がいまの在庫、黄が発注済みで届く分、赤が足りていない分＝今日発注する分です。')
    ws[f'AN{bar_sec + 0}'] = f'=IF($AN$32=0,0,N(INDEX({col(S_CALC, "V", CALC_FIRST, CALC_LAST)},$AN$32)))'
    ws[f'AN{bar_sec + 1}'] = f'=IF($AN$32=0,0,MAX(0,N(INDEX({col(S_CALC, "K", CALC_FIRST, CALC_LAST)},$AN$32))))'
    ws[f'AN{bar_sec + 2}'] = f'=IF($AN$32=0,0,N(INDEX({col(S_CALC, "W", CALC_FIRST, CALC_LAST)},$AN$32)))'
    ws[f'AN{bar_sec + 3}'] = (f'=IF($AN${bar_sec}<=0,36,'
                                   f'MIN(36,ROUND($AN${bar_sec + 1}/$AN${bar_sec}*36,0)))')
    ws[f'AN{bar_sec + 4}'] = (f'=IF($AN${bar_sec}<=0,36,MIN(36,ROUND('
                                   f'($AN${bar_sec + 1}+$AN${bar_sec + 2})/$AN${bar_sec}*36,0)))')
    for r in range(bar_sec, bar_sec + 5):
        ws[f'AN{r}'].font = F_NOTE

    for i in range(36):
        cell = ws.cell(bar_row, 2 + i)
        cell.value = i + 1
        cell.font = Font(name=FONT, size=1, color='FFFFFFFF')
        cell.fill = FILL_TRACK
        cell.border = Border(right=Side(style='thin', color='FFFFFFFF'))
    ws.row_dimensions[bar_row].height = 26
    bar = f'B{bar_row}:{get_column_letter(37)}{bar_row}'
    for limit, color in ((f'$AN${bar_sec + 3}', C_OK),
                         (f'$AN${bar_sec + 4}', C_BUTTER), ('36', C_CRIT)):
        ws.conditional_formatting.add(bar, FormulaRule(
            formula=[f'B{bar_row}<={limit}'], stopIfTrue=True,
            fill=PatternFill('solid', fgColor=color)))

    legend = [('B', 'H', C_OK, 'いまの在庫'), ('J', 'P', C_BUTTER, '発注済みで届く分'),
              ('R', 'X', C_CRIT, '今日発注する分'), ('Z', 'AL', None,
              '帯の右端 ＝ 持っておくべき数（発注点）')]
    for left, right, color, text in legend:
        if color:
            ws[f'{left}{legend_row}'].fill = PatternFill('solid', fgColor=color)
            ws[f'{left}{legend_row}'].border = BORDER
            nxt = get_column_letter(ws[f'{left}{legend_row}'].column + 1)
            mrg(f'{nxt}{legend_row}:{right}{legend_row}', text, font=F_NOTE)
        else:
            mrg(f'{left}{legend_row}:{right}{legend_row}', text, font=F_NOTE, align='right')

    # ================= 補足 =================
    section(faq_sec, 'よくある質問')
    faq = [
        ('購買率が分からないのに、その日の動員と減り方が連動するのですか？',
         '連動します。購買率は「人が知っている必要」も「入力する必要」もありません。'
         '前回と今日の在庫の差＝期間消費数と、その期間の平均動員が分かれば、'
         '商品ごとの実測PI値（来場者100人あたり何個減ったか）は割り算で出ます。'
         'あとは、これからの動員がその平均の何倍かを掛けるだけです。'
         '式にすると 1日に減る数 ＝ 実績消費 ×（これからの動員 ÷ 期間の平均動員）。'
         'GWで動員が1.3倍なら、消費も1.3倍で見積もります。'),
        ('動員に合わせるなら、毎日在庫を数えるか、在庫一覧を毎日貼る必要があるのでは？',
         'どちらも要りません。発注数は「定数 − 在庫 − 発注済み」で決まるので、'
         '消費量そのものを使いません。'
         '消費量は定数を決めるときの目安と残り日数の表示に使うだけで、'
         'そこでも「その日の消費 ÷ その日の動員」ではなく'
         '「期間の合計 ÷ 期間の合計」で出すため、日ごとの対応づけは不要です。'
         'さらに、発注数に効くのは購買率の絶対値ではなく'
         '「これからの動員 ÷ 期間の平均動員」の比だけです。'
         '動員の数字を実績も予測もまとめて2倍にしても、発注数は1品目も変わりません。'
         '在庫CSVは発注する日に貼りますが、それは発注そのものに必要な作業なので、'
         '動員に連動させるために増える手間はありません。'),
        ('日々の動員実績は入力しなくていいのですか？',
         '計算だけなら無くても動きますが、入れたほうが確実に良くなります。'
         '3つ効きます。(1) 期間の平均動員が自動で出るので、手入力の打ち間違いが消える。'
         '(2) 曜日ごとの平均が出るので、予測を入れていない日の見積もりが'
         '「劇場の標準来場者数」ではなく「その曜日の実績平均」になる。'
         '映画館は曜日で動員が3倍近く違うので、ここがいちばん効きます。'
         '(3) 予測と実績のズレが見えるので、読みの癖が分かります。'
         '毎日でなくても、月1回まとめて貼るだけで構いません。'),
        ('発注できる曜日が商品によって違いますが？',
         '定数マスタ ⑥発注サイクルにグループごとの曜日を入れてあります。'
         '冷凍品とドリンクシロップは週3回（月水金）、常温包材は木曜の1回だけです。'
         '週1回の商品は次の便まで7日あるため、そのぶん多めに持つ計算になります。'),
        ('現物の在庫は確認しなくていいのですか？',
         '確認してください。CSVの在庫は理論値（仕入 − 減算）でしかなく、'
         'ロス・破損・減算マスタのズレ・納品の未計上で現物とずれます。'
         'そのまま発注すると、ズレがそのまま発注数のズレになります。'
         'ただし全商品を毎日数えるのは無理なので、'
         '「実棚を入れる」シートが今日数えるべき商品に★を付けます。'
         '理論値がマイナスの商品、今日発注する商品、在庫金額の大きい商品の3つだけです。'),
        ('1ヶ月前・2ヶ月前の在庫CSVは何のために貼るのですか？',
         '取扱商品が変わっていないかを確認するためです。消費量を出すためではありません。'
         '半年ごとに商品が入れ替わるので、当日のCSVだけ見ていると'
         '「先月まで扱っていた商品」が発注リストから静かに消えます。'
         '3枚を突き合わせれば、新商品・終売・復活が今月からか先月からかまで分かります。'
         'それぞれのシートの上に、その結果を出しています。'),
        ('消費量の計算は使わないのですか？',
         '発注数の計算には使いません。'
         '定数 − 今の在庫 − 発注済み、それだけです。'
         '消費量（前回在庫＋納品−当日在庫）は、発注管理に入荷が'
         'すべて記録されていることが前提で、記録が抜けると消費が過小に出ます。'
         '定数運用ならその影響を受けません。'
         '消費量は、定数をいくつにするか決めるときの目安（商品マスタの「目安」列）と、'
         '「②発注数を決める」の残り日数と帯の表示にだけ使います。'),
        ('在庫がマイナスになっています',
         '納品がまだシステムに入っていない状態です。'
         'そのまま発注すると、実際には棚にあるものを大量に頼むことになります。'
         '現物を数えて「実棚を入れる」に入れてください。'
         '在庫システム側も直して、次回のCSVから正しくなるようにしてください。'),
    ]
    r = faq_sec + 1
    for q, a in faq:
        # 回答の長さで行数を変える。2行に詰め込むと文字が重なって読めなくなる
        lines = max(2, -(-len(a) // 58))
        mrg(f'B{r}:AL{r}', 'Q. ' + q, font=F_BOLD)
        mrg(f'B{r + 1}:AL{r + lines}', 'A. ' + a,
            font=Font(name=FONT, size=10, color=C_TEXT2))
        for k in range(1, lines + 1):
            ws.row_dimensions[r + k].height = 15
        r += lines + 2

    # ================= 入力の手間 =================
    # 「結局どれだけ入れないといけないのか」が分からないと現場は不安になる。
    # 頻度と、入れなかったときに何が起きるかを一覧にしておく。
    effort_sec = r + 1
    section(effort_sec, '入力の手間はこれだけです',
            '毎日やるのは①だけです。ほかは月1回か、必要なときだけです。')
    eff_heads = [('B', 'J', '入れるもの'), ('K', 'R', '入れる場所'),
                 ('S', 'X', '頻度'), ('Y', 'AL', '入れないとどうなるか')]
    for left, right, text in eff_heads:
        mrg(f'{left}{effort_sec + 2}:{right}{effort_sec + 2}', text,
            font=F_HEAD, fill=FILL_HEAD, align='center', border=True)
    efforts = [
        ('当日の在庫CSV', S_CUR, '発注する日ごと',
         '発注できません。これは発注そのものに必要な作業です', True),
        ('1ヶ月前・2ヶ月前の在庫CSV', S_MID + ' ／ ' + S_PRV, '月1回',
         '取扱商品の入れ替わりに気づけず、発注リストから静かに漏れます', False),
        ('動員の実績', S_PLAN, '月1回まとめてで可',
         '曜日ごとの差が出ません。平日を多めに見積もってしまいます', False),
        ('動員の予測', S_PLAN, '読みが変わる日だけ',
         'その曜日の実績平均で代用されます。ふだんはそれで足ります', False),
        ('実在庫（数えた数）', S_COUNT, '★が付いた商品だけ',
         '理論値のズレがそのまま発注数のズレになります', False),
        ('商品ごとの最大保管数', S_ITEM, '棚割りを変えたとき',
         '暫定で消費量から計算します。定数を入れるまでの仮の数字です', False),
        ('購買率（PI値）', '－', '入力不要',
         '在庫の減りと動員から自動で出ます', False),
    ]
    for i, (what, where, freq, effect, daily) in enumerate(efforts):
        rr = effort_sec + 3 + i
        mrg(f'B{rr}:J{rr}', what,
            font=F_BOLD if daily else Font(name=FONT, size=10, color=C_INK),
            fill=FILL_ACCENT if daily else None, border=True)
        mrg(f'K{rr}:R{rr}', where, font=Font(name=FONT, size=10, color=C_TEXT2),
            fill=FILL_INPUT if where != '－' else None, align='center', border=True)
        mrg(f'S{rr}:X{rr}', freq,
            font=Font(name=FONT, size=10, bold=True,
                      color=C_CRIT if daily else C_TEXT2),
            align='center', border=True)
        mrg(f'Y{rr}:AL{rr}', effect, font=F_NOTE, border=True)
        ws.row_dimensions[rr].height = 20
    last = effort_sec + 3 + len(efforts)
    mrg(f'B{last + 1}:AL{last + 1}',
        '※ 購買率は「期間の消費合計 ÷ 期間の動員合計」で出しています。'
        'その日の消費とその日の動員を突き合わせているわけではないので、毎日数える必要はありません。',
        font=F_NOTE)

    ws.print_area = f'B2:AL{last + 2}'
    ws.page_setup.orientation = 'landscape'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    return ws


def build_count_sheet(wb):
    """
    実棚を入れるシート。

    在庫CSVの値は理論値（仕入 − 減算）でしかなく、ロス・破損・減算マスタのズレ・
    納品の未計上で現物とずれる。理論値だけで発注すると、そのズレがそのまま
    発注数のズレになるため、数えた数で上書きできるようにしている。

    ただし全商品を毎日数えるのは無理なので、「今日数えるべき商品」を出す。
    理論値がマイナスの商品、今日発注する商品、在庫金額の大きい商品の3つに絞れば、
    数えるのは十数品目で済む。
    """
    ws = wb.create_sheet(S_COUNT)
    sheet_title(ws, '実棚を入れる',
                'CSVの在庫は理論値です。数えた数を入れると、その商品はそちらを使って発注量を計算します。')

    calc_stock = col(S_CALC, 'AL', CALC_FIRST, CALC_LAST)      # 理論値在庫（CSV）
    calc_amount = col(S_CALC, 'AA', CALC_FIRST, CALC_LAST)     # 在庫金額
    now_col = col(S_CONST, CYCLE_LBL_NOW, CYCLE_FIRST, CYCLE_LAST)

    reason_col = col(S_COUNT, 'K', COUNT_FIRST, COUNT_LAST)
    status_col = col(S_COUNT, 'J', COUNT_FIRST, COUNT_LAST)
    gap_col = col(S_COUNT, 'I', COUNT_FIRST, COUNT_LAST)

    # ---- まとめ ----
    # 差はプラスとマイナスで意味が違う。合算すると相殺されて何も見えなくなるため分ける。
    #   マイナス＝理論値より現物が少ない → ロス・廃棄・減算のズレ
    #   プラス　＝理論値より現物が多い　 → 納品の未計上が濃厚
    summary = [
        ('B', 'D', '今日数える商品', f'=COUNTIF({reason_col},"★*")', FMT_INT,
         '「数える理由」に★が付いた商品だけ数えれば足ります'),
        ('F', 'H', 'カウント済み', f'=COUNTIF({status_col},"採用中")', FMT_INT,
         '今日のカウントとして使われている商品数'),
        ('J', 'L', '不足（ロス）', f'=SUMIF({gap_col},"<0")', FMT_YEN,
         '理論値より現物が少ない分。廃棄・破損・減算のズレ'),
        ('N', 'P', '超過（未計上）', f'=SUMIF({gap_col},">0")', FMT_YEN,
         '理論値より現物が多い分。納品がシステム未計上の疑い'),
    ]
    for left, right, label, formula, fmt, note in summary:
        block_cell(ws, f'{left}4:{right}4', label, F_HEAD, FILL_HEAD)
        v = block_cell(ws, f'{left}5:{right}5', formula,
                       Font(name=FONT, size=20, bold=True, color=C_INK), FILL_AUTO)
        v.number_format = fmt
        block_cell(ws, f'{left}6:{right}6', note, F_NOTE, FILL_AUTO)
    ws.row_dimensions[5].height = 28

    # 「金額が大きい」の線引き。在庫金額の上位10番目を1セルで出しておく
    # （行ごとに LARGE を回すと400×400になるため）
    ws['R4'] = '金額上位10位のライン'
    ws['R5'] = f'=IFERROR(LARGE({calc_amount},10),0)'
    ws['R4'].font = F_NOTE
    ws['R5'].font = F_NOTE
    ws['R5'].number_format = FMT_YEN

    headers = ['No', '商品名', '発注グループ', '理論値在庫', '数えた数', 'カウント日',
               '差（数えた数−理論値）', '差の金額', '状態', '数える理由']
    kinds = ['自動', '自動', '自動', '自動', '入力', '入力', '自動', '自動', '自動', '自動']
    put_headers(ws, COUNT_FIRST - 1, headers, kinds)

    for r in range(COUNT_FIRST, COUNT_LAST + 1):
        c = CALC_FIRST + (r - COUNT_FIRST)
        ws[f'B{r}'] = r - COUNT_FIRST + 1
        ws[f'C{r}'] = f'=IF({Q_CALC}!$C{c}="","",{Q_CALC}!$D{c})'
        ws[f'D{r}'] = f'=IF({Q_CALC}!$C{c}="","",{Q_CALC}!$AF{c})'
        ws[f'E{r}'] = f'=IF({Q_CALC}!$C{c}="","",{Q_CALC}!$AL{c})'
        ws[f'H{r}'] = f'=IF(OR($C{r}="",$F{r}=""),"",N($F{r})-N($E{r}))'
        ws[f'I{r}'] = (f'=IF(OR($H{r}="",$J{r}<>"採用中"),"",'
                       f'N($H{r})*N({Q_CALC}!$H{c}))')
        # 数えた数を使うかどうかの判定。古いカウントを黙って使うと事故になる
        ws[f'J{r}'] = (f'=IF($C{r}="","",IF($F{r}="","未入力",'
                       f'IF($G{r}="","カウント日を入れてください",'
                       f'IF($G{r}<{Q_SET}!$C$6,"日付が古い（使いません）","採用中"))))')
        # 数える理由。全部数えるのは無理なので、優先順位を言葉で出す
        ws[f'K{r}'] = (
            f'=IF($C{r}="","",'
            f'IF(N($E{r})<0,"★ 理論値がマイナス",'
            f'IF(AND(COUNTIF({now_col},$D{r})>0,N({Q_CALC}!$Y{c})>0),"★ 今日発注する商品",'
            f'IF(AND(N({Q_CALC}!$AA{c})>0,N({Q_CALC}!$AA{c})>=N($R$5)),"金額が大きい","－"))))')
        style_row(ws, r, 'BCDEHIJK', fill=FILL_AUTO)
        style_row(ws, r, 'FG', fill=FILL_INPUT)
        ws[f'F{r}'].font = F_INPUT
        ws[f'G{r}'].font = F_INPUT
        for letter in 'EFH':
            ws[f'{letter}{r}'].number_format = FMT_INT
        ws[f'G{r}'].number_format = FMT_DATE
        ws[f'I{r}'].number_format = FMT_YEN
        for letter in 'JK':
            ws[f'{letter}{r}'].font = Font(name=FONT, size=9, color=C_TEXT2)

    ws.conditional_formatting.add(
        f'K{COUNT_FIRST}:K{COUNT_LAST}',
        FormulaRule(formula=[f'LEFT($K{COUNT_FIRST},1)="★"'], stopIfTrue=True,
                    fill=FILL_ALERT,
                    font=Font(name=FONT, size=9, bold=True, color=C_CRIT)))
    ws.conditional_formatting.add(
        f'J{COUNT_FIRST}:J{COUNT_LAST}',
        CellIsRule(operator='equal', formula=['"採用中"'],
                   fill=PatternFill('solid', bgColor=C_OK_SOFT),
                   font=Font(name=FONT, size=9, bold=True, color=C_OK)))
    ws.conditional_formatting.add(
        f'J{COUNT_FIRST}:J{COUNT_LAST}',
        CellIsRule(operator='equal', formula=['"日付が古い（使いません）"'],
                   fill=PatternFill('solid', bgColor=C_WARN_SOFT),
                   font=Font(name=FONT, size=9, bold=True, color=C_WARN)))
    ws.conditional_formatting.add(
        f'H{COUNT_FIRST}:H{COUNT_LAST}',
        CellIsRule(operator='lessThan', formula=['0'], fill=FILL_ALERT))

    for letter, w in zip('BCDEFGHIJKLMNOPQR',
                         [5, 30, 15, 12, 11, 13, 18, 13, 20, 20, 4, 4, 10, 8, 8, 3, 20]):
        ws.column_dimensions[letter].width = w
    ws.freeze_panes = 'C10'

    note = ws[f'B{COUNT_LAST + 2}']
    note.value = ('※ カウント日が当日基準日より前の行は使いません。'
                  '前に数えた数がそのまま残って発注量を狂わせるのを防ぐためです。'
                  'CSVを貼り替えたら、数え直した商品だけカウント日を今日に更新してください。')
    note.font = F_NOTE
    return ws


def build_const_master(wb):
    """
    定数マスタ。劇場・規模区分・月別係数・特別期間を1枚に集めている。
    マスタが分かれるほど「どこを直せばいいか」が分からなくなるため、
    人が更新する定数はすべてこのシートに置く。
    """
    ws = wb.create_sheet(S_CONST)
    # ブロック見出しは行3に置く。行4は put_headers が「入力／自動」を書く行なので使わない。
    sheet_title(ws, '定数マスタ')
    # 6ブロックを横に並べているため、どこに何があるかを先頭に書いておく
    ws['E2'] = ('ブロックの場所　① 劇場マスタ＝B列　／　② 規模区分＝N列　／　③ 月別係数＝T列'
                '　／　④ 特別期間＝W列　／　⑤ 保管設備＝110行目　／　⑥ 発注サイクル＝AB列')
    ws['E2'].font = F_NOTE

    # ---------- ① 劇場マスタ（B〜J列） ----------
    ws['B3'] = '① 劇場マスタ　劇場ごとの区分と、標準値を上書きしたいときの欄'
    ws['B3'].font = F_BOLD
    headers = ['劇場コード', '劇場名', '規模区分', 'ドリンク提供方式',
               '来場者数/日(個別)', '安全在庫日数(個別)', '発注サイクル(個別)',
               '常温 収容ケース', '冷蔵 収容ケース', '冷凍 収容ケース', '備考',
               'CSV照合']
    kinds = ['入力', '入力', '選択', '選択', '任意', '任意', '任意',
             '自動', '自動', '自動', '入力', '自動']
    put_headers(ws, 5, headers, kinds)
    for r in range(THEATER_FIRST, THEATER_LAST + 1):
        # 収容ケース数は⑤保管設備の合計。設備を足せば自動で増える
        for letter, kind in zip('IJK', STORAGE_KINDS):
            ws[f'{letter}{r}'] = (f'=IF($B{r}="","",SUMIFS({equip_col("G")},'
                                  f'{equip_col("B")},$B{r},{equip_col("C")},"{kind}"))')
            ws[f'{letter}{r}'].number_format = FMT_INT
        # 劇場コードが在庫CSVに実在するかの照合。コードの取り違えは
        # 「その劇場だけ何も出ない」という形で出るため、一覧で潰せるようにする
        ws[f'M{r}'] = (
            f'=IF($B{r}="","",IF(COUNTIF({csv_col(S_CUR, "B")},$B{r})=0,"✖ CSVに無し",'
            f'IF(INDEX({csv_col(S_CUR, "C")},MATCH($B{r},{csv_col(S_CUR, "B")},0))=$C{r},'
            f'"✔ 一致","△ CSVでは "&INDEX({csv_col(S_CUR, "C")},'
            f'MATCH($B{r},{csv_col(S_CUR, "B")},0)))))')
        style_row(ws, r, 'BCDEFGHL', fill=FILL_INPUT)
        style_row(ws, r, 'IJKM', fill=FILL_AUTO)
        ws[f'B{r}'].number_format = '@'
        ws[f'F{r}'].number_format = FMT_INT
        ws[f'M{r}'].font = Font(name=FONT, size=9, color=C_TEXT2)

    # 公式サイトで確認できた劇場を初期登録しておく。
    # 規模区分は暫定値（スクリーン数からの機械的な割り当て）なので、現場で直す前提。
    for i, (code3, name, pref, screens, seats, note) in enumerate(THEATER_SEED):
        r = THEATER_FIRST + i
        if r > THEATER_LAST:
            break
        ws[f'B{r}'] = code3 + '1'
        ws[f'C{r}'] = name
        ws[f'D{r}'] = seed_size(screens, seats)
        ws[f'E{r}'] = DRINK_STYLES[0]
        spec = '／'.join(x for x in (
            pref,
            f'{screens}スクリーン' if screens else '',
            f'約{seats:,}席' if seats else '',
            f'公式コード{code3}', note) if x)
        ws[f'L{r}'] = spec

    ws[f'B{THEATER_LAST + 2}'] = (
        '※ 劇場コードは「公式サイトの3桁コード＋1」で暫定登録しています'
        '（新宿 076→0761 が在庫CSVと一致するため）。'
        'CSVを貼るとM列で全劇場ぶん照合できます。「✖」「△」の行は自社コードに直してください。')
    ws[f'B{THEATER_LAST + 2}'].font = F_NOTE
    ws[f'B{THEATER_LAST + 3}'] = (
        '※ 規模区分（大／中／小）はスクリーン数からの暫定値です。'
        '実際の動員・売上に合わせて直してください。ドリンク提供方式も既定は「1杯売り」です。')
    ws[f'B{THEATER_LAST + 3}'].font = F_NOTE
    add_validation(ws, '"' + ','.join(SIZES) + '"', f'D{THEATER_FIRST}:D{THEATER_LAST}')
    add_validation(ws, '"' + ','.join(DRINK_STYLES) + '"', f'E{THEATER_FIRST}:E{THEATER_LAST}')

    # ---------- ② 規模区分の標準値（L〜P列） ----------
    ws['N3'] = '② 規模区分の標準値　劇場マスタの個別欄が空のときはこの値を使う'
    ws['N3'].font = F_BOLD
    put_headers(ws, 5, ['規模区分', '来場者数/日', '安全在庫日数', '発注サイクル', '備考'],
                ['固定', '入力', '入力', '入力', '入力'], start_col=14)
    for i, row in enumerate([
        ('大規模', 3000, 4, 7, 'スクリーン数が多く来場者の多い劇場'),
        ('中規模', 1500, 3, 7, '標準的な規模の劇場'),
        ('小規模', 700, 3, 14, '来場者が少なく発注頻度を落とす劇場'),
    ]):
        r = 6 + i
        for j, v in enumerate(row):
            ws.cell(r, 14 + j, v)
        style_row(ws, r, 'NOPQR', fill=FILL_INPUT)
        ws[f'N{r}'].fill = FILL_AUTO
        ws[f'O{r}'].number_format = FMT_INT

    # ---------- ③ 月別の季節係数（R〜S列） ----------
    ws['T3'] = '③ 月別の係数'
    ws['T3'].font = F_BOLD
    put_headers(ws, 5, ['月', '係数'], ['固定', '入力'], start_col=20)
    for i in range(12):
        r = 6 + i
        ws.cell(r, 20, i + 1)
        ws.cell(r, 21, 1.0)
        style_row(ws, r, 'TU', fill=FILL_INPUT)
        ws[f'T{r}'].fill = FILL_AUTO
        ws[f'U{r}'].number_format = '0%'

    # ---------- ④ 特別期間（U〜X列） ----------
    ws['W3'] = '④ 特別期間　この期間は月別の係数より優先される'
    ws['W3'].font = F_BOLD
    put_headers(ws, 5, ['期間名', '開始日', '終了日', '係数'],
                ['入力', '入力', '入力', '入力'], start_col=23)
    for i in range(20):
        r = 6 + i
        style_row(ws, r, 'WXYZ', fill=FILL_INPUT)
        ws[f'X{r}'].number_format = FMT_DATE
        ws[f'Y{r}'].number_format = FMT_DATE
        ws[f'Z{r}'].number_format = '0%'
    for i, (name, rate) in enumerate([('年末年始', 1.3), ('大型作品公開週', 1.5), ('閑散期', 0.85)]):
        ws.cell(6 + i, 23, name)
        ws.cell(6 + i, 26, rate)
    ws['W27'] = '※ 開始日・終了日が空欄の行は無視されます。期間が重なるときは係数が最大の行を使います。'
    ws['W27'].font = F_NOTE
    # ---------- ⑤ 保管設備（下段） ----------
    ws[f'B{EQUIP_FIRST - 3}'] = ('⑤ 保管設備　冷蔵庫・ストッカーなど。'
                                 'ここの合計が①の「収容ケース数」になり、発注量の上限判定に使われます')
    ws[f'B{EQUIP_FIRST - 3}'].font = F_BOLD
    put_headers(ws, EQUIP_FIRST - 1,
                ['劇場コード', '保管区分', '設備名', '台数', '1台あたり収容ケース', '小計', '備考'],
                ['入力', '選択', '入力', '入力', '入力', '自動', '入力'])
    for r in range(EQUIP_FIRST, EQUIP_LAST + 1):
        ws[f'G{r}'] = f'=IF($B{r}="","",N($E{r})*N($F{r}))'
        style_row(ws, r, 'BCDEFH', fill=FILL_INPUT)
        style_row(ws, r, 'G', fill=FILL_AUTO)
        ws[f'B{r}'].number_format = '@'
        for letter in 'EFG':
            ws[f'{letter}{r}'].number_format = FMT_INT
    add_validation(ws, '"' + ','.join(STORAGE_KINDS) + '"', f'C{EQUIP_FIRST}:C{EQUIP_LAST}')
    ws[f'B{EQUIP_LAST + 2}'] = ('※ 1台あたりの収容ケース数は、実際に積める数を入れてください。'
                               'メーカー公称の容量ではなく、運用上の目安で構いません。')
    ws[f'B{EQUIP_LAST + 2}'].font = F_NOTE

    # ---------- ⑥ 発注サイクル（商品グループ別） ----------
    d_first, d_last = WEEKDAY_COLS[0], WEEKDAY_COLS[-1]
    ws[f'{CYCLE_NAME}3'] = ('⑥ 発注サイクル　商品グループごとの発注曜日。'
                            'ここが発注点の計算に効きます')
    ws[f'{CYCLE_NAME}3'].font = F_BOLD
    put_headers(ws, CYCLE_FIRST - 1,
                ['発注グループ', '月', '火', '水', '木', '金', '土', '日',
                 '週何回', 'サイクル日数', '今日', '備考'],
                ['入力', '入力', '入力', '入力', '入力', '入力', '入力', '入力',
                 '自動', '自動', '自動', '入力'], start_col=28)
    for r in range(CYCLE_FIRST, CYCLE_LAST + 1):
        ws[f'{CYCLE_CNT}{r}'] = (f'=IF(${CYCLE_NAME}{r}="","",'
                                 f'COUNTIF(${d_first}{r}:${d_last}{r},"○"))')
        ws[f'{CYCLE_DAYS}{r}'] = (f'=IF(OR(${CYCLE_NAME}{r}="",N(${CYCLE_CNT}{r})=0),"",'
                                  f'ROUND(7/N(${CYCLE_CNT}{r}),1))')
        # 基準日の曜日に○が付いているグループが、今日の発注対象
        ws[f'{CYCLE_TODAY}{r}'] = (
            f'=IF(OR(${CYCLE_NAME}{r}="",{Q_SET}!$C$6=""),"",'
            f'IF(INDEX(${d_first}{r}:${d_last}{r},WEEKDAY({Q_SET}!$C$6,2))="○","● 今日",""))')
        style_row(ws, r, [CYCLE_NAME] + WEEKDAY_COLS + [CYCLE_NOTE], fill=FILL_INPUT)
        style_row(ws, r, [CYCLE_CNT, CYCLE_DAYS, CYCLE_TODAY], fill=FILL_AUTO)
        for letter in WEEKDAY_COLS:
            ws[f'{letter}{r}'].alignment = Alignment(horizontal='center')
        ws[f'{CYCLE_CNT}{r}'].number_format = FMT_INT
        ws[f'{CYCLE_DAYS}{r}'].number_format = FMT_DEC
        ws[f'{CYCLE_TODAY}{r}'].font = Font(name=FONT, size=10, bold=True, color=C_CRIT)
        ws[f'{CYCLE_TODAY}{r}'].alignment = Alignment(horizontal='center')
    for i, (name, _label, days, note) in enumerate(ORDER_GROUPS):
        r = CYCLE_FIRST + i
        ws[f'{CYCLE_NAME}{r}'] = name
        ws[f'{CYCLE_NOTE}{r}'] = note
        for day in days:
            ws.cell(r, CYCLE_DAY1 + '月火水木金土日'.index(day), '○')
    ws[f'{CYCLE_NAME}{CYCLE_LAST + 2}'] = (
        '※ 発注する曜日に「○」を入れてください。'
        'サイクル日数（＝7÷週の回数）が発注点の計算に使われます。')
    ws[f'{CYCLE_NAME}{CYCLE_LAST + 2}'].font = F_NOTE

    # 「次にこのグループを発注できるのはいつか」を求める補助列（AN〜AX）。
    # 基準日から0〜6日先を順に見て、○の付いた最初の曜日までの日数を取る。
    # ②発注数を決めるシートの見出し帯が、ここの結果を文章にして表示する。
    for r in range(CYCLE_FIRST, CYCLE_LAST + 1):
        for k in range(7):
            letter = get_column_letter(CYCLE_OFF1 + k)
            ws[f'{letter}{r}'] = (
                f'=IF(OR(${CYCLE_NAME}{r}="",{Q_SET}!$C$6=""),99,'
                f'IF(INDEX(${d_first}{r}:${d_last}{r},'
                f'WEEKDAY({Q_SET}!$C$6+{k},2))="○",{k},99))')
        off_first = get_column_letter(CYCLE_OFF1)
        off_last = get_column_letter(CYCLE_OFF1 + 6)
        ws[f'{CYCLE_WAIT}{r}'] = (f'=IF(${CYCLE_NAME}{r}="","",'
                                  f'MIN(${off_first}{r}:${off_last}{r}))')
        ws[f'{CYCLE_NEXT}{r}'] = (f'=IF(OR(${CYCLE_NAME}{r}="",N(${CYCLE_WAIT}{r})>=99),"",'
                                  f'{Q_SET}!$C$6+${CYCLE_WAIT}{r})')
        ws[f'{CYCLE_NEXT}{r}'].number_format = FMT_DATE
        # 今日発注できるグループ／待ちのグループを、そのまま文章に使える形で持つ
        ws[f'{CYCLE_LBL_NOW}{r}'] = (f'=IF(AND(${CYCLE_NAME}{r}<>"",N(${CYCLE_WAIT}{r})=0),'
                                     f'${CYCLE_NAME}{r},"")')
        ws[f'{CYCLE_LBL_WAIT}{r}'] = (
            f'=IF(OR(${CYCLE_NAME}{r}="",N(${CYCLE_WAIT}{r})=0,N(${CYCLE_WAIT}{r})>=99),"",'
            f'${CYCLE_NAME}{r}&" は "&TEXT(${CYCLE_NEXT}{r},"m/d")&"（"&'
            f'{weekday_jp(f"${CYCLE_NEXT}{r}")}&"）から")')
        for letter in (CYCLE_WAIT, CYCLE_NEXT, CYCLE_LBL_NOW, CYCLE_LBL_WAIT):
            ws[f'{letter}{r}'].font = F_NOTE
    put_headers(ws, CYCLE_FIRST - 1, ['次の発注日まで', '次の発注日', '今日', '待ち'],
                ['自動'] * 4, start_col=CYCLE_OFF1 + 7)

    # 注記は劇場ブロック（6〜105行）の外に置く。中に書くと劇場を1つ潰してしまう
    ws[f'B{THEATER_LAST + 4}'] = (
        '人が決める数字はすべてこの1枚にあります。黄色のセルだけ触ってください。'
        'マスタを分けるほど「どこを直せばいいか」が分からなくなるため、1枚に集めています。')
    ws[f'B{THEATER_LAST + 4}'].font = F_NOTE

    # 列幅はシート単位なので、ブロックごとに列範囲を分けたうえで幅を決める。
    #   B〜M ＝ ①劇場マスタ（B〜Hは⑤保管設備と共用）／N〜R ＝ ②規模区分
    #   T〜U ＝ ③月別係数／W〜Z ＝ ④特別期間／AB〜AM ＝ ⑥発注サイクル
    widths = {'B': 12, 'C': 22, 'D': 11, 'E': 17, 'F': 18, 'G': 18, 'H': 17,
              'I': 15, 'J': 15, 'K': 15, 'L': 38, 'M': 20,
              'N': 10, 'O': 13, 'P': 13, 'Q': 13, 'R': 30,
              'T': 5, 'U': 8, 'W': 22, 'X': 13, 'Y': 13, 'Z': 9,
              'AB': 17, 'AJ': 7, 'AK': 12, 'AL': 6, 'AM': 22}
    for c, w in widths.items():
        ws.column_dimensions[c].width = w
    for letter in WEEKDAY_COLS:
        ws.column_dimensions[letter].width = 3.5
    ws.freeze_panes = 'B6'
    return ws


def build_item_master(wb, items):
    ws = wb.create_sheet(S_ITEM)
    sheet_title(ws, '商品マスタ',
                'PI値は来場者100人あたりの消費数です。取り扱いのないプロファイルは空欄のままで構いません。')
    headers = ['商品コード', '商品名', '商品分類名', '支払先名', '入数', '税抜単価',
               'L/T日数', '最低ロット', '保管区分', '発注グループ'] + \
              [f'PI値 {p}' for p in PROFILES] + \
              ['最大保管数(ケース)', '目安(自動)', '備考']
    kinds = ['入力'] * 10 + ['入力'] * len(PROFILES) + ['入力', '自動', '入力']
    put_headers(ws, 5, headers, kinds)

    last_col_idx = 2 + len(headers) - 1
    letters = [get_column_letter(i) for i in range(2, last_col_idx + 1)]
    for r in range(MASTER_FIRST, MASTER_LAST + 1):
        style_row(ws, r, letters, fill=FILL_INPUT)
        ws[f'B{r}'].number_format = '@'
        ws[f'F{r}'].number_format = FMT_INT
        ws[f'G{r}'].number_format = FMT_YEN
        for i in range(len(PROFILES)):
            ws.cell(r, PI_FIRST_COL + i).number_format = FMT_DEC

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
        ws[f'J{r}'] = '常温'
        ws[f'K{r}'] = ORDER_GROUPS[2][0]

    # 「この商品は最低何ケース置ければ回るか」を発注計算から引いてくる。
    # 132商品ぶんの定数を手で決めるのは無理なので、目安を隣に出しておく。
    const_col = get_column_letter(PI_FIRST_COL + len(PROFILES))          # 最大保管数
    guide_col = get_column_letter(PI_FIRST_COL + len(PROFILES) + 1)      # 目安
    for r in range(MASTER_FIRST, MASTER_LAST + 1):
        ws[f'{guide_col}{r}'] = (
            f'=IF($B{r}="","",IFERROR(INDEX('
            f'{col(S_CALC, "AU", CALC_FIRST, CALC_LAST)},'
            f'MATCH($B{r},{col(S_CALC, "C", CALC_FIRST, CALC_LAST)},0)),""))')
        ws[f'{guide_col}{r}'].fill = FILL_AUTO
        ws[f'{guide_col}{r}'].font = F_NOTE
        ws[f'{guide_col}{r}'].number_format = FMT_INT
        ws[f'{const_col}{r}'].number_format = FMT_INT

    widths = [16, 30, 16, 22, 8, 12, 9, 11, 10, 16] + [14] * len(PROFILES) + [16, 11, 24]
    for c, w in zip(letters, widths):
        ws.column_dimensions[c].width = w
    add_validation(ws, '"' + ','.join(STORAGE_KINDS) + '"',
                   f'J{MASTER_FIRST}:J{MASTER_LAST}')
    add_validation(ws, f'={cycle_col(CYCLE_NAME)}', f'K{MASTER_FIRST}:K{MASTER_LAST}')
    ws.freeze_panes = 'D6'
    return ws


def build_csv_sheet(wb, name, others, label, note='', paste_hint=''):
    """
    CSVを貼るシート。1〜4行目は案内と取込判定に使い、5行目が見出し、6行目からがデータ。
    貼り付け範囲に案内文を置くと、貼った瞬間に消えてしまうため上に逃がしている。
    """
    ws = wb.create_sheet(name)
    ws['B2'] = label
    ws['B2'].font = F_TITLE
    ws['B3'] = note
    ws['B3'].font = F_NOTE
    ws['B4'] = paste_hint
    ws['B4'].font = Font(name=FONT, size=11, bold=True, color=C_INK)

    # 貼れているかどうかを、いちばん目立つところに出す
    hit = f'COUNTIF($B${CSV_FIRST}:$B${CSV_LAST},{Q_SET}!$C$4)'
    ws['H2'] = (f'=IF({hit}=0,"✖ まだ貼れていません",ّ"✔ "&{hit}&"行 読み込めました")'
                .replace('ّ', ''))
    ws['H2'].font = Font(name=FONT, size=16, bold=True, color=C_INK)
    ws['H2'].alignment = Alignment(horizontal='left', vertical='center')
    ws['H3'] = (f'=IF({hit}=0,"CSVの見出し行を除いたデータを A{CSV_FIRST} セルに貼り付けてください。'
                f'貼ったのに0のままなら、劇場コードが数値になっています。",'
                f'"対象劇場：" & {Q_SET}!$C$4 & "　" & {Q_SET}!$C$5)')
    ws['H3'].font = F_NOTE
    ws.row_dimensions[2].height = 24
    ws.conditional_formatting.add('H2', FormulaRule(
        formula=[f'{hit}=0'],
        fill=PatternFill('solid', bgColor=C_CRIT_SOFT),
        font=Font(name=FONT, size=16, bold=True, color=C_CRIT), stopIfTrue=True))
    ws.conditional_formatting.add('H2', FormulaRule(
        formula=['TRUE'],
        fill=PatternFill('solid', bgColor=C_OK_SOFT),
        font=Font(name=FONT, size=16, bold=True, color=C_OK), stopIfTrue=True))

    # 過去のCSVシートは「取扱商品が変わっていないか」を見るためのもの。
    # 何のために貼っているのかが分かるよう、結果をこのシートの上に出す。
    if others:
        change = col(S_CALC, 'AS', CALC_FIRST, CALC_LAST)
        ws['B2'] = label
        checks = [
            ('N', '新商品（今月から）'), ('P', '新商品（先月から）'),
            ('R', '終売（今月から）'), ('T', '終売（先月から）'),
            ('V', '復活（先月は無し）'),
        ]
        ws['N1'] = '取扱の変化（3枚を突き合わせた結果）'
        ws['N1'].font = Font(name=FONT, size=10, bold=True, color=C_INK)
        for letter, name in checks:
            nxt = get_column_letter(ws[f'{letter}1'].column + 1)
            c1 = ws[f'{letter}2']
            c1.value = f'=COUNTIF({change},"{name}")'
            c1.font = Font(name=FONT, size=18, bold=True, color=C_INK)
            c1.alignment = Alignment(horizontal='center')
            c2 = ws[f'{letter}3']
            c2.value = name
            c2.font = F_NOTE
            c2.alignment = Alignment(horizontal='center')
            for cc in (c1, c2):
                cc.fill = FILL_HEAD
            ws.column_dimensions[letter].width = 6
            ws.column_dimensions[nxt].width = 12
        ws['B4'] = paste_hint

    for i, h in enumerate(CSV_HEADERS):
        cell = ws.cell(CSV_HEAD, 1 + i, h)
        cell.font = F_HEAD
        cell.fill = FILL_HEAD
        cell.border = BORDER
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
    helper = ws.cell(CSV_HEAD, 21, '連番(自動)')
    helper.font = F_HEAD
    helper.fill = FILL_ACCENT
    helper.border = BORDER

    for r in range(CSV_FIRST, CSV_LAST + 1):
        if not others:
            # 当日CSV: 対象劇場の行に通し番号を振る
            ws[f'{CSV_HELPER}{r}'] = f'=IF($B{r}={Q_SET}!$C$4,N({CSV_HELPER}{r-1})+1,N({CSV_HELPER}{r-1}))'
        else:
            # 過去のCSV: 対象劇場の行のうち、より新しいどのCSVにも無い商品だけに番号を振る。
            # これで3枚を順に足すと、重複なしの和集合＝取扱商品の全体になる。
            dup = '+'.join(f'COUNTIFS({csv_col(o, "B")},$B{r},{csv_col(o, "I")},$I{r})'
                           for o in others)
            ws[f'{CSV_HELPER}{r}'] = (
                f'=IF($B{r}<>{Q_SET}!$C$4,N({CSV_HELPER}{r-1}),'
                f'IF({dup}>0,N({CSV_HELPER}{r-1}),N({CSV_HELPER}{r-1})+1))')
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
    ws.freeze_panes = f'A{CSV_FIRST}'
    ws.sheet_view.showGridLines = False
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
    mid_date = cur_date.replace(S_CUR, S_MID)

    def const(individual, std_col):
        """劇場個別の値があればそれを、無ければ規模マスタの標準値を返す。"""
        return (f'=IF($C$4="","",IF(IFERROR(INDEX({theater_col(individual)},'
                f'MATCH($C$4,{theater_col("B")},0)),"")<>"",'
                f'INDEX({theater_col(individual)},MATCH($C$4,{theater_col("B")},0)),'
                f'IFERROR(INDEX({size_col(std_col)},MATCH($C$10,{size_col("B")},0)),"")))')

    special = (f'SUMPRODUCT(MAX(({special_col("F")}<=$C$6)*'
               f'({special_col("G")}>=$C$6)*{special_col("H")}))')

    rows = [
        ('対象劇場コード', None, '発注計算・ダッシュボードの対象になる劇場です。劇場マスタから選択します。', FILL_INPUT, '@'),
        ('劇場名', f'=IFERROR(INDEX({theater_col("C")},MATCH($C$4,{theater_col("B")},0)),"")',
         '劇場マスタから自動取得します。', FILL_AUTO, None),
        ('当日基準日', f'={cur_date}', '②在庫CSV_当日の対象日付から自動取得します。', FILL_AUTO, FMT_DATE),
        ('2ヶ月前の基準日', f'={prv_date}', '①在庫CSV_2ヶ月前の対象日付から自動取得します。', FILL_AUTO, FMT_DATE),
        ('期間日数', '=IF(OR($C$6="",$C$7=""),"",$C$6-$C$7)', '実績消費/日の算出に使う日数です。', FILL_AUTO, FMT_INT),
        ('算出基準', None,
         '動員連動＝実績消費に動員の増減を掛ける（推奨）、実績消費＝在庫2時点の差のみ、'
         'PI値予測＝手入力のPI値、最大値＝いちばん大きい値。',
         FILL_INPUT, None),
        ('規模区分', f'=IFERROR(INDEX({theater_col("D")},MATCH($C$4,{theater_col("B")},0)),"")',
         '劇場マスタから自動取得します。', FILL_AUTO, None),
        ('PIプロファイル', f'=IFERROR(INDEX({theater_col("E")},MATCH($C$4,{theater_col("B")},0)),"")',
         'ドリンクの提供方式です。商品マスタのどのPI値列を使うかが決まります。', FILL_AUTO, None),
        ('安全在庫日数', const('G', 'D'), '劇場個別の設定があればそちらが優先されます。', FILL_AUTO, FMT_DEC),
        ('発注サイクル日数', const('H', 'E'), '次回発注までの日数。発注点の算出に使います。', FILL_AUTO, FMT_INT),
        ('標準来場者数/日', const('F', 'C'), '動員計画が空欄の日に使われる既定値です。', FILL_AUTO, FMT_INT),
        ('季節係数', f'=IF($C$6="",1,IF({special}=0,'
                     f'IFERROR(INDEX({month_col("C")},MONTH($C$6)),1),{special}))',
         '当日基準日が特別期間に該当すればその係数、なければ月別係数です。', FILL_AUTO, '0%'),
        ('発注担当者', None, '発注書・発注管理に記録する担当者名です。', FILL_INPUT, None),
        ('当日CSV該当行数', f'=COUNTIF({csv_col(S_CUR, "B")},$C$4)',
         '0のままなら劇場コードが一致していません（先頭ゼロが落ちていないか確認）。', FILL_AUTO, FMT_INT),
        ('2ヶ月前CSV該当行数', f'=COUNTIF({csv_col(S_PRV, "B")},$C$4)',
         '前回CSVを貼っていない場合は0です（実績消費は算出できません）。', FILL_AUTO, FMT_INT),
        ('取扱商品数', f'=MAX({csv_col(S_CUR, CSV_HELPER)})+MAX({csv_col(S_MID, CSV_HELPER)})'
                       f'+MAX({csv_col(S_PRV, CSV_HELPER)})',
         '3枚のCSVの和集合です。発注計算シートの行数と一致します。', FILL_AUTO, FMT_INT),
        ('動員計画の入力日数', f'=COUNT({col(S_PLAN, "D", PLAN_FIRST, PLAN_LAST)})',
         '0でも動きます（標準来場者数を使用）。読みが変わる日だけ入れてください。', FILL_AUTO, FMT_INT),
        ('動員実績の入力日数', f'=COUNT({col(S_PLAN, "E", PLAN_FIRST, PLAN_LAST)})',
         '実績を入れるほど、曜日ごとの傾向と実測PI値が正確になります。', FILL_AUTO, FMT_INT),
        ('期間の平均動員/日（実績から）',
         f'=IFERROR(AVERAGEIFS({col(S_PLAN, "E", PLAN_FIRST, PLAN_LAST)},'
         f'{col(S_PLAN, "B", PLAN_FIRST, PLAN_LAST)},">="&$C$7,'
         f'{col(S_PLAN, "B", PLAN_FIRST, PLAN_LAST)},"<="&$C$6),"")',
         '動員シートの実績から自動で出ます。実績を入れていなければ空欄です。',
         FILL_AUTO, FMT_INT),
        ('期間の平均動員/日（手入力）', None,
         '実績を入れていないときの代わりです。上の自動値があればそちらが優先されます。',
         FILL_INPUT, FMT_INT),
        ('採用する期間平均動員', '=IF(N($C$22)>0,$C$22,IF(N($C$23)>0,$C$23,N($C$14)))',
         '動員連動の分母です。これからの動員がこの値の何倍かで、消費量を伸縮させます。',
         FILL_AUTO, FMT_INT),
        # ここから下は1ヶ月前CSVを足したときに増えた項目。
        # 既存の $C$4〜$C$24 を動かすと参照が全部ずれるため、末尾に足している。
        ('1ヶ月前の基準日', f'={mid_date}',
         '月1回_1ヶ月前の在庫を貼る の対象日付から自動取得します。', FILL_AUTO, FMT_DATE),
        ('1ヶ月前CSV該当行数', f'=COUNTIF({csv_col(S_MID, "B")},$C$4)',
         '3枚そろうと、期中に入れ替わった商品を取りこぼしません。', FILL_AUTO, FMT_INT),
        ('直近1ヶ月の日数', '=IF(OR($C$6="",$C$25=""),"",$C$6-$C$25)',
         '1ヶ月前の基準日から当日基準日までの日数です。', FILL_AUTO, FMT_INT),
        ('その前1ヶ月の日数', '=IF(OR($C$25="",$C$7=""),"",$C$25-$C$7)',
         '2ヶ月前の基準日から1ヶ月前の基準日までの日数です。', FILL_AUTO, FMT_INT),
        ('実績消費の集計期間', None,
         '2ヶ月＝ならして安定させる（推奨）。直近1ヶ月＝最近の動きに素早く追随する。',
         FILL_INPUT, None),
        # 空行の期間納品数は "" で、Excelでは ""＞0 が TRUE になるため
        # SUMPRODUCT ではなく COUNTIF を使う（文字列は数値条件に一致しない）
        ('入荷記録のある商品数',
         f'=COUNTIF({col(S_CALC, "M", CALC_FIRST, CALC_LAST)},">0")',
         '発注管理に入荷が記録されている商品数です。'
         '消費量（前回在庫＋納品−当日在庫）は、この記録が揃っている商品でしか正しく出ません。',
         FILL_AUTO, FMT_INT),
        ('最大保管数の設定済み商品数',
         f'=SUMPRODUCT(--({col(S_ITEM, get_column_letter(PI_FIRST_COL + len(PROFILES)), MASTER_FIRST, MASTER_LAST)}<>""))',
         '商品マスタの「最大保管数(ケース)」です。空欄の商品は上限なしで計算します。',
         FILL_AUTO, FMT_INT),
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
    ws['C9'] = BASIS_OPTIONS[0]
    ws['C29'] = SPAN_OPTIONS[0]
    add_validation(ws, f'={theater_col("B")}', 'C4')
    add_validation(ws, '"' + ','.join(BASIS_OPTIONS) + '"', 'C9')
    add_validation(ws, '"' + ','.join(SPAN_OPTIONS) + '"', 'C29')

    ws['B34'] = '※ 在庫CSVを貼り替えたら、この画面の該当行数と基準日が想定どおりか必ず確認してください。'
    ws['B35'].font = F_NOTE
    ws['B35'] = ('※ 動員シートに実績を入れておくと、期間の平均動員が自動で出て、'
                 '商品ごとの実測PI値（来場者100人あたり何個売れたか）も正確になります。'
                 '購買率を人が管理する必要はありません。')
    ws['B35'].font = F_NOTE
    return ws


def build_calc(wb):
    ws = wb.create_sheet(S_CALC)
    sheet_title(ws, '発注計算',
                '設定シートで選んだ劇場の商品別計算です。行は在庫CSV（当日＋前回）から自動生成されます。')
    headers = ['No', '商品コード', '商品名', '商品分類名', '支払先名', '入数', '税抜単価',
               'L/T日数', '最低ロット', '当日在庫(採用)', '前回理論在庫', '期間納品数', '期間消費数',
               '実績消費/日', 'PI値', 'PI予測消費/日', '規定数', '採用消費/日', '在庫日数',
               '安全在庫数量', '発注点', '発注残', '推奨発注数(バラ)', '推奨発注数(ケース)',
               '発注要否', '在庫金額', '状態', '推奨発注金額', '保管区分', '在庫ケース数',
               '発注グループ', '発注サイクル', '図解用', '実測PI値', '動員連動消費/日', '動員/日(カバー期間)',
               '理論値在庫(CSV)', '実棚(有効)', '1ヶ月前在庫', '直近1ヶ月消費数',
               '直近1ヶ月消費/日', '前1ヶ月消費/日', '消費の変化', '取扱の変化',
               '最大保管数(個)', '定数の目安(ケース)']
    kinds = ['自動'] * len(headers)
    put_headers(ws, 5, headers, kinds)

    n_cur = f'MAX({csv_col(S_CUR, CSV_HELPER)})'
    n_mid = f'MAX({csv_col(S_MID, CSV_HELPER)})'
    letters = [get_column_letter(i) for i in range(2, 2 + len(headers))]

    for r in range(CALC_FIRST, CALC_LAST + 1):
        no = r - CALC_FIRST + 1
        ws[f'B{r}'] = no
        # 商品コード: 当日CSVの対象劇場行 → 足りない分を前回CSVの未登場商品で埋める
        # 当日 → 1ヶ月前（当日に無いもの）→ 2ヶ月前（どちらにも無いもの）の順に積む
        ws[f'C{r}'] = (
            f'=IFERROR(IF($B{r}<={n_cur},'
            f'INDEX({csv_col(S_CUR, "I")},MATCH($B{r},{csv_col(S_CUR, CSV_HELPER)},0)),'
            f'IF($B{r}<={n_cur}+{n_mid},'
            f'INDEX({csv_col(S_MID, "I")},MATCH($B{r}-{n_cur},{csv_col(S_MID, CSV_HELPER)},0)),'
            f'INDEX({csv_col(S_PRV, "I")},'
            f'MATCH($B{r}-{n_cur}-{n_mid},{csv_col(S_PRV, CSV_HELPER)},0)))),"")')

        def from_master_or_csv(master_col, csv_letter):
            return (f'=IF($C{r}="","",IFERROR(INDEX({item_col(master_col)},'
                    f'MATCH($C{r},{item_col("B")},0)),'
                    f'IFERROR(INDEX({csv_col(S_CUR, csv_letter)},'
                    f'MATCH($C{r},{csv_col(S_CUR, "I")},0)),'
                    f'IFERROR(INDEX({csv_col(S_MID, csv_letter)},'
                    f'MATCH($C{r},{csv_col(S_MID, "I")},0)),'
                    f'IFERROR(INDEX({csv_col(S_PRV, csv_letter)},'
                    f'MATCH($C{r},{csv_col(S_PRV, "I")},0)),"")))))')

        ws[f'D{r}'] = from_master_or_csv('C', 'J')     # 商品名
        ws[f'E{r}'] = from_master_or_csv('D', 'H')     # 商品分類名
        ws[f'F{r}'] = from_master_or_csv('E', 'E')     # 支払先名
        ws[f'G{r}'] = from_master_or_csv('F', 'K')     # 入数
        ws[f'H{r}'] = from_master_or_csv('G', 'Q')     # 税抜単価
        ws[f'I{r}'] = (f'=IF($C{r}="","",IFERROR(INDEX({item_col("H")},'
                       f'MATCH($C{r},{item_col("B")},0)),3))')
        ws[f'J{r}'] = (f'=IF($C{r}="","",MAX(1,IFERROR(INDEX({item_col("I")},'
                       f'MATCH($C{r},{item_col("B")},0)),1)))')

        # AL＝CSVの理論値、AM＝有効な実棚。K（＝以降すべてが使う在庫）は実棚を優先する。
        # 現物を数えたならそれが正しいので、理論値より実棚を採る。
        cr = COUNT_FIRST + (r - CALC_FIRST)
        ws[f'AL{r}'] = (f'=IF($C{r}="","",SUMIFS({csv_col(S_CUR, "N")},'
                        f'{csv_col(S_CUR, "B")},{Q_SET}!$C$4,{csv_col(S_CUR, "I")},$C{r}))')
        ws[f'AL{r}'].number_format = FMT_INT
        ws[f'AM{r}'] = (f'=IF($C{r}="","",IF({Q_COUNT}!$F{cr}="","",'
                        f'IF({Q_COUNT}!$G{cr}="","",'
                        f'IF({Q_COUNT}!$G{cr}<{Q_SET}!$C$6,"",{Q_COUNT}!$F{cr}))))')
        ws[f'AM{r}'].number_format = FMT_INT
        ws[f'K{r}'] = f'=IF($C{r}="","",IF($AM{r}<>"",N($AM{r}),N($AL{r})))'
        ws[f'L{r}'] = (f'=IF($C{r}="","",SUMIFS({csv_col(S_PRV, "N")},'
                       f'{csv_col(S_PRV, "B")},{Q_SET}!$C$4,{csv_col(S_PRV, "I")},$C{r}))')
        # 期間納品数: 発注管理の入荷実績（前回基準日〜当日基準日）
        ws[f'M{r}'] = (f'=IF($C{r}="","",IF(OR({Q_SET}!$C$6="",{Q_SET}!$C$7=""),0,'
                       f'SUMIFS({po_col("S")},{po_col("C")},{Q_SET}!$C$4,{po_col("F")},$C{r},'
                       f'{po_col("R")},">="&{Q_SET}!$C$7,{po_col("R")},"<="&{Q_SET}!$C$6)))')
        # 前回CSVが貼られていなければ期間消費は出せない（前回在庫が0のままだと嘘の値になる）
        ws[f'N{r}'] = (f'=IF(OR($C{r}="",{Q_SET}!$C$18=0),"",$L{r}+$M{r}-$K{r})')
        ws[f'O{r}'] = (f'=IF({Q_SET}!$C$29="直近1ヶ月",N($AP{r}),'
                       f'IF(OR($N{r}="",{Q_SET}!$C$8="",{Q_SET}!$C$8<=0),"",'
                       f'MAX(0,$N{r})/{Q_SET}!$C$8))')
        # PI値: 商品を縦、PIプロファイルを横に検索する。
        # 列見出しは「PI値 <プロファイル名>」なので、検索値も同じ形に組み立てる。
        pi_first = get_column_letter(PI_FIRST_COL)
        pi_last = get_column_letter(PI_FIRST_COL + len(PROFILES) - 1)
        ws[f'P{r}'] = (f'=IF($C{r}="","",IFERROR(INDEX('
                       f'{col(S_ITEM, pi_first, MASTER_FIRST, MASTER_LAST)}:'
                       f'${pi_last}${MASTER_LAST},'
                       f'MATCH($C{r},{item_col("B")},0),'
                       f'MATCH("PI値 "&{Q_SET}!$C$11,'
                       f'{Q_ITEM}!${pi_first}$5:${pi_last}$5,0)),""))')
        # カバー期間（リードタイム＋発注サイクル）の動員合計から1日あたりを出す
        cover = f'(N($I{r})+N($AG{r}))'
        plan_sum = (f'SUMIFS({col(S_PLAN, "H", PLAN_FIRST, PLAN_LAST)},'
                    f'{col(S_PLAN, "B", PLAN_FIRST, PLAN_LAST)},">="&{Q_SET}!$C$6,'
                    f'{col(S_PLAN, "B", PLAN_FIRST, PLAN_LAST)},"<="&{Q_SET}!$C$6+{cover}-1)')
        plan_avg = f'IFERROR({plan_sum}/{cover},0)'
        ws[f'Q{r}'] = (f'=IF(OR($C{r}="",$P{r}="",$P{r}=0,{cover}<=0),"",'
                       f'{plan_avg}*$P{r}/100)')
        ws[f'R{r}'] = (f'=IF($C{r}="","",SUMIFS({csv_col(S_CUR, "O")},'
                       f'{csv_col(S_CUR, "B")},{Q_SET}!$C$4,{csv_col(S_CUR, "I")},$C{r}))')
        ws[f'S{r}'] = (f'=IF($C{r}="","",'
                       f'IF({Q_SET}!$C$9="実績消費",N($O{r}),'
                       f'IF({Q_SET}!$C$9="動員連動",IF(N($AJ{r})>0,N($AJ{r}),N($O{r})),'
                       f'IF({Q_SET}!$C$9="PI値予測",N($Q{r}),'
                       f'MAX(N($O{r}),N($Q{r}),N($AJ{r}))))))')
        ws[f'T{r}'] = f'=IF(OR($S{r}="",$S{r}<=0),"",$K{r}/$S{r})'
        ws[f'U{r}'] = f'=IF($S{r}="","",ROUNDUP(N($S{r})*{Q_SET}!$C$12,0))'
        ws[f'V{r}'] = (f'=IF($C{r}="","",'
                       f'ROUNDUP(N($S{r})*(N($I{r})+N($AG{r}))+N($U{r}),0))')
        ws[f'W{r}'] = (f'=IF($C{r}="","",SUMIFS({po_col("L")},{po_col("C")},{Q_SET}!$C$4,'
                       f'{po_col("F")},$C{r},{po_col("Q")},"発注済"))')
        # ---- 最大保管数（定数）。棚に置ける上限であり、補充の目標でもある ----
        # 商品マスタの設定（ケース）＞ 在庫CSVの規定数（個）＞ 未設定 の順に採る
        ws[f'AT{r}'] = (
            f'=IF($C{r}="","",'
            f'IFERROR(IF(INDEX({item_col(get_column_letter(PI_FIRST_COL + len(PROFILES)))},'
            f'MATCH($C{r},{item_col("B")},0))="","",'
            f'INDEX({item_col(get_column_letter(PI_FIRST_COL + len(PROFILES)))},'
            f'MATCH($C{r},{item_col("B")},0))*MAX(1,N($G{r}))),'
            f'IF(N($R{r})>0,N($R{r}),"")))')
        ws[f'AT{r}'].number_format = FMT_INT
        # 定数の目安＝発注点をケース単位に切り上げた数。これ未満だと切らしやすい
        ws[f'AU{r}'] = (f'=IF($C{r}="","",ROUNDUP(N($V{r})/MAX(1,N($G{r})),0))')
        ws[f'AU{r}'].number_format = FMT_INT

        # need＝定数未設定のときの暫定値、room＝定数までの空き（これが発注数）
        need = f'MAX(0,N($V{r})-N($K{r})-N($W{r}))'
        room = f'MAX(0,N($AT{r})-N($K{r})-N($W{r}))'
        # 定数があればそれだけで決まる。無い商品は暫定で消費量から出し、
        # 状態欄に「定数未設定」と出して、定数を入れるべきことが分かるようにする
        ws[f'X{r}'] = f'=IF($C{r}="","",IF($AT{r}="",{need},{room}))'
        ws[f'Y{r}'] = (f'=IF($C{r}="","",IF(N($X{r})<=0,0,'
                       f'CEILING($X{r}/MAX(1,N($G{r})),MAX(1,N($J{r})))))')
        ws[f'Z{r}'] = f'=IF($C{r}="","",IF(N($Y{r})>0,"要発注","正常"))'
        ws[f'AA{r}'] = f'=IF($C{r}="","",N($K{r})*N($H{r}))'
        ws[f'AB{r}'] = (
            f'=IF($C{r}="","",_xlfn.TEXTJOIN(" / ",TRUE,'
            f'IF(COUNTIF({item_col("B")},$C{r})=0,"マスタ未登録",""),'
            f'IF(OR($AS{r}="継続",$AS{r}="－"),"",$AS{r}),'
            f'IF(N($K{r})<0,"在庫マイナス",""),'
            f'IF($AT{r}="","定数未設定（暫定計算）",""),'
            f'IF(AND($AT{r}<>"",N($AT{r})<N($V{r})),"定数が小さい（次の便まで持たない）",""),'
            f'IF(AND({Q_SET}!$C$9="PI値予測",$P{r}=""),"PI値未設定","")))')
        # 集計用（空文字を含まない数値列にしておき、ダッシュボードから合計する）
        ws[f'AC{r}'] = f'=N($Y{r})*N($H{r})'
        ws[f'AC{r}'].number_format = FMT_YEN
        # 保管スペースの判定に使う。ケース単位に丸めた在庫量
        ws[f'AD{r}'] = (f'=IF($C{r}="","",IFERROR(INDEX({item_col("J")},'
                       f'MATCH($C{r},{item_col("B")},0)),"常温"))')
        ws[f'AE{r}'] = f'=IF($C{r}="","",MAX(0,ROUNDUP(N($K{r})/MAX(1,N($G{r})),0)))'
        ws[f'AE{r}'].number_format = FMT_INT
        # 発注サイクルは商品グループごとに違う（冷凍は週3回、包材は週1回など）。
        # グループが未設定の商品だけ、劇場の既定サイクルにフォールバックする
        ws[f'AF{r}'] = (f'=IF($C{r}="","",IFERROR(INDEX({item_col("K")},'
                       f'MATCH($C{r},{item_col("B")},0)),""))')
        ws[f'AG{r}'] = (f'=IF($C{r}="","",IFERROR(INDEX({cycle_col(CYCLE_DAYS)},'
                       f'MATCH($AF{r},{cycle_col(CYCLE_NAME)},0)),N({Q_SET}!$C$13)))')
        ws[f'AG{r}'].number_format = FMT_DEC
        # 「発注のしくみ」シートの既定表示を選ぶためだけの列。
        # 在庫があって発注も要る商品＝図として一番わかりやすい商品を選びたい
        ws[f'AH{r}'] = f'=IF(AND(N($K{r})>0,N($Y{r})>0),N($Y{r}),0)'
        # 実測PI値＝この商品が来場者100人あたり何個減ったか。
        # 購買率を人が入れるのではなく、期間消費と期間の平均動員から測る。
        ws[f'AI{r}'] = (f'=IF(OR($C{r}="",$O{r}="",N({Q_SET}!$C$24)<=0),"",'
                        f'N($O{r})/(N({Q_SET}!$C$24)/100))')
        ws[f'AI{r}'].number_format = FMT_PI
        # 動員連動の消費/日＝実測PI値 × これからの1日あたり動員 ÷ 100。
        # 式を展開すると 実績消費/日 ×（これからの動員 ÷ 期間の平均動員）になる。
        ws[f'AJ{r}'] = (f'=IF(OR($AI{r}="",$AI{r}=0,{cover}<=0),"",'
                        f'{plan_avg}*$AI{r}/100)')
        ws[f'AJ{r}'].number_format = FMT_DEC
        # 「発注のしくみ」で動員の効き方を見せるための、カバー期間の1日あたり動員
        ws[f'AK{r}'] = f'=IF(OR($C{r}="",{cover}<=0),"",{plan_avg})'
        ws[f'AK{r}'].number_format = FMT_INT

        # ---- 1ヶ月前スナップショット。取扱の変化と、消費の急変を見る ----
        ws[f'AN{r}'] = (f'=IF($C{r}="","",SUMIFS({csv_col(S_MID, "N")},'
                        f'{csv_col(S_MID, "B")},{Q_SET}!$C$4,{csv_col(S_MID, "I")},$C{r}))')
        deliv = (f'SUMIFS({po_col("S")},{po_col("C")},{Q_SET}!$C$4,{po_col("F")},$C{r},'
                 f'{po_col("R")},">="&{{a}},{po_col("R")},"<="&{{b}})')
        ws[f'AO{r}'] = (f'=IF(OR($C{r}="",{Q_SET}!$C$26=0),"",'
                        f'$AN{r}+' + deliv.format(a=f'{Q_SET}!$C$25', b=f'{Q_SET}!$C$6') +
                        f'-$K{r})')
        ws[f'AP{r}'] = (f'=IF(OR($AO{r}="",N({Q_SET}!$C$27)<=0),"",'
                        f'MAX(0,$AO{r})/N({Q_SET}!$C$27))')
        ws[f'AQ{r}'] = (f'=IF(OR($C{r}="",{Q_SET}!$C$26=0,{Q_SET}!$C$18=0,'
                        f'N({Q_SET}!$C$28)<=0),"",'
                        f'MAX(0,$L{r}+' + deliv.format(a=f'{Q_SET}!$C$7', b=f'{Q_SET}!$C$25') +
                        f'-$AN{r})/N({Q_SET}!$C$28))')
        ws[f'AR{r}'] = (f'=IF(OR($AP{r}="",$AQ{r}="",N($AQ{r})<=0),"",'
                        f'N($AP{r})/N($AQ{r})-1)')
        # 3枚のどこに載っていたかで、取扱の変化を言葉にする
        in_cur = f'COUNTIFS({csv_col(S_CUR, "B")},{Q_SET}!$C$4,{csv_col(S_CUR, "I")},$C{r})>0'
        in_mid = f'COUNTIFS({csv_col(S_MID, "B")},{Q_SET}!$C$4,{csv_col(S_MID, "I")},$C{r})>0'
        in_prv = f'COUNTIFS({csv_col(S_PRV, "B")},{Q_SET}!$C$4,{csv_col(S_PRV, "I")},$C{r})>0'
        ws[f'AS{r}'] = (
            f'=IF($C{r}="","",'
            f'IF(AND({in_cur},{in_mid},{in_prv}),"継続",'
            f'IF(AND({in_cur},{in_mid}),"新商品（先月から）",'
            f'IF({in_cur},IF({in_prv},"復活（先月は無し）","新商品（今月から）"),'
            f'IF({in_mid},"終売（今月から）",'
            f'IF({in_prv},"終売（先月から）","－"))))))')
        for letter in ('AN', 'AO'):
            ws[f'{letter}{r}'].number_format = FMT_INT
        for letter in ('AP', 'AQ'):
            ws[f'{letter}{r}'].number_format = FMT_DEC
        ws[f'AR{r}'].number_format = '+0%;-0%;0%'

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
              11, 13, 11, 10, 15, 17, 10, 13, 26, 15, 10, 13, 16, 13, 8, 11, 15, 17, 15, 12, 13, 15, 15, 15, 12, 18, 14, 15]
    for c, w in zip(letters, widths):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = 'D6'
    return ws


def block_cell(ws, cell_range, value, font, fill, align='center'):
    """結合セルに値とスタイルを入れる。ダッシュボードと同じ組み方。"""
    ws.merge_cells(cell_range)
    c = ws[cell_range.split(':')[0]]
    c.value = value
    c.font = font
    c.fill = fill
    c.alignment = Alignment(horizontal=align, vertical='center')
    c.border = BORDER
    return c


def build_timetable(wb):
    """
    在庫を上映スケジュール表の形で見せる画面。
    横軸は今日から14日。各商品の帯は在庫が尽きるまでの日数で、
    色が状態（余裕／要発注／今日が期限）、太い縦罫線が発注デッドラインを示す。
    発注数を入れると、伸びた分がバター色で帯の先に足される。
    """
    ws = wb.create_sheet(S_TT)
    ws.sheet_view.showGridLines = False
    ws['B2'] = '発注数を決める'
    ws['B2'].font = F_TITLE

    state_col = col(S_TT, TT_COL['state'], TT_FIRST, TT_LAST)
    amount_col = col(S_TT, TT_COL['amt'], TT_FIRST, TT_LAST)

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
    ws['B7'] = (f'=IF($B$5=0,"すべての商品が次の発注日を越える在庫を持っています。",'
                f'"うち "&COUNTIF({state_col},"今日が期限")&'
                f'"品目は今日発注しないとリードタイムに間に合いません。 "&'
                f'COUNTIF({state_col},"要発注")&"品目は次の発注日までに在庫が切れます。")')
    ws['B7'].font = Font(name=FONT, size=10, color=C_TEXT2)

    # ---- 今日はどのグループを発注する日か ----
    # 冷凍は月水金、包材は木曜だけ、というように便が違う。
    # 「今日そもそも発注できるのか」が分からないと現場が迷うため、いちばん上に出す。
    now_col = col(S_CONST, CYCLE_LBL_NOW, CYCLE_FIRST, CYCLE_LAST)
    wait_col = col(S_CONST, CYCLE_LBL_WAIT, CYCLE_FIRST, CYCLE_LAST)
    joined_now = f'_xlfn.TEXTJOIN("・",TRUE,{now_col})'
    ws.merge_cells('D2:AA2')
    ws['D2'] = (f'=IF({Q_SET}!$C$6="","在庫CSVを貼ると、今日の発注対象がここに出ます。",'
                f'IF({joined_now}="",'
                f'"今日 "&TEXT({Q_SET}!$C$6,"m/d")&"（"&{weekday_jp(f"{Q_SET}!$C$6")}&'
                f'"）は、どのグループも発注日ではありません。",'
                f'"今日 "&TEXT({Q_SET}!$C$6,"m/d")&"（"&{weekday_jp(f"{Q_SET}!$C$6")}&'
                f'"）に発注できるのは　"&{joined_now}))')
    ws['D2'].font = Font(name=FONT, size=12, bold=True, color=C_INK)
    ws['D2'].alignment = Alignment(horizontal='left', vertical='center')
    ws.merge_cells('D3:AA3')
    ws['D3'] = (f'=IF({Q_SET}!$C$6="","",'
                f'IF(_xlfn.TEXTJOIN("　",TRUE,{wait_col})="","",'
                f'"次に発注できる日　"&_xlfn.TEXTJOIN("　／　",TRUE,{wait_col})))')
    ws['D3'].font = F_NOTE
    ws['D3'].alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[2].height = 20

    # 保管スペースの埋まり具合。発注しながら「入りきるか」が見える
    ws['N4'] = '保管スペースの埋まり具合（今の在庫＋発注ぶん ÷ 収容ケース数）'
    ws['N4'].font = F_NOTE
    calc_kind = col(S_CALC, 'AD', CALC_FIRST, CALC_LAST)
    calc_case = col(S_CALC, 'AE', CALC_FIRST, CALC_LAST)
    tt_qty = col(S_TT, TT_COL['qty'], TT_FIRST, TT_LAST)
    for i, (kind, cap_col) in enumerate(zip(STORAGE_KINDS, 'IJK')):
        left = get_column_letter(14 + i * 4)
        right = get_column_letter(17 + i * 4)
        cap = (f'IFERROR(INDEX({theater_col(cap_col)},'
               f'MATCH({Q_SET}!$C$4,{theater_col("B")},0)),0)')
        used = (f'(SUMIF({calc_kind},"{kind}",{calc_case})'
                f'+SUMIF({calc_kind},"{kind}",{tt_qty}))')
        head = block_cell(ws, f'{left}5:{right}5', kind,
                          Font(name=FONT, size=10, bold=True, color=C_TEXT2), FILL_HEAD)
        rate = block_cell(ws, f'{left}6:{right}6',
                          f'=IF({cap}=0,"収容未設定",{used}/{cap})',
                          Font(name=FONT, size=14, bold=True, color=C_INK), FILL_AUTO)
        rate.number_format = '0%'
        block_cell(ws, f'{left}7:{right}7',
                   f'=IF({cap}=0,"定数マスタ⑤に設備を登録してください",'
                   f'{used}&" / "&{cap}&"ケース")',
                   F_NOTE, FILL_AUTO)
        ws.conditional_formatting.add(
            f'{left}6', CellIsRule(operator='greaterThan', formula=['1'],
                                   fill=PatternFill('solid', bgColor=C_CRIT_SOFT),
                                   font=Font(name=FONT, size=14, bold=True, color=C_CRIT)))
        ws.conditional_formatting.add(
            f'{left}6', CellIsRule(operator='greaterThan', formula=['0.85'],
                                   fill=PatternFill('solid', bgColor=C_WARN_SOFT),
                                   font=Font(name=FONT, size=14, bold=True, color=C_WARN)))

    ws.merge_cells('G4:M4')
    ws['G4'] = '発注予定額（税抜）'
    ws['G4'].font = F_NOTE
    ws['G4'].alignment = Alignment(horizontal='right')
    ws.merge_cells('G5:M5')
    ws['G5'] = f'=SUM({amount_col})'
    ws['G5'].font = Font(name=FONT, size=16, bold=True, color=C_INK)
    ws['G5'].number_format = FMT_YEN
    ws['G5'].alignment = Alignment(horizontal='right', vertical='center')

    # ---- 見出し ----
    headers = ['No', '商品名', '発注グループ', '定数', '在庫', '残り日数', 'デッドライン',
               '状態', '発注後日数', '推奨(ケース)', '発注数', '金額']
    kinds = ['自動'] * 9 + ['自動', '入力', '自動']
    put_headers(ws, 9, headers, kinds)

    # 日付の目盛り。行9に日付、行8に日番号（条件付き書式が参照する）
    for d in range(SPAN_DAYS):
        letter = get_column_letter(TT_DAY_COL + d)
        head = ws[f'{letter}9']
        head.value = f'=IF({Q_SET}!$C$6="","",{Q_SET}!$C$6+{d})'
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
        ws[f'C{r}'] = f'=IF({Q_CALC}!$C{c}="","",{Q_CALC}!$D{c})'
        # 分類より「どの便で発注する商品か」のほうが現場の判断に効く
        ws[f'D{r}'] = f'=IF({Q_CALC}!$C{c}="","",{Q_CALC}!$AF{c})'
        # 定数は棚の単位（ケース）で見せる。個数だと現場の感覚と合わない
        ws[f'E{r}'] = (f'=IF({Q_CALC}!$C{c}="","",IF({Q_CALC}!$AT{c}="","未設定",'
                       f'ROUND(N({Q_CALC}!$AT{c})/MAX(1,N({Q_CALC}!$G{c})),0)))')
        ws[f'F{r}'] = (f'=IF({Q_CALC}!$C{c}="","",'
                       f'ROUND(N({Q_CALC}!$K{c})/MAX(1,N({Q_CALC}!$G{c})),1))')
        ws[f'G{r}'] = f'=IF({Q_CALC}!$C{c}="","",{Q_CALC}!$T{c})'
        ws[f'H{r}'] = f'=IF(ISNUMBER($G{r}),$G{r}-N({Q_CALC}!$I{c}),"")'
        ws[f'I{r}'] = (f'=IF({Q_CALC}!$C{c}="","",IF(NOT(ISNUMBER($H{r})),"データなし",'
                       f'IF($H{r}<=0,"今日が期限",'
                       f'IF($H{r}<=N({Q_CALC}!$AG{c}),"要発注","余裕"))))')
        ws[f'J{r}'] = (f'=IF(OR({Q_CALC}!$C{c}="",NOT(ISNUMBER({Q_CALC}!$S{c})),{Q_CALC}!$S{c}<=0),"",'
                       f'({Q_CALC}!$K{c}+N($L{r})*N({Q_CALC}!$G{c}))/{Q_CALC}!$S{c})')
        ws[f'K{r}'] = f'=IF({Q_CALC}!$C{c}="","",{Q_CALC}!$Y{c})'
        ws[f'M{r}'] = f'=IF($L{r}="","",N($L{r})*N({Q_CALC}!$H{c}))'
        # 発注書シートの抽出用。発注数が入っていて、選んだ仕入先に一致する行に番号を振る
        ws[f'AB{r}'] = (f'=IF(OR(N($L{r})<=0,AND({Q_SHEET}!$C$5<>"すべて",'
                        f'{Q_CALC}!$F{c}<>{Q_SHEET}!$C$5)),N(AB{r - 1}),N(AB{r - 1})+1)')
        ws[f'AB{r}'].font = F_NOTE

        style_row(ws, r, 'BCDEFGHIJKLM', fill=FILL_AUTO)
        ws[f'L{r}'].fill = FILL_INPUT
        ws[f'L{r}'].font = F_INPUT
        ws[f'E{r}'].font = F_BOLD
        ws[f'E{r}'].number_format = FMT_INT
        ws[f'F{r}'].number_format = FMT_DEC
        for letter in 'GHJ':
            ws[f'{letter}{r}'].number_format = FMT_DAYS
        ws[f'K{r}'].number_format = FMT_INT
        ws[f'L{r}'].number_format = FMT_INT
        ws[f'M{r}'].number_format = FMT_YEN
        ws[f'I{r}'].alignment = Alignment(horizontal='center')
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
        formula=[f'AND(ISNUMBER($H{TT_FIRST}),$H{TT_FIRST}>0,'
                 f'{day_ref}=ROUNDUP($H{TT_FIRST},0))'],
        border=Border(left=Side(style='medium', color=C_INK)), stopIfTrue=False))
    ws.conditional_formatting.add(band, rule(
        f'AND(ISNUMBER($G{TT_FIRST}),{day_ref}<=$G{TT_FIRST},$I{TT_FIRST}="今日が期限")', C_CRIT))
    ws.conditional_formatting.add(band, rule(
        f'AND(ISNUMBER($G{TT_FIRST}),{day_ref}<=$G{TT_FIRST},$I{TT_FIRST}="要発注")', C_WARN))
    ws.conditional_formatting.add(band, rule(
        f'AND(ISNUMBER($G{TT_FIRST}),{day_ref}<=$G{TT_FIRST},$I{TT_FIRST}="余裕")', C_OK))
    # 発注すると伸びる分
    ws.conditional_formatting.add(band, rule(
        f'AND(ISNUMBER($J{TT_FIRST}),{day_ref}>N($G{TT_FIRST}),{day_ref}<=$J{TT_FIRST})', C_BUTTER))

    # 今日が発注日でないグループは、発注グループ列を薄くして「今日は入れられない」と分かるようにする。
    # 行ごと消さないのは、切れそうなのに次の便まで待つしかない商品こそ気づいてほしいため。
    ws.conditional_formatting.add(
        f'D{TT_FIRST}:D{TT_LAST}',
        FormulaRule(formula=[f'AND($D{TT_FIRST}<>"",COUNTIF({now_col},$D{TT_FIRST})=0)'],
                    font=Font(name=FONT, size=10, color=C_TEXT3),
                    fill=PatternFill('solid', bgColor='FFF4F1EA'), stopIfTrue=True))
    ws.conditional_formatting.add(
        f'D{TT_FIRST}:D{TT_LAST}',
        FormulaRule(formula=[f'$D{TT_FIRST}<>""'],
                    font=Font(name=FONT, size=10, bold=True, color=C_INK), stopIfTrue=True))

    # 定数が未設定の商品は、暫定計算になっていることが分かるようにする
    ws.conditional_formatting.add(
        f'E{TT_FIRST}:E{TT_LAST}',
        CellIsRule(operator='equal', formula=['"未設定"'],
                   fill=PatternFill('solid', bgColor=C_CRIT_SOFT),
                   font=Font(name=FONT, size=10, bold=True, color=C_CRIT)))

    # 状態列は色だけに頼らず、文字でも読めるようにする
    for text, fill, font_color in (('今日が期限', C_CRIT_SOFT, C_CRIT),
                                   ('要発注', C_WARN_SOFT, C_WARN),
                                   ('余裕', C_OK_SOFT, C_OK)):
        ws.conditional_formatting.add(
            f'I{TT_FIRST}:I{TT_LAST}',
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
        cell = ws.cell(legend_row + i, 14)
        cell.fill = PatternFill('solid', fgColor=color)
        label = ws.cell(legend_row + i, 15, text)
        label.font = F_NOTE
    ws.cell(legend_row + len(legend), 14, '│').font = Font(name=FONT, size=9, bold=True, color=C_INK)
    ws.cell(legend_row + len(legend), 15, '発注デッドライン（在庫が尽きる日 − リードタイム）').font = F_NOTE
    axis = ws.cell(legend_row + len(legend) + 2, 14,
                   '横軸は当日基準日からの14日間です。帯の長さがその商品の在庫が持つ日数を表します。')
    axis.font = F_NOTE
    ws.cell(legend_row + len(legend) + 3, 14,
            '発注数を入れると、伸びた分がバター色で帯の先に足されます。').font = F_NOTE
    ws.cell(legend_row + len(legend) + 5, 14,
            '「定数」が未設定（赤）の商品は、暫定で消費量から計算しています。'
            '商品マスタに最大保管数を入れてください。').font = F_NOTE
    ws.cell(legend_row + len(legend) + 6, 14,
            '在庫が定数に足りているのに帯が赤い商品は、定数が小さすぎます。'
            '棚を広げるか、発注の回数を増やしてください。').font = F_NOTE
    ws.cell(legend_row + len(legend) + 4, 14,
            '発注グループが薄いグレーの行は、今日が発注日ではない商品です'
            '（次に発注できる日は画面上部に出ています）。'
            '赤い帯なのにグレーの行は、次の便まで在庫が持たないということなので、'
            '早めに支配人へ相談してください。').font = F_NOTE

    ws['AB9'] = '発注書用'
    ws['AB9'].font = F_NOTE
    # デッドラインと発注後日数は帯で見えるので、列としては隠す
    for letter in ('H', 'J', 'AB'):
        ws.column_dimensions[letter].hidden = True
    ws.print_area = f'B2:AA{TT_LAST}'
    ws.page_setup.orientation = 'landscape'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    for c, w in zip('BCDEFGHIJKLM',
                    [5, 30, 15, 8, 8, 11, 13, 12, 12, 12, 11, 13]):
        ws.column_dimensions[c].width = w
    ws.column_dimensions['AB'].width = 9
    ws.freeze_panes = f'C{TT_FIRST}'
    return ws


def build_plan(wb):
    """
    動員を入れる場所。予測と実績の両方をここに置く。

    実績を入れると3つ効く。
      1. 期間の平均動員（実測PI値の分母）が自動で出る。手入力の打ち間違いが消える
      2. 曜日ごとの平均が出るので、予測を入れていない日の既定値が
         「劇場の標準来場者数」ではなく「その曜日の実績平均」になる。
         映画館は曜日で動員が3倍近く違うため、ここが効く
      3. 予測と実績の差が見えるので、読みの癖が分かる

    日付は「表の開始日」から作る。当日基準日に紐づけると、CSVを貼り替えるたびに
    行がずれて、入力した予測が別の日の値になってしまう。
    """
    ws = wb.create_sheet(S_PLAN)
    sheet_title(ws, '動員を入れる',
                '予測と実績の両方をここに入れます。空欄の日はその曜日の実績平均が使われます。')

    ws['B4'] = '入力のしかた'
    ws['B4'].font = F_BOLD
    for i, text in enumerate([
        '・「実績来場者数」は毎日でも、月1回まとめてでも構いません。入れるほど精度が上がります。',
        '・「予測来場者数」は読みが変わる日だけで十分です。空欄の日はその曜日の実績平均を使います。',
        '・実績が入っている日は実績をそのまま使います（季節係数は掛けません）。',
        '・予測・曜日平均・標準来場者数を使う日には、季節係数が掛かります。',
        '・発注計算は「当日基準日から リードタイム＋発注サイクル 日分」の動員を見ます。',
    ]):
        ws.cell(5 + i, 2, text).font = F_NOTE

    ws['B11'] = '表の開始日'
    ws['B11'].font = F_BOLD
    ws['B11'].fill = FILL_HEAD
    ws['B11'].border = BORDER
    # 曜日列(C)は幅6しかないので、日付が「###」に潰れないよう結合して幅を確保する
    ws.merge_cells('C11:D11')
    ws['C11'].fill = FILL_INPUT
    ws['C11'].border = BORDER
    ws['C11'].font = F_INPUT
    ws['C11'].number_format = FMT_DATE
    ws['C11'].alignment = Alignment(horizontal='center', vertical='center')
    ws['D11'].border = BORDER
    ws['E11'] = ('← 1回だけ入れてください。空欄なら「2ヶ月前の基準日」から始まります。'
                 'ここを動かすと入力済みの行がずれるので、決めたら触らないでください。')
    ws['E11'].font = F_NOTE

    date_col = col(S_PLAN, 'B', PLAN_FIRST, PLAN_LAST)
    dow_col = col(S_PLAN, 'C', PLAN_FIRST, PLAN_LAST)
    actual_col = col(S_PLAN, 'E', PLAN_FIRST, PLAN_LAST)
    in_period = (f'{date_col},">="&{Q_SET}!$C$7,{date_col},"<="&{Q_SET}!$C$6')

    # ---- 曜日別の実績平均。予測を入れていない日の既定値になる ----
    ws['N4'] = '曜日別の実績平均（前回基準日〜当日基準日）'
    ws['N4'].font = F_BOLD
    put_headers(ws, PLAN_DOW_ROW - 1, ['曜日', '実績の平均', '日数'],
                ['自動'] * 3, start_col=14)
    for i, day in enumerate('月火水木金土日'):
        r = PLAN_DOW_ROW + i
        ws[f'N{r}'] = day
        ws[f'O{r}'] = (f'=IFERROR(AVERAGEIFS({actual_col},{dow_col},$N{r},{in_period}),"")')
        ws[f'P{r}'] = f'=COUNTIFS({dow_col},$N{r},{in_period},{actual_col},">0")'
        style_row(ws, r, 'NOP', fill=FILL_AUTO)
        ws[f'N{r}'].alignment = Alignment(horizontal='center')
        ws[f'O{r}'].number_format = FMT_INT
        ws[f'P{r}'].number_format = FMT_INT
    dow_key = col(S_PLAN, 'N', PLAN_DOW_ROW, PLAN_DOW_ROW + 6)
    dow_avg = col(S_PLAN, 'O', PLAN_DOW_ROW, PLAN_DOW_ROW + 6)
    ws[f'N{PLAN_DOW_ROW + 8}'] = ('※ 日数が0の曜日は、実績がまだ入っていない曜日です。'
                                  'その曜日は標準来場者数が使われます。')
    ws[f'N{PLAN_DOW_ROW + 8}'].font = F_NOTE

    headers = ['日付', '曜日', '予測来場者数', '実績来場者数', '主な作品・備考',
               '季節係数', '採用来場者数', '予測とのズレ']
    kinds = ['自動', '自動', '入力', '入力', '入力', '自動', '自動', '自動']
    put_headers(ws, PLAN_FIRST - 1, headers, kinds)

    for r in range(PLAN_FIRST, PLAN_LAST + 1):
        offset = r - PLAN_FIRST
        special = (f'SUMPRODUCT(MAX(({special_col("F")}<=$B{r})*'
                   f'({special_col("G")}>=$B{r})*{special_col("H")}))')
        ws[f'B{r}'] = (f'=IF($C$11<>"",$C$11+{offset},'
                       f'IF({Q_SET}!$C$7="","",{Q_SET}!$C$7+{offset}))')
        # 曜日は環境の言語設定に左右されないよう CHOOSE で組み立てる
        ws[f'C{r}'] = f'=IF($B{r}="","",{weekday_jp(f"$B{r}")})'
        ws[f'G{r}'] = (f'=IF($B{r}="","",IF({special}=0,'
                       f'IFERROR(INDEX({month_col("C")},MONTH($B{r})),1),{special}))')
        # 採用来場者数の決め方。実績＞予測＞その曜日の実績平均＞標準来場者数 の順。
        # 実績は実際に来た数なので、季節係数は掛けない（もう入っている）。
        fallback = (f'IF(N(IFERROR(INDEX({dow_avg},MATCH($C{r},{dow_key},0)),0))>0,'
                    f'IFERROR(INDEX({dow_avg},MATCH($C{r},{dow_key},0)),0),'
                    f'N({Q_SET}!$C$14))')
        ws[f'H{r}'] = (f'=IF($B{r}="","",IF(ISNUMBER($E{r}),$E{r},'
                       f'ROUND(IF(ISNUMBER($D{r}),$D{r},{fallback})*N($G{r}),0)))')
        ws[f'I{r}'] = (f'=IF(OR(NOT(ISNUMBER($D{r})),NOT(ISNUMBER($E{r})),$D{r}=0),"",'
                       f'$E{r}/$D{r}-1)')
        # グラフの目盛り用。日付を軸に使うと環境によって目盛りがずれるため文字にする
        ws[f'K{r}'] = f'=IF($B{r}="","",TEXT($B{r},"m/d")&"("&$C{r}&")")'
        ws[f'K{r}'].font = F_NOTE
        style_row(ws, r, 'BCDEFGHI', fill=FILL_AUTO)
        for letter in 'DEF':
            ws[f'{letter}{r}'].fill = FILL_INPUT
            ws[f'{letter}{r}'].font = F_INPUT
        ws[f'B{r}'].number_format = FMT_DATE
        for letter in 'DEH':
            ws[f'{letter}{r}'].number_format = FMT_INT
        ws[f'G{r}'].number_format = '0%'
        ws[f'I{r}'].number_format = '+0%;-0%;0%'
        ws[f'C{r}'].alignment = Alignment(horizontal='center')

    # 土日を薄く塗って週のリズムを見せる
    ws.conditional_formatting.add(
        f'B{PLAN_FIRST}:I{PLAN_LAST}',
        FormulaRule(formula=[f'OR(WEEKDAY($B{PLAN_FIRST})=1,WEEKDAY($B{PLAN_FIRST})=7)'],
                    fill=PatternFill('solid', bgColor='FFF4EFE2'), stopIfTrue=False))
    # 当日基準日より前＝実績を入れるべき行。まだ空なら薄く警告する
    ws.conditional_formatting.add(
        f'E{PLAN_FIRST}:E{PLAN_LAST}',
        FormulaRule(formula=[f'AND($B{PLAN_FIRST}<>"",$B{PLAN_FIRST}<{Q_SET}!$C$6,'
                             f'$E{PLAN_FIRST}="")'],
                    fill=PatternFill('solid', bgColor=C_WARN_SOFT), stopIfTrue=True))
    # 読みが大きく外れた日を目立たせる
    for op, val in (('greaterThan', '0.2'), ('lessThan', '-0.2')):
        ws.conditional_formatting.add(
            f'I{PLAN_FIRST}:I{PLAN_LAST}',
            CellIsRule(operator=op, formula=[val],
                       font=Font(name=FONT, size=10, bold=True, color=C_WARN)))

    widths = {'B': 13, 'C': 6, 'D': 15, 'E': 15, 'F': 34, 'G': 11, 'H': 15, 'I': 13,
              'J': 3, 'K': 14, 'L': 3, 'M': 3, 'N': 8, 'O': 12, 'P': 8}
    for c, w in widths.items():
        ws.column_dimensions[c].width = w
    ws.freeze_panes = f'B{PLAN_FIRST}'
    return ws


def build_order_sheet(wb, vendors):
    """
    発注書。日々の作業の終着点で、ここを印刷して仕入先に送る。
    タイムテーブルで入れた発注数のうち、選んだ仕入先の分だけを並べる。
    """
    ws = wb.create_sheet(S_SHEET)
    ws.sheet_view.showGridLines = False
    ws['B2'] = '発 注 書'
    ws['B2'].font = Font(name=FONT, size=18, bold=True, color=C_INK)
    ws['B3'] = 'タイムテーブルで入れた発注数がここに集まります。発注先を選んで印刷してください。'
    ws['B3'].font = F_NOTE

    meta = [
        ('発注先', None, FILL_INPUT, '← ここを選ぶ'),
        ('発注日', f'=IF({Q_SET}!$C$6="","",TEXT({Q_SET}!$C$6,"yyyy/m/d")&"（"&'
                  f'CHOOSE(WEEKDAY({Q_SET}!$C$6),"日","月","火","水","木","金","土")&"）")',
         FILL_AUTO, '自動'),
        ('発注元', f'=IF({Q_SET}!$C$5="","",{Q_SET}!$C$4&"　"&{Q_SET}!$C$5)', FILL_AUTO, '自動'),
        ('担当者', f'={Q_SET}!$C$16', FILL_AUTO, '自動'),
    ]
    for i, (label, formula, fill, kind) in enumerate(meta):
        r = 5 + i
        ws[f'B{r}'] = label
        ws[f'B{r}'].font = F_HEAD
        ws[f'B{r}'].fill = FILL_HEAD
        ws[f'B{r}'].border = BORDER
        ws.merge_cells(f'C{r}:F{r}')
        cell = ws[f'C{r}']
        if formula:
            cell.value = formula
        cell.fill = fill
        cell.border = BORDER
        cell.font = F_INPUT if fill is FILL_INPUT else F_BASE
        ws[f'G{r}'] = kind
        ws[f'G{r}'].font = F_NOTE
    ws['C5'] = 'すべて'

    # 発注先の選択肢。数式で重複を除くとCSE前提の式になるため、生成時に書き出す
    ws['N4'] = '発注先の選択肢'
    ws['N4'].font = F_NOTE
    ws['N5'] = 'すべて'
    ws['N5'].font = F_NOTE
    for i, v in enumerate(vendors[:24]):
        c = ws.cell(6 + i, 14, v)
        c.font = F_NOTE
    add_validation(ws, f'={col(S_SHEET, "N", 5, 5 + max(1, len(vendors[:24])))}', 'C5')

    headers = ['No', '仕入先', '商品コード', '商品名', '発注単位', '入数', '発注数', '単価', '金額']
    put_headers(ws, 10, headers, ['自動'] * len(headers))

    first, last = 11, 110
    tt_pick = col(S_TT, TT_COL['pick'], TT_FIRST, TT_LAST)
    for r in range(first, last + 1):
        k = r - first + 1
        pick = f'MATCH({k},{tt_pick},0)'
        ws[f'B{r}'] = f'=IF({k}<=MAX({tt_pick}),{k},"")'
        ws[f'C{r}'] = (f'=IF($B{r}="","",IFERROR(INDEX({col(S_CALC, "F", CALC_FIRST, CALC_LAST)},'
                       f'{pick}),""))')
        ws[f'D{r}'] = (f'=IF($B{r}="","",IFERROR(INDEX({col(S_CALC, "C", CALC_FIRST, CALC_LAST)},'
                       f'{pick})&"",""))')
        ws[f'E{r}'] = (f'=IF($B{r}="","",IFERROR(INDEX({col(S_TT, TT_COL["name"], TT_FIRST, TT_LAST)},'
                       f'{pick})&"",""))')
        # 入数1の商品（BIBシロップ・樽・缶など）に「ケース」と印字すると
        # 仕入先に別の数量として伝わる。入数から単位を組み立てる。
        ws[f'F{r}'] = (f'=IF($B{r}="","",IF(N($G{r})<=1,"単品",'
                       f'"ケース("&TEXT(N($G{r}),"#,##0")&"入)"))')
        ws[f'G{r}'] = (f'=IF($B{r}="","",IFERROR(INDEX({col(S_CALC, "G", CALC_FIRST, CALC_LAST)},'
                       f'{pick}),""))')
        ws[f'H{r}'] = (f'=IF($B{r}="","",IFERROR(INDEX({col(S_TT, TT_COL["qty"], TT_FIRST, TT_LAST)},'
                       f'{pick}),""))')
        ws[f'I{r}'] = (f'=IF($B{r}="","",IFERROR(INDEX({col(S_CALC, "H", CALC_FIRST, CALC_LAST)},'
                       f'{pick}),""))')
        ws[f'J{r}'] = f'=IF($B{r}="","",N($H{r})*N($I{r}))'
        # 罫線と塗りは条件付き書式で「明細がある行」だけに付ける
        style_row(ws, r, 'BCDEFGHIJ', border=False)
        ws[f'B{r}'].number_format = FMT_INT
        ws[f'G{r}'].number_format = FMT_INT
        ws[f'H{r}'].number_format = FMT_INT
        ws[f'H{r}'].font = F_BOLD
        ws[f'I{r}'].number_format = FMT_YEN
        ws[f'J{r}'].number_format = FMT_YEN

    total = last + 1
    ws.merge_cells(f'B{total}:I{total}')
    ws[f'B{total}'] = '合計（税抜）'
    ws[f'B{total}'].font = F_BOLD
    ws[f'B{total}'].fill = FILL_HEAD
    ws[f'B{total}'].alignment = Alignment(horizontal='right', vertical='center')
    ws[f'B{total}'].border = BORDER
    ws[f'J{total}'] = f'=SUM(J{first}:J{last})'
    ws[f'J{total}'].font = Font(name=FONT, size=12, bold=True, color=C_INK)
    ws[f'J{total}'].number_format = FMT_YEN
    ws[f'J{total}'].fill = FILL_HEAD
    ws[f'J{total}'].border = BORDER

    ws[f'B{total + 2}'] = '※ 発注数はタイムテーブルシートで入力します。ここは印刷用の清書です。'
    ws[f'B{total + 2}'].font = F_NOTE
    ws[f'B{total + 3}'] = '※ 発注先を切り替えると、その仕入先の明細だけに絞り込まれます。'
    ws[f'B{total + 3}'].font = F_NOTE

    # 明細がある行にだけ枠と塗りを付ける。印刷したときに空の升目が並ばない
    ws.conditional_formatting.add(
        f'B{first}:J{last}',
        FormulaRule(formula=[f'$B{first}<>""'],
                    fill=PatternFill('solid', bgColor='FFFBFAF7'),
                    border=Border(left=THIN, right=THIN, top=THIN, bottom=THIN),
                    stopIfTrue=True))

    ws.print_area = f'B2:J{total}'
    ws.page_setup.orientation = 'portrait'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    for c, w in zip('BCDEFGHIJ', [5, 24, 16, 30, 10, 8, 9, 11, 13]):
        ws.column_dimensions[c].width = w
    ws.column_dimensions['N'].width = 26
    ws.freeze_panes = f'B{first}'
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
        ws[f'J{r}'] = (f'=IF(OR($F{r}="",$C{r}<>{Q_SET}!$C$4),"",'
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
        ws[f'B{r}'] = f'=IF({Q_THEATER}!$B{src}="","",{Q_THEATER}!$B{src})'
        ws[f'C{r}'] = f'=IF($B{r}="","",{Q_THEATER}!$C{src})'
        ws[f'D{r}'] = f'=IF($B{r}="","",{Q_THEATER}!$D{src})'
        ws[f'E{r}'] = f'=IF($B{r}="","",{Q_THEATER}!$E{src})'
        ws[f'F{r}'] = (f'=IF($B{r}="","",COUNTIFS({csv_col(S_CUR, "B")},$B{r},'
                       f'{csv_col(S_CUR, "I")},"<>"))')
        ws[f'G{r}'] = (f'=IF($B{r}="","",SUMPRODUCT(({csv_col(S_CUR, "B")}=$B{r})*'
                       f'{csv_col(S_CUR, "N")}*{csv_col(S_CUR, "Q")}))')
        ws[f'H{r}'] = (f'=IF($B{r}="","",COUNTIFS({csv_col(S_CUR, "B")},$B{r},'
                       f'{csv_col(S_CUR, "N")},"<0"))')
        ws[f'I{r}'] = (f'=IF($B{r}="","",COUNTIFS({po_col("C")},$B{r},{po_col("Q")},"発注済"))')
        ws[f'J{r}'] = (f'=IF(OR($B{r}="",{Q_SET}!$C$6=""),"",'
                       f'SUMIFS({po_col("N")},{po_col("C")},$B{r},'
                       f'{po_col("E")},">="&DATE(YEAR({Q_SET}!$C$6),MONTH({Q_SET}!$C$6),1),'
                       f'{po_col("E")},"<="&DATE(YEAR({Q_SET}!$C$6),MONTH({Q_SET}!$C$6)+1,0)))')
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


def build_dashboard(wb, categories, vendors):
    """
    figure-first のダッシュボード。
    細い列を並べた方眼をキャンバスにして、結合セルでブロックを組む
    （Excelのダッシュボードでよく使われる作り方）。
    グラフとセル塗りだけで構成し、マクロやアドインは使わない。
    """
    ws = wb.create_sheet(S_DASH)
    ws.sheet_view.showGridLines = False

    CANVAS_LAST = 30                       # B列から30列を方眼として使う
    for i in range(2, 2 + CANVAS_LAST):
        ws.column_dimensions[get_column_letter(i)].width = 3.3

    tt_state = col(S_TT, TT_COL['state'], TT_FIRST, TT_LAST)
    tt_days = col(S_TT, TT_COL['days'], TT_FIRST, TT_LAST)
    # ②のD列は発注グループに変えたので、分類別の集計は発注計算から直接引く。
    # 行数が同じ（400行）なので、シートをまたいでもCOUNTIFSで突き合わせられる。
    tt_cat = col(S_CALC, 'E', CALC_FIRST, CALC_LAST)
    calc_vendor = col(S_CALC, 'F', CALC_FIRST, CALC_LAST)
    calc_amount = col(S_CALC, 'AC', CALC_FIRST, CALC_LAST)
    calc_stock = col(S_CALC, 'AA', CALC_FIRST, CALC_LAST)

    def block(cell_range, value, fmt=None, font=None, fill=None, align='left'):
        ws.merge_cells(cell_range)
        top = cell_range.split(':')[0]
        c = ws[top]
        c.value = value
        if font:
            c.font = font
        if fill:
            c.fill = fill
        if fmt:
            c.number_format = fmt
        c.alignment = Alignment(horizontal=align, vertical='center')
        return c

    def section(row, start_col, text):
        letter = get_column_letter(start_col)
        c = ws[f'{letter}{row}']
        c.value = text
        c.font = Font(name=FONT, size=11, bold=True, color=C_INK)

    # ================= 見出し =================
    block('B2:N2', 'ダッシュボード', font=F_TITLE)
    block('B3:N3',
          f'=IF({Q_SET}!$C$4="","設定シートで対象劇場を選んでください。",'
          f'{Q_SET}!$C$4&"　"&{Q_SET}!$C$5&"　／　基準日 "&TEXT({Q_SET}!$C$6,"yyyy/m/d")&'
          f'"　／　算出基準 "&{Q_SET}!$C$9)',
          font=Font(name=FONT, size=10, color=C_TEXT2))

    # ================= 信号ボード =================
    # 数の大小より「どの色がどれだけあるか」が先に目に入るようにする
    signals = [
        ('B', 'G', '欠品リスク', f'=COUNTIF({tt_state},"今日が期限")', '今日発注しないと間に合わない',
         C_CRIT, C_CRIT_SOFT),
        ('I', 'N', '要発注', f'=COUNTIF({tt_state},"要発注")', '発注サイクル内に切れる',
         C_WARN, C_WARN_SOFT),
        ('P', 'U', '在庫に余裕', f'=COUNTIF({tt_state},"余裕")', 'しばらく発注不要',
         C_OK, C_OK_SOFT),
    ]
    for left, right, label, formula, note, color, soft in signals:
        block(f'{left}5:{right}5', label, font=Font(name=FONT, size=10, bold=True, color=color),
              fill=PatternFill('solid', fgColor=soft))
        big = block(f'{left}6:{right}8', formula, fmt=FMT_INT,
                    font=Font(name=FONT, size=30, bold=True, color=color),
                    fill=PatternFill('solid', fgColor=soft))
        big.alignment = Alignment(horizontal='center', vertical='center')
        block(f'{left}9:{right}9', note, font=F_NOTE,
              fill=PatternFill('solid', fgColor=soft), align='center')

    block('W5:AB5', '推奨発注額（税抜）', font=Font(name=FONT, size=10, bold=True, color=C_TEXT2),
          fill=FILL_HEAD)
    amt = block('W6:AB8', f'=SUM({calc_amount})', fmt=FMT_YEN,
                font=Font(name=FONT, size=18, bold=True, color=C_INK), fill=FILL_HEAD)
    amt.alignment = Alignment(horizontal='center', vertical='center')
    block('W9:AB9', f'="在庫金額 "&TEXT(SUM({calc_stock}),"¥#,##0")',
          font=F_NOTE, fill=FILL_HEAD, align='center')

    for r in (5, 9):
        ws.row_dimensions[r].height = 16
    for r in (6, 7, 8):
        ws.row_dimensions[r].height = 15

    # ================= ワッフル（状態の構成比） =================
    section(11, 2, '在庫状況の内訳')
    ws['B12'] = '1マス＝全体の1%。100マスで取扱商品全体を表しています。'
    ws['B12'].font = F_NOTE

    # 累積のしきい値。ワッフルの塗り分けに使う
    ws['AE13'] = '欠品リスク%'
    ws['AF13'] = f'=IFERROR(COUNTIF({tt_state},"今日が期限")/$AF$18*100,0)'
    ws['AE14'] = '＋要発注%'
    ws['AF14'] = f'=$AF$13+IFERROR(COUNTIF({tt_state},"要発注")/$AF$18*100,0)'
    ws['AE15'] = '＋余裕%'
    ws['AF15'] = f'=$AF$14+IFERROR(COUNTIF({tt_state},"余裕")/$AF$18*100,0)'
    ws['AE18'] = '取扱商品数'
    ws['AF18'] = f'=COUNTIF({tt_state},"?*")'
    for r in (13, 14, 15, 18):
        ws[f'AE{r}'].font = F_NOTE
        ws[f'AF{r}'].font = F_NOTE

    for row in range(10):
        for c in range(10):
            cell = ws.cell(14 + row, 2 + c)
            cell.value = row * 10 + c + 1
            cell.font = Font(name=FONT, size=1, color='FFFFFFFF')
            cell.fill = FILL_TRACK
            cell.border = Border(right=Side(style='thin', color='FFFFFFFF'),
                                 bottom=Side(style='thin', color='FFFFFFFF'))
        ws.row_dimensions[14 + row].height = 15

    waffle = 'B14:K23'
    for threshold, color in (('$AF$13', C_CRIT), ('$AF$14', C_WARN), ('$AF$15', C_OK)):
        ws.conditional_formatting.add(waffle, FormulaRule(
            formula=[f'B14<={threshold}'],
            fill=PatternFill('solid', bgColor=color), stopIfTrue=True))
    ws.conditional_formatting.add(waffle, FormulaRule(
        formula=['TRUE'], fill=PatternFill('solid', bgColor='FFD8D4CB'), stopIfTrue=True))

    legend = [('欠品リスク', C_CRIT), ('要発注', C_WARN), ('余裕', C_OK), ('データなし', 'FFD8D4CB')]
    for i, (text, color) in enumerate(legend):
        c = ws.cell(25, 2 + i * 3)
        c.fill = PatternFill('solid', fgColor=color)
        lab = ws.cell(25, 3 + i * 3, text)
        lab.font = F_NOTE
        lab.alignment = Alignment(horizontal='left')

    # ================= 在庫日数の分布 =================
    section(11, 13, '在庫日数の分布')
    # 横棒グラフは先頭のカテゴリが下に来るので、危ない側が上に出るよう逆順に置く
    buckets = [
        ('14日以上', f'=COUNTIFS({tt_days},">14")'),
        ('7〜14日', f'=COUNTIFS({tt_days},">7",{tt_days},"<=14")'),
        ('3〜7日', f'=COUNTIFS({tt_days},">3",{tt_days},"<=7")'),
        ('1〜3日', f'=COUNTIFS({tt_days},">1",{tt_days},"<=3")'),
        ('0〜1日', f'=COUNTIFS({tt_days},">0",{tt_days},"<=1")'),
        ('切れている', f'=COUNTIFS({tt_days},"<=0")'),
    ]
    ws['AE21'] = '在庫日数'
    ws['AF21'] = '品目数'
    for i, (label, formula) in enumerate(buckets):
        ws.cell(22 + i, 31, label).font = F_NOTE
        c = ws.cell(22 + i, 32, formula)
        c.font = F_NOTE
    dist = BarChart()
    dist.type = 'bar'
    dist.title = None
    dist.legend = None
    dist.height, dist.width = 6.6, 11.5
    dist.add_data(Reference(ws, min_col=32, min_row=22, max_row=27), titles_from_data=False)
    dist.set_categories(Reference(ws, min_col=31, min_row=22, max_row=27))
    # 分布は重症度の軸なので、棒ごとに状態の色を与える
    for i, color in enumerate([C_OK, C_OK, C_WARN, C_WARN, C_CRIT, C_CRIT]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = color[2:]
        pt.graphicalProperties.line.noFill = True
        dist.series[0].data_points.append(pt)
    dist.gapWidth = 40
    dist.plotVisOnly = False        # 元データを隠し列に置いているため
    ws.add_chart(dist, 'M13')

    # ================= 商品分類 × 状態 =================
    section(27, 2, '商品分類 × 状態')
    states = ['今日が期限', '要発注', '余裕']
    colors = [C_CRIT, C_WARN, C_OK]
    block('B28:G28', '商品分類', font=F_HEAD, fill=FILL_HEAD, align='center')
    for i, (st, cl) in enumerate(zip(states, colors)):
        left = get_column_letter(8 + i * 2)
        right = get_column_letter(9 + i * 2)
        block(f'{left}28:{right}28', st, font=Font(name=FONT, size=9, bold=True, color=cl),
              fill=FILL_HEAD, align='center')
    block('N28:P28', '在庫金額', font=F_HEAD, fill=FILL_HEAD, align='center')

    for i, name in enumerate(categories):
        r = 29 + i
        block(f'B{r}:G{r}', name, font=Font(name=FONT, size=9, color=C_INK))
        for j, st in enumerate(states):
            left = get_column_letter(8 + j * 2)
            right = get_column_letter(9 + j * 2)
            c = block(f'{left}{r}:{right}{r}',
                      f'=COUNTIFS({tt_cat},$B{r},{tt_state},"{st}")',
                      fmt=FMT_INT, font=Font(name=FONT, size=9, color=C_INK), align='center')
        block(f'N{r}:P{r}', f'=SUMIF({col(S_CALC, "E", CALC_FIRST, CALC_LAST)},$B{r},{calc_stock})',
              fmt=FMT_YEN, font=Font(name=FONT, size=9, color=C_TEXT2), align='right')
        ws.row_dimensions[r].height = 14

    heat_last = 28 + len(categories)
    for j, color in enumerate(colors):
        left = get_column_letter(8 + j * 2)
        right = get_column_letter(9 + j * 2)
        ws.conditional_formatting.add(
            f'{left}29:{right}{heat_last}',
            ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFFFF',
                           end_type='max', end_color=color))
    # 在庫金額はデータバーで大小を見せる
    ws.conditional_formatting.add(f'N29:P{heat_last}',
                                  DataBarRule(start_type='num', start_value=0, end_type='max',
                                              color=C_BUTTER[2:], showValue=True))

    # ================= 仕入先別 発注予定額 =================
    section(27, 19, '仕入先別 推奨発注額')
    ws['AE31'] = '仕入先'
    ws['AF31'] = '推奨発注額'
    for i, v in enumerate(vendors[:12]):
        ws.cell(32 + i, 31, v).font = F_NOTE
        ws.cell(32 + i, 32, f'=SUMIF({calc_vendor},$AE${32 + i},{calc_amount})').font = F_NOTE
        ws.cell(32 + i, 32).number_format = FMT_YEN
    vend_chart = BarChart()
    vend_chart.type = 'bar'
    vend_chart.title = None
    vend_chart.legend = None
    vend_chart.height, vend_chart.width = 8.4, 11.5
    n = min(12, len(vendors))
    vend_chart.add_data(Reference(ws, min_col=32, min_row=32, max_row=31 + n), titles_from_data=False)
    vend_chart.set_categories(Reference(ws, min_col=31, min_row=32, max_row=31 + n))
    vend_chart.series[0].graphicalProperties.solidFill = C_BUTTER[2:]
    vend_chart.gapWidth = 45
    vend_chart.x_axis.numFmt = '#,##0,,"百万"'   # 桁が長いと軸ラベルが斜めになるため
    vend_chart.plotVisOnly = False
    ws.add_chart(vend_chart, 'S29')

    # ================= 今後14日の入荷予定 =================
    first_row = heat_last + 3
    section(first_row, 2, '今後14日の入荷予定（発注済で未入荷のもの）')
    ws.cell(first_row + 1, 2, '納品予定日ごとの金額です。色が濃い日ほど入荷が集中しています。').font = F_NOTE
    for d in range(SPAN_DAYS):
        left = get_column_letter(2 + d * 2)
        right = get_column_letter(3 + d * 2)
        head = block(f'{left}{first_row + 2}:{right}{first_row + 2}',
                     f'=IF({Q_SET}!$C$6="","",{Q_SET}!$C$6+{d})',
                     fmt='m/d', font=F_DAY, fill=FILL_HEAD, align='center')
        val = block(f'{left}{first_row + 3}:{right}{first_row + 3}',
                    f'=IF({Q_SET}!$C$6="",0,SUMIFS({po_col("N")},{po_col("C")},{Q_SET}!$C$4,'
                    f'{po_col("P")},{Q_SET}!$C$6+{d},{po_col("Q")},"発注済"))',
                    fmt='#,##0;;"-"', font=Font(name=FONT, size=9, color=C_INK), align='center')
        ws.row_dimensions[first_row + 3].height = 26
    ws.conditional_formatting.add(
        f'B{first_row + 3}:{get_column_letter(1 + SPAN_DAYS * 2)}{first_row + 3}',
        ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFFFF',
                       end_type='max', end_color=C_BUTTER))

    # ================= 今後14日の動員予測 =================
    plan_row = first_row + 6
    section(plan_row, 2, '今後14日の動員予測')
    ws.cell(plan_row + 1, 2,
            '動員計画シートの数字です。山の来る日に在庫が持つかを見てください。').font = F_NOTE
    plan_chart = BarChart()
    plan_chart.type = 'col'
    plan_chart.title = None
    plan_chart.legend = None
    plan_chart.height, plan_chart.width = 6.4, 13.0
    plan_chart.add_data(Reference(ws.parent[S_PLAN], min_col=8,
                                  min_row=PLAN_FIRST, max_row=PLAN_FIRST + 13),
                        titles_from_data=False)
    plan_chart.set_categories(Reference(ws.parent[S_PLAN], min_col=11,
                                        min_row=PLAN_FIRST, max_row=PLAN_FIRST + 13))
    plan_chart.series[0].graphicalProperties.solidFill = C_INK[2:]
    plan_chart.gapWidth = 35
    plan_chart.plotVisOnly = False
    ws.add_chart(plan_chart, f'B{plan_row + 2}')

    # ================= 商品の入れ替わり（発注漏れの防止） =================
    section(plan_row, 20, '商品の入れ替わり')
    ws.cell(plan_row + 1, 20,
            '3枚の在庫CSV（当日・1ヶ月前・2ヶ月前）を突き合わせた結果です。').font = F_NOTE
    calc_status = col(S_CALC, 'AB', CALC_FIRST, CALC_LAST)
    calc_change = col(S_CALC, 'AS', CALC_FIRST, CALC_LAST)
    changes = [
        ('新商品（今月から）', f'=COUNTIF({calc_change},"新商品（今月から）")',
         '当日だけにある。マスタ登録がまだの可能性', C_BUTTER),
        ('新商品（先月から）', f'=COUNTIF({calc_change},"新商品（先月から）")',
         '1ヶ月前から扱い開始。2ヶ月前には無い', C_BUTTER),
        ('終売（今月から）', f'=COUNTIF({calc_change},"終売（今月から）")',
         '1ヶ月前まではあった。当日で消えた', C_TEXT3),
        ('終売（先月から）', f'=COUNTIF({calc_change},"終売（先月から）")',
         '2ヶ月前にだけあった。もう扱っていない', C_TEXT3),
        ('復活（先月は無し）', f'=COUNTIF({calc_change},"復活（先月は無し）")',
         '2ヶ月前と当日にあり、1ヶ月前に無い。要確認', C_WARN),
        ('定数が発注点より小さい', f'=COUNTIF({calc_status},"*定数が発注点より小さい*")',
         '棚を満たしても次の便まで持たない。棚割りか発注頻度の見直しが要る', C_WARN),
        ('マスタ未登録', f'=COUNTIF({calc_status},"*マスタ未登録*")',
         'L/T・ロット・PI値が既定値のまま', C_WARN),
        ('在庫マイナス', f'=COUNTIF({calc_status},"*在庫マイナス*")',
         '納品の未計上などデータ側のずれ', C_CRIT),
    ]
    for i, (label, formula, note, color) in enumerate(changes):
        r = plan_row + 2 + i * 2
        block(f'T{r}:X{r}', label, font=Font(name=FONT, size=10, bold=True, color=color),
              fill=FILL_HEAD)
        block(f'Y{r}:Z{r}', formula, fmt=FMT_INT,
              font=Font(name=FONT, size=14, bold=True, color=color), align='center')
        block(f'T{r + 1}:Z{r + 1}', note, font=F_NOTE)

    # グラフの元データはキャンバス（B〜AC）の外に置く。
    # 列を隠すと既定ではグラフに描画されなくなるため、隠さず印刷範囲の外に出す。
    ws.print_area = f'B2:AC{plan_row + 17}'
    ws.page_setup.orientation = 'landscape'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1

    note_row = plan_row + 12
    ws.cell(note_row, 2,
            '※ 数字は「発注計算」シートの推奨発注数にもとづく金額です（税抜）。').font = F_NOTE
    ws.cell(note_row + 1, 2,
            '※ グラフの元データは AE 列以降（非表示）にあります。'
            '商品分類や仕入先が増えたらそこに追記してください。').font = F_NOTE
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
    categories, vendor_count = [], {}
    for r in rows:
        if r['商品分類名'] not in categories:
            categories.append(r['商品分類名'])
        vendor_count[r['支払先名']] = vendor_count.get(r['支払先名'], 0) + 1
    # 取扱点数の多い順に並べる。グラフでは主要な仕入先が上に来る
    vendors = sorted(vendor_count, key=lambda v: -vendor_count[v])
    return items, categories, vendors


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path('売店発注ツール.xlsx')
    src = Path(sys.argv[2]) if len(sys.argv) > 2 else None
    items, categories, vendors = (load_items(src) if src
                                  else ([], ['コンセ包材'], ['仕入先A']))

    wb = Workbook()
    wb.remove(wb.active)
    build_intro(wb)
    build_settings(wb)
    build_calc(wb)
    build_plan(wb)
    build_timetable(wb)
    build_order_sheet(wb, vendors)
    build_dashboard(wb, categories, vendors)
    build_po(wb)
    build_csv_sheet(wb, S_CUR, [], '今日の在庫を貼る',
                    '在庫システムから出力した、今日の在庫一覧CSVです。これが理論在庫になります。',
                    f'CSVの見出し行を除いたデータを、A{CSV_FIRST} セルに貼り付けてください。')
    build_csv_sheet(wb, S_MID, [S_CUR], '1ヶ月前の在庫を貼る（月1回）',
                    '目的は「取扱商品が変わっていないか」の確認です。'
                    '1ヶ月前の1日時点の在庫一覧を貼ってください。'
                    '当日と2ヶ月前のあいだを埋めることで、1ヶ月だけ扱った商品も拾えます。',
                    f'月に1回、A{CSV_FIRST} セルに貼り替えてください。')
    build_csv_sheet(wb, S_PRV, [S_CUR, S_MID], '2ヶ月前の在庫を貼る（月1回）',
                    '目的は「取扱商品が変わっていないか」の確認です。'
                    '当月を含まない2ヶ月前の1日時点の在庫一覧を貼ってください。'
                    '半年ごとに入れ替わる商品も、これで発注リストから漏れません。',
                    f'月に1回、A{CSV_FIRST} セルに貼り替えてください。')
    build_item_master(wb, items)
    build_const_master(wb)
    build_mechanism(wb)
    build_count_sheet(wb)
    build_all_theaters(wb)

    # 毎日の3枚を左に固め、たまに触るもの・管理者向けを右に置く。
    # 計算だけのシートは非表示にして、触る場所を減らす。
    order = [S_INTRO, S_MECH, S_CUR, S_TT, S_SHEET, S_COUNT, S_MID, S_PRV, S_PLAN,
             S_DASH, S_PO, S_ITEM, S_CONST, S_SET, S_CALC, S_ALL]
    wb._sheets = [wb[n] for n in order]
    for ws in wb.worksheets:
        if ws.title == S_SHEET:
            ws.sheet_properties.tabColor = C_CRIT[2:]      # 終着点
        elif ws.title in (S_INTRO, S_MECH):
            ws.sheet_properties.tabColor = C_OK[2:]        # 読み物
        elif ws.title in (S_CUR, S_TT):
            ws.sheet_properties.tabColor = C_BUTTER[2:]    # 毎日さわる
        elif ws.title in (S_CALC, S_ALL):
            ws.sheet_properties.tabColor = C_INK[2:]
            ws.sheet_state = 'hidden'                      # 計算専用。触る必要がない

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f'saved: {out}  sheets={len(wb.worksheets)}  items={len(items)}')


if __name__ == '__main__':
    main()
