#!/usr/bin/env python3
"""
再計算済みのxlsxを読み、Excelの数式とは別にPythonで同じ計算をやり直して突き合わせる。

  python3 verify.py 売店発注ツール_サンプルデータ入り.xlsx

見るもの
  ・商品の並び（商品マスタの順 → 当日CSVだけ → 1ヶ月前だけ → 2ヶ月前だけ）
  ・在庫（CSVの総数、数えた在庫の上書き）、定数（マスタ → CSVの規定数）
  ・14日の在庫予測（減り数＝100人あたりの減り数×動員÷100、0で止める、入庫は2便まで）
  ・切れる日（このままだと／入庫後）、目安（切り上げ）、目安超え・定数超え
  ・「このままだと」「入庫を入れると」の記号（● △ ○ ？ － ⚠）
  ・③発注書の明細（1便目→2便目の順）と合計、上部の「発注予定額」
"""
import datetime
import math
import sys
from collections import OrderedDict

import openpyxl

DAYS = 14
ORD_FIRST = 10
GRID_FIRST = 16                    # P列
COL = {'name': 3, 'par': 4, 'stock': 5, 'use': 6, 'cut': 7, 'status': 8, 'in_date': 9, 'in_qty': 10,
       'unit': 11, 'guide': 12, 'in_date2': 13, 'in_qty2': 14, 'after': 15,
       'counted': 30, 'counted_date': 31, 'code': 32, 'vendor': 33}


def as_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    return v


def num(v, default=0):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else default


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def main(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    st = wb['設定']
    theater = str(st['C4'].value)
    base = as_date(st['C6'].value)
    lt = int(num(st['C7'].value, 0))
    std = num(st['C8'].value, 0)
    assert isinstance(base, datetime.date), f'当日基準日が取れていません: {base!r}'

    def csv_rows(name):
        ws = wb[name]
        out = []
        for r in range(6, ws.max_row + 1):
            th = ws.cell(r, 2).value
            if th in (None, ''):
                continue
            th = str(th).zfill(4) if isinstance(th, int) else str(th)
            code = ws.cell(r, 9).value
            code = str(code) if not isinstance(code, float) else str(int(code))
            out.append({'th': th, 'code': code, 'name': ws.cell(r, 10).value, 'cat': ws.cell(r, 8).value,
                        'vendor': ws.cell(r, 5).value, 'pack': num(ws.cell(r, 11).value),
                        'total': num(ws.cell(r, 14).value), 'kitei': num(ws.cell(r, 15).value),
                        'price': num(ws.cell(r, 17).value)})
        return out

    sheets = {'cur': csv_rows('①在庫を貼る'), 'mid': csv_rows('月1回_1ヶ月前の在庫'), 'prv': csv_rows('月1回_2ヶ月前の在庫')}
    mine = {k: [x for x in v if x['th'] == theater] for k, v in sheets.items()}
    seen_any = {x['code'] for v in mine.values() for x in v}

    # 商品マスタ
    im = wb['商品マスタ']
    master = OrderedDict()
    for r in range(6, im.max_row + 1):
        code = im.cell(r, 2).value
        if code in (None, ''):
            continue
        master[str(code)] = {'name': im.cell(r, 3).value, 'cat': im.cell(r, 4).value, 'vendor': im.cell(r, 5).value,
                             'pack': im.cell(r, 6).value, 'price': im.cell(r, 7).value,
                             'par_case': im.cell(r, 8).value, 'pi': im.cell(r, 10).value}

    # 並び：マスタの取扱中 → 当日だけ → 1ヶ月前だけ → 2ヶ月前だけ
    order = [c for c in master if c in seen_any]
    for key in ('cur', 'mid', 'prv'):
        for x in mine[key]:
            if x['code'] not in master and x['code'] not in order:
                order.append(x['code'])

    def csv_text(code, field):
        for key in ('cur', 'mid', 'prv'):
            for x in sheets[key]:            # 文字は劇場を問わず最初の行（数式と同じ）
                if x['code'] == code:
                    return x[field]
        return ''

    def csv_num(code, field):
        for key in ('cur', 'mid', 'prv'):
            hits = [x[field] for x in mine[key] if x['code'] == code]
            if hits:
                return sum(hits) / len(hits)
        return 0

    # 動員
    pl = wb['動員を入れる']
    att = {}
    for r in range(16, 416):
        d = as_date(pl.cell(r, 2).value)
        if isinstance(d, datetime.date):
            att[d] = num(pl.cell(r, 7).value)
    days = [base + datetime.timedelta(days=d) for d in range(DAYS)]
    day_att = [att.get(d, std) for d in days]
    att_sum = sum(day_att)

    o = wb['②発注する']
    problems = []
    checked = 0
    sheet1, sheet2 = [], []
    amount = 0
    ordered = 0
    for i, code in enumerate(order):
        if i >= 400:
            break
        r = ORD_FIRST + i
        m = master.get(code, {})
        in_cur = any(x['code'] == code for x in mine['cur'])

        def mval(field, csv_field, numeric):
            v = m.get(field)
            if v not in (None, ''):
                return v
            return csv_num(code, csv_field) if numeric else csv_text(code, csv_field)
        name = mval('name', 'name', False)
        vendor = mval('vendor', 'vendor', False)
        pack = max(1, num(mval('pack', 'pack', True)))
        price = num(mval('price', 'price', True))
        pi = m.get('pi')
        pi = pi if is_num(pi) and pi > 0 else None
        par_case = m.get('par_case')
        if is_num(par_case) and par_case > 0:
            par = par_case * pack
        elif csv_num(code, 'kitei') > 0:
            par = csv_num(code, 'kitei')
        else:
            par = None
        stock = sum(x['total'] for x in mine['cur'] if x['code'] == code) if in_cur else 0
        cell = lambda n: o.cell(r, COL[n]).value       # noqa: E731
        counted, counted_date = cell('counted'), as_date(cell('counted_date'))
        if is_num(counted) and isinstance(counted_date, datetime.date) and counted_date >= base:
            stock = counted
        d1, q1 = as_date(cell('in_date')), cell('in_qty')
        d2, q2 = as_date(cell('in_date2')), cell('in_qty2')
        u1 = max(0, q1) * pack if isinstance(d1, datetime.date) and is_num(q1) else 0
        u2 = max(0, q2) * pack if isinstance(d2, datetime.date) and is_num(q2) else 0
        start = max(0, stock)
        use = [max(0, pi) * max(0, a) / 100 if pi is not None else 0 for a in day_att]
        b, a = [], []
        pb = pa = start
        for d in range(DAYS):
            arrive = (u1 if d1 == days[d] else 0) + (u2 if d2 == days[d] else 0)
            pb = max(0, pb - use[d])
            pa = max(0, pa + arrive - use[d])
            b.append(pb)
            a.append(pa)
        n_before = sum(1 for x in b if x >= 1) if pi is not None else None
        cut_before = base + datetime.timedelta(days=n_before) if n_before is not None and n_before < DAYS else None
        last_in = max([x for x in (d1, d2) if isinstance(x, datetime.date)], default=None)
        post_idx = [d for d in range(DAYS) if a[d] < 1 and (last_in is None or days[d] >= last_in)]
        n_post = min(post_idx) if pi is not None and post_idx else None
        cut_post = base + datetime.timedelta(days=n_post) if n_post is not None else None
        # 目安
        target = d1 if isinstance(d1, datetime.date) else base + datetime.timedelta(days=lt)
        prev1 = None
        if target in days:
            k = days.index(target)
            prev1 = start if k == 0 else b[k - 1]
        guide = None
        if par is not None and prev1 is not None:
            guide = 0 if par - prev1 <= 0 else math.ceil((par - prev1) / pack - 1e-9)
        over1 = max(0, q1 - guide) if guide is not None and is_num(q1) else 0
        over2 = 0
        if par is not None and u2 > 0 and d2 in days:
            k = days.index(d2)
            prev2 = start if k == 0 else a[k - 1]
            over2 = math.ceil(max(0, prev2 + u2 - par) / pack - 1e-9)

        def pair_err(d, q, over):
            if d in (None, '') and q in (None, ''):
                return ''
            if d in (None, '') or not isinstance(d, datetime.date):
                return '⚠'
            if d < base or d > days[-1]:
                return '⚠'
            if num(q) <= 0:
                return '⚠'
            if over > 0:
                return '⚠'
            return ''
        err1, err2 = pair_err(d1, q1, over1), pair_err(d2, q2, over2)
        valid1 = err1 == '' and num(q1) > 0 and isinstance(d1, datetime.date)
        valid2 = err2 == '' and num(q2) > 0 and isinstance(d2, datetime.date)
        amt = (u1 if valid1 else 0) * price + (u2 if valid2 else 0) * price
        amount += amt
        ordered += 1 if (valid1 or valid2) else 0

        if not in_cur:
            sym = '－'
        elif stock <= 0:
            sym = '●'
        elif pi is None or att_sum == 0:
            sym = '？'
        elif cut_before is None:
            sym = '○'
        elif n_before <= lt:
            sym = '●'
        else:
            sym = '△'
        if all(x in (None, '') for x in (d1, q1, d2, q2)):
            after = ''
        elif d1 in (None, '') and q1 in (None, ''):
            after = '⚠'
        elif err1 or err2:
            after = '⚠'
        elif pi is None or att_sum == 0:
            after = '？'
        elif cut_before is not None and d1 > cut_before:
            after = '●' if d1 > base + datetime.timedelta(days=lt) else '△'
        elif cut_post is None:
            after = '○'
        elif cut_post - datetime.timedelta(days=lt) <= base:
            after = '●'
        else:
            after = '△'

        got = {k: cell(k) for k in ('name', 'par', 'stock', 'guide', 'code')}
        got['cut'] = as_date(cell('cut'))
        got['status'] = cell('status') or ''
        got['after'] = cell('after') or ''
        got['grid'] = [o.cell(r, GRID_FIRST + d).value for d in range(DAYS)]

        def bad(what, exp, act):
            problems.append(f'行{r} {code} {name}: {what} 期待={exp!r} 実際={act!r}')

        if str(got['code']) != code:
            bad('商品コード', code, got['code'])
        if got['name'] != name:
            bad('商品名', name, got['name'])
        if abs(num(got['stock']) - stock) > 1e-6:
            bad('在庫', stock, got['stock'])
        if par is None:
            if got['par'] not in (None, ''):
                bad('定数', '', got['par'])
        elif abs(num(got['par']) - par) > 1e-6:
            bad('定数', par, got['par'])
        if got['cut'] != cut_before:
            bad('切れる日', cut_before, got['cut'])
        if not got['status'].startswith(sym):
            bad('このままだと', sym, got['status'])
        if sym in '●△' and stock > 0 and cut_before is not None and cut_before.strftime('%-m/%-d') not in got['status']:
            bad('このままだと の日付', cut_before, got['status'])
        if after == '':
            if got['after'] != '':
                bad('入庫を入れると', '', got['after'])
        elif not got['after'].startswith(after):
            bad('入庫を入れると', after, got['after'])
        if guide is None:
            if got['guide'] not in (None, ''):
                bad('目安', '', got['guide'])
        elif num(got['guide'], -1) != guide:
            bad('目安', guide, got['guide'])
        if pi is None:
            if any(x not in (None, '') for x in got['grid']):
                bad('14日の表(減り数未設定)', '空', got['grid'])
        else:
            for d in range(DAYS):
                if abs(num(got['grid'][d]) - a[d]) > 1e-6:
                    bad(f'14日の表 {d}日目', round(a[d], 3), got['grid'][d])
                    break
        if valid1:
            sheet1.append((vendor, code, d1, q1, u1 * price))
        if valid2:
            sheet2.append((vendor, code, d2, q2, u2 * price))
        checked += 1

    # ---- ③発注書 ----
    sh = wb['③発注書']
    sel = sh['C4'].value or 'すべて'
    exp_rows = [x for x in sheet1 + sheet2 if sel in ('すべて', '') or x[0] == sel]
    got_rows = []
    for r in range(11, 111):
        if sh.cell(r, 2).value in (None, ''):
            continue
        got_rows.append((sh.cell(r, 3).value, str(sh.cell(r, 4).value), as_date(sh.cell(r, 6).value),
                         sh.cell(r, 8).value, sh.cell(r, 10).value))
    if len(exp_rows) != len(got_rows):
        problems.append(f'③発注書 行数 期待={len(exp_rows)} 実際={len(got_rows)}')
    else:
        for e, g in zip(exp_rows, got_rows):
            if (e[0], e[1], e[2], e[3]) != (g[0], g[1], g[2], g[3]) or abs(e[4] - num(g[4])) > 0.5:
                problems.append(f'③発注書 明細 期待={e} 実際={g}')
                break
    total_exp = sum(x[4] for x in exp_rows)
    if abs(total_exp - num(sh['J8'].value)) > 0.5:
        problems.append(f'③発注書 合計 期待={total_exp} 実際={sh["J8"].value}')
    if abs(amount - num(o['M5'].value)) > 0.5:
        problems.append(f'② 発注予定額 期待={amount} 実際={o["M5"].value}')
    if ordered != num(o['I5'].value):
        problems.append(f'② 入庫を入れた 期待={ordered} 実際={o["I5"].value}')
    crit_exp = sum(1 for r in range(ORD_FIRST, ORD_FIRST + 400) if str(o.cell(r, COL['status']).value or '').startswith('●'))
    if num(o['C5'].value) != crit_exp:
        problems.append(f'まとめ ● 期待={crit_exp} 実際={o["C5"].value}')

    print(f'{path}: 商品 {checked} 件を突き合わせ　動員合計 {att_sum:.0f}　発注書 {len(got_rows)} 行　問題 {len(problems)} 件')
    for p in problems[:30]:
        print('  ', p)
    return len(problems)


if __name__ == '__main__':
    sys.exit(1 if main(sys.argv[1] if len(sys.argv) > 1 else '売店発注ツール_サンプルデータ入り.xlsx') else 0)
