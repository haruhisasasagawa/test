/* =========================================================
   月次予算割振アプリ

   主役は「項目別の週別割振」。
   各項目は明細（時給・時間・人数）に対して「どの週に何回実施するか」を
   入力すると、週別金額・月計・予算比・差額まで自動で出る。

   週の区切り : 金曜〜木曜（月初・月末は端数週）
   WEEK番号   : 3/1 を含む週を第1週とする年度通し番号
   ========================================================= */

'use strict';

/* ---------------------------------------------------------
   定数・既定値
   --------------------------------------------------------- */

const STORAGE_PREFIX = 'toho-budget:';
const DEPTS = ['FLOOR', 'CONCESSION', 'STORE', 'OFFICE'];
const WDAY = ['日', '月', '火', '水', '木', '金', '土'];

const DEFAULT_DEPT_RATIO = {
  FLOOR: 0.442736342367074,
  CONCESSION: 0.4321965545697439,
  STORE: 0.03629238541840829,
  OFFICE: 0.08877471764477365,
};

/** 出力ファイル名のベース（日本語名はブラウザによって欠落するためASCIIで固定） */
const FILE_BASE = {
  weekSummaryTable: 'weekly-attendance',
  hallTable: 'hall-allocation',
  summaryTable: 'summary-by-item',
  summaryHallTable: 'summary-hall',
  planTable: 'plan-detail',
};

const MODES = [
  { value: 'plan', label: '実施予定から積み上げ' },
  { value: 'ratio', label: '月予算を動員構成比で配分' },
  { value: 'even', label: '月予算を週均等で配分' },
  { value: 'manual', label: '週別に直接入力' },
];

/**
 * 既定の項目。init は初回に週別回数へ展開する指定。
 *   {everyWeek: n} … 毎週 n 回（＝週n日）
 *   {atWeek: i}    … i 番目の週に 1 回
 */
function defaultCategories() {
  const it = (label, wage, hours, people, init) => ({ label, wage, hours, people, init: init || null });
  return [
    { name: '障がい者雇用', mode: 'plan', budget: 0, items: [it('駒澤さん', 1236, 5, 1, { everyWeek: 3 })] },
    { name: '新規採用研修（トレーニー）', mode: 'plan', budget: 0, items: [it('週2〜3名', 1226, 5.5, 2, null), it('カンポリ参加', 1226, 2.15, 3, null)] },
    { name: '新規採用研修（トレーナー）', mode: 'plan', budget: 0, items: [it('週2〜3名', 1251, 5.5, 1, null), it('通常', 1251, 5, 4, null)] },
    { name: '既存スタッフ研修', mode: 'plan', budget: 0, items: [it('週1', 1251, 1, 1, { atWeek: 0 }), it('週2', 1251, 1.5, 3, { atWeek: 1 }), it('週3', 1251, 1.75, 6, { atWeek: 2 }), it('週4', 1251, 1.5, 6, { atWeek: 3 })] },
    { name: '防災訓練', mode: 'plan', budget: 0, items: [it('週1', 1251, 0.5, 10, { atWeek: 0 }), it('週2', 1251, 0.5, 15, { atWeek: 1 }), it('週3', 1251, 0.75, 8, { atWeek: 2 }), it('週4', 1251, 0.75, 2, { atWeek: 3 }), it('総合訓練', 1251, 1, 8, null)] },
    { name: '棚卸（予算分）', mode: 'plan', budget: 0, items: [it('棚卸', 1251, 5, 5, null)] },
    { name: 'ストア準備・返品（予算分）', mode: 'plan', budget: 0, items: [it('※平均時給以外は変更不可', 1251, 2, 1, { everyWeek: 1 })] },
    { name: 'ストア準備・返品（追加分）', mode: 'plan', budget: 0, items: [it('準備', 1251, 2, 2, null), it('返品', 1251, 6, 2, null), it('陳列', 1251, 6, 4, null)] },
    { name: 'その他', mode: 'plan', budget: 0, items: [it('', 1251, 0, 0, null)] },
    { name: 'リーダー手当', mode: 'manual', budget: 0, items: [], weekly: [] },
  ];
}

/* ---------------------------------------------------------
   ユーティリティ
   --------------------------------------------------------- */

const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

function el(tag, attrs = {}, children = []) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null || v === false) continue;
    if (k === 'class') node.className = v;
    else if (k === 'text') node.textContent = v;
    else if (k === 'dataset') Object.assign(node.dataset, v);
    else if (k.startsWith('on')) node.addEventListener(k.slice(2), v);
    else node.setAttribute(k, v);
  }
  for (const c of [].concat(children)) {
    if (c === null || c === undefined) continue;
    node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
  }
  return node;
}

const toNum = (v) => {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  if (v === null || v === undefined) return 0;
  const n = parseFloat(String(v).replace(/[,\s¥円%]/g, ''));
  return isFinite(n) ? n : 0;
};

const fmt = (n) => (isFinite(n) ? Math.round(n).toLocaleString('ja-JP') : '-');
const fmtRaw = (n) => (isFinite(n) ? String(Math.round(n * 100) / 100) : '');
const fmtPct = (n, digits = 1) => (isFinite(n) ? (n * 100).toFixed(digits) + '%' : '-');

/** 合計が total にぴったり一致する整数配分（最大剰余法） */
function allocate(total, weights) {
  const n = weights.length;
  if (n === 0) return [];
  let w = weights.map((x) => (isFinite(x) && x > 0 ? x : 0));
  let sum = w.reduce((a, b) => a + b, 0);
  if (sum <= 0) { w = new Array(n).fill(1); sum = n; }
  const target = Math.round(total);
  const exact = w.map((x) => (target * x) / sum);
  const base = exact.map((x) => Math.floor(x));
  let rest = target - base.reduce((a, b) => a + b, 0);
  const order = exact.map((x, i) => ({ i, frac: x - Math.floor(x) })).sort((a, b) => b.frac - a.frac);
  for (let k = 0; k < order.length && rest > 0; k++, rest--) base[order[k].i] += 1;
  for (let k = order.length - 1; k >= 0 && rest < 0; k--, rest++) base[order[k].i] -= 1;
  return base;
}

const ymKey = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

/* ---------------------------------------------------------
   週の計算（金〜木、月初/月末は端数）
   --------------------------------------------------------- */

/** 3/1 を含む週（金〜木）を第1週とする年度通し WEEK 番号 */
function fiscalWeekNo(date) {
  const thu = new Date(date);
  thu.setDate(thu.getDate() + ((4 - thu.getDay() + 7) % 7));
  for (const y of [thu.getFullYear(), thu.getFullYear() - 1]) {
    const mar1 = new Date(y, 2, 1);
    const firstThu = new Date(mar1);
    firstThu.setDate(firstThu.getDate() + ((4 - firstThu.getDay() + 7) % 7));
    if (thu >= firstThu) return Math.floor((thu - firstThu) / 604800000) + 1;
  }
  return 1;
}

function buildWeeks(ym, startNo) {
  const [y, m] = ym.split('-').map(Number);
  const last = new Date(y, m, 0).getDate();
  const weeks = [];
  let cur = [];
  for (let d = 1; d <= last; d++) {
    const dt = new Date(y, m - 1, d);
    cur.push(dt);
    if (dt.getDay() === 4 || d === last) { weeks.push(cur); cur = []; }
  }
  return weeks.map((dates, i) => ({
    index: i,
    no: startNo + i,
    dates,
    label: `W${startNo + i}`,
    range: `${dates[0].getMonth() + 1}/${dates[0].getDate()}〜${dates[dates.length - 1].getMonth() + 1}/${dates[dates.length - 1].getDate()}`,
  }));
}

/* ---------------------------------------------------------
   State
   --------------------------------------------------------- */

function emptyState(ym) {
  const [y, m] = ym.split('-').map(Number);
  return {
    ym,
    weekStartNo: fiscalWeekNo(new Date(y, m - 1, 1)),
    hallBudget: 0,
    hallWeeklyManual: false,
    hallWeekly: [],
    deptRatio: { ...DEFAULT_DEPT_RATIO },
    attendance: {},
    extra: { FLOOR: [], CONCESSION: [], STORE: [], OFFICE: [] },
    categories: defaultCategories(),
    leave: { unit: 979, hours: 5, budget: 0, people: [] },
    changing: { unit: 1132, minutes: 9, base: 81178, people: [] },
  };
}

let state = null;
let weeks = [];
let outs = new Map();

/** 明細の週別回数を、週数が確定したこの時点で初期化・旧形式から移行する */
function normalizeItems() {
  for (const cat of state.categories) {
    if (!Array.isArray(cat.items)) cat.items = [];
    for (const item of cat.items) {
      if (Array.isArray(item.weekCounts)) continue;
      const counts = weeks.map(() => 0);
      if (item.init && item.init.everyWeek) counts.forEach((_, i) => { counts[i] = item.init.everyWeek; });
      else if (item.init && Number.isInteger(item.init.atWeek)) {
        if (item.init.atWeek < counts.length) counts[item.init.atWeek] = 1;
      } else if (toNum(item.weeks) > 0) {
        // 旧形式（週＝回数）からの移行：先頭の週から1回ずつ置く
        let n = Math.round(toNum(item.weeks));
        for (let i = 0; i < counts.length && n > 0; i++, n--) counts[i] = 1;
      }
      item.weekCounts = counts;
      delete item.init;
      delete item.weeks;
    }
    if (!MODES.some((m) => m.value === cat.mode)) cat.mode = cat.items.length ? 'plan' : 'manual';
  }
}

function refreshOuts() {
  outs = new Map();
  $$('[data-out]').forEach((n) => outs.set(n.dataset.out, n));
}

function setOut(key, text, cls) {
  const node = outs.get(key);
  if (!node) return;
  node.textContent = text;
  node.classList.remove('pos', 'neg');
  if (cls) node.classList.add(cls);
}

/* ---------------------------------------------------------
   永続化
   --------------------------------------------------------- */

function savedMonths() {
  return Object.keys(localStorage)
    .filter((k) => k.startsWith(STORAGE_PREFIX))
    .map((k) => k.slice(STORAGE_PREFIX.length))
    .filter((k) => /^\d{4}-\d{2}$/.test(k))
    .sort()
    .reverse();
}

function saveState(silent) {
  localStorage.setItem(STORAGE_PREFIX + state.ym, JSON.stringify(state));
  localStorage.setItem(STORAGE_PREFIX + '@last', state.ym);
  const t = new Date();
  $('#savedState').textContent = `保存済み ${String(t.getHours()).padStart(2, '0')}:${String(t.getMinutes()).padStart(2, '0')}`;
  refreshSavedMonths();
  if (!silent) toast('保存しました');
}

function loadState(ym) {
  const raw = localStorage.getItem(STORAGE_PREFIX + ym);
  if (!raw) return emptyState(ym);
  try {
    return Object.assign(emptyState(ym), JSON.parse(raw), { ym });
  } catch (e) {
    return emptyState(ym);
  }
}

function refreshSavedMonths() {
  const sel = $('#savedMonths');
  const cur = sel.value;
  sel.innerHTML = '';
  sel.appendChild(el('option', { value: '', text: '-' }));
  for (const m of savedMonths()) {
    const [y, mm] = m.split('-');
    sel.appendChild(el('option', { value: m, text: `${y}年${Number(mm)}月` }));
  }
  sel.value = cur;
}

let saveTimer = null;
function autoSave() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => saveState(true), 600);
}

function toast(msg) {
  const t = $('#toast');
  t.textContent = msg;
  t.hidden = false;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => { t.hidden = true; }, 2400);
}

/* ---------------------------------------------------------
   計算
   --------------------------------------------------------- */

const itemUnit = (it) => toNum(it.wage) * toNum(it.hours) * toNum(it.people);

function calc() {
  const weekAtt = weeks.map((w) => w.dates.reduce((s, d) => s + toNum(state.attendance[ymKey(d)]), 0));
  const attTotal = weekAtt.reduce((a, b) => a + b, 0);
  const ratio = weekAtt.map((v) => (attTotal > 0 ? v / attTotal : 0));

  // ---- 項目別 ----
  const cats = state.categories.map((c) => {
    const items = c.items.map((it) => {
      const unit = itemUnit(it);
      const counts = weeks.map((_, i) => toNum((it.weekCounts || [])[i]));
      const perWeek = counts.map((n) => unit * n);
      return { ref: it, unit, counts, perWeek, times: counts.reduce((a, b) => a + b, 0), total: perWeek.reduce((a, b) => a + b, 0) };
    });
    let weekly;
    if (c.mode === 'plan') {
      // 週別の端数は最大剰余法で処理し、月計が実際の合計（四捨五入）と一致するようにする
      const exact = weeks.map((_, i) => items.reduce((s, it) => s + it.perWeek[i], 0));
      weekly = allocate(exact.reduce((a, b) => a + b, 0), exact);
    }
    else if (c.mode === 'ratio') weekly = allocate(toNum(c.budget), weekAtt);
    else if (c.mode === 'even') weekly = allocate(toNum(c.budget), weeks.map(() => 1));
    else weekly = weeks.map((_, i) => toNum((c.weekly || [])[i]));
    const total = weekly.reduce((a, b) => a + b, 0);
    const budget = toNum(c.budget);
    return { ref: c, items, weekly, total, budget, rate: budget > 0 ? total / budget : NaN, diff: budget - total };
  });

  const itemsWeekly = weeks.map((_, i) => cats.reduce((s, cc) => s + cc.weekly[i], 0));
  const itemsTotal = itemsWeekly.reduce((a, b) => a + b, 0);
  const itemsBudget = cats.reduce((s, cc) => s + cc.budget, 0);

  // ---- 接客部門 ----
  const hallWeekly = state.hallWeeklyManual
    ? weeks.map((_, i) => toNum(state.hallWeekly[i]))
    : allocate(state.hallBudget, attTotal > 0 ? weekAtt : weeks.map(() => 1));

  const deptAlloc = {};
  const deptExtra = {};
  const deptTotal = {};
  const ratios = DEPTS.map((d) => toNum(state.deptRatio[d]));
  weeks.forEach((_, i) => {
    const split = allocate(hallWeekly[i], ratios);
    DEPTS.forEach((d, di) => {
      (deptAlloc[d] = deptAlloc[d] || [])[i] = split[di];
      const ex = toNum((state.extra[d] || [])[i]);
      (deptExtra[d] = deptExtra[d] || [])[i] = ex;
      (deptTotal[d] = deptTotal[d] || [])[i] = split[di] + ex;
    });
  });
  const hallWeekTotal = weeks.map((_, i) => DEPTS.reduce((s, d) => s + deptTotal[d][i], 0));
  const hallMonthTotal = hallWeekTotal.reduce((a, b) => a + b, 0);
  const hallRate = state.hallBudget > 0 ? hallMonthTotal / state.hallBudget : NaN;

  // ---- 有休 ----
  const lv = state.leave;
  const leaveRows = weeks.map((_, i) => {
    const people = toNum(lv.people[i]);
    return { people, amount: toNum(lv.unit) * toNum(lv.hours) * people };
  });
  const leaveBase = leaveRows.reduce((s, r) => s + r.amount, 0);
  const leaveBudget = toNum(lv.budget);
  const leaveGap = leaveBudget - leaveBase;
  const leaveShare = weeks.length > 0 ? leaveGap / weeks.length : 0;
  leaveRows.forEach((r) => {
    r.adjusted = r.amount + leaveShare;
    r.perDay = r.people > 0 ? r.adjusted / r.people : NaN;
  });
  const leaveTotal = leaveRows.reduce((s, r) => s + r.adjusted, 0);

  // ---- 更衣時間 ----
  const cg = state.changing;
  const changeRows = weeks.map((_, i) => {
    const people = toNum(cg.people[i]);
    return { people, amount: toNum(cg.unit) * (toNum(cg.minutes) / 60) * people };
  });
  const changeTotal = changeRows.reduce((s, r) => s + r.amount, 0);
  const changeBudget = toNum(cg.base) * (isFinite(hallRate) ? hallRate : 0);

  return {
    weekAtt, attTotal, ratio,
    cats, itemsWeekly, itemsTotal, itemsBudget,
    hallWeekly, deptAlloc, deptExtra, deptTotal, hallWeekTotal, hallMonthTotal, hallRate,
    leaveRows, leaveBase, leaveBudget, leaveGap, leaveShare, leaveTotal,
    changeRows, changeTotal, changeBudget,
  };
}

/* ---------------------------------------------------------
   入力セル生成ヘルパ
   --------------------------------------------------------- */

function numInput(value, onChange, opts = {}) {
  const input = el('input', {
    type: 'text',
    inputmode: opts.decimal ? 'decimal' : 'numeric',
    class: 'num' + (opts.class ? ' ' + opts.class : ''),
    value: value === '' || value === undefined || value === null ? '' : String(value),
  });
  input.addEventListener('input', () => onChange(toNum(input.value)));
  input.addEventListener('change', () => {
    if (input.value.trim() === '') { onChange(0); return; }
    input.value = fmtRaw(toNum(input.value));
    onChange(toNum(input.value));
  });
  return input;
}

function textInput(value, onChange) {
  const input = el('input', { type: 'text', class: 'text', value: value ?? '' });
  input.addEventListener('input', () => onChange(input.value));
  return input;
}

function cellOut(key, cls) {
  return el('td', { class: 'out' + (cls ? ' ' + cls : ''), dataset: { out: key }, text: '-' });
}

/* ---------------------------------------------------------
   画面構築
   --------------------------------------------------------- */

function buildAll() {
  weeks = buildWeeks(state.ym, toNum(state.weekStartNo) || 1);
  normalizeItems();

  const [y, m] = state.ym.split('-').map(Number);
  buildYearOptions(y);
  $('#selYear').value = String(y);
  $('#selMonth').value = String(m);
  $('#weekStartNo').value = state.weekStartNo;
  $('#monthTitle').textContent = `${y}年${m}月（${weeks.length}週：${weeks[0].label}〜${weeks[weeks.length - 1].label}）`;

  $('#hallBudget').value = state.hallBudget ? fmtRaw(state.hallBudget) : '';
  $('#hallWeeklyManual').checked = !!state.hallWeeklyManual;
  $('#leaveUnit').value = fmtRaw(state.leave.unit);
  $('#leaveHours').value = fmtRaw(state.leave.hours);
  $('#leaveBudget').value = state.leave.budget ? fmtRaw(state.leave.budget) : '';
  $('#changeUnit').value = fmtRaw(state.changing.unit);
  $('#changeMinutes').value = fmtRaw(state.changing.minutes);
  $('#changeBase').value = fmtRaw(state.changing.base);

  buildWeekStrip();
  buildCategories();
  buildDeptRatio();
  buildAttendance();
  buildWeekSummary();
  buildWeekList();
  buildHall();
  buildLeave();
  buildChanging();
  buildSummary();

  refreshOuts();
  recalc();
}

function buildYearOptions(current) {
  const sel = $('#selYear');
  const now = new Date().getFullYear();
  const years = new Set();
  for (let y = now - 3; y <= now + 2; y++) years.add(y);
  years.add(current);
  savedMonths().forEach((m) => years.add(Number(m.split('-')[0])));
  sel.innerHTML = '';
  Array.from(years).sort().forEach((y) => sel.appendChild(el('option', { value: String(y), text: `${y}年` })));
}

function buildWeekStrip() {
  const thead = $('#weekStripTable thead');
  const tbody = $('#weekStripTable tbody');
  thead.innerHTML = '';
  tbody.innerHTML = '';
  thead.appendChild(el('tr', {}, [
    el('th', { class: 'label', text: '週' }),
    ...weeks.map((w) => el('th', { text: w.label })),
    el('th', { text: '月計' }),
  ]));
  tbody.appendChild(el('tr', {}, [
    el('th', { class: 'label', text: '期間' }),
    ...weeks.map((w) => el('td', { class: 'center', text: w.range })),
    el('td', { class: 'center', text: '-' }),
  ]));
  tbody.appendChild(el('tr', {}, [
    el('th', { class: 'label', text: '動員構成比' }),
    ...weeks.map((_, i) => cellOut(`strip.ratio.${i}`)),
    cellOut('strip.ratio.total'),
  ]));
  tbody.appendChild(el('tr', { class: 'total-row' }, [
    el('th', { class: 'label', text: '項目合計' }),
    ...weeks.map((_, i) => cellOut(`strip.items.${i}`)),
    cellOut('strip.items.total'),
  ]));
}

function buildDeptRatio() {
  const tb = $('#deptRatioTable tbody');
  tb.innerHTML = '';
  for (const d of DEPTS) {
    tb.appendChild(el('tr', {}, [
      el('th', { class: 'label', text: d }),
      el('td', {}, [numInput(state.deptRatio[d], (v) => { state.deptRatio[d] = v; onChange(); }, { decimal: true })]),
      cellOut(`deptRatioPct.${d}`),
    ]));
  }
}

/* ---------- 項目カード ---------- */

function buildCategories() {
  const list = $('#categoryList');
  list.innerHTML = '';
  state.categories.forEach((cat, ci) => list.appendChild(buildCategoryCard(cat, ci)));
}

function buildCategoryCard(cat, ci) {
  const modeSelect = el('select', {}, MODES.map((m) => el('option', { value: m.value, text: m.label })));
  modeSelect.value = cat.mode;
  modeSelect.addEventListener('change', () => {
    cat.mode = modeSelect.value;
    rebuildDynamic();
  });

  const head = el('div', { class: 'cat-head' }, [
    el('div', { class: 'cat-name' }, [textInput(cat.name, (v) => { cat.name = v; onChange(); })]),
    el('label', { class: 'field' }, [el('span', { text: '月予算' }), numInput(cat.budget || '', (v) => { cat.budget = v; onChange(); })]),
    el('label', { class: 'field' }, [el('span', { text: '割振方法' }), modeSelect]),
    el('div', { class: 'stat' }, [el('span', { text: '月計' }), el('strong', { dataset: { out: `cat.${ci}.total` }, text: '-' })]),
    el('div', { class: 'stat' }, [el('span', { text: '予算比' }), el('strong', { dataset: { out: `cat.${ci}.rate` }, text: '-' })]),
    el('div', { class: 'stat' }, [el('span', { text: '差額' }), el('strong', { dataset: { out: `cat.${ci}.diff` }, text: '-' })]),
    el('div', { class: 'gauge', title: '予算に対する消化率' }, [el('span', { class: 'gauge-fill', dataset: { out: `cat.${ci}.gauge` } })]),
    el('button', {
      type: 'button', class: 'btn ghost mini', text: '削除',
      onclick: () => {
        if (!confirm(`「${cat.name}」を削除しますか？`)) return;
        state.categories.splice(ci, 1);
        rebuildDynamic();
      },
    }),
  ]);

  const body = cat.mode === 'plan' ? buildPlanTable(cat, ci) : buildAllocRow(cat, ci);
  return el('div', { class: 'cat-card' }, [head, el('div', { class: 'cat-body' }, [body])]);
}

/** 実施予定テーブル（明細 × 週別回数） */
function buildPlanTable(cat, ci) {
  const tbody = el('tbody');

  cat.items.forEach((item, ii) => {
    const cells = [
      el('td', {}, [textInput(item.label, (v) => { item.label = v; onChange(); })]),
      el('td', {}, [numInput(item.wage ?? '', (v) => { item.wage = v; onChange(); })]),
      el('td', {}, [numInput(item.hours ?? '', (v) => { item.hours = v; onChange(); }, { decimal: true })]),
      el('td', {}, [numInput(item.people ?? '', (v) => { item.people = v; onChange(); }, { decimal: true })]),
    ];
    weeks.forEach((_, wi) => {
      cells.push(el('td', { class: 'wk' }, [
        numInput((item.weekCounts || [])[wi] || '', (v) => {
          item.weekCounts = item.weekCounts || [];
          item.weekCounts[wi] = v;
          onChange();
        }, { decimal: true }),
      ]));
    });
    cells.push(cellOut(`cat.${ci}.item.${ii}.times`));
    cells.push(cellOut(`cat.${ci}.item.${ii}.total`));
    cells.push(el('td', { class: 'ops' }, [
      el('button', {
        type: 'button', class: 'btn ghost mini', text: '全週', title: '全ての週に1回ずつ入れる',
        onclick: () => { item.weekCounts = weeks.map(() => 1); rebuildDynamic(); },
      }),
      el('button', {
        type: 'button', class: 'btn ghost mini', text: '×', title: 'この行を削除',
        onclick: () => { cat.items.splice(ii, 1); rebuildDynamic(); },
      }),
    ]));
    tbody.appendChild(el('tr', {}, cells));
  });

  const table = el('table', { class: 'grid plan' }, [
    el('thead', {}, [
      el('tr', {}, [
        el('th', { class: 'label', rowspan: 2, text: '内容' }),
        el('th', { rowspan: 2, text: '時給' }),
        el('th', { rowspan: 2, text: '時間' }),
        el('th', { rowspan: 2, text: '人数' }),
        el('th', { class: 'center wk-group', colspan: weeks.length, text: '週ごとの回数（日数）' }),
        el('th', { rowspan: 2, text: '回数計' }),
        el('th', { rowspan: 2, text: '金額' }),
        el('th', { rowspan: 2, text: '' }),
      ]),
      el('tr', {}, weeks.map((w) => el('th', { class: 'wk', title: w.range, text: w.label }))),
    ]),
    tbody,
    el('tfoot', {}, [
      el('tr', { class: 'total-row' }, [
        el('th', { class: 'label', colspan: 4, text: '週別金額' }),
        ...weeks.map((_, i) => cellOut(`cat.${ci}.week.${i}`, 'wk')),
        el('td', {}),
        cellOut(`cat.${ci}.weekTotal`),
        el('td', {}),
      ]),
    ]),
  ]);

  return el('div', {}, [
    el('div', { class: 'table-scroll' }, [table]),
    el('div', { class: 'row wrap card-actions' }, [
      el('button', {
        type: 'button', class: 'btn ghost mini', text: '行を追加',
        onclick: () => { cat.items.push({ label: '', wage: 1251, hours: 0, people: 1, weekCounts: weeks.map(() => 0) }); rebuildDynamic(); },
      }),
      el('button', {
        type: 'button', class: 'btn ghost mini', text: '回数をすべてクリア',
        onclick: () => { cat.items.forEach((it) => { it.weekCounts = weeks.map(() => 0); }); rebuildDynamic(); },
      }),
      el('button', {
        type: 'button', class: 'btn ghost mini', text: '月計を月予算にセット',
        onclick: () => {
          const c = calc();
          cat.budget = c.cats[ci].total;
          rebuildDynamic();
        },
      }),
      el('span', { class: 'hint inline', text: '回数＝その週に何回（何日）実施するか。金額 ＝ 時給 × 時間 × 人数 × 回数' }),
    ]),
  ]);
}

/** 月予算を配分するモード用の週別金額行 */
function buildAllocRow(cat, ci) {
  const manual = cat.mode === 'manual';
  const table = el('table', { class: 'grid' }, [
    el('thead', {}, [el('tr', {}, [
      el('th', { class: 'label', text: '週' }),
      ...weeks.map((w) => el('th', { title: w.range, text: w.label })),
      el('th', { text: '月計' }),
    ])]),
    el('tbody', {}, [el('tr', {}, [
      el('th', { class: 'label', text: '金額' }),
      ...weeks.map((_, i) => (manual
        ? el('td', {}, [numInput((cat.weekly || [])[i] || '', (v) => {
            cat.weekly = cat.weekly || [];
            cat.weekly[i] = v;
            onChange();
          })])
        : cellOut(`cat.${ci}.week.${i}`))),
      cellOut(`cat.${ci}.weekTotal`),
    ])]),
  ]);

  const actions = [];
  if (manual) {
    actions.push(
      el('button', {
        type: 'button', class: 'btn ghost mini', text: '動員構成比で埋める',
        onclick: () => { cat.weekly = allocate(toNum(cat.budget), calc().weekAtt); rebuildDynamic(); },
      }),
      el('button', {
        type: 'button', class: 'btn ghost mini', text: '均等で埋める',
        onclick: () => { cat.weekly = allocate(toNum(cat.budget), weeks.map(() => 1)); rebuildDynamic(); },
      }),
      el('button', {
        type: 'button', class: 'btn ghost mini', text: '差額を各週へ均等調整',
        onclick: () => {
          const cc = calc().cats[ci];
          const add = allocate(cc.diff, weeks.map(() => 1));
          cat.weekly = weeks.map((_, i) => toNum((cat.weekly || [])[i]) + add[i]);
          rebuildDynamic();
        },
      }),
      el('button', {
        type: 'button', class: 'btn ghost mini', text: 'クリア',
        onclick: () => { cat.weekly = []; rebuildDynamic(); },
      })
    );
  } else {
    actions.push(el('span', { class: 'hint inline', text: '月予算を自動配分しています（合計は月予算とぴったり一致します）。' }));
  }
  if (cat.items.length) {
    actions.push(el('span', { class: 'hint inline', text: `※「実施予定から積み上げ」に切り替えると明細 ${cat.items.length} 行を使えます` }));
  }

  return el('div', {}, [
    el('div', { class: 'table-scroll' }, [table]),
    el('div', { class: 'row wrap card-actions' }, actions),
  ]);
}

/* ---------- 動員 ---------- */

function buildAttendance() {
  const tb = $('#attendanceTable tbody');
  tb.innerHTML = '';
  weeks.forEach((w) => {
    w.dates.forEach((d) => {
      const key = ymKey(d);
      const wd = d.getDay();
      tb.appendChild(el('tr', {}, [
        el('td', { class: 'label', text: `${d.getMonth() + 1}/${d.getDate()}` }),
        el('td', { class: 'center' + (wd === 0 || wd === 6 ? ' weekend' : ''), text: WDAY[wd] }),
        el('td', { class: 'center', text: w.label }),
        el('td', {}, [numInput(state.attendance[key] ?? '', (v) => { state.attendance[key] = v; onChange(); })]),
        cellOut(`attPct.${key}`),
      ]));
    });
  });
}

function buildWeekSummary() {
  const thead = $('#weekSummaryTable thead');
  const tb = $('#weekSummaryTable tbody');
  thead.innerHTML = '';
  tb.innerHTML = '';
  thead.appendChild(el('tr', {}, [
    el('th', { class: 'label', text: '週' }),
    ...weeks.map((w) => el('th', { text: `${w.label} (${w.range})` })),
    el('th', { text: '月計' }),
  ]));
  for (const [label, key] of [['動員', 'weekAtt'], ['構成比', 'weekRatio'], ['日数', 'weekDays']]) {
    tb.appendChild(el('tr', {}, [
      el('th', { class: 'label', text: label }),
      ...weeks.map((_, i) => cellOut(`${key}.${i}`)),
      cellOut(`${key}.total`),
    ]));
  }
}

function buildWeekList() {
  const tb = $('#weekListTable tbody');
  tb.innerHTML = '';
  weeks.forEach((w, i) => {
    tb.appendChild(el('tr', {}, [
      el('th', { class: 'label', text: w.label }),
      el('td', { class: 'center', text: w.range }),
      el('td', { text: String(w.dates.length) }),
      cellOut(`wl.att.${i}`),
      cellOut(`wl.ratio.${i}`),
    ]));
  });
}

/* ---------- 接客部門 ---------- */

function buildHall() {
  const thead = $('#hallTable thead');
  const tb = $('#hallTable tbody');
  thead.innerHTML = '';
  tb.innerHTML = '';

  thead.appendChild(el('tr', {}, [
    el('th', { class: 'label', text: '項目' }),
    ...weeks.map((w) => el('th', { text: `${w.label} ${w.range}` })),
    el('th', { text: '月計' }),
  ]));

  const groupRow = (text) => el('tr', { class: 'group-head' }, [el('th', { class: 'label', colspan: weeks.length + 2, text })]);

  tb.appendChild(groupRow('週予算'));
  tb.appendChild(el('tr', {}, [
    el('th', { class: 'label', text: '週予算' }),
    ...weeks.map((_, i) => (state.hallWeeklyManual
      ? el('td', {}, [numInput(state.hallWeekly[i] ?? '', (v) => { state.hallWeekly[i] = v; onChange(); })])
      : cellOut(`hallWeekly.${i}`))),
    cellOut('hallWeekly.total'),
  ]));

  tb.appendChild(groupRow('予算配分（週予算 × 部門構成比）'));
  for (const d of DEPTS) {
    tb.appendChild(el('tr', {}, [
      el('th', { class: 'label', text: d }),
      ...weeks.map((_, i) => cellOut(`deptAlloc.${d}.${i}`)),
      cellOut(`deptAlloc.${d}.total`),
    ]));
  }

  tb.appendChild(groupRow('予算外（ミーティング・棚卸・返品など）'));
  for (const d of DEPTS) {
    tb.appendChild(el('tr', {}, [
      el('th', { class: 'label', text: d }),
      ...weeks.map((_, i) => el('td', {}, [
        numInput((state.extra[d] || [])[i] ?? '', (v) => {
          state.extra[d] = state.extra[d] || [];
          state.extra[d][i] = v;
          onChange();
        }),
      ])),
      cellOut(`deptExtra.${d}.total`),
    ]));
  }

  tb.appendChild(groupRow('合計（予算配分＋予算外）'));
  for (const d of DEPTS) {
    tb.appendChild(el('tr', { class: 'sub-total' }, [
      el('th', { class: 'label', text: d }),
      ...weeks.map((_, i) => cellOut(`deptTotal.${d}.${i}`)),
      cellOut(`deptTotal.${d}.total`),
    ]));
  }
  tb.appendChild(el('tr', { class: 'total-row' }, [
    el('th', { class: 'label', text: '週合計' }),
    ...weeks.map((_, i) => cellOut(`hallWeekTotal.${i}`)),
    cellOut('hallWeekTotal.total'),
  ]));
}

/* ---------- 有休・更衣 ---------- */

function buildLeave() {
  const tb = $('#leaveTable tbody');
  const tf = $('#leaveTable tfoot');
  tb.innerHTML = '';
  tf.innerHTML = '';
  weeks.forEach((w, i) => {
    tb.appendChild(el('tr', {}, [
      el('th', { class: 'label', text: `${w.label} (${w.range})` }),
      cellOut(`leave.unit.${i}`),
      cellOut(`leave.hours.${i}`),
      el('td', {}, [numInput(state.leave.people[i] ?? '', (v) => { state.leave.people[i] = v; onChange(); }, { decimal: true })]),
      cellOut(`leave.amount.${i}`),
      cellOut(`leave.adjusted.${i}`),
      cellOut(`leave.perDay.${i}`),
    ]));
  });
  tf.appendChild(el('tr', { class: 'total-row' }, [
    el('th', { class: 'label', colspan: 3, text: '合計' }),
    cellOut('leave.peopleTotal'),
    cellOut('leave.amountTotal'),
    cellOut('leave.adjustedTotal'),
    el('td', {}),
  ]));
  tf.appendChild(el('tr', {}, [
    el('th', { class: 'label', colspan: 4, text: '予算 − 単純計算額（各週へ均等配分）' }),
    cellOut('leave.gap'),
    cellOut('leave.share'),
    el('td', {}),
  ]));
}

function buildChanging() {
  const tb = $('#changeTable tbody');
  const tf = $('#changeTable tfoot');
  tb.innerHTML = '';
  tf.innerHTML = '';
  weeks.forEach((w, i) => {
    tb.appendChild(el('tr', {}, [
      el('th', { class: 'label', text: `${w.label} (${w.range})` }),
      cellOut(`change.unit.${i}`),
      cellOut(`change.min.${i}`),
      el('td', {}, [numInput(state.changing.people[i] ?? '', (v) => { state.changing.people[i] = v; onChange(); }, { decimal: true })]),
      cellOut(`change.amount.${i}`),
    ]));
  });
  tf.appendChild(el('tr', { class: 'total-row' }, [
    el('th', { class: 'label', colspan: 3, text: '合計' }),
    cellOut('change.peopleTotal'),
    cellOut('change.amountTotal'),
  ]));
  tf.appendChild(el('tr', {}, [
    el('th', { class: 'label', colspan: 4, text: '予算（予算基準額 × 接客部門予算比）' }),
    cellOut('change.budget'),
  ]));
  tf.appendChild(el('tr', {}, [
    el('th', { class: 'label', colspan: 4, text: '差（合計 − 予算）' }),
    cellOut('change.diff'),
  ]));
}

/* ---------- 集計 ---------- */

function summaryRowDefs() {
  const rows = state.categories.map((c, i) => ({ key: `cat${i}`, name: c.name }));
  rows.push({ key: 'leave', name: '有給休暇' });
  rows.push({ key: 'changing', name: '更衣時間' });
  return rows;
}

function buildSummary() {
  const thead = $('#summaryTable thead');
  const tb = $('#summaryTable tbody');
  const tf = $('#summaryTable tfoot');
  thead.innerHTML = '';
  tb.innerHTML = '';
  tf.innerHTML = '';

  thead.appendChild(el('tr', {}, [
    el('th', { class: 'label', text: '項目' }),
    ...weeks.map((w) => el('th', { text: w.label })),
    el('th', { text: '月計' }), el('th', { text: '予算' }), el('th', { text: '予算比' }), el('th', { text: '差額' }),
  ]));

  summaryRowDefs().forEach((r) => {
    tb.appendChild(el('tr', {}, [
      el('th', { class: 'label', dataset: { out: `sum.${r.key}.name` }, text: r.name }),
      ...weeks.map((_, i) => cellOut(`sum.${r.key}.${i}`)),
      cellOut(`sum.${r.key}.total`),
      cellOut(`sum.${r.key}.budget`),
      cellOut(`sum.${r.key}.rate`),
      cellOut(`sum.${r.key}.diff`),
    ]));
  });

  tf.appendChild(el('tr', { class: 'total-row' }, [
    el('th', { class: 'label', text: '合計（有休除く）' }),
    ...weeks.map((_, i) => cellOut(`sum.total.${i}`)),
    cellOut('sum.total.total'),
    cellOut('sum.total.budget'),
    cellOut('sum.total.rate'),
    cellOut('sum.total.diff'),
  ]));

  // 接客部門
  const thead2 = $('#summaryHallTable thead');
  const tb2 = $('#summaryHallTable tbody');
  thead2.innerHTML = '';
  tb2.innerHTML = '';
  thead2.appendChild(el('tr', {}, [
    el('th', { class: 'label', text: '部門' }),
    ...weeks.map((w) => el('th', { text: w.label })),
    el('th', { text: '月計' }),
  ]));
  for (const d of DEPTS) {
    tb2.appendChild(el('tr', {}, [
      el('th', { class: 'label', text: d }),
      ...weeks.map((_, i) => cellOut(`sumHall.${d}.${i}`)),
      cellOut(`sumHall.${d}.total`),
    ]));
  }
  tb2.appendChild(el('tr', { class: 'total-row' }, [
    el('th', { class: 'label', text: '合計' }),
    ...weeks.map((_, i) => cellOut(`sumHall.total.${i}`)),
    cellOut('sumHall.total.total'),
  ]));

  buildPlanExport();
}

/** 実施予定の明細一覧（出力用） */
function buildPlanExport() {
  const thead = $('#planTable thead');
  const tb = $('#planTable tbody');
  thead.innerHTML = '';
  tb.innerHTML = '';
  thead.appendChild(el('tr', {}, [
    el('th', { class: 'label', text: '項目' }),
    el('th', { class: 'label', text: '内容' }),
    el('th', { text: '時給' }), el('th', { text: '時間' }), el('th', { text: '人数' }),
    ...weeks.map((w) => el('th', { text: w.label })),
    el('th', { text: '回数計' }), el('th', { text: '金額' }),
  ]));
  state.categories.forEach((cat, ci) => {
    if (cat.mode !== 'plan') return;
    cat.items.forEach((item, ii) => {
      tb.appendChild(el('tr', {}, [
        el('th', { class: 'label', dataset: { out: `plan.${ci}.${ii}.cat` }, text: cat.name }),
        el('td', { class: 'label', dataset: { out: `plan.${ci}.${ii}.label` }, text: item.label || '' }),
        cellOut(`plan.${ci}.${ii}.wage`),
        cellOut(`plan.${ci}.${ii}.hours`),
        cellOut(`plan.${ci}.${ii}.people`),
        ...weeks.map((_, wi) => cellOut(`plan.${ci}.${ii}.w${wi}`)),
        cellOut(`plan.${ci}.${ii}.times`),
        cellOut(`plan.${ci}.${ii}.total`),
      ]));
    });
  });
  if (!tb.children.length) {
    tb.appendChild(el('tr', {}, [el('td', { class: 'label', colspan: weeks.length + 7, text: '「実施予定から積み上げ」の項目がありません' })]));
  }
}

/* ---------------------------------------------------------
   再計算（出力セルのみ更新）
   --------------------------------------------------------- */

function recalc() {
  const c = calc();
  const signed = (v) => (v < 0 ? 'neg' : v > 0 ? 'pos' : undefined);

  // 週ストリップ
  weeks.forEach((_, i) => {
    setOut(`strip.ratio.${i}`, c.attTotal > 0 ? fmtPct(c.ratio[i], 1) : '-');
    setOut(`strip.items.${i}`, fmt(c.itemsWeekly[i]));
    setOut(`wl.att.${i}`, fmt(c.weekAtt[i]));
    setOut(`wl.ratio.${i}`, c.attTotal > 0 ? fmtPct(c.ratio[i], 2) : '-');
  });
  setOut('strip.ratio.total', c.attTotal > 0 ? '100.0%' : '-');
  setOut('strip.items.total', fmt(c.itemsTotal));
  setOut('items.total', fmt(c.itemsTotal));
  setOut('items.budget', fmt(c.itemsBudget));
  setOut('items.rate', c.itemsBudget > 0 ? fmtPct(c.itemsTotal / c.itemsBudget, 1) : '-');
  setOut('items.diff', fmt(c.itemsBudget - c.itemsTotal), signed(c.itemsBudget - c.itemsTotal));

  // 部門構成比
  let ratioSum = 0;
  for (const d of DEPTS) {
    const v = toNum(state.deptRatio[d]);
    ratioSum += v;
    setOut(`deptRatioPct.${d}`, fmtPct(v, 2));
  }
  setOut('deptRatioSum', ratioSum.toFixed(6));
  setOut('deptRatioSumPct', fmtPct(ratioSum, 2), Math.abs(ratioSum - 1) < 1e-6 ? undefined : 'neg');

  // 動員
  weeks.forEach((w, i) => {
    w.dates.forEach((d) => {
      const key = ymKey(d);
      setOut(`attPct.${key}`, c.attTotal > 0 ? fmtPct(toNum(state.attendance[key]) / c.attTotal, 2) : '-');
    });
    setOut(`weekAtt.${i}`, fmt(c.weekAtt[i]));
    setOut(`weekRatio.${i}`, fmtPct(c.ratio[i], 2));
    setOut(`weekDays.${i}`, String(w.dates.length));
  });
  setOut('weekAtt.total', fmt(c.attTotal));
  setOut('weekRatio.total', fmtPct(c.ratio.reduce((a, b) => a + b, 0), 2));
  setOut('weekDays.total', String(weeks.reduce((s, w) => s + w.dates.length, 0)));
  setOut('attTotal', fmt(c.attTotal));
  setOut('attTotalPct', c.attTotal > 0 ? '100.00%' : '-');

  // 項目別
  c.cats.forEach((cc, ci) => {
    cc.items.forEach((it, ii) => {
      setOut(`cat.${ci}.item.${ii}.times`, it.times ? fmtRaw(it.times) : '-');
      setOut(`cat.${ci}.item.${ii}.total`, fmt(it.total));
      setOut(`plan.${ci}.${ii}.cat`, cc.ref.name);
      setOut(`plan.${ci}.${ii}.label`, it.ref.label || '');
      setOut(`plan.${ci}.${ii}.wage`, fmt(it.ref.wage));
      setOut(`plan.${ci}.${ii}.hours`, fmtRaw(toNum(it.ref.hours)));
      setOut(`plan.${ci}.${ii}.people`, fmtRaw(toNum(it.ref.people)));
      weeks.forEach((_, wi) => setOut(`plan.${ci}.${ii}.w${wi}`, it.counts[wi] ? fmtRaw(it.counts[wi]) : ''));
      setOut(`plan.${ci}.${ii}.times`, fmtRaw(it.times));
      setOut(`plan.${ci}.${ii}.total`, fmt(it.total));
    });
    weeks.forEach((_, i) => {
      setOut(`cat.${ci}.week.${i}`, fmt(cc.weekly[i]));
      setOut(`sum.cat${ci}.${i}`, fmt(cc.weekly[i]));
    });
    setOut(`cat.${ci}.weekTotal`, fmt(cc.total));
    setOut(`cat.${ci}.total`, fmt(cc.total));
    setOut(`cat.${ci}.rate`, isFinite(cc.rate) ? fmtPct(cc.rate, 1) : '-');
    setOut(`cat.${ci}.diff`, fmt(cc.diff), signed(cc.diff));
    const gauge = outs.get(`cat.${ci}.gauge`);
    if (gauge) {
      const pct = isFinite(cc.rate) ? Math.max(0, Math.min(1.2, cc.rate)) : 0;
      gauge.style.width = `${(pct / 1.2) * 100}%`;
      gauge.className = 'gauge-fill' + (cc.rate > 1.001 ? ' over' : cc.rate >= 0.98 ? ' good' : '');
      gauge.textContent = '';
    }
    setOut(`sum.cat${ci}.name`, cc.ref.name);
    setOut(`sum.cat${ci}.total`, fmt(cc.total));
    setOut(`sum.cat${ci}.budget`, fmt(cc.budget));
    setOut(`sum.cat${ci}.rate`, isFinite(cc.rate) ? fmtPct(cc.rate, 1) : '-');
    setOut(`sum.cat${ci}.diff`, fmt(cc.diff), signed(cc.diff));
  });

  // 接客部門
  weeks.forEach((_, i) => {
    setOut(`hallWeekly.${i}`, fmt(c.hallWeekly[i]));
    for (const d of DEPTS) {
      setOut(`deptAlloc.${d}.${i}`, fmt(c.deptAlloc[d][i]));
      setOut(`deptTotal.${d}.${i}`, fmt(c.deptTotal[d][i]));
      setOut(`sumHall.${d}.${i}`, fmt(c.deptTotal[d][i]));
    }
    setOut(`hallWeekTotal.${i}`, fmt(c.hallWeekTotal[i]));
    setOut(`sumHall.total.${i}`, fmt(c.hallWeekTotal[i]));
  });
  setOut('hallWeekly.total', fmt(c.hallWeekly.reduce((a, b) => a + b, 0)));
  for (const d of DEPTS) {
    const alloc = c.deptAlloc[d].reduce((a, b) => a + b, 0);
    const extra = c.deptExtra[d].reduce((a, b) => a + b, 0);
    setOut(`deptAlloc.${d}.total`, fmt(alloc));
    setOut(`deptExtra.${d}.total`, fmt(extra));
    setOut(`deptTotal.${d}.total`, fmt(alloc + extra));
    setOut(`sumHall.${d}.total`, fmt(alloc + extra));
  }
  setOut('hallWeekTotal.total', fmt(c.hallMonthTotal));
  setOut('sumHall.total.total', fmt(c.hallMonthTotal));
  setOut('hallTotal', fmt(c.hallMonthTotal));
  setOut('hallBudgetOut', fmt(state.hallBudget));
  setOut('hallRate', isFinite(c.hallRate) ? fmtPct(c.hallRate, 2) : '-');
  setOut('hallDiff', fmt(state.hallBudget - c.hallMonthTotal), signed(state.hallBudget - c.hallMonthTotal));

  // 有休
  weeks.forEach((_, i) => {
    const r = c.leaveRows[i];
    setOut(`leave.unit.${i}`, fmt(state.leave.unit));
    setOut(`leave.hours.${i}`, fmtRaw(state.leave.hours));
    setOut(`leave.amount.${i}`, fmt(r.amount));
    setOut(`leave.adjusted.${i}`, fmt(r.adjusted));
    setOut(`leave.perDay.${i}`, isFinite(r.perDay) ? fmt(r.perDay) : '-');
    setOut(`sum.leave.${i}`, fmt(r.adjusted));
  });
  setOut('leave.peopleTotal', fmt(c.leaveRows.reduce((s, r) => s + r.people, 0)));
  setOut('leave.amountTotal', fmt(c.leaveBase));
  setOut('leave.adjustedTotal', fmt(c.leaveTotal));
  setOut('leave.gap', fmt(c.leaveGap), signed(c.leaveGap));
  setOut('leave.share', fmt(c.leaveShare), signed(c.leaveShare));
  setOut('sum.leave.total', fmt(c.leaveTotal));
  setOut('sum.leave.budget', fmt(c.leaveBudget));
  setOut('sum.leave.rate', c.leaveBudget > 0 ? fmtPct(c.leaveTotal / c.leaveBudget, 1) : '-');
  setOut('sum.leave.diff', fmt(c.leaveBudget - c.leaveTotal), signed(c.leaveBudget - c.leaveTotal));

  // 更衣
  weeks.forEach((_, i) => {
    const r = c.changeRows[i];
    setOut(`change.unit.${i}`, fmt(state.changing.unit));
    setOut(`change.min.${i}`, fmtRaw(state.changing.minutes));
    setOut(`change.amount.${i}`, fmt(r.amount));
    setOut(`sum.changing.${i}`, fmt(r.amount));
  });
  setOut('change.peopleTotal', fmt(c.changeRows.reduce((s, r) => s + r.people, 0)));
  setOut('change.amountTotal', fmt(c.changeTotal));
  setOut('change.budget', fmt(c.changeBudget));
  setOut('change.diff', fmt(c.changeTotal - c.changeBudget), signed(c.changeBudget - c.changeTotal));
  setOut('sum.changing.total', fmt(c.changeTotal));
  setOut('sum.changing.budget', fmt(c.changeBudget));
  setOut('sum.changing.rate', c.changeBudget > 0 ? fmtPct(c.changeTotal / c.changeBudget, 1) : '-');
  setOut('sum.changing.diff', fmt(c.changeBudget - c.changeTotal), signed(c.changeBudget - c.changeTotal));

  // 合計（有休除く）
  const totalWeekly = weeks.map((_, i) => c.itemsWeekly[i] + c.changeRows[i].amount);
  const totalMonth = totalWeekly.reduce((a, b) => a + b, 0);
  const totalBudget = c.itemsBudget + c.changeBudget;
  weeks.forEach((_, i) => setOut(`sum.total.${i}`, fmt(totalWeekly[i])));
  setOut('sum.total.total', fmt(totalMonth));
  setOut('sum.total.budget', fmt(totalBudget));
  setOut('sum.total.rate', totalBudget > 0 ? fmtPct(totalMonth / totalBudget, 1) : '-');
  setOut('sum.total.diff', fmt(totalBudget - totalMonth), signed(totalBudget - totalMonth));

  renderChecks(c, ratioSum);
}

function renderChecks(c, ratioSum) {
  const ul = $('#checkList');
  if (!ul) return;
  const checks = [];

  c.cats.forEach((cc) => {
    if (cc.budget > 0 && Math.abs(cc.diff) > Math.max(1, cc.budget * 0.005)) {
      checks.push({ cls: cc.diff < 0 ? 'ng' : 'warn', text: `${cc.ref.name}：予算比 ${fmtPct(cc.rate, 1)}（${cc.diff < 0 ? '超過' : '残'} ${fmt(Math.abs(cc.diff))} 円）` });
    } else if (cc.budget > 0) {
      checks.push({ cls: 'ok', text: `${cc.ref.name}：予算どおり（予算比 ${fmtPct(cc.rate, 1)}）` });
    } else if (cc.total > 0) {
      checks.push({ cls: 'warn', text: `${cc.ref.name}：月予算が未入力です（月計 ${fmt(cc.total)} 円）` });
    }
  });

  if (c.leaveBudget > 0 && Math.abs(c.leaveGap) > c.leaveBudget * 0.2) {
    checks.push({ cls: 'warn', text: `有給休暇：単純計算額と予算の差が大きい（差 ${fmt(c.leaveGap)} 円）。人数を確認してください` });
  }

  if (state.hallBudget > 0) {
    const diff = state.hallBudget - c.hallMonthTotal;
    checks.push(Math.round(diff) === 0
      ? { cls: 'ok', text: '接客部門：予算どおり（差 0 円）' }
      : { cls: diff < 0 ? 'ng' : 'warn', text: `接客部門：${fmt(Math.abs(diff))} 円 ${diff < 0 ? '超過' : '残'}（予算比 ${fmtPct(c.hallRate, 2)}）` });
    if (Math.abs(ratioSum - 1) > 1e-6) {
      checks.push({ cls: 'ng', text: `部門構成比の合計が ${fmtPct(ratioSum, 2)} です（100% になるよう調整してください）` });
    }
  }

  if (c.attTotal === 0) {
    checks.push({ cls: 'warn', text: '日別動員が未入力です（動員構成比による配分を使う場合は入力してください）' });
  }

  ul.innerHTML = '';
  if (!checks.length) checks.push({ cls: 'ok', text: '入力待ちです。項目の月予算と実施予定を入力してください' });
  checks.forEach((x) => ul.appendChild(el('li', { class: x.cls, text: x.text })));
}

/* ---------------------------------------------------------
   変更ハンドラ
   --------------------------------------------------------- */

function onChange() {
  recalc();
  autoSave();
}

function rebuildDynamic() {
  buildCategories();
  buildSummary();
  refreshOuts();
  recalc();
  autoSave();
}

/* ---------------------------------------------------------
   コピー / ダウンロード
   --------------------------------------------------------- */

function tableToRows(tableId) {
  const table = document.getElementById(tableId);
  return $$('tr', table).map((tr) =>
    $$('th,td', tr).map((cell) => {
      const input = $('input,select', cell);
      const text = input ? input.value : cell.textContent;
      return String(text).replace(/\s+/g, ' ').trim();
    })
  );
}

/**
 * ファイル保存。
 * ローカル（file://）はブラウザの通常ダウンロード、
 * claude.ai で公開したページでは downloads 機能経由で保存する。
 * 公開ページで CSV が許可されていない場合は TSV(.txt) で保存し直す。
 */
async function saveFile(filename, text, fallback) {
  const api = window.claude;
  if (api && typeof api.use === 'function') {
    let downloads = null;
    try { downloads = await api.use('downloads'); } catch (e) { downloads = null; }

    if (!downloads) {
      // 公開ページだがファイル保存が使えない環境 → 画面表示に切り替える
      showTextPanel(filename, (fallback && fallback.text) || text,
        'この環境ではファイル保存が使えないため、内容を表示しています。選択してコピーし、Excelに貼り付けてください。');
      return;
    }

    try {
      await downloads.save({ filename, data: text });
      toast('ダウンロードしました');
      return;
    } catch (err) {
      const code = err && err.code;
      if (code === 'declined') return;
      if (code === 'rate_limited') { toast('少し待ってからもう一度お試しください'); return; }
      if ((code === 'extension_not_enabled' || code === 'rejected_extension') && fallback) {
        try {
          await downloads.save({ filename: fallback.filename, data: fallback.text });
          toast(`${fallback.filename} でダウンロードしました（Excelにそのまま貼付・取込できます）`);
          return;
        } catch (err2) {
          if (err2 && err2.code === 'declined') return;
        }
      }
      showTextPanel(filename, (fallback && fallback.text) || text,
        'ダウンロードできなかったため、内容を表示しています。選択してコピーし、Excelに貼り付けてください。');
      return;
    }
  }

  try {
    const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = el('a', { href: url, download: filename });
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    toast(`${filename} をダウンロードしました`);
  } catch (e) {
    showTextPanel(filename, (fallback && fallback.text) || text);
  }
}

/**
 * どの環境でも必ずデータを取り出せる最後の受け皿。
 * ダウンロードもクリップボードも使えないとき、本文を画面に出して手動コピーできるようにする。
 */
function showTextPanel(filename, text, note) {
  const area = el('textarea', { class: 'panel-text', readonly: 'readonly', rows: 14, spellcheck: 'false' });
  area.value = text;

  const overlay = el('div', { class: 'modal-overlay' });
  const close = () => overlay.remove();
  overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });

  overlay.appendChild(el('div', { class: 'modal' }, [
    el('h3', { text: filename }),
    el('p', { class: 'hint', text: note || 'この内容を選択してコピーし、Excelに貼り付けてください（タブ区切り）。' }),
    area,
    el('div', { class: 'row wrap' }, [
      el('button', {
        type: 'button', class: 'btn', text: 'すべて選択',
        onclick: () => { area.focus(); area.select(); },
      }),
      el('button', {
        type: 'button', class: 'btn ghost', text: 'コピーを試す',
        onclick: () => {
          area.focus();
          area.select();
          let ok = false;
          try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
          toast(ok ? 'コピーしました' : '手動で選択してコピーしてください');
        },
      }),
      el('button', { type: 'button', class: 'btn ghost', text: '閉じる', onclick: close }),
    ]),
  ]));

  document.body.appendChild(overlay);
  area.focus();
  area.select();
}

function copyTable(tableId) {
  const tsv = tableToRows(tableId).map((r) => r.join('\t')).join('\n');
  const base = `${FILE_BASE[tableId] || 'export'}_${state.ym}`;
  const legacyCopy = () => {
    const ta = el('textarea', { style: 'position:fixed;opacity:0' });
    ta.value = tsv;
    document.body.appendChild(ta);
    ta.select();
    let ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    ta.remove();
    if (ok) toast('コピーしました（Excelに貼り付けできます）');
    else saveFile(`${base}.txt`, tsv);
  };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(tsv).then(() => toast('コピーしました（Excelに貼り付けできます）'), legacyCopy);
  } else {
    legacyCopy();
  }
}

function downloadCsv(tableId) {
  const rows = tableToRows(tableId);
  const csv = rows.map((r) => r.map((v) => (/[",]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v)).join(',')).join('\r\n');
  const tsv = rows.map((r) => r.join('\t')).join('\r\n');
  const base = `${FILE_BASE[tableId] || 'export'}_${state.ym}`;
  saveFile(`${base}.csv`, '﻿' + csv, { filename: `${base}.txt`, text: tsv });
}

/* ---------------------------------------------------------
   動員の貼り付け取込
   --------------------------------------------------------- */

function importPaste(text) {
  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter((l) => l !== '');
  if (lines.length === 0) { toast('取込むデータがありません'); return; }
  const allDates = weeks.flatMap((w) => w.dates);
  const [y, m] = state.ym.split('-').map(Number);
  let count = 0;

  lines.forEach((line, idx) => {
    const cols = line.split(/\t|,|\s{2,}/).map((s) => s.trim()).filter((s) => s !== '');
    let date = null;
    let value = null;
    if (cols.length >= 2) {
      const dm = cols[0].match(/^(?:(\d{4})[-/])?(\d{1,2})[-/](\d{1,2})/);
      if (dm) {
        date = new Date(dm[1] ? Number(dm[1]) : y, Number(dm[2]) - 1, Number(dm[3]));
        value = toNum(cols[cols.length - 1]);
      }
    }
    if (date === null) {
      value = toNum(cols[cols.length - 1]);
      date = allDates[idx] || null;
    }
    if (!date || date.getFullYear() !== y || date.getMonth() !== m - 1) return;
    state.attendance[ymKey(date)] = value;
    count += 1;
  });

  buildAttendance();
  refreshOuts();
  recalc();
  autoSave();
  toast(`${count} 日分を取込みました`);
}

/* ---------------------------------------------------------
   初期化
   --------------------------------------------------------- */

function switchMonth(ym) {
  saveState(true);
  state = loadState(ym);
  buildAll();
}

function init() {
  const now = new Date();
  const initialYm = localStorage.getItem(STORAGE_PREFIX + '@last')
    || savedMonths()[0]
    || `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

  const selMonth = $('#selMonth');
  for (let m = 1; m <= 12; m++) selMonth.appendChild(el('option', { value: String(m), text: `${m}月` }));

  state = loadState(initialYm);
  buildAll();
  refreshSavedMonths();

  // タブ
  $$('#tabbar .tab').forEach((tab) => {
    tab.addEventListener('click', () => {
      $$('#tabbar .tab').forEach((t) => t.classList.toggle('active', t === tab));
      $$('main .panel').forEach((p) => p.classList.toggle('active', p.id === tab.dataset.target));
      window.scrollTo({ top: 0 });
    });
  });

  // 月の選択
  const applyMonth = () => {
    const y = Number($('#selYear').value);
    const m = Number($('#selMonth').value);
    switchMonth(`${y}-${String(m).padStart(2, '0')}`);
  };
  $('#selYear').addEventListener('change', applyMonth);
  $('#selMonth').addEventListener('change', applyMonth);
  $('#btnPrevMonth').addEventListener('click', () => {
    const [y, m] = state.ym.split('-').map(Number);
    const d = new Date(y, m - 2, 1);
    switchMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
  });
  $('#btnNextMonth').addEventListener('click', () => {
    const [y, m] = state.ym.split('-').map(Number);
    const d = new Date(y, m, 1);
    switchMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
  });

  $('#weekStartNo').addEventListener('change', (e) => {
    state.weekStartNo = toNum(e.target.value) || 1;
    buildAll();
    autoSave();
  });
  $('#btnAutoWeek').addEventListener('click', () => {
    const [y, m] = state.ym.split('-').map(Number);
    state.weekStartNo = fiscalWeekNo(new Date(y, m - 1, 1));
    buildAll();
    autoSave();
    toast(`開始WEEKを ${state.weekStartNo} にしました`);
  });

  // 項目
  $('#btnAddCategory').addEventListener('click', () => {
    state.categories.push({ name: '新しい項目', mode: 'plan', budget: 0, items: [{ label: '', wage: 1251, hours: 0, people: 1, weekCounts: weeks.map(() => 0) }], weekly: [] });
    rebuildDynamic();
  });
  $('#btnRestoreCategories').addEventListener('click', () => {
    if (!confirm('項目を既定の内容に戻します。入力済みの内容は失われます。よろしいですか？')) return;
    state.categories = defaultCategories();
    normalizeItems();
    rebuildDynamic();
  });

  // 接客部門
  $('#hallBudget').addEventListener('input', (e) => { state.hallBudget = toNum(e.target.value); onChange(); });
  $('#hallWeeklyManual').addEventListener('change', (e) => {
    state.hallWeeklyManual = e.target.checked;
    if (state.hallWeeklyManual && (!state.hallWeekly || state.hallWeekly.length === 0)) {
      state.hallWeekly = calc().hallWeekly.slice();
    }
    buildHall();
    refreshOuts();
    recalc();
    autoSave();
  });

  // 動員
  $('#btnPaste').addEventListener('click', () => {
    importPaste($('#pasteBox').value);
    $('#pasteBox').value = '';
  });
  $('#btnClearAttendance').addEventListener('click', () => {
    if (!confirm('日別動員をすべてクリアしますか？')) return;
    state.attendance = {};
    buildAttendance();
    refreshOuts();
    recalc();
    autoSave();
  });

  // 有休・更衣
  const bindNum = (sel, apply) => $(sel).addEventListener('input', (e) => { apply(toNum(e.target.value)); onChange(); });
  bindNum('#leaveUnit', (v) => { state.leave.unit = v; });
  bindNum('#leaveHours', (v) => { state.leave.hours = v; });
  bindNum('#leaveBudget', (v) => { state.leave.budget = v; });
  bindNum('#changeUnit', (v) => { state.changing.unit = v; });
  bindNum('#changeMinutes', (v) => { state.changing.minutes = v; });
  bindNum('#changeBase', (v) => { state.changing.base = v; });

  // 保存
  $('#btnSave').addEventListener('click', () => saveState(false));
  $('#btnLoadMonth').addEventListener('click', () => {
    const ym = $('#savedMonths').value;
    if (!ym) return;
    switchMonth(ym);
    toast(`${ym} を読込みました`);
  });
  $('#btnCopyPrev').addEventListener('click', () => {
    const [y, m] = state.ym.split('-').map(Number);
    const prev = new Date(y, m - 2, 1);
    const prevYm = `${prev.getFullYear()}-${String(prev.getMonth() + 1).padStart(2, '0')}`;
    const raw = localStorage.getItem(STORAGE_PREFIX + prevYm);
    if (!raw) { toast(`${prevYm} の保存データがありません`); return; }
    if (!confirm(`${prevYm} の内容（項目・単価・予算）を複製します。動員は引き継ぎません。よろしいですか？`)) return;
    const cur = state.ym;
    state = Object.assign(emptyState(cur), JSON.parse(raw), {
      ym: cur,
      attendance: {},
      weekStartNo: fiscalWeekNo(new Date(y, m - 1, 1)),
    });
    buildAll();
    saveState(true);
    toast(`${prevYm} から複製しました`);
  });
  $('#btnResetMonth').addEventListener('click', () => {
    if (!confirm(`${state.ym} の入力をすべて消去します。よろしいですか？`)) return;
    localStorage.removeItem(STORAGE_PREFIX + state.ym);
    state = emptyState(state.ym);
    buildAll();
    refreshSavedMonths();
    toast('リセットしました');
  });
  $('#btnExportJson').addEventListener('click', () => saveFile(`budget_${state.ym}.json`, JSON.stringify(state, null, 2)));
  $('#btnImportJson').addEventListener('click', () => $('#fileJson').click());
  $('#fileJson').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result);
        state = Object.assign(emptyState(data.ym || state.ym), data);
        buildAll();
        saveState(true);
        toast('JSONを読込みました');
      } catch (err) {
        toast('JSONの読込に失敗しました');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  });

  // コピー / CSV
  document.addEventListener('click', (e) => {
    const ds = e.target.dataset;
    if (!ds) return;
    if (ds.copy) copyTable(ds.copy);
    if (ds.csv) downloadCsv(ds.csv);
    if (ds.show) {
      const tsv = tableToRows(ds.show).map((r) => r.join('\t')).join('\n');
      showTextPanel(`${FILE_BASE[ds.show] || 'export'}_${state.ym}`, tsv);
    }
  });

  window.addEventListener('beforeunload', () => saveState(true));
}

document.addEventListener('DOMContentLoaded', init);
