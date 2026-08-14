/* =========================================================
   月次予算割振アプリ
   Excel「予算割振」ブックの計算ロジックをブラウザ上で再現する。

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

const MODES = [
  { value: 'unit', label: '単価表（毎週同額）' },
  { value: 'ratio', label: '動員構成比で按分' },
  { value: 'even', label: '週均等で按分' },
  { value: 'manual', label: '手入力' },
];

function defaultCategories() {
  const item = (label, wage, hours, weeks, people) => ({ label, wage, hours, weeks, people });
  return [
    { name: '障がい者雇用', mode: 'unit', budget: 0, items: [item('', 1236, 5, 3, 1)] },
    { name: '新規採用研修（トレーニー）', mode: 'ratio', budget: 0, items: [item('週2〜3名', 1226, 5.5, 1, 2), item('カンポリ参加', 1226, 2.15, 1, 3)] },
    { name: '新規採用研修（トレーナー）', mode: 'ratio', budget: 0, items: [item('週2〜3名', 1251, 5.5, 1, 1), item('', 1251, 5, 2, 4)] },
    { name: '既存スタッフ研修', mode: 'manual', budget: 0, items: [item('週1', 1251, 1, 1, 1), item('週2', 1251, 1.5, 1, 3), item('週3', 1251, 1.75, 1, 6), item('週4', 1251, 1.5, 1, 6)] },
    { name: '防災訓練', mode: 'manual', budget: 0, items: [item('週1', 1251, 0.5, 1, 10), item('週2', 1251, 0.5, 1, 15), item('週3', 1251, 0.75, 1, 8), item('週4', 1251, 0.75, 1, 2), item('総合訓練', 1251, 1, 1, 8)] },
    { name: '棚卸（予算分）', mode: 'manual', budget: 0, items: [item('', 1251, 5, 1, 5)] },
    { name: 'ストア準備・返品（予算分）', mode: 'unit', budget: 0, items: [item('※平均時給以外は変更不可', 1251, 2, 1, 1)] },
    { name: 'ストア準備・返品（追加分）', mode: 'manual', budget: 0, items: [item('', 1251, 2, 1, 2), item('', 1251, 6, 1, 2), item('', 1251, 6, 1, 4)] },
    { name: 'その他', mode: 'manual', budget: 0, items: [item('', 1251, 0, 0, 0)] },
    { name: 'リーダー手当', mode: 'manual', budget: 0, items: [] },
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
    else if (k === 'html') node.innerHTML = v;
    else if (k.startsWith('on')) node.addEventListener(k.slice(2), v);
    else if (k === 'dataset') Object.assign(node.dataset, v);
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
  const order = exact
    .map((x, i) => ({ i, frac: x - Math.floor(x) }))
    .sort((a, b) => b.frac - a.frac);
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
  thu.setDate(thu.getDate() + ((4 - thu.getDay() + 7) % 7)); // その週の木曜
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
    const data = JSON.parse(raw);
    return Object.assign(emptyState(ym), data, { ym });
  } catch (e) {
    return emptyState(ym);
  }
}

function refreshSavedMonths() {
  const sel = $('#savedMonths');
  const cur = sel.value;
  sel.innerHTML = '';
  sel.appendChild(el('option', { value: '', text: '-' }));
  for (const m of savedMonths()) sel.appendChild(el('option', { value: m, text: m }));
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
  toast._t = setTimeout(() => { t.hidden = true; }, 2200);
}

/* ---------------------------------------------------------
   計算
   --------------------------------------------------------- */

function calc() {
  const weekAtt = weeks.map((w) => w.dates.reduce((s, d) => s + toNum(state.attendance[ymKey(d)]), 0));
  const attTotal = weekAtt.reduce((a, b) => a + b, 0);
  const ratio = weekAtt.map((v) => (attTotal > 0 ? v / attTotal : 0));

  // 接客部門の週予算
  let hallWeekly;
  if (state.hallWeeklyManual) {
    hallWeekly = weeks.map((_, i) => toNum(state.hallWeekly[i]));
  } else {
    hallWeekly = allocate(state.hallBudget, attTotal > 0 ? weekAtt : weeks.map(() => 1));
  }

  // 部門別割振（週予算 × 構成比）＋ 予算外
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

  // 項目別
  const cats = state.categories.map((c) => {
    const items = c.items.map((it) => ({
      ...it,
      amount: toNum(it.wage) * toNum(it.hours) * toNum(it.weeks) * toNum(it.people),
    }));
    const unitTotal = items.reduce((s, it) => s + it.amount, 0);
    let weekly;
    if (c.mode === 'unit') weekly = weeks.map(() => Math.round(unitTotal));
    else if (c.mode === 'ratio') weekly = allocate(toNum(c.budget), weekAtt);
    else if (c.mode === 'even') weekly = allocate(toNum(c.budget), weeks.map(() => 1));
    else weekly = weeks.map((_, i) => toNum((c.weekly || [])[i]));
    const total = weekly.reduce((a, b) => a + b, 0);
    const budget = toNum(c.budget);
    return { ref: c, items, unitTotal, weekly, total, budget, rate: budget > 0 ? total / budget : NaN, diff: budget - total };
  });

  // 有休
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

  // 更衣時間
  const cg = state.changing;
  const changeRows = weeks.map((_, i) => {
    const people = toNum(cg.people[i]);
    return { people, amount: toNum(cg.unit) * (toNum(cg.minutes) / 60) * people };
  });
  const changeTotal = changeRows.reduce((s, r) => s + r.amount, 0);
  const changeBudget = toNum(cg.base) * (isFinite(hallRate) ? hallRate : 0);

  return {
    weekAtt, attTotal, ratio, hallWeekly,
    deptAlloc, deptExtra, deptTotal, hallWeekTotal, hallMonthTotal, hallRate,
    cats,
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
  if (opts.disabled) input.disabled = true;
  input.addEventListener('change', () => {
    const v = input.value.trim() === '' ? 0 : toNum(input.value);
    input.value = v === 0 && input.value.trim() === '' ? '' : fmtRaw(v);
    onChange(v);
  });
  input.addEventListener('input', () => onChange(toNum(input.value)));
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
  $('#ym').value = state.ym;
  $('#weekStartNo').value = state.weekStartNo;
  $('#hallBudget').value = state.hallBudget ? fmtRaw(state.hallBudget) : '';
  $('#hallWeeklyManual').checked = !!state.hallWeeklyManual;
  $('#leaveUnit').value = fmtRaw(state.leave.unit);
  $('#leaveHours').value = fmtRaw(state.leave.hours);
  $('#leaveBudget').value = state.leave.budget ? fmtRaw(state.leave.budget) : '';
  $('#changeUnit').value = fmtRaw(state.changing.unit);
  $('#changeMinutes').value = fmtRaw(state.changing.minutes);
  $('#changeBase').value = fmtRaw(state.changing.base);

  buildDeptRatio();
  buildAttendance();
  buildWeekSummary();
  buildHall();
  buildCategories();
  buildLeave();
  buildChanging();
  buildSummary();

  refreshOuts();
  recalc();
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
  const rows = [
    ['動員', 'weekAtt'],
    ['構成比', 'weekRatio'],
    ['日数', 'weekDays'],
  ];
  for (const [label, key] of rows) {
    tb.appendChild(el('tr', {}, [
      el('th', { class: 'label', text: label }),
      ...weeks.map((_, i) => cellOut(`${key}.${i}`)),
      cellOut(`${key}.total`),
    ]));
  }
}

function buildHall() {
  const thead = $('#hallTable thead');
  const tb = $('#hallTable tbody');
  thead.innerHTML = '';
  tb.innerHTML = '';

  thead.appendChild(el('tr', {}, [
    el('th', { class: 'label', text: '項目' }),
    ...weeks.map((w) => el('th', { text: `${w.label}\n${w.range}` })),
    el('th', { text: '月計' }),
  ]));

  const groupRow = (text) => el('tr', { class: 'group-head' }, [
    el('th', { class: 'label', colspan: weeks.length + 2, text }),
  ]);

  // 週予算
  tb.appendChild(groupRow('週予算'));
  tb.appendChild(el('tr', {}, [
    el('th', { class: 'label', text: '週予算' }),
    ...weeks.map((_, i) =>
      state.hallWeeklyManual
        ? el('td', {}, [numInput(state.hallWeekly[i] ?? '', (v) => { state.hallWeekly[i] = v; onChange(); })])
        : cellOut(`hallWeekly.${i}`)
    ),
    cellOut('hallWeekly.total'),
  ]));

  // 予算配分
  tb.appendChild(groupRow('予算配分（週予算 × 部門構成比）'));
  for (const d of DEPTS) {
    tb.appendChild(el('tr', {}, [
      el('th', { class: 'label', text: d }),
      ...weeks.map((_, i) => cellOut(`deptAlloc.${d}.${i}`)),
      cellOut(`deptAlloc.${d}.total`),
    ]));
  }

  // 予算外
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

  // 合計
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

function buildCategories() {
  const list = $('#categoryList');
  list.innerHTML = '';
  state.categories.forEach((cat, ci) => list.appendChild(buildCategoryCard(cat, ci)));
}

function buildCategoryCard(cat, ci) {
  const modeSelect = el('select', {}, MODES.map((m) => el('option', { value: m.value, text: m.label, selected: cat.mode === m.value })));
  modeSelect.value = cat.mode;
  modeSelect.addEventListener('change', () => {
    cat.mode = modeSelect.value;
    buildCategories();
    refreshOuts();
    recalc();
    autoSave();
  });

  const head = el('div', { class: 'cat-head' }, [
    el('div', { class: 'cat-name' }, [textInput(cat.name, (v) => { cat.name = v; onChange(); buildSummaryLabels(); })]),
    el('label', { class: 'field' }, [
      el('span', { text: '月予算' }),
      numInput(cat.budget || '', (v) => { cat.budget = v; onChange(); }),
    ]),
    el('label', { class: 'field' }, [el('span', { text: '配分方法' }), modeSelect]),
    el('div', { class: 'stat' }, [el('span', { text: '月計' }), el('strong', { dataset: { out: `cat.${ci}.total` }, text: '-' })]),
    el('div', { class: 'stat' }, [el('span', { text: '予算比' }), el('strong', { dataset: { out: `cat.${ci}.rate` }, text: '-' })]),
    el('div', { class: 'stat' }, [el('span', { text: '差額' }), el('strong', { dataset: { out: `cat.${ci}.diff` }, text: '-' })]),
    el('div', { class: 'spacer' }),
    el('button', {
      type: 'button', class: 'btn ghost mini', text: '項目削除',
      onclick: () => {
        if (!confirm(`「${cat.name}」を削除しますか？`)) return;
        state.categories.splice(ci, 1);
        rebuildDynamic();
      },
    }),
  ]);

  // --- 単価表 ---
  const unitBody = el('tbody');
  cat.items.forEach((it, ii) => {
    unitBody.appendChild(el('tr', {}, [
      el('td', {}, [textInput(it.label, (v) => { it.label = v; onChange(); })]),
      el('td', {}, [numInput(it.wage ?? '', (v) => { it.wage = v; onChange(); })]),
      el('td', {}, [numInput(it.hours ?? '', (v) => { it.hours = v; onChange(); }, { decimal: true })]),
      el('td', {}, [numInput(it.weeks ?? '', (v) => { it.weeks = v; onChange(); }, { decimal: true })]),
      el('td', {}, [numInput(it.people ?? '', (v) => { it.people = v; onChange(); }, { decimal: true })]),
      cellOut(`cat.${ci}.item.${ii}`),
      el('td', {}, [el('button', {
        type: 'button', class: 'btn ghost mini', text: '×',
        onclick: () => { cat.items.splice(ii, 1); rebuildDynamic(); },
      })]),
    ]));
  });

  const unitTable = el('table', { class: 'grid' }, [
    el('thead', {}, [el('tr', {}, [
      el('th', { class: 'label', text: '内容' }),
      el('th', { text: '時給' }), el('th', { text: '時間' }), el('th', { text: '週' }), el('th', { text: '人数' }),
      el('th', { text: '金額' }), el('th', { text: '' }),
    ])]),
    unitBody,
    el('tfoot', {}, [el('tr', {}, [
      el('th', { class: 'label', colspan: 5, text: '1週あたり合計' }),
      cellOut(`cat.${ci}.unitTotal`),
      el('td', {}),
    ])]),
  ]);

  const unitBlock = el('div', {}, [
    el('h4', { text: '単価表（時給 × 時間 × 週 × 人数）' }),
    el('div', { class: 'table-scroll' }, [unitTable]),
    el('div', { class: 'row wrap', style: 'margin-top:8px' }, [
      el('button', {
        type: 'button', class: 'btn ghost mini', text: '行を追加',
        onclick: () => { cat.items.push({ label: '', wage: 1251, hours: 0, weeks: 1, people: 1 }); rebuildDynamic(); },
      }),
      el('button', {
        type: 'button', class: 'btn ghost mini', text: '1週あたり合計 × 週数 を月予算にセット',
        onclick: () => {
          const unitTotal = cat.items.reduce((s, it) => s + toNum(it.wage) * toNum(it.hours) * toNum(it.weeks) * toNum(it.people), 0);
          cat.budget = Math.round(unitTotal * weeks.length);
          rebuildDynamic();
        },
      }),
    ]),
  ]);

  // --- 週配分 ---
  const weekBody = el('tbody', {}, [
    el('tr', {}, [
      el('th', { class: 'label', text: '金額' }),
      ...weeks.map((_, i) =>
        cat.mode === 'manual'
          ? el('td', {}, [numInput((cat.weekly || [])[i] ?? '', (v) => {
              cat.weekly = cat.weekly || [];
              cat.weekly[i] = v;
              onChange();
            })])
          : cellOut(`cat.${ci}.week.${i}`)
      ),
      cellOut(`cat.${ci}.weekTotal`),
    ]),
  ]);

  const weekTable = el('table', { class: 'grid' }, [
    el('thead', {}, [el('tr', {}, [
      el('th', { class: 'label', text: '週' }),
      ...weeks.map((w) => el('th', { text: w.label })),
      el('th', { text: '月計' }),
    ])]),
    weekBody,
  ]);

  const weekBlockChildren = [
    el('h4', { text: '週別配分' }),
    el('div', { class: 'table-scroll' }, [weekTable]),
  ];

  if (cat.mode === 'manual') {
    weekBlockChildren.push(el('div', { class: 'row wrap', style: 'margin-top:8px' }, [
      el('button', {
        type: 'button', class: 'btn ghost mini', text: '動員構成比で埋める',
        onclick: () => {
          const c = calc();
          cat.weekly = allocate(toNum(cat.budget), c.weekAtt);
          rebuildDynamic();
        },
      }),
      el('button', {
        type: 'button', class: 'btn ghost mini', text: '均等で埋める',
        onclick: () => { cat.weekly = allocate(toNum(cat.budget), weeks.map(() => 1)); rebuildDynamic(); },
      }),
      el('button', {
        type: 'button', class: 'btn ghost mini', text: '1週あたり合計で埋める',
        onclick: () => {
          const unitTotal = cat.items.reduce((s, it) => s + toNum(it.wage) * toNum(it.hours) * toNum(it.weeks) * toNum(it.people), 0);
          cat.weekly = weeks.map(() => Math.round(unitTotal));
          rebuildDynamic();
        },
      }),
      el('button', {
        type: 'button', class: 'btn ghost mini', text: '単価表の行を週順に割当',
        onclick: () => {
          cat.weekly = weeks.map((_, i) => {
            const it = cat.items[i];
            return it ? Math.round(toNum(it.wage) * toNum(it.hours) * toNum(it.weeks) * toNum(it.people)) : 0;
          });
          rebuildDynamic();
        },
      }),
      el('button', {
        type: 'button', class: 'btn ghost mini', text: 'クリア',
        onclick: () => { cat.weekly = []; rebuildDynamic(); },
      }),
    ]));
  } else {
    weekBlockChildren.push(el('p', { class: 'hint', text: '配分方法を「手入力」にすると週ごとに直接入力できます。' }));
  }

  return el('div', { class: 'cat-card' }, [head, el('div', { class: 'cat-body' }, [unitBlock, el('div', {}, weekBlockChildren)])]);
}

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

  // 接客部門サマリ
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
    ...weeks.map((_, i) => cellOut('sumHall.total.' + i)),
    cellOut('sumHall.total.total'),
  ]));
}

function buildSummaryLabels() {
  state.categories.forEach((c, i) => setOut(`sum.cat${i}.name`, c.name));
}

/* ---------------------------------------------------------
   再計算（出力セルのみ更新）
   --------------------------------------------------------- */

function recalc() {
  const c = calc();
  const signed = (v) => (v < 0 ? 'neg' : v > 0 ? 'pos' : undefined);

  setOut('weekCount', `${weeks.length} 週`);

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
      const v = toNum(state.attendance[key]);
      setOut(`attPct.${key}`, c.attTotal > 0 ? fmtPct(v / c.attTotal, 2) : '-');
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

  // 項目別
  c.cats.forEach((cc, ci) => {
    cc.items.forEach((it, ii) => setOut(`cat.${ci}.item.${ii}`, fmt(it.amount)));
    setOut(`cat.${ci}.unitTotal`, fmt(cc.unitTotal));
    weeks.forEach((_, i) => {
      setOut(`cat.${ci}.week.${i}`, fmt(cc.weekly[i]));
      setOut(`sum.cat${ci}.${i}`, fmt(cc.weekly[i]));
    });
    setOut(`cat.${ci}.weekTotal`, fmt(cc.total));
    setOut(`cat.${ci}.total`, fmt(cc.total));
    setOut(`cat.${ci}.rate`, isFinite(cc.rate) ? fmtPct(cc.rate, 1) : '-');
    setOut(`cat.${ci}.diff`, fmt(cc.diff), signed(cc.diff));
    setOut(`sum.cat${ci}.name`, cc.ref.name);
    setOut(`sum.cat${ci}.total`, fmt(cc.total));
    setOut(`sum.cat${ci}.budget`, fmt(cc.budget));
    setOut(`sum.cat${ci}.rate`, isFinite(cc.rate) ? fmtPct(cc.rate, 1) : '-');
    setOut(`sum.cat${ci}.diff`, fmt(cc.diff), signed(cc.diff));
  });

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
  const totalWeekly = weeks.map((_, i) =>
    c.cats.reduce((s, cc) => s + cc.weekly[i], 0) + c.changeRows[i].amount);
  const totalMonth = totalWeekly.reduce((a, b) => a + b, 0);
  const totalBudget = c.cats.reduce((s, cc) => s + cc.budget, 0) + c.changeBudget;
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

  checks.push(Math.abs(ratioSum - 1) < 1e-6
    ? { cls: 'ok', text: '部門構成比の合計は 100% です。' }
    : { cls: 'ng', text: `部門構成比の合計が ${fmtPct(ratioSum, 2)} です（100% になるよう調整してください）。` });

  checks.push(c.attTotal > 0
    ? { cls: 'ok', text: `動員入力済み（月計 ${fmt(c.attTotal)}）。` }
    : { cls: 'warn', text: '日別動員が未入力です。動員構成比による按分ができません。' });

  if (state.hallBudget > 0) {
    const diff = state.hallBudget - c.hallMonthTotal;
    if (Math.round(diff) === 0) {
      checks.push({ cls: 'ok', text: '接客部門は予算どおりです（差 0 円）。' });
    } else {
      checks.push({
        cls: Math.abs(diff) <= state.hallBudget * 0.001 ? 'warn' : (diff < 0 ? 'ng' : 'warn'),
        text: `接客部門は予算に対して ${fmt(Math.abs(diff))} 円 ${diff < 0 ? '超過' : '残'} です（予算比 ${fmtPct(c.hallRate, 2)}）。`,
      });
    }
  } else {
    checks.push({ cls: 'warn', text: '接客部門の月間予算が未入力です。' });
  }

  c.cats.forEach((cc) => {
    if (cc.budget > 0 && Math.abs(cc.diff) > Math.max(1, cc.budget * 0.005)) {
      checks.push({ cls: cc.diff < 0 ? 'ng' : 'warn', text: `${cc.ref.name}：予算比 ${fmtPct(cc.rate, 1)}（差 ${fmt(cc.diff)} 円）。` });
    }
  });

  if (c.leaveBudget > 0 && Math.abs(c.leaveGap) > c.leaveBudget * 0.2) {
    checks.push({ cls: 'warn', text: `有休：単純計算額と予算の差が大きい（差 ${fmt(c.leaveGap)} 円）。人数を確認してください。` });
  }

  ul.innerHTML = '';
  checks.forEach((c2) => ul.appendChild(el('li', { class: c2.cls, text: c2.text })));
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
   コピー / CSV
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

function copyTable(tableId) {
  const tsv = tableToRows(tableId).map((r) => r.join('\t')).join('\n');
  navigator.clipboard.writeText(tsv).then(
    () => toast('コピーしました（Excelに貼り付けできます）'),
    () => {
      const ta = el('textarea', { style: 'position:fixed;opacity:0' });
      ta.value = tsv;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      ta.remove();
      toast('コピーしました');
    }
  );
}

function downloadCsv(tableId, name) {
  const csv = tableToRows(tableId)
    .map((r) => r.map((v) => (/[",]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v)).join(','))
    .join('\r\n');
  const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' });
  const a = el('a', { href: URL.createObjectURL(blob), download: `${name}_${state.ym}.csv` });
  document.body.appendChild(a);
  a.click();
  a.remove();
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

function init() {
  const now = new Date();
  const initialYm = localStorage.getItem(STORAGE_PREFIX + '@last')
    || savedMonths()[0]
    || `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  state = loadState(initialYm);
  buildAll();
  refreshSavedMonths();

  // タブ
  $$('#tabbar .tab').forEach((tab) => {
    tab.addEventListener('click', () => {
      $$('#tabbar .tab').forEach((t) => t.classList.toggle('active', t === tab));
      $$('main .panel').forEach((p) => p.classList.toggle('active', p.id === tab.dataset.target));
    });
  });

  // ヘッダ
  $('#ym').addEventListener('change', (e) => {
    if (!e.target.value) return;
    saveState(true);
    state = loadState(e.target.value);
    buildAll();
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
    toast(`開始WEEKを ${state.weekStartNo} に設定しました`);
  });

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

  // 項目
  $('#btnAddCategory').addEventListener('click', () => {
    state.categories.push({ name: '新しい項目', mode: 'manual', budget: 0, items: [], weekly: [] });
    rebuildDynamic();
  });
  $('#btnRestoreCategories').addEventListener('click', () => {
    if (!confirm('項目を既定の内容に戻します。入力済みの金額は失われます。よろしいですか？')) return;
    state.categories = defaultCategories();
    rebuildDynamic();
  });

  // 保存関連
  $('#btnSave').addEventListener('click', () => saveState(false));
  $('#btnLoadMonth').addEventListener('click', () => {
    const ym = $('#savedMonths').value;
    if (!ym) return;
    saveState(true);
    state = loadState(ym);
    buildAll();
    toast(`${ym} を読込みました`);
  });
  $('#btnCopyPrev').addEventListener('click', () => {
    const [y, m] = state.ym.split('-').map(Number);
    const prev = new Date(y, m - 2, 1);
    const prevYm = `${prev.getFullYear()}-${String(prev.getMonth() + 1).padStart(2, '0')}`;
    const raw = localStorage.getItem(STORAGE_PREFIX + prevYm);
    if (!raw) { toast(`${prevYm} の保存データがありません`); return; }
    if (!confirm(`${prevYm} の設定（予算・単価表・構成比）を複製します。動員は引き継ぎません。よろしいですか？`)) return;
    const prevState = JSON.parse(raw);
    const cur = state.ym;
    state = Object.assign(emptyState(cur), prevState, {
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
  $('#btnExportJson').addEventListener('click', () => {
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
    const a = el('a', { href: URL.createObjectURL(blob), download: `budget_${state.ym}.json` });
    document.body.appendChild(a);
    a.click();
    a.remove();
  });
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
    const copyId = e.target.dataset && e.target.dataset.copy;
    const csvId = e.target.dataset && e.target.dataset.csv;
    if (copyId) copyTable(copyId);
    if (csvId) downloadCsv(csvId, csvId === 'summaryTable' ? '項目別集計' : '接客部門');
  });

  window.addEventListener('beforeunload', () => saveState(true));
}

document.addEventListener('DOMContentLoaded', init);
