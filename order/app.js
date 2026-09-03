/* =========================================================
 *  売店発注ツール — ランウェイ・ビュー
 *
 *  発注の判断は「在庫があと何日もつか」と「リードタイムに間に合うか」の
 *  勝負なので、表ではなく時間軸で見せる。各商品のバーは在庫が尽きるまでの
 *  滑走路で、縦線はそれを過ぎると欠品する発注デッドライン。
 * =======================================================*/

const STORAGE_KEY = 'concession-order-tool-v1';
const SPAN_DAYS = 14;                 // タイムテーブルに表示する日数

const CATEGORY_LABEL = { food: 'フード', drink: 'ドリンク', packaging: '包材' };
const CATEGORY_KEY = { 'フード': 'food', 'ドリンク': 'drink', '包材': 'packaging' };
const STATUS_LABEL = { ordered: '発注済', received: '入荷済', cancelled: '取消' };
const WEEK = ['日', '月', '火', '水', '木', '金', '土'];

let state = null;
let draft = {};                        // 商品ID -> 発注単位数
let filters = { cat: '', scope: 'need', stockCat: '', stockStorage: '' };

/* ---------------------------------------------------------
 * 保存と読み込み
 * -------------------------------------------------------*/

function defaultState() {
  return {
    settings: {
      theater: '', staff: '', cycleDays: 7, safetyDays: 3, seasonRate: 100,
      attendance: SEED_ATTENDANCE.slice(), theme: 'auto'
    },
    items: SEED_ITEMS.map((it, i) => ({
      id: 'item-' + (i + 1), code: it.code, name: it.name, category: it.category,
      vendor: it.vendor, unit: it.unit, packSize: it.packSize, price: it.price,
      pi: it.pi, lead: it.lead, lot: it.lot, storage: it.storage, active: true
    })),
    stock: {}, orders: [], seq: SEED_ITEMS.length
  };
}

function load() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultState();
    const p = JSON.parse(raw);
    const base = defaultState();
    return {
      settings: Object.assign(base.settings, p.settings || {}),
      items: Array.isArray(p.items) ? p.items : base.items,
      stock: p.stock || {},
      orders: Array.isArray(p.orders) ? p.orders : [],
      seq: p.seq || (p.items ? p.items.length : base.seq)
    };
  } catch (e) {
    console.error('保存データを読めませんでした', e);
    return defaultState();
  }
}

function save() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (e) {
    toast('保存できませんでした。ブラウザの空き容量を確認してください。');
  }
}

/* ---------------------------------------------------------
 * 小道具
 * -------------------------------------------------------*/

const $ = id => document.getElementById(id);

const yen = n => '¥' + Math.round(n).toLocaleString('ja-JP');
const num = (n, d = 0) => Number(n).toLocaleString('ja-JP', { maximumFractionDigits: d });

function esc(s) {
  return String(s).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function parseDate(s) { const [y, m, d] = String(s).split('-').map(Number); return new Date(y, m - 1, d); }
function fmtDate(d) { const p = n => String(n).padStart(2, '0'); return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`; }
function addDays(d, n) { const r = new Date(d.getTime()); r.setDate(r.getDate() + n); return r; }
function fmtDateJa(s) { if (!s) return '-'; const d = parseDate(s); return `${d.getMonth() + 1}/${d.getDate()}(${WEEK[d.getDay()]})`; }

let toastTimer = null;
function toast(msg) {
  const el = $('toast');
  el.textContent = msg;
  el.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.hidden = true; }, 2600);
}

function download(filename, content, mime) {
  const url = URL.createObjectURL(new Blob([content], { type: mime || 'text/plain;charset=utf-8' }));
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** Excelで文字化けしないようBOM付きUTF-8で書き出す */
function downloadCsv(filename, rows) {
  const body = rows.map(row => row.map(cell => {
    const s = cell == null ? '' : String(cell);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }).join(',')).join('\r\n');
  download(filename, '﻿' + body, 'text/csv;charset=utf-8');
}

function parseCsv(text) {
  const rows = [];
  let row = [], field = '', quoted = false;
  const src = text.replace(/^﻿/, '');
  for (let i = 0; i < src.length; i++) {
    const c = src[i];
    if (quoted) {
      if (c === '"') { if (src[i + 1] === '"') { field += '"'; i++; } else quoted = false; }
      else field += c;
    } else if (c === '"') quoted = true;
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
    else if (c !== '\r') field += c;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows.filter(r => r.some(c => c.trim() !== ''));
}

/* ---------------------------------------------------------
 * 需要とランウェイの計算
 * -------------------------------------------------------*/

const season = () => (Number(state.settings.seasonRate) || 100) / 100;

function avgAttendance() {
  const a = state.settings.attendance;
  return (a.reduce((t, n) => t + (Number(n) || 0), 0) / 7) * season();
}

function forecastAttendance(start, days) {
  let total = 0;
  for (let i = 0; i < days; i++) total += Number(state.settings.attendance[addDays(start, i).getDay()]) || 0;
  return total * season();
}

const stockQty = id => (state.stock[id] ? Number(state.stock[id].qty) || 0 : 0);

function incomingQty(id) {
  let t = 0;
  state.orders.forEach(o => {
    if (o.status !== 'ordered') return;
    o.lines.forEach(l => { if (l.itemId === id) t += l.qty * l.packSize; });
  });
  return t;
}

const ceilToLot = (units, lot) => { const l = Math.max(1, Number(lot) || 1); return Math.ceil(units / l) * l; };

/**
 * 1商品分の計算。
 *   ランウェイ  = (在庫 + 入荷予定) ÷ 1日あたり消費数
 *   デッドライン = ランウェイ − リードタイム（この日を過ぎると欠品する）
 *   推奨発注数  = 発注点 − 在庫 − 入荷予定 を入数で割り、ロット単位に切り上げ
 */
function calc(item, today) {
  const cycle = Number(state.settings.cycleDays) || 1;
  const lead = Number(item.lead) || 0;
  const pi = Number(item.pi) || 0;
  const packSize = Math.max(1, Number(item.packSize) || 1);

  const cover = lead + cycle;
  const demand = forecastAttendance(addDays(today, 1), cover) * pi / 100;
  const dailyUse = avgAttendance() * pi / 100;
  const safety = dailyUse * (Number(state.settings.safetyDays) || 0);

  const onHand = stockQty(item.id);
  const incoming = incomingQty(item.id);
  const counted = !!state.stock[item.id];

  const gap = demand + safety - onHand - incoming;
  const recommend = gap > 0 ? ceilToLot(Math.ceil(gap / packSize), item.lot) : 0;

  const runway = dailyUse > 0 ? (onHand + incoming) / dailyUse : Infinity;
  const deadline = runway - lead;

  const ordered = draft[item.id] || 0;
  const runwayAfter = dailyUse > 0 ? (onHand + incoming + ordered * packSize) / dailyUse : Infinity;

  let level;
  if (!counted) level = 'none';
  else if (!isFinite(runway)) level = 'ok';
  else if (deadline <= 0) level = 'crit';
  else if (deadline <= cycle) level = 'warn';
  else level = 'ok';

  return { item, dailyUse, onHand, incoming, counted, safety, recommend,
           runway, runwayAfter, deadline, level, ordered, packSize, cover };
}

const activeItems = () => state.items.filter(i => i.active !== false);

const LEVEL_TEXT = { crit: '今日が期限', warn: '要発注', ok: '余裕', none: '在庫未入力' };
const LEVEL_ORDER = { crit: 0, warn: 1, none: 2, ok: 3 };

function currentRows() {
  const today = new Date();
  return activeItems()
    .filter(i => !filters.cat || i.category === filters.cat)
    .map(i => calc(i, today))
    .filter(r => filters.scope === 'all' || r.level === 'crit' || r.level === 'warn' ||
                 r.recommend > 0 || r.ordered > 0)
    .sort((a, b) => LEVEL_ORDER[a.level] - LEVEL_ORDER[b.level] ||
                    a.deadline - b.deadline ||
                    a.item.name.localeCompare(b.item.name, 'ja'));
}

/* ---------------------------------------------------------
 * タイムテーブル
 * -------------------------------------------------------*/

const pct = d => Math.max(0, Math.min(SPAN_DAYS, d)) / SPAN_DAYS * 100;

function renderScale() {
  const today = new Date();
  let html = '';
  for (let d = 0; d <= SPAN_DAYS; d += 2) {
    const date = addDays(today, d);
    const label = d === 0 ? '今日' : `${date.getMonth() + 1}/${date.getDate()}`;
    html += `<span style="left:${d / SPAN_DAYS * 100}%">${label}</span>`;
  }
  $('ttScale').innerHTML = html;

  const end = addDays(today, SPAN_DAYS);
  $('ttScaleMobile').innerHTML =
    `<span style="left:0">今日</span>` +
    `<span style="left:50%">${addDays(today, 7).getMonth() + 1}/${addDays(today, 7).getDate()}</span>` +
    `<span style="left:100%">${end.getMonth() + 1}/${end.getDate()}</span>`;
}

function runwayHtml(r) {
  const barW = isFinite(r.runway) ? pct(r.runway) : 100;
  const afterW = isFinite(r.runwayAfter) ? pct(r.runwayAfter) : 100;
  const showGhost = r.ordered > 0 && afterW > barW + 0.5;
  const dl = r.deadline;

  let marker = '';
  if (isFinite(dl) && dl > 0 && dl <= SPAN_DAYS) {
    marker = `<div class="runway__deadline" style="left:${pct(dl)}%"
                title="この日までに発注しないと間に合いません"></div>`;
  }

  // 残り日数はバーの終端に置く。バーが長いときは内側に入れて枠外に逃がさない
  let over = '';
  if (!isFinite(r.runway)) over = `<div class="runway__over" style="left:0">消費予測なし</div>`;
  else if (r.runway > SPAN_DAYS) over = `<div class="runway__over is-inside" style="right:0">${SPAN_DAYS}日以上</div>`;
  else if (r.counted) {
    const inside = barW > 78;
    over = `<div class="runway__over${inside ? ' is-inside' : ''}" style="${inside ? `right:${100 - barW}%` : `left:${barW}%`}">残り ${num(r.runway, 1)}日</div>`;
  }

  return `
    <div class="runway">
      <div class="runway__grid"></div>
      ${showGhost ? `<div class="runway__ghost" style="left:calc(${barW}% + 2px);width:calc(${afterW - barW}% - 2px)"></div>` : ''}
      <div class="runway__bar is-${r.level === 'none' ? 'warn' : r.level}" style="width:${r.counted ? barW : 0}%"></div>
      ${marker}
      ${over}
    </div>`;
}

function renderRunway() {
  renderScale();
  const rows = currentRows();
  const today = new Date();

  $('ttBody').innerHTML = rows.map(r => {
    const it = r.item;
    const amount = r.ordered * it.price;
    return `
    <div class="tt__row" data-id="${it.id}">
      <div class="tt__name">
        <b>${esc(it.name)}</b>
        <small>${esc(it.code)}・${CATEGORY_LABEL[it.category]}・${esc(it.vendor)}</small>
      </div>
      <span class="state state--${r.level}">${LEVEL_TEXT[r.level]}</span>
      ${runwayHtml(r)}
      <div class="tt__rec">${r.recommend > 0 ? `<b>${num(r.recommend)}</b>${esc(it.unit)}` : '—'}</div>
      <div class="stepper">
        <button type="button" data-step="-1" aria-label="${esc(it.name)}の発注数を減らす">−</button>
        <input type="number" class="num" min="0" step="${Math.max(1, it.lot)}"
               value="${r.ordered}" data-qty aria-label="${esc(it.name)}の発注数">
        <button type="button" data-step="1" aria-label="${esc(it.name)}の発注数を増やす">＋</button>
      </div>
      <div class="tt__amount num ${r.ordered ? 'is-set' : ''}">${r.ordered ? yen(amount) : '—'}</div>
    </div>`;
  }).join('');

  $('ttEmpty').hidden = rows.length > 0;
  renderVerdict(today);
  renderDock();
}

function renderVerdict(today) {
  const all = activeItems().map(i => calc(i, today));
  const crit = all.filter(r => r.level === 'crit');
  const warn = all.filter(r => r.level === 'warn');
  const none = all.filter(r => r.level === 'none');
  const need = crit.length + warn.length;

  const count = $('vCount');
  count.textContent = need;
  count.className = 'verdict__count num ' + (crit.length ? 'is-crit' : need ? 'is-warn' : 'is-ok');

  $('vUnit').textContent = need ? '品目に発注が必要です' : '品目。今日の発注は不要です';

  const parts = [];
  if (crit.length) parts.push(`うち<b>${crit.length}品目</b>は今日発注しないとリードタイムに間に合いません。`);
  if (warn.length) parts.push(`${warn.length}品目は次の発注サイクル（${state.settings.cycleDays}日）内に在庫が切れます。`);
  if (none.length) parts.push(`${none.length}品目は在庫が未入力のため判定できません。`);
  if (!parts.length) parts.push('すべての商品が発注サイクルを越える在庫を持っています。');
  $('vSub').innerHTML = parts.join(' ');

  let total = 0;
  state.items.forEach(i => { total += (draft[i.id] || 0) * i.price; });
  $('vAmount').textContent = yen(total);
}

function renderDock() {
  let count = 0, total = 0;
  state.items.forEach(i => {
    const q = draft[i.id] || 0;
    if (q > 0) { count++; total += q * i.price; }
  });
  $('dockCount').textContent = count;
  $('dockTotal').textContent = yen(total);
  $('dock').classList.toggle('is-up', count > 0);
}

function setQty(id, qty) {
  const item = state.items.find(i => i.id === id);
  if (!item) return;
  const q = Math.max(0, Math.round(Number(qty) || 0));
  if (q > 0) draft[id] = q; else delete draft[id];

  const row = document.querySelector(`.tt__row[data-id="${id}"]`);
  if (row) {
    const r = calc(item, new Date());
    const input = row.querySelector('[data-qty]');
    if (input && Number(input.value) !== q) input.value = q;
    const bar = row.querySelector('.runway__bar');
    const ghost = row.querySelector('.runway__ghost');
    const barW = isFinite(r.runway) ? pct(r.runway) : 100;
    const afterW = isFinite(r.runwayAfter) ? pct(r.runwayAfter) : 100;
    if (ghost) {
      const show = q > 0 && afterW > barW + 0.5;
      ghost.style.display = show ? '' : 'none';
      ghost.style.left = `calc(${barW}% + 2px)`;
      ghost.style.width = `calc(${afterW - barW}% - 2px)`;
    } else if (q > 0 && afterW > barW + 0.5 && bar) {
      const el = document.createElement('div');
      el.className = 'runway__ghost';
      el.style.left = `calc(${barW}% + 2px)`;
      el.style.width = `calc(${afterW - barW}% - 2px)`;
      bar.parentNode.insertBefore(el, bar);
    }
    const amt = row.querySelector('.tt__amount');
    amt.textContent = q ? yen(q * item.price) : '—';
    amt.classList.toggle('is-set', q > 0);
  }
  renderVerdict(new Date());
  renderDock();
}

/* ---------------------------------------------------------
 * 発注の確定・書き出し
 * -------------------------------------------------------*/

function draftLines() {
  const lines = [];
  state.items.forEach(i => {
    const qty = Number(draft[i.id]) || 0;
    if (qty > 0) lines.push({ itemId: i.id, code: i.code, name: i.name, vendor: i.vendor,
                              unit: i.unit, packSize: i.packSize, price: i.price, qty });
  });
  return lines;
}

function groupByVendor(lines) {
  const map = new Map();
  lines.forEach(l => { if (!map.has(l.vendor)) map.set(l.vendor, []); map.get(l.vendor).push(l); });
  return map;
}

const orderTotal = o => o.lines.reduce((t, l) => t + l.qty * l.price, 0);

function submitOrder() {
  const lines = draftLines();
  if (!lines.length) { toast('発注数が入っていません。'); return; }
  const groups = groupByVendor(lines);
  if (!confirm(`${groups.size}件の発注（${[...groups.keys()].join('、')}）を確定します。よろしいですか？`)) return;

  const today = fmtDate(new Date());
  groups.forEach((vLines, vendor) => {
    const maxLead = Math.max(...vLines.map(l => {
      const it = state.items.find(i => i.id === l.itemId);
      return it ? Number(it.lead) || 0 : 0;
    }));
    state.orders.push({
      id: 'order-' + Date.now() + '-' + Math.random().toString(36).slice(2, 7),
      createdAt: Date.now(), orderDate: today,
      deliveryDate: fmtDate(addDays(parseDate(today), maxLead)),
      vendor, staff: state.settings.staff, status: 'ordered', lines: vLines
    });
  });

  draft = {};
  save();
  renderRunway();
  toast(`${groups.size}件の発注を登録しました。`);
}

function exportOrderCsv() {
  const lines = draftLines();
  if (!lines.length) { toast('発注数が入っていません。'); return; }
  const today = fmtDate(new Date());
  const rows = [['発注日', '仕入先', '商品コード', '商品名', '発注単位', '入数', '発注数', '単価', '金額']];
  lines.forEach(l => rows.push([today, l.vendor, l.code, l.name, l.unit, l.packSize, l.qty, l.price, l.qty * l.price]));
  rows.push([], ['', '', '', '', '', '', '', '合計', lines.reduce((t, l) => t + l.qty * l.price, 0)]);
  downloadCsv(`発注データ_${today}.csv`, rows);
}

function printOrder() {
  const lines = draftLines();
  if (!lines.length) { toast('発注数が入っていません。'); return; }
  const today = fmtDate(new Date());
  let html = '';
  groupByVendor(lines).forEach((vLines, vendor) => {
    const maxLead = Math.max(...vLines.map(l => {
      const it = state.items.find(i => i.id === l.itemId);
      return it ? Number(it.lead) || 0 : 0;
    }));
    const total = vLines.reduce((t, l) => t + l.qty * l.price, 0);
    html += `
    <div class="sheet">
      <h1>発 注 書</h1>
      <div class="sheet__meta">
        <div><span>発注先</span><strong>${esc(vendor)} 御中</strong></div>
        <div><span>発注日</span>${today}</div>
        <div><span>納品希望日</span>${fmtDate(addDays(parseDate(today), maxLead))}</div>
        <div><span>発注元</span>${esc(state.settings.theater || '（劇場名未設定）')}</div>
        <div><span>担当者</span>${esc(state.settings.staff || '-')}</div>
      </div>
      <table>
        <thead><tr><th>商品コード</th><th>商品名</th><th>単位</th><th class="r">入数</th><th class="r">発注数</th><th class="r">単価</th><th class="r">金額</th></tr></thead>
        <tbody>${vLines.map(l => `<tr>
          <td>${esc(l.code)}</td><td>${esc(l.name)}</td><td>${esc(l.unit)}</td>
          <td class="r">${num(l.packSize)}</td><td class="r">${num(l.qty)}</td>
          <td class="r">${yen(l.price)}</td><td class="r">${yen(l.qty * l.price)}</td></tr>`).join('')}</tbody>
        <tfoot><tr><td colspan="6" class="r">合計（税抜）</td><td class="r">${yen(total)}</td></tr></tfoot>
      </table>
    </div>`;
  });
  $('printArea').innerHTML = html;
  window.print();
}

/* ---------------------------------------------------------
 * 棚卸
 * -------------------------------------------------------*/

function renderStock() {
  const items = activeItems()
    .filter(i => !filters.stockCat || i.category === filters.stockCat)
    .filter(i => !filters.stockStorage || i.storage === filters.stockStorage);

  $('stockBody').innerHTML = items.map(it => {
    const qty = stockQty(it.id);
    const packSize = Math.max(1, it.packSize);
    const cases = Math.floor(qty / packSize);
    const daily = avgAttendance() * (Number(it.pi) || 0) / 100;
    const safety = daily * (Number(state.settings.safetyDays) || 0);
    const updated = state.stock[it.id] ? state.stock[it.id].updatedAt : '';
    return `
    <tr data-id="${it.id}">
      <td><b>${esc(it.name)}</b><br><small style="color:var(--text-3)">${esc(it.code)}・1${esc(it.unit)}=${num(packSize)}</small></td>
      <td>${esc(it.storage)}</td>
      <td class="r"><input type="number" class="stock-case num" min="0" value="${cases}"></td>
      <td class="r"><input type="number" class="stock-rest num" min="0" value="${qty - cases * packSize}"></td>
      <td class="r num" data-total>${num(qty)}</td>
      <td class="r num" style="color:${qty < safety ? 'var(--crit)' : 'var(--text-3)'}">${num(safety)}</td>
      <td style="color:var(--text-3)">${updated ? esc(updated) : '未入力'}</td>
    </tr>`;
  }).join('');
}

function saveStock() {
  const now = fmtDate(new Date());
  document.querySelectorAll('#stockBody tr').forEach(tr => {
    const it = state.items.find(i => i.id === tr.dataset.id);
    if (!it) return;
    const cases = Number(tr.querySelector('.stock-case').value) || 0;
    const rest = Number(tr.querySelector('.stock-rest').value) || 0;
    state.stock[it.id] = { qty: cases * Math.max(1, it.packSize) + rest, updatedAt: now };
  });
  save();
  renderStock();
  toast('在庫を保存しました。');
}

/* ---------------------------------------------------------
 * 商品マスタ
 * -------------------------------------------------------*/

function renderItems() {
  $('itemsBody').innerHTML = state.items.map(it => `
    <tr${it.active === false ? ' style="opacity:.5"' : ''}>
      <td>${esc(it.code)}</td>
      <td><b>${esc(it.name)}</b></td>
      <td>${CATEGORY_LABEL[it.category] || '-'}</td>
      <td>${esc(it.vendor)}</td>
      <td>${esc(it.unit)}</td>
      <td class="r num">${num(it.packSize)}</td>
      <td class="r num">${yen(it.price)}</td>
      <td class="r num">${num(it.pi, 1)}</td>
      <td class="r num">${num(it.lead)}日</td>
      <td class="r num">${num(it.lot)}</td>
      <td>${esc(it.storage)}</td>
      <td class="r">
        <button class="btn btn--quiet" data-edit="${it.id}">編集</button>
        <button class="btn btn--quiet btn--danger" data-delete="${it.id}">削除</button>
      </td>
    </tr>`).join('');
  refreshVendors();
}

function refreshVendors() {
  const vendors = [...new Set(state.items.map(i => i.vendor).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'ja'));
  $('vendorList').innerHTML = vendors.map(v => `<option value="${esc(v)}">`).join('');
}

let editingId = null;

function openItemModal(id) {
  editingId = id;
  const it = id ? state.items.find(i => i.id === id) : null;
  $('itemModalTitle').textContent = it ? '商品を編集' : '商品を追加';
  $('f_code').value = it ? it.code : '';
  $('f_name').value = it ? it.name : '';
  $('f_category').value = it ? it.category : 'food';
  $('f_vendor').value = it ? it.vendor : '';
  $('f_unit').value = it ? it.unit : 'ケース';
  $('f_packSize').value = it ? it.packSize : 1;
  $('f_price').value = it ? it.price : 0;
  $('f_pi').value = it ? it.pi : 0;
  $('f_lead').value = it ? it.lead : 3;
  $('f_lot').value = it ? it.lot : 1;
  $('f_storage').value = it ? it.storage : '常温';
  $('f_active').checked = it ? it.active !== false : true;
  $('itemModal').hidden = false;
  $('f_name').focus();
}

function closeItemModal() { $('itemModal').hidden = true; editingId = null; }

function saveItem() {
  const name = $('f_name').value.trim();
  if (!name) { toast('商品名を入れてください。'); return; }
  const data = {
    code: $('f_code').value.trim() || '-', name,
    category: $('f_category').value,
    vendor: $('f_vendor').value.trim() || '未設定',
    unit: $('f_unit').value.trim() || 'ケース',
    packSize: Math.max(1, Number($('f_packSize').value) || 1),
    price: Math.max(0, Number($('f_price').value) || 0),
    pi: Math.max(0, Number($('f_pi').value) || 0),
    lead: Math.max(0, Number($('f_lead').value) || 0),
    lot: Math.max(1, Number($('f_lot').value) || 1),
    storage: $('f_storage').value,
    active: $('f_active').checked
  };
  if (editingId) Object.assign(state.items.find(i => i.id === editingId), data);
  else {
    state.seq++;
    state.items.push(Object.assign({ id: 'item-' + state.seq + '-' + Date.now().toString(36) }, data));
  }
  save();
  closeItemModal();
  renderItems();
  toast('商品マスタを更新しました。');
}

function deleteItem(id) {
  const it = state.items.find(i => i.id === id);
  if (!it || !confirm(`「${it.name}」を削除します。よろしいですか？\n（発注履歴は残ります）`)) return;
  state.items = state.items.filter(i => i.id !== id);
  delete state.stock[id];
  delete draft[id];
  save();
  renderItems();
  toast('商品を削除しました。');
}

const ITEM_HEADER = ['商品コード', '商品名', 'カテゴリ', '仕入先', '発注単位', '入数', '単価',
                     'PI値', 'リードタイム', 'ロット', '保管場所', '発注対象'];

function exportItems() {
  const rows = [ITEM_HEADER];
  state.items.forEach(i => rows.push([i.code, i.name, CATEGORY_LABEL[i.category] || i.category,
    i.vendor, i.unit, i.packSize, i.price, i.pi, i.lead, i.lot, i.storage,
    i.active === false ? '×' : '○']));
  downloadCsv('商品マスタ.csv', rows);
}

function importItems(file) {
  const reader = new FileReader();
  reader.onload = () => {
    const rows = parseCsv(String(reader.result));
    if (rows.length < 2) { toast('読み込める行がありません。'); return; }
    if (!confirm('CSVの内容で商品マスタを置き換えます。よろしいですか？\n（在庫と発注履歴は残ります）')) return;
    const items = [];
    rows.slice(1).forEach((r, idx) => {
      const name = (r[1] || '').trim();
      if (!name) return;
      const code = (r[0] || '-').trim();
      const existing = state.items.find(i => i.code === code && i.name === name);
      items.push({
        id: existing ? existing.id : 'item-import-' + idx + '-' + Date.now().toString(36),
        code, name,
        category: CATEGORY_KEY[(r[2] || '').trim()] || 'food',
        vendor: (r[3] || '未設定').trim(),
        unit: (r[4] || 'ケース').trim(),
        packSize: Math.max(1, Number(r[5]) || 1),
        price: Math.max(0, Number(r[6]) || 0),
        pi: Math.max(0, Number(r[7]) || 0),
        lead: Math.max(0, Number(r[8]) || 0),
        lot: Math.max(1, Number(r[9]) || 1),
        storage: (r[10] || '常温').trim(),
        active: (r[11] || '○').trim() !== '×'
      });
    });
    if (!items.length) { toast('読み込める商品がありませんでした。'); return; }
    state.items = items;
    save();
    renderItems();
    toast(`${items.length}件の商品を読み込みました。`);
  };
  reader.readAsText(file, 'UTF-8');
}

/* ---------------------------------------------------------
 * 発注履歴
 * -------------------------------------------------------*/

function renderHistory() {
  const orders = state.orders.slice().sort((a, b) => b.createdAt - a.createdAt);
  const opened = new Set([...document.querySelectorAll('.order-item[open]')].map(d => d.dataset.id));
  $('historyEmpty').hidden = orders.length > 0;
  $('historyList').innerHTML = orders.map(o => `
    <details class="order-item" data-id="${o.id}" ${opened.has(o.id) ? 'open' : ''}>
      <summary>
        <span class="tag tag--${o.status}">${STATUS_LABEL[o.status]}</span>
        <b>${fmtDateJa(o.orderDate)}</b>
        <span>${esc(o.vendor)}</span>
        <span class="order-item__meta">納品予定 ${fmtDateJa(o.deliveryDate)}／${o.lines.length}品目／${yen(orderTotal(o))}</span>
      </summary>
      <div class="table-wrap">
        <table class="grid">
          <thead><tr><th>コード</th><th>商品名</th><th class="r">発注数</th><th class="r">入数</th><th class="r">単価</th><th class="r">金額</th></tr></thead>
          <tbody>${o.lines.map(l => `<tr>
            <td>${esc(l.code)}</td><td>${esc(l.name)}</td>
            <td class="r num">${num(l.qty)}${esc(l.unit)}</td>
            <td class="r num">${num(l.packSize)}</td>
            <td class="r num">${yen(l.price)}</td>
            <td class="r num">${yen(l.qty * l.price)}</td></tr>`).join('')}</tbody>
        </table>
      </div>
      <div class="btn-row">
        ${o.status === 'ordered' ? `<button class="btn btn--primary" data-receive="${o.id}">入荷済にする</button>
        <button class="btn" data-cancel="${o.id}">取消</button>` : ''}
        <button class="btn" data-csv="${o.id}">CSV</button>
        <button class="btn btn--danger" data-remove="${o.id}">履歴を削除</button>
      </div>
    </details>`).join('');
}

function receiveOrder(id) {
  const o = state.orders.find(x => x.id === id);
  if (!o || o.status !== 'ordered') return;
  if (!confirm('入荷済にして、発注数を在庫に加算します。よろしいですか？')) return;
  const now = fmtDate(new Date());
  o.lines.forEach(l => { state.stock[l.itemId] = { qty: stockQty(l.itemId) + l.qty * l.packSize, updatedAt: now }; });
  o.status = 'received';
  o.receivedAt = now;
  save();
  renderHistory();
  toast('入荷処理をして在庫に反映しました。');
}

function cancelOrder(id) {
  const o = state.orders.find(x => x.id === id);
  if (!o || !confirm('この発注を取消にします。よろしいですか？')) return;
  o.status = 'cancelled';
  save();
  renderHistory();
  toast('発注を取消しました。');
}

function removeOrder(id) {
  if (!confirm('この発注履歴を削除します。よろしいですか？')) return;
  state.orders = state.orders.filter(o => o.id !== id);
  save();
  renderHistory();
  toast('履歴を削除しました。');
}

function exportHistoryCsv(id) {
  const o = state.orders.find(x => x.id === id);
  if (!o) return;
  const rows = [['発注日', '仕入先', '納品予定日', '状態', '商品コード', '商品名', '発注単位', '入数', '発注数', '単価', '金額']];
  o.lines.forEach(l => rows.push([o.orderDate, o.vendor, o.deliveryDate, STATUS_LABEL[o.status],
    l.code, l.name, l.unit, l.packSize, l.qty, l.price, l.qty * l.price]));
  downloadCsv(`発注書_${o.vendor}_${o.orderDate}.csv`, rows);
}

/* ---------------------------------------------------------
 * 設定・テーマ
 * -------------------------------------------------------*/

function renderSettings() {
  const s = state.settings;
  $('setTheater').value = s.theater;
  $('setStaff').value = s.staff;
  $('setCycle').value = s.cycleDays;
  $('setSafety').value = s.safetyDays;
  $('setSeason').value = s.seasonRate;
  $('weekGrid').innerHTML = WEEK.map((w, i) => `
    <label class="field">
      <span>${w}曜日</span>
      <input type="number" class="attendance num" min="0" step="10" data-day="${i}" value="${s.attendance[i]}">
    </label>`).join('');
}

function saveSettings() {
  const s = state.settings;
  s.theater = $('setTheater').value.trim();
  s.staff = $('setStaff').value.trim();
  s.cycleDays = Math.max(1, Number($('setCycle').value) || 7);
  s.safetyDays = Math.max(0, Number($('setSafety').value) || 0);
  s.seasonRate = Math.max(10, Number($('setSeason').value) || 100);
  document.querySelectorAll('.attendance').forEach(i => {
    s.attendance[Number(i.dataset.day)] = Math.max(0, Number(i.value) || 0);
  });
  save();
  applyTheater();
  toast('設定を保存しました。');
}

function applyTheater() {
  $('theaterName').textContent = state.settings.theater || '劇場名未設定';
}

function applyTheme() {
  const t = state.settings.theme || 'auto';
  if (t === 'auto') document.documentElement.removeAttribute('data-theme');
  else document.documentElement.setAttribute('data-theme', t);
}

function cycleTheme() {
  const order = ['auto', 'light', 'dark'];
  const next = order[(order.indexOf(state.settings.theme || 'auto') + 1) % 3];
  state.settings.theme = next;
  applyTheme();
  save();
  toast({ auto: '表示テーマ：端末に合わせる', light: '表示テーマ：明るい', dark: '表示テーマ：暗い' }[next]);
}

function exportJson() {
  download(`発注ツールデータ_${fmtDate(new Date())}.json`, JSON.stringify(state, null, 2), 'application/json');
}

function importJson(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const p = JSON.parse(String(reader.result));
      if (!p.items) throw new Error('形式が違います');
      if (!confirm('いまのデータをファイルの内容で置き換えます。よろしいですか？')) return;
      state = Object.assign(defaultState(), p);
      draft = {};
      save();
      boot();
      toast('データを戻しました。');
    } catch (e) {
      toast('ファイルを読み込めませんでした。');
    }
  };
  reader.readAsText(file, 'UTF-8');
}

function resetAll() {
  if (!confirm('商品マスタ・在庫・発注履歴・設定をすべて初期状態に戻します。\nこの操作は元に戻せません。よろしいですか？')) return;
  state = defaultState();
  draft = {};
  save();
  boot();
  toast('初期状態に戻しました。');
}

/* ---------------------------------------------------------
 * 画面の切り替えとイベント
 * -------------------------------------------------------*/

function switchView(name) {
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('is-active', t.dataset.view === name));
  document.querySelectorAll('.view').forEach(v => v.classList.toggle('is-active', v.id === 'view-' + name));
  $('dock').style.display = name === 'runway' ? '' : 'none';
  if (name === 'runway') renderRunway();
  if (name === 'stock') renderStock();
  if (name === 'items') renderItems();
  if (name === 'history') renderHistory();
  if (name === 'settings') renderSettings();
}

function chipGroup(containerId, key, onChange) {
  $(containerId).addEventListener('click', e => {
    const chip = e.target.closest('.chip');
    if (!chip) return;
    [...chip.parentNode.children].forEach(c => c.classList.toggle('is-on', c === chip));
    filters[key] = chip.dataset.cat ?? chip.dataset.scope ?? chip.dataset.storage ?? '';
    onChange();
  });
}

function bind() {
  $('tabs').addEventListener('click', e => {
    const tab = e.target.closest('.tab');
    if (tab) switchView(tab.dataset.view);
  });
  $('themeToggle').addEventListener('click', cycleTheme);

  chipGroup('catChips', 'cat', renderRunway);
  chipGroup('scopeChips', 'scope', renderRunway);
  chipGroup('stockCatChips', 'stockCat', renderStock);
  chipGroup('stockStorageChips', 'stockStorage', renderStock);

  $('ttBody').addEventListener('input', e => {
    const input = e.target.closest('[data-qty]');
    if (input) setQty(input.closest('.tt__row').dataset.id, input.value);
  });
  $('ttBody').addEventListener('click', e => {
    const btn = e.target.closest('[data-step]');
    if (!btn) return;
    const row = btn.closest('.tt__row');
    const item = state.items.find(i => i.id === row.dataset.id);
    const lot = Math.max(1, item ? item.lot : 1);
    setQty(row.dataset.id, (draft[row.dataset.id] || 0) + Number(btn.dataset.step) * lot);
  });

  $('applyRec').addEventListener('click', () => {
    currentRows().forEach(r => { if (r.recommend > 0) draft[r.item.id] = r.recommend; });
    renderRunway();
    toast('推奨数を入れました。');
  });
  $('clearQty').addEventListener('click', () => { draft = {}; renderRunway(); });

  $('submitOrder').addEventListener('click', submitOrder);
  $('exportOrderCsv').addEventListener('click', exportOrderCsv);
  $('printOrder').addEventListener('click', printOrder);

  $('stockBody').addEventListener('input', e => {
    const tr = e.target.closest('tr');
    if (!tr || !e.target.matches('.stock-case, .stock-rest')) return;
    const it = state.items.find(i => i.id === tr.dataset.id);
    if (!it) return;
    const cases = Number(tr.querySelector('.stock-case').value) || 0;
    const rest = Number(tr.querySelector('.stock-rest').value) || 0;
    tr.querySelector('[data-total]').textContent = num(cases * Math.max(1, it.packSize) + rest);
  });
  $('saveStock').addEventListener('click', saveStock);

  $('addItem').addEventListener('click', () => openItemModal(null));
  $('itemsBody').addEventListener('click', e => {
    const edit = e.target.closest('[data-edit]');
    const del = e.target.closest('[data-delete]');
    if (edit) openItemModal(edit.dataset.edit);
    if (del) deleteItem(del.dataset.delete);
  });
  $('saveItem').addEventListener('click', saveItem);
  document.querySelectorAll('[data-close]').forEach(el => el.addEventListener('click', closeItemModal));
  $('exportItems').addEventListener('click', exportItems);
  $('importItems').addEventListener('change', e => {
    if (e.target.files[0]) importItems(e.target.files[0]);
    e.target.value = '';
  });

  $('historyList').addEventListener('click', e => {
    const r = e.target.closest('[data-receive]'), c = e.target.closest('[data-cancel]');
    const v = e.target.closest('[data-csv]'), d = e.target.closest('[data-remove]');
    if (r) receiveOrder(r.dataset.receive);
    if (c) cancelOrder(c.dataset.cancel);
    if (v) exportHistoryCsv(v.dataset.csv);
    if (d) removeOrder(d.dataset.remove);
  });

  $('saveSettings').addEventListener('click', saveSettings);
  $('exportJson').addEventListener('click', exportJson);
  $('importJson').addEventListener('change', e => {
    if (e.target.files[0]) importJson(e.target.files[0]);
    e.target.value = '';
  });
  $('resetAll').addEventListener('click', resetAll);

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && !$('itemModal').hidden) closeItemModal();
  });
}

function boot() {
  applyTheme();
  applyTheater();
  refreshVendors();
  renderSettings();
  switchView(document.querySelector('.tab.is-active').dataset.view);
}

state = load();
bind();
boot();
