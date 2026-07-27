/* =========================================================
 *  売店発注ツール
 *  フード / ドリンク / 包材の発注量を、来場者予測とPI値から算出する。
 *  データはブラウザの localStorage に保存する。
 * =======================================================*/

const STORAGE_KEY = 'concession-order-tool-v1';

const CATEGORY_LABEL = { food: 'フード', drink: 'ドリンク', packaging: '包材' };
const CATEGORY_KEY = { 'フード': 'food', 'ドリンク': 'drink', '包材': 'packaging' };
const STATUS_LABEL = { ordered: '発注済', received: '入荷済', cancelled: '取消' };
const WEEK_LABEL = ['日', '月', '火', '水', '木', '金', '土'];

let state = null;
/** 発注作成タブの入力値（商品ID -> 発注単位数） */
let draft = {};

/* ---------------------------------------------------------
 * 初期化
 * -------------------------------------------------------*/

function defaultState() {
  return {
    settings: {
      theater: '',
      staff: '',
      cycleDays: 7,
      safetyDays: 3,
      seasonRate: 100,
      attendance: SEED_ATTENDANCE.slice()
    },
    items: SEED_ITEMS.map((item, i) => ({
      id: 'item-' + (i + 1),
      code: item.code,
      name: item.name,
      category: item.category,
      vendor: item.vendor,
      unit: item.unit,
      packSize: item.packSize,
      price: item.price,
      pi: item.pi,
      lead: item.lead,
      lot: item.lot,
      storage: item.storage,
      active: true
    })),
    stock: {},
    orders: [],
    seq: SEED_ITEMS.length
  };
}

function load() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultState();
    const parsed = JSON.parse(raw);
    const base = defaultState();
    return {
      settings: Object.assign(base.settings, parsed.settings || {}),
      items: Array.isArray(parsed.items) ? parsed.items : base.items,
      stock: parsed.stock || {},
      orders: Array.isArray(parsed.orders) ? parsed.orders : [],
      seq: parsed.seq || (parsed.items ? parsed.items.length : base.seq)
    };
  } catch (e) {
    console.error('保存データの読み込みに失敗しました', e);
    return defaultState();
  }
}

function save() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (e) {
    toast('保存に失敗しました。ブラウザの空き容量を確認してください。');
  }
}

/* ---------------------------------------------------------
 * 汎用ユーティリティ
 * -------------------------------------------------------*/

function $(id) { return document.getElementById(id); }

function parseDate(s) {
  const [y, m, d] = String(s).split('-').map(Number);
  return new Date(y, m - 1, d);
}

function fmtDate(d) {
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function fmtDateJa(s) {
  if (!s) return '-';
  const d = parseDate(s);
  return `${d.getMonth() + 1}/${d.getDate()}(${WEEK_LABEL[d.getDay()]})`;
}

function addDays(d, n) {
  const r = new Date(d.getTime());
  r.setDate(r.getDate() + n);
  return r;
}

function yen(n) { return '¥' + Math.round(n).toLocaleString('ja-JP'); }

function num(n, digits = 0) {
  return Number(n).toLocaleString('ja-JP', { maximumFractionDigits: digits });
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}

let toastTimer = null;
function toast(message) {
  const el = $('toast');
  el.textContent = message;
  el.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.hidden = true; }, 2600);
}

function download(filename, content, mime) {
  const blob = new Blob([content], { type: mime || 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** Excelで文字化けしないようBOM付きUTF-8で出力する */
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
      if (c === '"') {
        if (src[i + 1] === '"') { field += '"'; i++; }
        else quoted = false;
      } else field += c;
    } else if (c === '"') {
      quoted = true;
    } else if (c === ',') {
      row.push(field); field = '';
    } else if (c === '\n') {
      row.push(field); rows.push(row); row = []; field = '';
    } else if (c !== '\r') {
      field += c;
    }
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows.filter(r => r.some(c => c.trim() !== ''));
}

/* ---------------------------------------------------------
 * 需要予測・発注数の計算
 * -------------------------------------------------------*/

function seasonFactor() {
  return (Number(state.settings.seasonRate) || 100) / 100;
}

/** 曜日別想定来場者数の平均（繁忙係数込み） */
function avgAttendance() {
  const a = state.settings.attendance;
  const sum = a.reduce((t, n) => t + (Number(n) || 0), 0);
  return (sum / 7) * seasonFactor();
}

/** startDate から days 日間の想定来場者数の合計 */
function forecastAttendance(startDate, days) {
  let total = 0;
  for (let i = 0; i < days; i++) {
    const d = addDays(startDate, i);
    total += Number(state.settings.attendance[d.getDay()]) || 0;
  }
  return total * seasonFactor();
}

function stockQty(itemId) {
  const s = state.stock[itemId];
  return s ? Number(s.qty) || 0 : 0;
}

/** 発注済み（未入荷）の数量を消費単位で集計 */
function incomingQty(itemId) {
  let total = 0;
  state.orders.forEach(order => {
    if (order.status !== 'ordered') return;
    order.lines.forEach(line => {
      if (line.itemId === itemId) total += line.qty * line.packSize;
    });
  });
  return total;
}

function ceilToLot(units, lot) {
  const l = Math.max(1, Number(lot) || 1);
  return Math.ceil(units / l) * l;
}

/**
 * 1商品分の発注計算。
 * カバー日数 = リードタイム + 発注サイクル
 * 不足数 = 予測消費 + 安全在庫 - 現在庫 - 入荷予定
 */
function calcRow(item, orderDateStr) {
  const start = addDays(parseDate(orderDateStr), 1);
  const coverDays = (Number(item.lead) || 0) + (Number(state.settings.cycleDays) || 1);
  const attendance = forecastAttendance(start, coverDays);
  const pi = Number(item.pi) || 0;

  const demand = attendance * pi / 100;
  const dailyUse = avgAttendance() * pi / 100;
  const safety = dailyUse * (Number(state.settings.safetyDays) || 0);
  const onHand = stockQty(item.id);
  const incoming = incomingQty(item.id);
  const gap = demand + safety - onHand - incoming;

  const packSize = Math.max(1, Number(item.packSize) || 1);
  const recommend = gap > 0 ? ceilToLot(Math.ceil(gap / packSize), item.lot) : 0;
  const stockDays = dailyUse > 0 ? (onHand + incoming) / dailyUse : Infinity;

  // 在庫を一度も入力していない商品は欠品と区別する
  const counted = !!state.stock[item.id];
  let level = 'ok';
  if (!counted) level = 'unset';
  else if (dailyUse > 0 && stockDays < (Number(item.lead) || 0)) level = 'danger';
  else if (recommend > 0) level = 'warn';

  return { item, coverDays, attendance, demand, safety, onHand, incoming, gap, recommend, stockDays, dailyUse, level, counted };
}

function activeItems() {
  return state.items.filter(i => i.active !== false);
}

/* ---------------------------------------------------------
 * タブ切り替え
 * -------------------------------------------------------*/

function switchTab(name) {
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('is-active', t.dataset.tab === name));
  document.querySelectorAll('.panel').forEach(p => p.classList.toggle('is-active', p.id === 'panel-' + name));
  if (name === 'dashboard') renderDashboard();
  if (name === 'order') renderOrder();
  if (name === 'stock') renderStock();
  if (name === 'master') renderMaster();
  if (name === 'history') renderHistory();
  if (name === 'settings') renderSettings();
}

/* ---------------------------------------------------------
 * ダッシュボード
 * -------------------------------------------------------*/

function renderDashboard() {
  const today = fmtDate(new Date());
  const rows = activeItems().map(item => calcRow(item, today));
  const need = rows.filter(r => r.recommend > 0);
  const danger = rows.filter(r => r.level === 'danger');

  $('kpiItems').textContent = activeItems().length;
  $('kpiNeed').textContent = need.length;
  $('kpiShort').textContent = danger.length;
  $('kpiOpen').textContent = state.orders.filter(o => o.status === 'ordered').length;

  const LEVEL_ORDER = { danger: 0, warn: 1, unset: 2, ok: 3 };
  const LEVEL_TEXT = { danger: '欠品リスク', warn: '要発注', unset: '在庫未入力' };
  const alerts = rows
    .filter(r => r.level !== 'ok')
    .sort((a, b) => LEVEL_ORDER[a.level] - LEVEL_ORDER[b.level] || a.stockDays - b.stockDays);

  $('alertBody').innerHTML = alerts.map(r => `
    <tr>
      <td><span class="badge badge--${r.level}">${LEVEL_TEXT[r.level]}</span></td>
      <td>${escapeHtml(r.item.name)}</td>
      <td>${CATEGORY_LABEL[r.item.category] || '-'}</td>
      <td class="num">${r.counted ? num(r.onHand) : '未入力'}</td>
      <td class="num">${!r.counted || r.stockDays === Infinity ? '—' : num(r.stockDays, 1) + '日'}</td>
      <td class="num">${r.gap > 0 ? num(r.gap) : '—'}</td>
      <td class="num">${r.recommend > 0 ? num(r.recommend) + r.item.unit : '—'}</td>
    </tr>`).join('');
  $('alertEmpty').hidden = alerts.length > 0;

  const recent = state.orders.slice().sort((a, b) => b.createdAt - a.createdAt).slice(0, 8);
  $('recentBody').innerHTML = recent.map(o => `
    <tr>
      <td>${fmtDateJa(o.orderDate)}</td>
      <td>${escapeHtml(o.vendor)}</td>
      <td>${fmtDateJa(o.deliveryDate)}</td>
      <td class="num">${o.lines.length}</td>
      <td class="num">${yen(orderTotal(o))}</td>
      <td><span class="badge badge--${o.status}">${STATUS_LABEL[o.status]}</span></td>
    </tr>`).join('');
  $('recentEmpty').hidden = recent.length > 0;
}

/* ---------------------------------------------------------
 * 発注作成
 * -------------------------------------------------------*/

function orderRows() {
  const date = $('orderDate').value || fmtDate(new Date());
  const vendor = $('filterVendor').value;
  const category = $('filterCategory').value;
  const needOnly = $('filterNeed').value === 'need';

  return activeItems()
    .filter(i => !vendor || i.vendor === vendor)
    .filter(i => !category || i.category === category)
    .map(i => calcRow(i, date))
    .filter(r => !needOnly || r.recommend > 0 || (draft[r.item.id] > 0))
    .sort((a, b) => a.item.vendor.localeCompare(b.item.vendor, 'ja') || a.item.code.localeCompare(b.item.code));
}

function renderOrder() {
  const date = $('orderDate').value || fmtDate(new Date());
  const cycle = Number(state.settings.cycleDays) || 1;
  const start = addDays(parseDate(date), 1);
  const attendance = forecastAttendance(start, cycle);
  $('forecastNote').innerHTML =
    `発注日翌日（${fmtDateJa(fmtDate(start))}）から ${cycle}日間の想定来場者数：<strong>${num(attendance)}人</strong>` +
    `／安全在庫 ${state.settings.safetyDays}日分・繁忙係数 ${state.settings.seasonRate}%`;

  const rows = orderRows();
  $('orderBody').innerHTML = rows.map(r => {
    const qty = draft[r.item.id] || 0;
    return `
    <tr data-id="${r.item.id}" class="row--${r.level}">
      <td>
        <div class="cell-main">${escapeHtml(r.item.name)}</div>
        <div class="cell-sub">${escapeHtml(r.item.code)}・${CATEGORY_LABEL[r.item.category]}・${escapeHtml(r.item.storage)}</div>
      </td>
      <td>${escapeHtml(r.item.vendor)}</td>
      <td class="num">${num(r.onHand)}</td>
      <td class="num">${num(r.demand)}</td>
      <td class="num">${num(r.safety)}</td>
      <td class="num">${r.incoming ? num(r.incoming) : '—'}</td>
      <td class="num ${r.gap > 0 ? 'is-short' : ''}">${r.gap > 0 ? '▲' + num(r.gap) : num(-r.gap)}</td>
      <td class="num">${num(r.item.packSize)}</td>
      <td class="num recommend">${num(r.recommend)}</td>
      <td class="num">
        <input type="number" class="qty" min="0" step="${Math.max(1, r.item.lot)}" value="${qty || ''}" data-id="${r.item.id}">
        <span class="unit">${escapeHtml(r.item.unit)}</span>
      </td>
      <td class="num amount">${qty ? yen(qty * r.item.price) : '—'}</td>
    </tr>`;
  }).join('');

  $('orderEmpty').hidden = rows.length > 0;
  updateOrderTotal();
}

function updateOrderTotal() {
  let count = 0, total = 0;
  state.items.forEach(item => {
    const qty = draft[item.id] || 0;
    if (qty > 0) { count++; total += qty * item.price; }
  });
  $('orderCount').textContent = count;
  $('orderTotal').textContent = yen(total);
}

function draftLines() {
  const lines = [];
  state.items.forEach(item => {
    const qty = Number(draft[item.id]) || 0;
    if (qty > 0) {
      lines.push({
        itemId: item.id, code: item.code, name: item.name, vendor: item.vendor,
        unit: item.unit, packSize: item.packSize, price: item.price, qty
      });
    }
  });
  return lines;
}

function groupByVendor(lines) {
  const map = new Map();
  lines.forEach(line => {
    if (!map.has(line.vendor)) map.set(line.vendor, []);
    map.get(line.vendor).push(line);
  });
  return map;
}

function submitOrder() {
  const lines = draftLines();
  if (!lines.length) { toast('発注数が入力されていません。'); return; }

  const orderDate = $('orderDate').value || fmtDate(new Date());
  const groups = groupByVendor(lines);
  const summary = [...groups.keys()].join('、');
  if (!confirm(`${groups.size}件の発注（${summary}）を確定します。よろしいですか？`)) return;

  groups.forEach((vendorLines, vendor) => {
    const maxLead = Math.max(...vendorLines.map(l => {
      const item = state.items.find(i => i.id === l.itemId);
      return item ? Number(item.lead) || 0 : 0;
    }));
    state.orders.push({
      id: 'order-' + Date.now() + '-' + Math.random().toString(36).slice(2, 7),
      createdAt: Date.now(),
      orderDate,
      deliveryDate: fmtDate(addDays(parseDate(orderDate), maxLead)),
      vendor,
      staff: state.settings.staff,
      status: 'ordered',
      lines: vendorLines
    });
  });

  draft = {};
  save();
  renderOrder();
  toast(`${groups.size}件の発注を登録しました。`);
}

function orderTotal(order) {
  return order.lines.reduce((t, l) => t + l.qty * l.price, 0);
}

function exportOrderCsv() {
  const lines = draftLines();
  if (!lines.length) { toast('発注数が入力されていません。'); return; }
  const orderDate = $('orderDate').value;
  const rows = [['発注日', '仕入先', '商品コード', '商品名', '発注単位', '入数', '発注数', '単価', '金額']];
  lines.forEach(l => {
    rows.push([orderDate, l.vendor, l.code, l.name, l.unit, l.packSize, l.qty, l.price, l.qty * l.price]);
  });
  rows.push([]);
  rows.push(['', '', '', '', '', '', '', '合計', lines.reduce((t, l) => t + l.qty * l.price, 0)]);
  downloadCsv(`発注データ_${orderDate}.csv`, rows);
}

function printOrder() {
  const lines = draftLines();
  if (!lines.length) { toast('発注数が入力されていません。'); return; }
  const orderDate = $('orderDate').value;
  const groups = groupByVendor(lines);
  let html = '';

  groups.forEach((vendorLines, vendor) => {
    const maxLead = Math.max(...vendorLines.map(l => {
      const item = state.items.find(i => i.id === l.itemId);
      return item ? Number(item.lead) || 0 : 0;
    }));
    const delivery = fmtDate(addDays(parseDate(orderDate), maxLead));
    const total = vendorLines.reduce((t, l) => t + l.qty * l.price, 0);

    html += `
    <div class="sheet">
      <h1>発 注 書</h1>
      <div class="sheet__meta">
        <div><span>発注先</span><strong>${escapeHtml(vendor)} 御中</strong></div>
        <div><span>発注日</span>${escapeHtml(orderDate)}</div>
        <div><span>納品希望日</span>${escapeHtml(delivery)}</div>
        <div><span>発注元</span>${escapeHtml(state.settings.theater || '（劇場名未設定）')}</div>
        <div><span>担当者</span>${escapeHtml(state.settings.staff || '-')}</div>
      </div>
      <table>
        <thead>
          <tr><th>商品コード</th><th>商品名</th><th>発注単位</th><th class="num">入数</th><th class="num">発注数</th><th class="num">単価</th><th class="num">金額</th></tr>
        </thead>
        <tbody>
          ${vendorLines.map(l => `
            <tr>
              <td>${escapeHtml(l.code)}</td>
              <td>${escapeHtml(l.name)}</td>
              <td>${escapeHtml(l.unit)}</td>
              <td class="num">${num(l.packSize)}</td>
              <td class="num">${num(l.qty)}</td>
              <td class="num">${yen(l.price)}</td>
              <td class="num">${yen(l.qty * l.price)}</td>
            </tr>`).join('')}
        </tbody>
        <tfoot>
          <tr><td colspan="6" class="num">合計（税抜）</td><td class="num">${yen(total)}</td></tr>
        </tfoot>
      </table>
    </div>`;
  });

  $('printArea').innerHTML = html;
  window.print();
}

/* ---------------------------------------------------------
 * 在庫入力
 * -------------------------------------------------------*/

function renderStock() {
  const category = $('stockCategory').value;
  const storage = $('stockStorage').value;
  const items = activeItems()
    .filter(i => !category || i.category === category)
    .filter(i => !storage || i.storage === storage);

  $('stockBody').innerHTML = items.map(item => {
    const qty = stockQty(item.id);
    const packSize = Math.max(1, item.packSize);
    const cases = Math.floor(qty / packSize);
    const rest = qty - cases * packSize;
    const dailyUse = avgAttendance() * (Number(item.pi) || 0) / 100;
    const safety = dailyUse * (Number(state.settings.safetyDays) || 0);
    const updated = state.stock[item.id] ? state.stock[item.id].updatedAt : '';
    return `
    <tr data-id="${item.id}">
      <td>
        <div class="cell-main">${escapeHtml(item.name)}</div>
        <div class="cell-sub">${escapeHtml(item.code)}・1${escapeHtml(item.unit)} = ${num(packSize)}</div>
      </td>
      <td>${escapeHtml(item.storage)}</td>
      <td class="num"><input type="number" class="stock-case" min="0" value="${cases}" data-id="${item.id}"><span class="unit">${escapeHtml(item.unit)}</span></td>
      <td class="num"><input type="number" class="stock-rest" min="0" value="${rest}" data-id="${item.id}"></td>
      <td class="num total" data-total="${item.id}">${num(qty)}</td>
      <td class="num ${qty < safety ? 'is-short' : ''}">${num(safety)}</td>
      <td class="cell-sub">${updated ? escapeHtml(updated) : '未入力'}</td>
    </tr>`;
  }).join('');
}

function saveStock() {
  const now = fmtDate(new Date());
  document.querySelectorAll('#stockBody tr').forEach(tr => {
    const id = tr.dataset.id;
    const item = state.items.find(i => i.id === id);
    if (!item) return;
    const cases = Number(tr.querySelector('.stock-case').value) || 0;
    const rest = Number(tr.querySelector('.stock-rest').value) || 0;
    const qty = cases * Math.max(1, item.packSize) + rest;
    const prev = stockQty(id);
    if (qty !== prev || !state.stock[id]) {
      state.stock[id] = { qty, updatedAt: now };
    }
  });
  save();
  renderStock();
  toast('在庫を保存しました。');
}

/* ---------------------------------------------------------
 * 商品マスタ
 * -------------------------------------------------------*/

function renderMaster() {
  $('masterBody').innerHTML = state.items.map(item => `
    <tr class="${item.active === false ? 'is-inactive' : ''}">
      <td>${escapeHtml(item.code)}</td>
      <td>${escapeHtml(item.name)}${item.active === false ? '<span class="tag">対象外</span>' : ''}</td>
      <td>${CATEGORY_LABEL[item.category] || '-'}</td>
      <td>${escapeHtml(item.vendor)}</td>
      <td>${escapeHtml(item.unit)}</td>
      <td class="num">${num(item.packSize)}</td>
      <td class="num">${yen(item.price)}</td>
      <td class="num">${num(item.pi, 1)}</td>
      <td class="num">${num(item.lead)}日</td>
      <td class="num">${num(item.lot)}</td>
      <td>${escapeHtml(item.storage)}</td>
      <td class="num">
        <button class="link" data-edit="${item.id}">編集</button>
        <button class="link link--danger" data-delete="${item.id}">削除</button>
      </td>
    </tr>`).join('');
  refreshVendorOptions();
}

function refreshVendorOptions() {
  const vendors = [...new Set(state.items.map(i => i.vendor).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'ja'));
  const select = $('filterVendor');
  const current = select.value;
  select.innerHTML = '<option value="">すべて</option>' +
    vendors.map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
  select.value = vendors.includes(current) ? current : '';
  $('vendorList').innerHTML = vendors.map(v => `<option value="${escapeHtml(v)}">`).join('');
}

let editingId = null;

function openItemModal(id) {
  editingId = id;
  const item = id ? state.items.find(i => i.id === id) : null;
  $('itemModalTitle').textContent = item ? '商品を編集' : '商品を追加';
  $('f_code').value = item ? item.code : '';
  $('f_name').value = item ? item.name : '';
  $('f_category').value = item ? item.category : 'food';
  $('f_vendor').value = item ? item.vendor : '';
  $('f_unit').value = item ? item.unit : 'ケース';
  $('f_packSize').value = item ? item.packSize : 1;
  $('f_price').value = item ? item.price : 0;
  $('f_pi').value = item ? item.pi : 0;
  $('f_lead').value = item ? item.lead : 3;
  $('f_lot').value = item ? item.lot : 1;
  $('f_storage').value = item ? item.storage : '常温';
  $('f_active').checked = item ? item.active !== false : true;
  $('itemModal').hidden = false;
}

function closeItemModal() {
  $('itemModal').hidden = true;
  editingId = null;
}

function saveItem() {
  const name = $('f_name').value.trim();
  if (!name) { toast('商品名を入力してください。'); return; }

  const data = {
    code: $('f_code').value.trim() || '-',
    name,
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

  if (editingId) {
    const item = state.items.find(i => i.id === editingId);
    Object.assign(item, data);
  } else {
    state.seq++;
    state.items.push(Object.assign({ id: 'item-' + state.seq + '-' + Date.now().toString(36) }, data));
  }
  save();
  closeItemModal();
  renderMaster();
  toast('商品マスタを更新しました。');
}

function deleteItem(id) {
  const item = state.items.find(i => i.id === id);
  if (!item) return;
  if (!confirm(`「${item.name}」を削除します。よろしいですか？\n（発注履歴は残ります）`)) return;
  state.items = state.items.filter(i => i.id !== id);
  delete state.stock[id];
  delete draft[id];
  save();
  renderMaster();
  toast('商品を削除しました。');
}

const MASTER_HEADER = ['商品コード', '商品名', 'カテゴリ', '仕入先', '発注単位', '入数', '単価', 'PI値', 'リードタイム', 'ロット', '保管場所', '発注対象'];

function exportMasterCsv() {
  const rows = [MASTER_HEADER];
  state.items.forEach(i => {
    rows.push([i.code, i.name, CATEGORY_LABEL[i.category] || i.category, i.vendor, i.unit,
      i.packSize, i.price, i.pi, i.lead, i.lot, i.storage, i.active === false ? '×' : '○']);
  });
  downloadCsv('商品マスタ.csv', rows);
}

function importMasterCsv(file) {
  const reader = new FileReader();
  reader.onload = () => {
    const rows = parseCsv(String(reader.result));
    if (rows.length < 2) { toast('取込できる行がありません。'); return; }
    if (!confirm('CSVの内容で商品マスタを置き換えます。よろしいですか？\n（在庫・発注履歴は保持されます）')) return;

    const items = [];
    rows.slice(1).forEach((r, idx) => {
      const name = (r[1] || '').trim();
      if (!name) return;
      const code = (r[0] || '-').trim();
      const existing = state.items.find(i => i.code === code && i.name === name);
      items.push({
        id: existing ? existing.id : 'item-import-' + idx + '-' + Date.now().toString(36),
        code,
        name,
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

    if (!items.length) { toast('取込できる商品がありませんでした。'); return; }
    state.items = items;
    save();
    renderMaster();
    toast(`${items.length}件の商品を取り込みました。`);
  };
  reader.readAsText(file, 'UTF-8');
}

/* ---------------------------------------------------------
 * 発注履歴
 * -------------------------------------------------------*/

function renderHistory() {
  const orders = state.orders.slice().sort((a, b) => b.createdAt - a.createdAt);
  // 再描画で開閉状態が戻らないようにする
  const opened = new Set([...document.querySelectorAll('.order-card[open]')].map(d => d.dataset.id));
  $('historyEmpty').hidden = orders.length > 0;
  $('historyList').innerHTML = orders.map(o => `
    <details class="order-card" data-id="${o.id}" ${opened.has(o.id) ? 'open' : ''}>
      <summary>
        <span class="badge badge--${o.status}">${STATUS_LABEL[o.status]}</span>
        <span class="order-card__date">${fmtDateJa(o.orderDate)}</span>
        <span class="order-card__vendor">${escapeHtml(o.vendor)}</span>
        <span class="order-card__meta">納品予定 ${fmtDateJa(o.deliveryDate)}／${o.lines.length}品目／${yen(orderTotal(o))}</span>
      </summary>
      <div class="table-wrap">
        <table class="table">
          <thead>
            <tr><th>コード</th><th>商品名</th><th class="num">発注数</th><th class="num">入数</th><th class="num">単価</th><th class="num">金額</th></tr>
          </thead>
          <tbody>
            ${o.lines.map(l => `
              <tr>
                <td>${escapeHtml(l.code)}</td>
                <td>${escapeHtml(l.name)}</td>
                <td class="num">${num(l.qty)}${escapeHtml(l.unit)}</td>
                <td class="num">${num(l.packSize)}</td>
                <td class="num">${yen(l.price)}</td>
                <td class="num">${yen(l.qty * l.price)}</td>
              </tr>`).join('')}
          </tbody>
        </table>
      </div>
      <div class="btn-row">
        ${o.status === 'ordered' ? `<button class="btn btn--primary" data-receive="${o.id}">入荷済にする（在庫に加算）</button>` : ''}
        ${o.status === 'ordered' ? `<button class="btn" data-cancel="${o.id}">発注を取消</button>` : ''}
        <button class="btn" data-csv="${o.id}">CSV出力</button>
        <button class="btn btn--danger" data-remove="${o.id}">履歴を削除</button>
      </div>
    </details>`).join('');
}

function receiveOrder(id) {
  const order = state.orders.find(o => o.id === id);
  if (!order || order.status !== 'ordered') return;
  if (!confirm('入荷済にして、発注数を在庫に加算します。よろしいですか？')) return;

  const now = fmtDate(new Date());
  order.lines.forEach(l => {
    const qty = stockQty(l.itemId) + l.qty * l.packSize;
    state.stock[l.itemId] = { qty, updatedAt: now };
  });
  order.status = 'received';
  order.receivedAt = now;
  save();
  renderHistory();
  toast('入荷処理を行い、在庫に反映しました。');
}

function cancelOrder(id) {
  const order = state.orders.find(o => o.id === id);
  if (!order) return;
  if (!confirm('この発注を取消状態にします。よろしいですか？')) return;
  order.status = 'cancelled';
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

function exportOrderHistoryCsv(id) {
  const o = state.orders.find(x => x.id === id);
  if (!o) return;
  const rows = [['発注日', '仕入先', '納品予定日', '状態', '商品コード', '商品名', '発注単位', '入数', '発注数', '単価', '金額']];
  o.lines.forEach(l => {
    rows.push([o.orderDate, o.vendor, o.deliveryDate, STATUS_LABEL[o.status], l.code, l.name, l.unit, l.packSize, l.qty, l.price, l.qty * l.price]);
  });
  downloadCsv(`発注書_${o.vendor}_${o.orderDate}.csv`, rows);
}

/* ---------------------------------------------------------
 * 設定
 * -------------------------------------------------------*/

function renderSettings() {
  const s = state.settings;
  $('setTheater').value = s.theater;
  $('setStaff').value = s.staff;
  $('setCycle').value = s.cycleDays;
  $('setSafety').value = s.safetyDays;
  $('attendanceGrid').innerHTML = WEEK_LABEL.map((label, i) => `
    <label class="field">
      <span>${label}曜日</span>
      <input type="number" min="0" step="10" class="attendance" data-day="${i}" value="${s.attendance[i]}">
    </label>`).join('');
}

function saveSettings() {
  const s = state.settings;
  s.theater = $('setTheater').value.trim();
  s.staff = $('setStaff').value.trim();
  s.cycleDays = Math.max(1, Number($('setCycle').value) || 7);
  s.safetyDays = Math.max(0, Number($('setSafety').value) || 0);
  document.querySelectorAll('.attendance').forEach(input => {
    s.attendance[Number(input.dataset.day)] = Math.max(0, Number(input.value) || 0);
  });
  save();
  applyHeader();
  toast('設定を保存しました。');
}

function applyHeader() {
  $('headerTheater').textContent = state.settings.theater || '劇場名未設定';
  $('seasonRate').value = state.settings.seasonRate;
}

function exportJson() {
  download(`発注ツールデータ_${fmtDate(new Date())}.json`, JSON.stringify(state, null, 2), 'application/json');
}

function importJson(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsed = JSON.parse(String(reader.result));
      if (!parsed.items) throw new Error('形式が不正です');
      if (!confirm('現在のデータをファイルの内容で置き換えます。よろしいですか？')) return;
      state = Object.assign(defaultState(), parsed);
      draft = {};
      save();
      initView();
      toast('データを復元しました。');
    } catch (e) {
      toast('ファイルを読み込めませんでした。');
    }
  };
  reader.readAsText(file, 'UTF-8');
}

function resetData() {
  if (!confirm('商品マスタ・在庫・発注履歴・設定をすべて初期状態に戻します。\nこの操作は取り消せません。よろしいですか？')) return;
  state = defaultState();
  draft = {};
  save();
  initView();
  toast('初期データに戻しました。');
}

/* ---------------------------------------------------------
 * イベント登録
 * -------------------------------------------------------*/

function bindEvents() {
  $('tabs').addEventListener('click', e => {
    const tab = e.target.closest('.tab');
    if (tab) switchTab(tab.dataset.tab);
  });

  // 発注作成
  ['orderDate', 'filterVendor', 'filterCategory', 'filterNeed'].forEach(id => {
    $(id).addEventListener('change', renderOrder);
  });
  $('seasonRate').addEventListener('change', () => {
    state.settings.seasonRate = Math.max(10, Number($('seasonRate').value) || 100);
    $('seasonRate').value = state.settings.seasonRate;
    save();
    renderOrder();
  });
  $('orderBody').addEventListener('input', e => {
    const input = e.target.closest('.qty');
    if (!input) return;
    const id = input.dataset.id;
    const qty = Math.max(0, Number(input.value) || 0);
    if (qty > 0) draft[id] = qty; else delete draft[id];
    const item = state.items.find(i => i.id === id);
    const cell = input.closest('tr').querySelector('.amount');
    cell.textContent = qty > 0 ? yen(qty * item.price) : '—';
    updateOrderTotal();
  });
  $('applyRecommend').addEventListener('click', () => {
    orderRows().forEach(r => {
      if (r.recommend > 0) draft[r.item.id] = r.recommend;
    });
    renderOrder();
    toast('推奨発注数を反映しました。');
  });
  $('clearQty').addEventListener('click', () => {
    draft = {};
    renderOrder();
  });
  $('submitOrder').addEventListener('click', submitOrder);
  $('exportOrderCsv').addEventListener('click', exportOrderCsv);
  $('printOrder').addEventListener('click', printOrder);

  // 在庫
  $('stockCategory').addEventListener('change', renderStock);
  $('stockStorage').addEventListener('change', renderStock);
  $('stockBody').addEventListener('input', e => {
    const tr = e.target.closest('tr');
    if (!tr || !e.target.matches('.stock-case, .stock-rest')) return;
    const item = state.items.find(i => i.id === tr.dataset.id);
    if (!item) return;
    const cases = Number(tr.querySelector('.stock-case').value) || 0;
    const rest = Number(tr.querySelector('.stock-rest').value) || 0;
    tr.querySelector('.total').textContent = num(cases * Math.max(1, item.packSize) + rest);
  });
  $('saveStock').addEventListener('click', saveStock);

  // マスタ
  $('addItem').addEventListener('click', () => openItemModal(null));
  $('masterBody').addEventListener('click', e => {
    const edit = e.target.closest('[data-edit]');
    const del = e.target.closest('[data-delete]');
    if (edit) openItemModal(edit.dataset.edit);
    if (del) deleteItem(del.dataset.delete);
  });
  $('saveItem').addEventListener('click', saveItem);
  document.querySelectorAll('[data-close-modal]').forEach(el => {
    el.addEventListener('click', closeItemModal);
  });
  $('exportMasterCsv').addEventListener('click', exportMasterCsv);
  $('importMasterCsv').addEventListener('change', e => {
    if (e.target.files[0]) importMasterCsv(e.target.files[0]);
    e.target.value = '';
  });

  // 履歴
  $('historyList').addEventListener('click', e => {
    const receive = e.target.closest('[data-receive]');
    const cancel = e.target.closest('[data-cancel]');
    const csv = e.target.closest('[data-csv]');
    const remove = e.target.closest('[data-remove]');
    if (receive) receiveOrder(receive.dataset.receive);
    if (cancel) cancelOrder(cancel.dataset.cancel);
    if (csv) exportOrderHistoryCsv(csv.dataset.csv);
    if (remove) removeOrder(remove.dataset.remove);
  });

  // 設定
  $('saveSettings').addEventListener('click', saveSettings);
  $('exportJson').addEventListener('click', exportJson);
  $('importJson').addEventListener('change', e => {
    if (e.target.files[0]) importJson(e.target.files[0]);
    e.target.value = '';
  });
  $('resetData').addEventListener('click', resetData);

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && !$('itemModal').hidden) closeItemModal();
  });
}

function initView() {
  $('orderDate').value = fmtDate(new Date());
  applyHeader();
  refreshVendorOptions();
  renderSettings();
  switchTab(document.querySelector('.tab.is-active').dataset.tab);
}

state = load();
bindEvents();
initView();
