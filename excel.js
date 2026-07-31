const INITIAL_ROWS = 4;
const INITIAL_COLS = 3;
const SHEET_NAME_MAX = 31;
const SHEET_NAME_FORBIDDEN = /[\\\/\?\*\[\]:]/g;

let tableData = [];

document.addEventListener('DOMContentLoaded', () => {
    resetTable();
    renderTable();
});

function resetTable() {
    tableData = [];
    for (let r = 0; r < INITIAL_ROWS; r++) {
        const row = [];
        for (let c = 0; c < INITIAL_COLS; c++) {
            row.push(r === 0 ? `列${c + 1}` : '');
        }
        tableData.push(row);
    }
}

function renderTable() {
    const table = document.getElementById('excelTable');
    table.innerHTML = '';

    const colCount = tableData[0] ? tableData[0].length : 0;

    // Column-delete button row
    const controlRow = document.createElement('tr');
    controlRow.className = 'control-row';
    controlRow.appendChild(document.createElement('th'));
    for (let c = 0; c < colCount; c++) {
        const th = document.createElement('th');
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'cell-delete-button';
        btn.textContent = '×';
        btn.title = 'この列を削除';
        btn.onclick = () => removeColumn(c);
        th.appendChild(btn);
        controlRow.appendChild(th);
    }
    table.appendChild(controlRow);

    // Data rows
    tableData.forEach((row, rIndex) => {
        const tr = document.createElement('tr');
        if (rIndex === 0) tr.className = 'header-row';

        const rowLabel = document.createElement('th');
        rowLabel.className = 'row-label';
        const rowDel = document.createElement('button');
        rowDel.type = 'button';
        rowDel.className = 'cell-delete-button';
        rowDel.textContent = '×';
        rowDel.title = 'この行を削除';
        rowDel.onclick = () => removeRow(rIndex);
        rowLabel.appendChild(rowDel);
        tr.appendChild(rowLabel);

        row.forEach((value, cIndex) => {
            const td = document.createElement('td');
            const input = document.createElement('input');
            input.type = 'text';
            input.value = value;
            input.className = 'cell-input';
            input.addEventListener('input', (e) => {
                tableData[rIndex][cIndex] = e.target.value;
            });
            td.appendChild(input);
            tr.appendChild(td);
        });

        table.appendChild(tr);
    });
}

function addRow() {
    const colCount = tableData[0] ? tableData[0].length : INITIAL_COLS;
    const newRow = new Array(colCount).fill('');
    tableData.push(newRow);
    renderTable();
}

function addColumn() {
    if (tableData.length === 0) {
        tableData.push(['']);
    } else {
        tableData.forEach((row, rIndex) => {
            row.push(rIndex === 0 ? `列${row.length + 1}` : '');
        });
    }
    renderTable();
}

function removeRow(index) {
    if (tableData.length <= 1) {
        setStatus('最低1行は必要です', 'error');
        return;
    }
    tableData.splice(index, 1);
    renderTable();
}

function removeColumn(index) {
    const colCount = tableData[0] ? tableData[0].length : 0;
    if (colCount <= 1) {
        setStatus('最低1列は必要です', 'error');
        return;
    }
    tableData.forEach(row => row.splice(index, 1));
    renderTable();
}

function clearTable() {
    resetTable();
    renderTable();
    setStatus('表をクリアしました', 'info');
}

function normalizeFilename(raw) {
    const trimmed = (raw || '').trim();
    if (!trimmed) return 'data.xlsx';
    return trimmed.toLowerCase().endsWith('.xlsx') ? trimmed : `${trimmed}.xlsx`;
}

function normalizeSheetName(raw) {
    const trimmed = (raw || '').trim();
    if (!trimmed) return 'Sheet1';
    const sanitized = trimmed.replace(SHEET_NAME_FORBIDDEN, '_');
    return sanitized.slice(0, SHEET_NAME_MAX);
}

function downloadExcel() {
    if (typeof XLSX === 'undefined') {
        setStatus('ライブラリの読み込みに失敗しました。ネットワーク接続を確認してください。', 'error');
        return;
    }

    if (tableData.length === 0 || tableData[0].length === 0) {
        setStatus('データが空です', 'error');
        return;
    }

    const filename = normalizeFilename(document.getElementById('excelFilename').value);
    const sheetName = normalizeSheetName(document.getElementById('excelSheetname').value);

    try {
        const ws = XLSX.utils.aoa_to_sheet(tableData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
        XLSX.writeFile(wb, filename);
        setStatus(`「${filename}」をダウンロードしました`, 'success');
    } catch (err) {
        setStatus(`生成に失敗しました: ${err.message}`, 'error');
    }
}

function setStatus(message, kind) {
    const el = document.getElementById('excelStatus');
    el.textContent = message;
    el.className = `status-message ${kind || ''}`;
}
