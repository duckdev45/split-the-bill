/**
 * 美西分帳 2026 — Google Apps Script Web App
 *
 * 部署步驟：
 *   1. 開啟目標 Google Spreadsheet → 擴充功能 → Apps Script
 *   2. 貼上此程式碼，儲存
 *   3. 部署 → 新增部署作業
 *      執行身分：我（自己的帳戶）
 *      存取權限：任何人
 *   4. 複製部署 URL → 貼到 App 設定頁的「Apps Script Web App URL」
 *
 * API：
 *   GET  ?action=pull&sheet=分帳記錄          → 取回所有費用
 *   POST body { action:"push",
 *               sheet:"分帳記錄",
 *               expenses:[...] }              → 寫入/更新費用
 *
 * Client 端 POST 需使用 Content-Type: text/plain 避免 CORS preflight。
 */

const COLS = [
  'id', 'date', 'desc', 'cat', 'paidBy', 'cur', 'amt',
  'among', 'splitType', 'custom', 'notes',
  'createdAt', 'updatedAt',
];

// ─────────────────────────────────────────────
//  ENTRY POINTS
// ─────────────────────────────────────────────

function doGet(e) {
  try {
    const p = e.parameter;
    const action = p.action || 'pull';
    if (action === 'pull') return json(pull(p.sheet));
    return json({ error: 'unknown action: ' + action });
  } catch (err) {
    return json({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'push') return json(push(body.sheet, body.expenses || []));
    return json({ error: 'unknown action: ' + body.action });
  } catch (err) {
    return json({ error: err.message });
  }
}

// ─────────────────────────────────────────────
//  PULL  GET ?action=pull
// ─────────────────────────────────────────────

function pull(sheetName) {
  const sheet = getOrCreate(sheetName);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { expenses: [] };

  const headers  = rows[0].map(String);
  const expenses = rows.slice(1).map((row, i) => {
    const obj = { gsRowId: i + 2 };
    headers.forEach((h, j) => {
      const v = row[j];
      if (h === 'among' || h === 'custom') {
        try   { obj[h] = typeof v === 'string' ? JSON.parse(v) : v; }
        catch (_) { obj[h] = h === 'among' ? [] : {}; }
      } else if (h === 'date') {
        obj[h] = v instanceof Date
          ? Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd')
          : String(v ?? '').slice(0, 10);
      } else if (h === 'amt') {
        obj[h] = Number(v) || 0;
      } else {
        obj[h] = v instanceof Date ? v.toISOString() : String(v ?? '');
      }
    });
    return obj;
  });

  return { expenses };
}

// ─────────────────────────────────────────────
//  PUSH  POST { action:"push", expenses:[...] }
// ─────────────────────────────────────────────

function push(sheetName, expenses) {
  const sheet = getOrCreate(sheetName);
  const rows  = sheet.getDataRange().getValues();

  // id → row number (1-indexed, row 1 = header)
  const idCol  = rows[0].indexOf('id');
  const idToRow = {};
  rows.slice(1).forEach((row, i) => {
    if (row[idCol]) idToRow[String(row[idCol])] = i + 2;
  });

  const now     = new Date().toISOString();
  const updated = [];

  expenses.forEach(exp => {
    const rowData = COLS.map(h => {
      if (h === 'among')     return JSON.stringify(exp.among  || []);
      if (h === 'custom')    return JSON.stringify(exp.custom || {});
      if (h === 'updatedAt') return exp[h] ? String(exp[h]) : now;
      return exp[h] !== undefined ? String(exp[h]) : '';
    });

    if (idToRow[exp.id]) {
      // ── 更新現有列 ──
      sheet.getRange(idToRow[exp.id], 1, 1, COLS.length).setValues([rowData]);
      updated.push({ id: exp.id, gsRowId: idToRow[exp.id], gsSyncedAt: now });
    } else {
      // ── 新增列 ──
      sheet.appendRow(rowData);
      const gsRowId     = sheet.getLastRow();
      idToRow[exp.id]   = gsRowId;
      updated.push({ id: exp.id, gsRowId, gsSyncedAt: now });
    }
  });

  return { ok: true, updated, syncedAt: now };
}

// ─────────────────────────────────────────────
//  UTILS
// ─────────────────────────────────────────────

function getOrCreate(name) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = name || '分帳記錄';
  let   sheet     = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(COLS);
    sheet.setFrozenRows(1);

    // header style
    const hdr = sheet.getRange(1, 1, 1, COLS.length);
    hdr.setBackground('#1a7050')
       .setFontColor('#ffffff')
       .setFontWeight('bold');

    // column widths
    const widths = {
      desc: 200, among: 180, custom: 180, notes: 200,
      createdAt: 160, updatedAt: 160,
    };
    COLS.forEach((col, i) => {
      if (widths[col]) sheet.setColumnWidth(i + 1, widths[col]);
    });
  }

  return sheet;
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
