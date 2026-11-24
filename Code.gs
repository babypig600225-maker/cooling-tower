// ================================================
//  冷卻水塔水質管理系統 - Apps Script 後端
// ================================================
const SHEET_NAME   = '冷卻水塔水質';
const SUMMARY_NAME = '取樣摘要';   // 首次儲存摘要時自動建立

function doGet(e) {
  const callback = e.parameter.callback || 'cb';
  const action   = e.parameter.action   || 'read';

  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('找不到工作表：' + SHEET_NAME);

    // ── 新增資料 ──
    if (action === 'write') {
      const rows = JSON.parse(e.parameter.rows);
      rows.forEach(row => sheet.appendRow(row));
      return jsonp(callback, { status: 'ok', added: rows.length });
    }

    // ── 刪除指定日期資料（含摘要）──
    if (action === 'delete') {
      const date  = e.parameter.date;
      const count = deleteRowsByDate(sheet, date);
      const sum   = ss.getSheetByName(SUMMARY_NAME);
      if (sum) deleteRowsByDate(sum, date);
      return jsonp(callback, { status: 'ok', deleted: count });
    }

    // ── 更新指定日期資料（先刪再寫，摘要保留）──
    if (action === 'update') {
      const date = e.parameter.date;
      const rows = JSON.parse(e.parameter.rows);
      deleteRowsByDate(sheet, date);
      rows.forEach(row => sheet.appendRow(row));
      return jsonp(callback, { status: 'ok', updated: rows.length });
    }

    // ── 儲存摘要（同日期覆蓋；空白則刪除）──
    if (action === 'summary') {
      const date = e.parameter.date;
      const text = e.parameter.text || '';
      let sum = ss.getSheetByName(SUMMARY_NAME);
      if (!sum) {
        sum = ss.insertSheet(SUMMARY_NAME);
        sum.appendRow(['取樣日期', '摘要', '更新時間']);
        sum.getRange('A:A').setNumberFormat('@');
      }
      deleteRowsByDate(sum, date);
      if (text.trim()) sum.appendRow([date, text, new Date()]);
      return jsonp(callback, { status: 'ok', saved: true });
    }

    // ── 讀取資料（預設）──
    const values  = sheet.getDataRange().getValues();
    const headers = values[0].map(String);
    const rows    = values.slice(1)
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i]; });
        return obj;
      })
      .filter(r => r['取樣日期']);

    const summaries = {};
    const sum = ss.getSheetByName(SUMMARY_NAME);
    if (sum) {
      sum.getDataRange().getValues().slice(1).forEach(r => {
        if (r[0]) summaries[dateKey(r[0])] = String(r[1]);
      });
    }

    return jsonp(callback, { status: 'ok', data: rows, summaries });

  } catch (err) {
    return jsonp(callback, { status: 'error', message: err.toString() });
  }
}

// 日期統一為 yyyy-MM-dd（依試算表時區）
function dateKey(cell) {
  if (cell instanceof Date) {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(cell, tz, 'yyyy-MM-dd');
  }
  return String(cell).slice(0, 10);
}

// 刪除指定日期的所有列（由下往上避免索引位移）
function deleteRowsByDate(sheet, targetDate) {
  const data = sheet.getDataRange().getValues();
  let count  = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if (dateKey(data[i][0]) === targetDate) {
      sheet.deleteRow(i + 1);
      count++;
    }
  }
  return count;
}

function jsonp(callback, obj) {
  return ContentService
    .createTextOutput(`${callback}(${JSON.stringify(obj)})`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
