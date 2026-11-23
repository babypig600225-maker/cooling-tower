// ================================================
//  冷卻水塔水質管理系統 - Apps Script 後端
// ================================================
const SHEET_NAME = '冷卻水塔水質';

function doGet(e) {
  const callback = e.parameter.callback || 'cb';
  const action   = e.parameter.action   || 'read';

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
                                .getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('找不到工作表：' + SHEET_NAME);

    // ── 新增資料 ──
    if (action === 'write') {
      const rows = JSON.parse(decodeURIComponent(e.parameter.rows));
      rows.forEach(row => sheet.appendRow(row));
      return jsonp(callback, { status: 'ok', added: rows.length });
    }

    // ── 刪除指定日期資料 ──
    if (action === 'delete') {
      const date = decodeURIComponent(e.parameter.date);
      const count = deleteRowsByDate(sheet, date);
      return jsonp(callback, { status: 'ok', deleted: count });
    }

    // ── 更新指定日期資料（先刪再寫）──
    if (action === 'update') {
      const date = decodeURIComponent(e.parameter.date);
      const rows = JSON.parse(decodeURIComponent(e.parameter.rows));
      deleteRowsByDate(sheet, date);
      rows.forEach(row => sheet.appendRow(row));
      return jsonp(callback, { status: 'ok', updated: rows.length });
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

    return jsonp(callback, { status: 'ok', data: rows });

  } catch (err) {
    return jsonp(callback, { status: 'error', message: err.toString() });
  }
}

// 刪除指定日期的所有列（由下往上避免索引位移）
function deleteRowsByDate(sheet, targetDate) {
  const data = sheet.getDataRange().getValues();
  const tz   = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  let count  = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    const cell = data[i][0];
    let rowDate = '';
    if (cell instanceof Date) {
      // 使用 Utilities.formatDate 確保時區正確
      rowDate = Utilities.formatDate(cell, tz, 'yyyy-MM-dd');
    } else {
      rowDate = String(cell).slice(0, 10);
    }
    if (rowDate === targetDate) {
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
