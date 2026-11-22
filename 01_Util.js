/**
 * 01_Util.js — 共用工具（無業務邏輯）。所有領域檔只透過這些函式存取試算表、鎖與權限。
 */

// ---------- 日期 / 型別 ----------
function getTz_() {
  return Session.getScriptTimeZone();
}

function normalizeYearMonth(rawYm, tz) {
  if (!rawYm) return "";
  if (rawYm instanceof Date) return Utilities.formatDate(rawYm, tz || getTz_(), "yyyy-MM");
  const s = String(rawYm).trim();
  const m = s.match(/(\d{4})[\/-](\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}`;
  return s.substring(0, 7);
}

function normalizeDateStr(rawD, tz) {
  if (!rawD) return "";
  if (rawD instanceof Date) return Utilities.formatDate(rawD, tz || getTz_(), "yyyy-MM-dd");
  const s = String(rawD).trim();
  const m = s.match(/(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  return s.substring(0, 10);
}

function toBool_(v) {
  return v === true || String(v).toUpperCase() === "TRUE";
}

function ymParts_(ym) {
  const [y, m] = ym.split("-").map(Number);
  return { y, m, days: new Date(y, m, 0).getDate() };
}

function dateStrOf_(y, m, d) {
  return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

function getPreviousYearMonth(ymStr) {
  const { y, m } = ymParts_(ymStr);
  const d = new Date(y, m - 2, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function nowStr_(tz) {
  return Utilities.formatDate(new Date(), tz || getTz_(), "yyyy-MM-dd HH:mm:ss");
}

function getShiftDisplayTitle(shiftId) {
  const s = CONFIG.SHIFTS.find(x => x.id === shiftId);
  return s ? s.name : shiftId;
}

/** 依日期再依班別順序排序 */
function compareDateShift_(a, b) {
  const order = id => { const i = CONFIG.SHIFTS.findIndex(s => s.id === id); return i < 0 ? 99 : i; };
  return a.date.localeCompare(b.date) || (order(a.shift_id) - order(b.shift_id));
}

function computeSHA256(input) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  return raw.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

// ---------- 試算表存取 ----------
function getSheet_(key) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS[key].name);
}

/** 取得工作表，不存在則依 CONFIG 建立（含表頭與顏色） */
function getOrCreateSheet_(key) {
  const spec = CONFIG.SHEETS[key];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(spec.name);
  if (!sheet) {
    sheet = ss.insertSheet(spec.name);
    sheet.appendRow(spec.headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, spec.headers.length).setBackground(spec.color).setFontColor("#ffffff").setFontWeight("bold");
  }
  return sheet;
}

/** 讀取資料列（不含表頭）；表不存在或空表回傳 [] */
function readRows_(key) {
  const sheet = getSheet_(key);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getDataRange().getValues().slice(1);
}

/** 將資料列補齊 / 截斷成該表欄數 */
function fitRow_(row, len) {
  const r = row.slice(0, len);
  while (r.length < len) r.push("");
  return r;
}

/**
 * 原地覆蓋整張表（含表頭），並裁切尾端多餘空白列。
 */
function inPlaceReplaceSheetData(sheet, fullRows, targetColCount) {
  if (!sheet || !fullRows || fullRows.length === 0) return;
  const colCount = targetColCount || fullRows[0].length;
  const newRowCount = fullRows.length;
  const oldMaxRows = sheet.getMaxRows();
  const oldLastRow = sheet.getLastRow();
  const oldMaxCols = sheet.getMaxColumns();
  const normalizedRows = fullRows.map(row => fitRow_(Array.isArray(row) ? row : [row], colCount));

  if (colCount > oldMaxCols) sheet.insertColumnsAfter(oldMaxCols, colCount - oldMaxCols);
  if (newRowCount > oldMaxRows) sheet.insertRowsAfter(oldMaxRows, newRowCount - oldMaxRows);
  sheet.getRange(1, 1, newRowCount, colCount).setValues(normalizedRows);

  const currentMax = sheet.getMaxRows();
  if (currentMax > newRowCount + 10) {
    try {
      sheet.deleteRows(newRowCount + 6, currentMax - (newRowCount + 5));
    } catch (e) {
      if (oldLastRow > newRowCount) sheet.getRange(newRowCount + 1, 1, oldLastRow - newRowCount, colCount).clearContent();
    }
  } else if (oldLastRow > newRowCount) {
    sheet.getRange(newRowCount + 1, 1, oldLastRow - newRowCount, colCount).clearContent();
  }
  if (sheet.getFrozenRows() < 1) sheet.setFrozenRows(1);
}

/** 以 CONFIG 表頭覆寫整張表（body 不含表頭） */
function writeAll_(key, sheet, bodyRows) {
  const headers = CONFIG.SHEETS[key].headers;
  inPlaceReplaceSheetData(sheet, [headers, ...bodyRows], headers.length);
}

// ---------- 鎖 / 權限 ----------
/** 在 ScriptLock 內執行；失敗回傳 { success:false, message: errPrefix + 錯誤 } */
function withLock_(ms, errPrefix, fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(ms);
    return fn();
  } catch (e) {
    return { success: false, message: errPrefix + e.message };
  } finally {
    lock.releaseLock();
  }
}

function isAdmin_(user) {
  return Boolean(user && user.role === CONFIG.ROLE_ADMIN);
}

/** 代理帳號轉換（MGR001 → EMP001） */
function effectiveEmpId_(user) {
  return CONFIG.PROXY_ACCOUNTS[user.employee_id] || user.employee_id;
}
