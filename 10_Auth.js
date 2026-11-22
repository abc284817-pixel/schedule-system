/**
 * 10_Auth.js — 網頁進入點、登入、憑證驗證、員工主檔快取
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('餐飲門市智慧排班系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 以樣板方式載入，支援巢狀 include 與 <?!= ?> 注入（例如 js_constants 的 APP_CONFIG） */
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

/** 員工主檔快取：{ staffMap: {id:{name,is_core,role}}, allStaffList: [...排除代理帳號] } */
function getCachedStaffMap() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CONFIG.CACHE.STAFF);
  if (cached) return JSON.parse(cached);
  const C = COL.STAFF;
  const staffMap = {};
  const allStaffList = [];
  readRows_("STAFF").forEach(r => {
    if (!toBool_(r[C.is_active])) return;
    const eId = r[C.employee_id];
    const isCore = toBool_(r[C.is_core]);
    staffMap[eId] = { name: r[C.name], is_core: isCore, role: r[C.role] };
    if (!CONFIG.PROXY_ACCOUNTS[eId]) {
      allStaffList.push({ employee_id: eId, name: r[C.name], is_core: isCore, pref: "Standard" });
    }
  });
  const result = { staffMap, allStaffList };
  try { cache.put(CONFIG.CACHE.STAFF, JSON.stringify(result), CONFIG.CACHE.TTL_SEC); } catch (e) {}
  return result;
}

function apiLogin(employeeId, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffSheet = ss.getSheetByName(CONFIG.SHEETS.STAFF.name);
  if (!staffSheet) return { success: false, message: "找不到 Staff_Master 資料表" };
  const C = COL.STAFF;
  const rows = staffSheet.getDataRange().getValues().slice(1);

  for (const r of rows) {
    const empId = r[C.employee_id];
    if (String(empId).trim() !== String(employeeId).trim() || !toBool_(r[C.is_active])) continue;
    if (computeSHA256(password + r[C.password_salt]) !== String(r[C.password_hash]).trim()) continue;

    const token = Utilities.getUuid();
    const now = new Date();
    const expire = new Date(now.getTime() + CONFIG.RULES.TOKEN_TTL_HOURS * 3600 * 1000);
    const role = r[C.role];
    getOrCreateSheet_("TOKENS").appendRow([token, empId, role, now, expire]);
    try {
      CacheService.getScriptCache().put(CONFIG.CACHE.TOKEN_PREFIX + token, JSON.stringify({ employee_id: empId, role: role }), CONFIG.CACHE.TTL_SEC);
    } catch (e) {}

    const user = { employee_id: empId, name: r[C.name], role: role, is_core: toBool_(r[C.is_core]) };
    user.effective_employee_id = effectiveEmpId_(user);
    return { success: true, token: token, user: user };
  }
  return { success: false, message: "員工代碼或密碼錯誤" };
}

/** 回傳 { employee_id, role } 或 null */
function validateToken(token) {
  if (!token) return null;
  const key = CONFIG.CACHE.TOKEN_PREFIX + token;
  try {
    const cached = CacheService.getScriptCache().get(key);
    if (cached) return JSON.parse(cached);
  } catch (e) {}

  const C = COL.TOKENS;
  const now = new Date();
  const row = readRows_("TOKENS").find(r => r[C.token] === token && new Date(r[C.expires_at]) > now);
  if (!row) return null;
  const userObj = { employee_id: row[C.employee_id], role: row[C.role] };
  try { CacheService.getScriptCache().put(key, JSON.stringify(userObj), CONFIG.CACHE.TTL_SEC); } catch (e) {}
  return userObj;
}
