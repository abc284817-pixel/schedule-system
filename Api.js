/**
 * 網頁進入點：啟用 HTML 樣板模組化包含 (include)
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('餐飲門市智慧排班系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 核心寫入引擎：原地覆蓋 + 實體幽靈列深度裁切 + 維度防禦校準 (防禦 11 欄與 9 欄不符)
 */
function inPlaceReplaceSheetData(sheet, fullRows, targetColCount) {
  if (!sheet || !fullRows || fullRows.length === 0) return;
  const colCount = targetColCount || fullRows[0].length;
  const newRowCount = fullRows.length;
  const oldMaxRows = sheet.getMaxRows();
  const oldLastRow = sheet.getLastRow();
  const oldMaxCols = sheet.getMaxColumns();

  // 1. 維度校準防禦：強制確保每一列的長度嚴格等於 colCount，避免維度不符崩潰
  const normalizedRows = fullRows.map(row => {
    let r = Array.isArray(row) ? row.slice(0, colCount) : [row];
    while (r.length < colCount) r.push("");
    return r;
  });

  // 2. 若欄位不足則補欄，若過多則安全相容
  if (colCount > oldMaxCols) {
    sheet.insertColumnsAfter(oldMaxCols, colCount - oldMaxCols);
  }

  // 3. 若列數增加則擴充列數
  if (newRowCount > oldMaxRows) {
    sheet.insertRowsAfter(oldMaxRows, newRowCount - oldMaxRows);
  }

  // 4. 原地覆蓋真實資料
  sheet.getRange(1, 1, newRowCount, colCount).setValues(normalizedRows);

  // 5. 實體裁切多餘空白行：若尾端多出超過 10 行空白，直接從試算表實體刪除以釋放儲存格配額
  const currentMax = sheet.getMaxRows();
  if (currentMax > newRowCount + 10) {
    const rowsToDelete = currentMax - (newRowCount + 5);
    try {
      sheet.deleteRows(newRowCount + 6, rowsToDelete);
    } catch (e) {
      if (oldLastRow > newRowCount) {
        sheet.getRange(newRowCount + 1, 1, oldLastRow - newRowCount, colCount).clearContent();
      }
    }
  } else if (oldLastRow > newRowCount) {
    sheet.getRange(newRowCount + 1, 1, oldLastRow - newRowCount, colCount).clearContent();
  }

  if (sheet.getFrozenRows() < 1) {
    sheet.setFrozenRows(1);
  }
}

/**
 * 員工主檔快取 (對標實體 7 欄結構)
 */
function getCachedStaffMap() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("CACHE_STAFF_MAP_REAL_7COL_FINAL");
  if (cached) return JSON.parse(cached);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffSheet = ss.getSheetByName("Staff_Master");
  const staffData = staffSheet ? staffSheet.getDataRange().getValues() : [];
  const staffMap = {};
  const allStaffList = [];
  for (let i = 1; i < staffData.length; i++) {
    const [eId, , , name, role, isCore, isActive] = staffData[i];
    if (isActive === true || String(isActive).toUpperCase() === "TRUE") {
      const isCoreBool = (isCore === true || String(isCore).toUpperCase() === "TRUE");
      staffMap[eId] = { name: name, is_core: isCoreBool, role: role };
      if (eId !== "MGR001") {
        allStaffList.push({ employee_id: eId, name: name, is_core: isCoreBool, pref: "Standard" });
      }
    }
  }
  const result = { staffMap, allStaffList };
  try { cache.put("CACHE_STAFF_MAP_REAL_7COL_FINAL", JSON.stringify(result), 21600); } catch (e) {}
  return result;
}

function apiGetHolidays(year) {
  year = Number(year) || new Date().getFullYear();
  const cache = CacheService.getScriptCache();
  const cacheKey = "HOLIDAYS_REAL_FINAL_" + year;
  const cached = cache.get(cacheKey);
  if (cached) return { success: true, holidays: JSON.parse(cached) };
  const holidays = {};
  const tz = Session.getScriptTimeZone();
  const cal = CalendarApp.getCalendarById("taiwan__zh_tw@holiday.calendar.google.com");
  if (cal) {
    const events = cal.getEvents(new Date(year, 0, 1), new Date(year, 11, 31, 23, 59, 59));
    const kw = ["元旦", "除夕", "春節", "初一", "初二", "初三", "初四", "初五", "和平紀念日", "二二八", "兒童節", "清明", "端午", "中秋", "國慶"];
    events.forEach(evt => {
      const title = evt.getTitle().trim();
      const dStr = Utilities.formatDate(evt.getStartTime(), tz, "yyyy-MM-dd");
      holidays[dStr] = { name: title, isDoublePay: kw.some(k => title.includes(k)) };
    });
  }
  holidays[`${year}-05-01`] = { name: "勞動節", isDoublePay: true };
  try { cache.put(cacheKey, JSON.stringify(holidays), 86400); } catch (e) {}
  return { success: true, holidays: holidays };
}

function apiLogin(employeeId, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffSheet = ss.getSheetByName("Staff_Master");
  if (!staffSheet) return { success: false, message: "找不到 Staff_Master 資料表" };
  const staffData = staffSheet.getDataRange().getValues();

  for (let i = 1; i < staffData.length; i++) {
    const [empId, salt, hash, name, role, isCore, isActive] = staffData[i];
    const isUserActive = (isActive === true || String(isActive).toUpperCase() === "TRUE");
    if (String(empId).trim() === String(employeeId).trim() && isUserActive) {
      if (computeSHA256(password + salt) === String(hash).trim()) {
        const token = Utilities.getUuid();
        const now = new Date();
        const expire = new Date(now.getTime() + 12 * 60 * 60 * 1000);

        let tokenSheet = ss.getSheetByName("Session_Tokens");
        if (!tokenSheet) {
          tokenSheet = ss.insertSheet("Session_Tokens");
          tokenSheet.appendRow(["token", "employee_id", "role", "last_active_at", "expires_at"]);
        }
        tokenSheet.appendRow([token, empId, role, now, expire]);

        const sessionUser = {
          employee_id: empId,
          name: name,
          role: role,
          is_core: (isCore === true || String(isCore).toUpperCase() === "TRUE")
        };

        try {
          const cache = CacheService.getScriptCache();
          cache.put("TOKEN_AUTH_" + token, JSON.stringify({ employee_id: empId, role: role }), 21600);
        } catch (e) {}

        return { success: true, token: token, user: sessionUser };
      }
    }
  }
  return { success: false, message: "員工代碼或密碼錯誤" };
}

function validateToken(token) {
  if (!token) return null;
  try {
    const cache = CacheService.getScriptCache();
    const cachedData = cache.get("TOKEN_AUTH_" + token);
    if (cachedData) return JSON.parse(cachedData);
  } catch (e) {}

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Session_Tokens");
  if (!sheet || sheet.getLastRow() <= 1) return null;
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token && new Date(data[i][4]) > now) {
      const userObj = { employee_id: data[i][1], role: data[i][2] };
      try {
        CacheService.getScriptCache().put("TOKEN_AUTH_" + token, JSON.stringify(userObj), 21600);
      } catch (err) {}
      return userObj;
    }
  }
  return null;
}

// 修正 1：月份解析器
function normalizeYearMonth(rawYm, tz) {
  if (!rawYm) return "";
  if (rawYm instanceof Date) {
    return Utilities.formatDate(rawYm, tz || Session.getScriptTimeZone(), "yyyy-MM");
  }
  const s = String(rawYm).trim();
  const match = s.match(/(\d{4})[\/-](\d{1,2})/);
  if (match) {
    // 🚨 補上 $，讓系統能正確拼出 "2026-10"
    return `${match[1]}-${match[2].padStart(2, '0')}`;
  }
  return s.substring(0, 7);
}

// 修正 2：日期解析器
function normalizeDateStr(rawD, tz) {
  if (!rawD) return "";
  if (rawD instanceof Date) {
    return Utilities.formatDate(rawD, tz || Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const s = String(rawD).trim();
  const match = s.match(/(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);
  if (match) {
    // 🚨 補上 $，讓系統能正確拼出 "2026-10-01"
    return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
  }
  return s.substring(0, 10);
}

function getShiftDisplayTitle(shiftId) {
  return shiftId === 'SHIFT_LUNCH' ? "午班" : (shiftId === 'SHIFT_DINNER' ? "晚班" : shiftId);
}

function apiCheckMonthScheduleExists(token, yearMonth) {
  const user = validateToken(token);
  if (!user || user.role !== 'Admin') return { success: false, message: "無權限！" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const schedSheet = ss.getSheetByName("Schedule_Final");
  if (!schedSheet || schedSheet.getLastRow() <= 1) {
    return { success: true, exists: false, count: 0 };
  }
  const data = schedSheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const dStr = normalizeDateStr(data[i][1], tz);
    if (dStr.startsWith(yearMonth)) count++;
  }
  return { success: true, exists: count > 0, count: count };
}

function apiGetMyRosterSchedule(token, yearMonth) {
  const user = validateToken(token);
  if (!user) return { success: false, message: "憑證無效或過期" };
  const effectiveEmpId = user.employee_id === "MGR001" ? "EMP001" : user.employee_id;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const { staffMap } = getCachedStaffMap();
  const schedSheet = ss.getSheetByName("Schedule_Final");
  const myRosterSchedule = [];
  if (schedSheet && schedSheet.getLastRow() > 1) {
    const sData = schedSheet.getDataRange().getValues();
    const shiftGroupMap = {};
    for (let i = 1; i < sData.length; i++) {
      const row = sData[i];
      const dateStr = normalizeDateStr(row[1], tz);
      if (!dateStr.startsWith(yearMonth)) continue;
      const sId = String(row[2]).trim();
      const empId = String(row[3]).trim();
      const dispatchType = row[5];
      const key = `${dateStr}___${sId}`;
      if (!shiftGroupMap[key]) shiftGroupMap[key] = [];
      shiftGroupMap[key].push({
        employee_id: empId,
        name: staffMap[empId] ? staffMap[empId].name : empId,
        is_core: (row[4] === true || String(row[4]).toUpperCase() === "TRUE"),
        dispatch_type: dispatchType
      });
    }
    for (const [key, staffInShift] of Object.entries(shiftGroupMap)) {
      const me = staffInShift.find(s => s.employee_id === effectiveEmpId && s.dispatch_type === "Assigned");
      if (me) {
        const [dateStr, sId] = key.split("___");
        const partners = staffInShift.filter(s => s.employee_id !== effectiveEmpId && s.dispatch_type === "Assigned");

        myRosterSchedule.push({
          date: dateStr,
          shift_id: sId,
          shift_name: getShiftDisplayTitle(sId),
          partners: partners
        });
      }
    }
  }
  const shiftPriority = { 'SHIFT_LUNCH': 1, 'SHIFT_DINNER': 2 };
  myRosterSchedule.sort((a, b) => {
    const dateComp = a.date.localeCompare(b.date);
    if (dateComp !== 0) return dateComp;
    return (shiftPriority[a.shift_id] || 99) - (shiftPriority[b.shift_id] || 99);
  });
  return { success: true, myRosterSchedule: myRosterSchedule };
}

function apiGetAvailabilityForMonth(token, yearMonth) {
  const user = validateToken(token);
  if (!user) return { success: false, message: "憑證無效" };
  const effectiveEmpId = user.employee_id === "MGR001" ? "EMP001" : user.employee_id;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const availSheet = ss.getSheetByName("Availability_Raw");
  const availData = availSheet && availSheet.getLastRow() > 1 ? availSheet.getDataRange().getValues() : [];

  let savedPreferences = null;
  let availabilityMap = {};

  for (let i = 1; i < availData.length; i++) {
    const row = availData[i];
    const ymStr = normalizeYearMonth(row[1], tz);
    const empId = String(row[2]).trim();
    if (ymStr === yearMonth && empId === effectiveEmpId) {
      savedPreferences = {
        work_volume: row[3] || "Standard",
        day_preference: String(row[4] || "Flexible").trim(),
        partner_id: String(row[5] || "").trim(),
        dislike_partner_id: String(row[6] || "").trim()
      };
      try {
        const rawJson = row[7];
        const parsed = typeof rawJson === "string" ? JSON.parse(rawJson) : (rawJson || {});
        for (const [dStr, shiftArr] of Object.entries(parsed)) {
          if (!availabilityMap[dStr]) availabilityMap[dStr] = { SHIFT_LUNCH: false, SHIFT_DINNER: false };
          if (Array.isArray(shiftArr)) {
            shiftArr.forEach(sId => {
              if (availabilityMap[dStr][sId] !== undefined) availabilityMap[dStr][sId] = true;
            });
          }
        }
      } catch (err) {}
      break;
    }
  }

  const unlockReg = getActiveUnlockRegistry(yearMonth);
  let isTempUnlocked = false;
  let myUnlockInfo = null;
  if (unlockReg["ALL"] || unlockReg[effectiveEmpId]) {
    isTempUnlocked = true;
    myUnlockInfo = unlockReg[effectiveEmpId] || unlockReg["ALL"];
  }

  return {
    success: true,
    savedAvailabilityMap: availabilityMap,
    savedPreferences: savedPreferences,
    isTempUnlocked: isTempUnlocked,
    myUnlockInfo: myUnlockInfo
  };
}

function apiGetMyShiftsForMonth(token, yearMonth) {
  const user = validateToken(token);
  if (!user) return { success: false, message: "憑證無效" };
  const effectiveEmpId = user.employee_id === "MGR001" ? "EMP001" : user.employee_id;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const schedSheet = ss.getSheetByName("Schedule_Final");
  const shifts = [];
  const todayStr = normalizeDateStr(new Date(), tz);
  if (!schedSheet || schedSheet.getLastRow() <= 1) {
    return { success: true, shifts: [], todayStr: todayStr };
  }
  const sData = schedSheet.getDataRange().getValues();
  const shiftRosterMap = {};
  const shiftCoreCountMap = {};
  const myCandidateRows = [];
  for (let i = 1; i < sData.length; i++) {
    const row = sData[i];
    const dateStr = normalizeDateStr(row[1], tz);
    if (!dateStr.startsWith(yearMonth)) continue;
    const sId = String(row[2]).trim();
    const empId = String(row[3]).trim();
    const isCore = (row[4] === true || String(row[4]).toUpperCase() === "TRUE");
    const dispatchType = row[5];
    if (dispatchType === "Assigned") {
      const shiftKey = `${dateStr}_${sId}`;
      if (!shiftRosterMap[shiftKey]) shiftRosterMap[shiftKey] = [];
      shiftRosterMap[shiftKey].push(empId);
      if (!shiftCoreCountMap[shiftKey]) shiftCoreCountMap[shiftKey] = 0;
      if (isCore) shiftCoreCountMap[shiftKey]++;
      if (empId === effectiveEmpId && dateStr >= todayStr) {
        myCandidateRows.push({ date: dateStr, shift_id: sId, is_core: isCore });
      }
    }
  }
  for (let i = 0; i < myCandidateRows.length; i++) {
    const item = myCandidateRows[i];
    const shiftKey = `${item.date}_${item.shift_id}`;
    const isSoleCore = item.is_core && (shiftCoreCountMap[shiftKey] <= 1);
    shifts.push({
      date: item.date,
      shift_id: item.shift_id,
      shift_name: getShiftDisplayTitle(item.shift_id),
      is_sole_core: isSoleCore,
      assigned_staff_in_shift: shiftRosterMap[shiftKey] || []
    });
  }
  const shiftPriority = { 'SHIFT_LUNCH': 1, 'SHIFT_DINNER': 2 };
  shifts.sort((a, b) => a.date.localeCompare(b.date) || ((shiftPriority[a.shift_id] || 99) - (shiftPriority[b.shift_id] || 99)));
  return { success: true, shifts: shifts, todayStr: todayStr };
}

function getActiveUnlockRegistry(yearMonth) {
  const cache = CacheService.getScriptCache();
  const raw = cache.get("UNLOCK_REG_" + yearMonth);
  if (!raw) return {};
  try {
    const reg = JSON.parse(raw);
    const nowMs = new Date().getTime();
    const validReg = {};
    for (const [k, v] of Object.entries(reg)) {
      if (v && v.expireAt > nowMs) {
        validReg[k] = v;
      }
    }
    return validReg;
  } catch (e) {
    return {};
  }
}

function apiGetBootstrapData(token, yearMonth) {
  const user = validateToken(token);
  if (!user) return { success: false, message: "憑證無效或過期" };
  const effectiveEmpId = user.employee_id === "MGR001" ? "EMP001" : user.employee_id;
  const { allStaffList } = getCachedStaffMap();
  const colleagues = allStaffList
    .filter(s => s.employee_id !== effectiveEmpId)
    .map(s => ({ employee_id: s.employee_id, name: s.name, is_core: s.is_core }));
  const availRes = apiGetAvailabilityForMonth(token, yearMonth);
  return {
    success: true,
    colleagues: colleagues,
    allStaffList: allStaffList,
    savedAvailabilityMap: availRes.savedAvailabilityMap,
    savedPreferences: availRes.savedPreferences,
    isTempUnlocked: availRes.isTempUnlocked,
    myUnlockInfo: availRes.myUnlockInfo,
    unlockRegistry: Object.values(getActiveUnlockRegistry(yearMonth))
  };
}

function apiSetUnlockFillPermission(token, targetYearMonth, durationHours, targetEmpId) {
  const user = validateToken(token);
  if (!user || user.role !== 'Admin') return { success: false, message: "權限不足：僅店長可執行此操作！" };
  const cache = CacheService.getScriptCache();
  const cacheKey = "UNLOCK_REG_" + targetYearMonth;
  const targetVal = targetEmpId ? targetEmpId : "ALL";
  const seconds = Number(durationHours) * 3600;
  const expireAtMs = new Date().getTime() + seconds * 1000;
  const { staffMap } = getCachedStaffMap();
  const targetName = targetVal === "ALL" ? "全體同仁" : (staffMap[targetVal] ? staffMap[targetVal].name : targetVal);
  const reg = getActiveUnlockRegistry(targetYearMonth);
  reg[targetVal] = {
    target: targetVal,
    targetName: targetName,
    expireAt: expireAtMs,
    yearMonth: targetYearMonth
  };
  cache.put(cacheKey, JSON.stringify(reg), 86400);
  return {
    success: true,
    unlockRegistry: Object.values(reg),
    message: `已成功為【${targetName}】解鎖 ${targetYearMonth} 意願填報權限，維持開放 ${durationHours} 小時！`
  };
}

function apiSubmitAvailability(token, payload) {
  const user = validateToken(token);
  if (!user) return { success: false, message: "憑證無效或過期" };
  const effectiveEmpId = user.employee_id === "MGR001" ? "EMP001" : user.employee_id;
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const currentY = now.getFullYear();
  const currentM = now.getMonth() + 1;
  const currentD = now.getDate();
  const currentYm = `${currentY}-${String(currentM).padStart(2, '0')}`;
  const [targetY, targetM] = payload.year_month.split("-").map(Number);
  const isCurrent = payload.year_month === currentYm;
  const isNext = (targetY > currentY) || (targetY === currentY && targetM === currentM + 1);
  const lastDay = new Date(currentY, currentM, 0).getDate();
  const isLockedTime = isCurrent || (isNext && currentD >= (lastDay - 2));

  if (isLockedTime) {
    const reg = getActiveUnlockRegistry(payload.year_month);
    let allowed = false;
    if (reg["ALL"] || reg[effectiveEmpId]) allowed = true;
    if (!allowed) {
      return { success: false, message: "安全鎖定：該月份處於唯讀狀態！若需補填請聯繫店長為您開放臨時權限。" };
    }
  }

  const matrixCompact = {};
  let totalActiveShiftsCount = 0;
  for (const [dateStr, shifts] of Object.entries(payload.availability || {})) {
    const activeShifts = [];
    if (shifts.SHIFT_LUNCH) activeShifts.push("SHIFT_LUNCH");
    if (shifts.SHIFT_DINNER) activeShifts.push("SHIFT_DINNER");
    if (activeShifts.length > 0) {
      matrixCompact[dateStr] = activeShifts;
      totalActiveShiftsCount += activeShifts.length;
    }
  }

  if (totalActiveShiftsCount === 0) {
    return {
      success: false,
      message: "⚠️ 防呆阻擋：您尚未勾選任何可出勤班次！系統禁止送出完全空白的整月意願（防止既有排班意願遭到意外清空）。若該月全月不便出勤，請聯繫店長備查。"
    };
  }

  const matrixJsonString = JSON.stringify(matrixCompact);
  if (matrixJsonString.length > 45000) {
    return {
      success: false,
      message: "⚠️ 儲存失敗：點陣資料體積超過系統安全閾值 (45,000 字元)，請重試或聯繫系統管理員！"
    };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Availability_Raw");
    const headers = [
      "record_id", "year_month", "employee_id", "pref_work_volume",
      "pref_day_preference", "pref_partner_id", "dislike_partner_id",
      "availability_matrix", "submitted_at"
    ];
    if (!sheet) {
      sheet = ss.insertSheet("Availability_Raw");
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
    }
    const data = sheet.getDataRange().getValues();
    const keepRows = [headers];

    for (let i = 1; i < data.length; i++) {
      const ymStr = normalizeYearMonth(data[i][1], tz);
      const empId = String(data[i][2]).trim();
      if (!(ymStr === payload.year_month && empId === effectiveEmpId)) {
        const r = data[i].slice(0, 9);
        while (r.length < 9) r.push("");
        keepRows.push(r);
      }
    }

    const nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    const recId = `AR_${payload.year_month.replace("-", "")}_${effectiveEmpId}`;
    keepRows.push([
      recId,
      payload.year_month,
      effectiveEmpId,
      payload.preferences.work_volume || "Standard",
      payload.preferences.day_preference || "Flexible",
      payload.preferences.partner_id || "",
      payload.preferences.dislike_partner_id || "",
      matrixJsonString,
      nowStr
    ]);

    inPlaceReplaceSheetData(sheet, keepRows, 9);
    return { success: true };
  } catch (e) {
    return { success: false, message: "伺服器忙碌，請重試：" + e.message };
  } finally {
    lock.releaseLock();
  }
}

function apiSubmitLeaveRequest(token, payload) {
  try {
    const user = validateToken(token);
    if (!user) return { success: false, message: "憑證無效或登入逾時，請重新整理登入！" };
    const effectiveEmpId = user.employee_id === "MGR001" ? "EMP001" : user.employee_id;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = Session.getScriptTimeZone();
    const schedSheet = ss.getSheetByName("Schedule_Final");
    const todayStr = normalizeDateStr(new Date(), tz);
    if (!payload || !payload.date || !payload.shift_id) {
      return { success: false, message: "請假班別參數缺失，請重新選擇！" };
    }
    if (payload.date < todayStr) {
      return { success: false, message: "法規防呆阻擋：該班次屬於過去歷史日期，不可提出事前請假！" };
    }
    let hasValidShift = false;
    if (schedSheet && schedSheet.getLastRow() > 1) {
      const sData = schedSheet.getDataRange().getValues();
      for (let i = 1; i < sData.length; i++) {
        const dStr = normalizeDateStr(sData[i][1], tz);
        const sId = String(sData[i][2]).trim();
        if (dStr === payload.date && sId === payload.shift_id && sData[i][3] === effectiveEmpId && sData[i][5] === "Assigned") {
          hasValidShift = true;
          break;
        }
      }
    }
    if (!hasValidShift) {
      return {
        success: false,
        isScheduleMutated: true,
        message: "⚠️ 班表即時異動告警：您選擇的班次在店長端已被調動或更換！系統已自動重新整理。"
      };
    }
    let sheet = ss.getSheetByName("Leave_Requests");
    if (!sheet) {
      sheet = ss.insertSheet("Leave_Requests");
      sheet.appendRow(["request_id", "employee_id", "date", "shift_id", "leave_type", "reason_detail", "has_substitute", "substitute_id", "status", "created_at"]);
      sheet.setFrozenRows(1);
    }
    const reqId = "LR_" + Utilities.getUuid().substring(0, 8);
    // 後端再防一次：沒有自覓代班時，一律不寫入代班人
    const hasSubstitute = Boolean(payload.has_substitute);
    const substituteId = hasSubstitute ? String(payload.substitute_id || "") : "";
    sheet.appendRow([
      reqId,
      effectiveEmpId,
      String(payload.date),
      String(payload.shift_id),
      String(payload.leave_type || "事假"),
      String(payload.reason_detail || ""),
      hasSubstitute,
      substituteId,
      "Pending",
      new Date()
    ]);
    return { success: true, message: "請假申請已成功送出並呈報店長審核！" };
  } catch (err) {
    return { success: false, message: "伺服器處理失敗: " + err.message };
  }
}

function apiGetMyLeaveRequests(token) {
  const user = validateToken(token);
  if (!user) return { success: false, message: "憑證無效" };
  const effectiveEmpId = user.employee_id === "MGR001" ? "EMP001" : user.employee_id;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const sheet = ss.getSheetByName("Leave_Requests");
  if (!sheet || sheet.getLastRow() <= 1) return { success: true, requests: [] };
  const { staffMap } = getCachedStaffMap();
  const data = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const [reqId, empId, rawD, shiftId, lType, rDetail, hasSub, subId, status] = data[i];
    if (empId === effectiveEmpId) {
      const dateStr = normalizeDateStr(rawD, tz);
      list.push({
        request_id: reqId, date: dateStr, shift_id: shiftId, leave_type: lType,
        reason_detail: rDetail, has_substitute: Boolean(hasSub),
        substitute_name: staffMap[subId] ? staffMap[subId].name : "無 (由店長調派)", status: status
      });
    }
  }
  return { success: true, requests: list.reverse() };
}

function getPreviousYearMonth(ymStr) {
  const [y, m] = ymStr.split("-").map(Number);
  const d = new Date(y, m - 2, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function getPrecedingMonthTailWorkDates(ss, targetYearMonth, tz) {
  const prevYm = getPreviousYearMonth(targetYearMonth);
  const schedSheet = ss.getSheetByName("Schedule_Final");
  const tailMap = {};
  if (!schedSheet || schedSheet.getLastRow() <= 1) return tailMap;
  const [y, m] = prevYm.split("-").map(Number);
  const lastDayOfPrevMonth = new Date(y, m, 0).getDate();
  const cutoffDay = Math.max(1, lastDayOfPrevMonth - 27);
  const cutoffDateStr = `${y}-${String(m).padStart(2, '0')}-${String(cutoffDay).padStart(2, '0')}`;
  const data = schedSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dStr = normalizeDateStr(row[1], tz);
    if (dStr.startsWith(prevYm) && dStr >= cutoffDateStr && row[5] === "Assigned") {
      const empId = String(row[3]).trim();
      if (!tailMap[empId]) tailMap[empId] = new Set();
      tailMap[empId].add(dStr);
    }
  }
  return tailMap;
}

/**
 * [局部高頻刷新引擎 V2] - 綁定於班表異動事件
 * 嚴格套用 9 欄雙薪邏輯，並掛載 Config_Holidays 動態節日碰撞。
 */
function syncStaffMonthlyStatsForMonth(ss, yearMonth, tz) {
  let statSheet = ss.getSheetByName("Staff_Monthly_Stats");
  // 🚨 嚴格對齊 9 欄表頭
  const headers = ["stat_id", "year_month", "employee_id", "total_shifts", "total_hours", "normal_hours", "double_pay_hours", "standby_hours", "updated_at"];
  
  if (!statSheet) {
    statSheet = ss.insertSheet("Staff_Monthly_Stats");
    statSheet.appendRow(headers);
    statSheet.setFrozenRows(1);
    statSheet.getRange(1, 1, 1, headers.length).setBackground("#27ae60").setFontColor("#ffffff").setFontWeight("bold");
  }

  const schedSheet = ss.getSheetByName("Schedule_Final");
  const { allStaffList } = getCachedStaffMap();
  
  // 1. 取得動態節日唯一真理
  const holidaysMap = apiGetDynamicHolidays();
  
  const statMap = {};
  allStaffList.forEach(s => {
    statMap[s.employee_id] = { assignedCount: 0, standbyCount: 0, normalHours: 0, doubleHours: 0, standbyHours: 0 };
  });

  // 2. 掃描當月班表，進行雙薪碰撞與水管分流
  if (schedSheet && schedSheet.getLastRow() > 1) {
    const sData = schedSheet.getDataRange().getValues();
    for (let i = 1; i < sData.length; i++) {
      const dStr = normalizeDateStr(sData[i][1], tz);
      if (!dStr.startsWith(yearMonth)) continue;
      
      const empId = String(sData[i][3]).trim();
      const dispatchType = String(sData[i][5]).trim();
      const status = String(sData[i][6]).trim();
      
      if (!statMap[empId] || status !== "Confirmed") continue;

      const hours = 4.0;
      if (dispatchType === "Assigned") {
        statMap[empId].assignedCount++;
        
        // 🚨 節日碰撞判定
        const isDouble = (holidaysMap[dStr] && holidaysMap[dStr].isDoublePay === true);
        if (isDouble) {
          statMap[empId].doubleHours += hours;
        } else {
          statMap[empId].normalHours += hours;
        }
      } else if (dispatchType === "Standby") {
        statMap[empId].standbyCount++;
        statMap[empId].standbyHours += hours;
      }
    }
  }

  // 3. 讀取統計表，剔除該月份舊資料
  const currentData = statSheet.getDataRange().getValues();
  const keepRows = [headers];
  for (let i = 1; i < currentData.length; i++) {
    const ym = normalizeYearMonth(currentData[i][1], tz);
    if (ym !== yearMonth && ym !== "") {
      // 維度校準：確保歷史資料也是 9 欄長度
      const r = currentData[i].slice(0, 9);
      while (r.length < 9) r.push("");
      keepRows.push(r);
    }
  }

  // 4. 封裝當月新資料 (嚴格防呆對齊 9 欄)
  const now = new Date();
  allStaffList.forEach(s => {
    const rec = statMap[s.employee_id] || { assignedCount: 0, standbyCount: 0, normalHours: 0, doubleHours: 0, standbyHours: 0 };
    const totalShifts = rec.assignedCount + rec.standbyCount;
    // 🚨 總工時絕不包含備援
    const totalHours = rec.normalHours + rec.doubleHours; 
    const statId = `STAT_${yearMonth.replace("-", "")}_${s.employee_id}`;
    
    // 只保留有出勤紀錄的員工
    if (totalShifts > 0) {
      keepRows.push([
        statId, yearMonth, s.employee_id, totalShifts, totalHours, rec.normalHours, rec.doubleHours, rec.standbyHours, now
      ]);
    }
  });

  // 5. 排序並寫入
  const bodyRows = keepRows.slice(1).sort((a, b) => {
    const ymComp = String(a[1]).localeCompare(String(b[1]));
    if (ymComp !== 0) return ymComp;
    return String(a[2]).localeCompare(String(b[2]));
  });
  
  const finalRows = [headers, ...bodyRows];
  inPlaceReplaceSheetData(statSheet, finalRows, headers.length);
}


/**
 * ⚡ 全面審計排班衝突與法規違規 (防禦 $ 符號遺失版 + 完整 5 大審核)
 */
function syncScheduleConflictsForMonth(ss, yearMonth, tz) {
  let confSheet = ss.getSheetByName("Schedule_Conflicts");
  if (!confSheet) {
    confSheet = ss.insertSheet("Schedule_Conflicts");
    confSheet.appendRow(["conflict_id", "year_month", "date", "shift_id", "reason", "created_at"]);
    confSheet.setFrozenRows(1);
    confSheet.getRange(1, 1, 1, 6).setBackground("#c0392b").setFontColor("#ffffff").setFontWeight("bold");
  }
  const schedSheet = ss.getSheetByName("Schedule_Final");
  if (!schedSheet || schedSheet.getLastRow() <= 1) return [];
  const { staffMap, allStaffList } = getCachedStaffMap();
  const cData = confSheet.getDataRange().getValues();
  const keepConflicts = [cData[0]];
  for (let i = 1; i < cData.length; i++) {
    const ymStr = normalizeYearMonth(cData[i][1], tz);
    if (ymStr !== yearMonth) keepConflicts.push(cData[i]);
  }
  const sData = schedSheet.getDataRange().getValues();
  const shiftMap = {};
  const empWorkDatesMap = getPrecedingMonthTailWorkDates(ss, yearMonth, tz);
  const empMonthAssignedHours = {};
  
  allStaffList.forEach(s => {
    empMonthAssignedHours[s.employee_id] = 0;
    if (!empWorkDatesMap[s.employee_id]) empWorkDatesMap[s.employee_id] = new Set();
  });
  const [year, month] = yearMonth.split("-").map(Number);
  const daysInMonth = new Date(year, month, 0).getDate();
  const shifts = ["SHIFT_LUNCH", "SHIFT_DINNER"];
  
  for (let d = 1; d <= daysInMonth; d++) {
    const dStr = year + "-" + String(month).padStart(2, '0') + "-" + String(d).padStart(2, '0');
    shifts.forEach(sId => {
      shiftMap[dStr + "___" + sId] = [];
    });
  }
  
  for (let i = 1; i < sData.length; i++) {
    const row = sData[i];
    const dStr = normalizeDateStr(row[1], tz);
    if (!dStr.startsWith(yearMonth)) continue;
    const sId = String(row[2]).trim();
    const empId = String(row[3]).trim();
    const isCore = (row[4] === true || String(row[4]).toUpperCase() === "TRUE");
    const dispatchType = String(row[5]).trim();
    const key = dStr + "___" + sId;
    if (!shiftMap[key]) shiftMap[key] = [];
    shiftMap[key].push({
      employee_id: empId,
      is_core: isCore,
      dispatch_type: dispatchType
    });
    if (dispatchType === "Assigned") {
      if (!empWorkDatesMap[empId]) empWorkDatesMap[empId] = new Set();
      empWorkDatesMap[empId].add(dStr);
      if (empMonthAssignedHours[empId] !== undefined) empMonthAssignedHours[empId] += 4.0;
      else empMonthAssignedHours[empId] = 4.0;
    }
  }
  
  let confSeq = 50000;
  const nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
  const newlyGeneratedConflicts = [];

  // ======================================================================
  // 1. 門市營運編制檢查 (⚠️ 補回遺失的缺工與核心人力偵測)
  // ======================================================================
  Object.entries(shiftMap).forEach(([key, staffList]) => {
    const [dStr, sId] = key.split("___");
    const assigned = staffList.filter(s => s.dispatch_type === "Assigned");
    const coreCount = assigned.filter(s => s.is_core).length;
    
// 檢查缺工 (少於5人)
    if (assigned.length < 5) {
      confSeq++;
      const missingCount = 5 - assigned.length;
      const reasonText = "出勤編制不足：排班人數不足 5 人 (僅 " + assigned.length + " 人，缺口 " + missingCount + " 人)";
      const rec = ["CONF_" + confSeq, yearMonth, dStr, sId, reasonText, nowStr];
      keepConflicts.push(rec);
      newlyGeneratedConflicts.push(rec);
    }
    
    // 檢查核心人力不足
    if (assigned.length > 0 && coreCount === 0) {
      confSeq++;
      const reasonText = "重大營運警示：排班後無核心人力出勤！";
      const rec = ["CONF_" + confSeq, yearMonth, dStr, sId, reasonText, nowStr];
      keepConflicts.push(rec);
      newlyGeneratedConflicts.push(rec);
    }
  });

  // ======================================================================
  // 2. 全域每日工時追蹤器 (解除月份封印，獲取跨月歷史資料)
  // ======================================================================
  const empDailyHoursMap = {};
  for (let i = 1; i < sData.length; i++) {
    const row = sData[i];
    const dStr = normalizeDateStr(row[1], tz);
    const empId = String(row[3]).trim();
    const dispatchType = String(row[5]).trim();
    if (dispatchType === "Assigned") {
      if (!empDailyHoursMap[empId]) empDailyHoursMap[empId] = {};
      if (!empDailyHoursMap[empId][dStr]) empDailyHoursMap[empId][dStr] = 0;
      empDailyHoursMap[empId][dStr] += 4.0; 
    }
  }

  // ======================================================================
  // 3. 七休一連勤違規檢查
  // ======================================================================
  Object.entries(empWorkDatesMap).forEach(([empId, dateSet]) => {
    const sortedDates = Array.from(dateSet).sort();
    let consecutiveCount = 1;
    for (let idx = 1; idx < sortedDates.length; idx++) {
      const prev = new Date(sortedDates[idx - 1]);
      const curr = new Date(sortedDates[idx]);
      const diffDays = Math.round((curr - prev) / (1000 * 3600 * 24));
      if (diffDays === 1) {
        consecutiveCount++;
        if (consecutiveCount > 6 && sortedDates[idx].startsWith(yearMonth)) {
          confSeq++;
          const empName = staffMap[empId] ? staffMap[empId].name : empId;
          const triggerDate = sortedDates[idx];
          const reasonText = "七休一違規：同仁 " + empName + "(" + empId + ") 至本日已連續出勤達 " + consecutiveCount + " 天，違反勞基法法定規定！";
          const rec = ["CONF_" + confSeq, yearMonth, triggerDate, "SHIFT_LUNCH", reasonText, nowStr];
          
          keepConflicts.push(rec);
          newlyGeneratedConflicts.push(rec);
        }
      } else {
        consecutiveCount = 1;
      }
    }
  });

  // ======================================================================
  // 4. 160小時工時超標違規 (真實 28 天滾動區間)
  // ======================================================================
  Object.entries(empDailyHoursMap).forEach(([empId, dailyHours]) => {
    const empName = staffMap[empId] ? staffMap[empId].name : empId;
    let worstViolation = null;
    
    for (let d = 1; d <= daysInMonth; d++) {
      const currDateStr = year + "-" + String(month).padStart(2, '0') + "-" + String(d).padStart(2, '0');
      
      if (!dailyHours[currDateStr]) continue;
      
      const currD = new Date(year, month - 1, d);
      const start28D = new Date(currD.getTime() - 27 * 24 * 3600 * 1000);
      const start28Str = Utilities.formatDate(start28D, tz, "yyyy-MM-dd");
      
      let windowHours = 0;
      for (let trackD = new Date(start28D); trackD <= currD; trackD.setDate(trackD.getDate() + 1)) {
         const trackStr = Utilities.formatDate(trackD, tz, "yyyy-MM-dd");
         if (dailyHours[trackStr]) windowHours += dailyHours[trackStr];
      }
      
      if (windowHours > 160.0) {
         if (!worstViolation || windowHours > worstViolation.hours) {
            worstViolation = { hours: windowHours, startStr: start28Str, endStr: currDateStr };
         }
      }
    }
    
    if (worstViolation) {
      confSeq++;
      const overHours = (worstViolation.hours - 160.0).toFixed(1);
      const reasonText = "工時超標違規：同仁 " + empName + "(" + empId + ") 於滾動區間 (" + worstViolation.startStr + " ~ " + worstViolation.endStr + ") 累計排班達 " + worstViolation.hours + " 小時（超標 " + overHours + " hr），突破法定上限！";
      const rec = ["CONF_" + confSeq, yearMonth, worstViolation.endStr, "SHIFT_LUNCH", reasonText, nowStr];
      
      keepConflicts.push(rec);
      newlyGeneratedConflicts.push(rec);
    }
  });

  // ======================================================================
  // 5. 嚴格法定四週 28 天滾動例休審計
  // ======================================================================
  Object.entries(empWorkDatesMap).forEach(([empId, dateSet]) => {
    const empName = staffMap[empId] ? staffMap[empId].name : empId;
    let worstViolation = null;
  
    for (let d = 1; d <= daysInMonth; d++) {
      const currDateStr = year + "-" + String(month).padStart(2, '0') + "-" + String(d).padStart(2, '0');
      if (!dateSet.has(currDateStr)) continue;
      
      const currD = new Date(year, month - 1, d);
      const start28D = new Date(currD.getTime() - 27 * 24 * 3600 * 1000);
      const start28Str = Utilities.formatDate(start28D, tz, "yyyy-MM-dd");
      
      let countIn28Days = 0;
      dateSet.forEach(workDate => {
        if (workDate >= start28Str && workDate <= currDateStr) countIn28Days++;
      });
  
      if (countIn28Days > 20) {
        if (!worstViolation || countIn28Days > worstViolation.count) {
          worstViolation = { count: countIn28Days, restDays: 28 - countIn28Days, startStr: start28Str, endStr: currDateStr };
        }
      }
    }
  
    if (worstViolation) {
      confSeq++;
      const reasonText = "四週例休不足：同仁 " + empName + "(" + empId + ") 於滾動區間 (" + worstViolation.startStr + " ~ " + worstViolation.endStr + ") 累計出勤達 " + worstViolation.count + " 天（例休僅 " + worstViolation.restDays + " 天），未滿法定 8 天！";
      const rec = ["CONF_" + confSeq, yearMonth, worstViolation.endStr, "SHIFT_LUNCH", reasonText, nowStr];
      
      keepConflicts.push(rec);
      newlyGeneratedConflicts.push(rec);
    }
  });

  inPlaceReplaceSheetData(confSheet, keepConflicts, 6);
  archiveScheduleConflictsAudit(ss, yearMonth, newlyGeneratedConflicts, tz);
  return newlyGeneratedConflicts;
}

function archiveScheduleConflictsAudit(ss, yearMonth, conflictRows, tz) {
  let auditSheet = ss.getSheetByName("Schedule_Conflicts_Audit");
  if (!auditSheet) {
    auditSheet = ss.insertSheet("Schedule_Conflicts_Audit");
    auditSheet.appendRow(["audit_id", "year_month", "date", "shift_id", "reason", "archived_at"]);
    auditSheet.setFrozenRows(1);
    auditSheet.getRange(1, 1, 1, 6).setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold");
  }
  const aData = auditSheet.getDataRange().getValues();
  const keepAuditRows = [aData[0]];
  for (let i = 1; i < aData.length; i++) {
    const ymStr = normalizeYearMonth(aData[i][1], tz);
    if (ymStr !== yearMonth) keepAuditRows.push(aData[i]);
  }
  const nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
  conflictRows.forEach(row => {
    keepAuditRows.push([
      "AUDIT_" + Utilities.getUuid().substring(0, 8), yearMonth, row[2], row[3], row[4], nowStr
    ]);
  });
  inPlaceReplaceSheetData(auditSheet, keepAuditRows, 6);
}

/**
 * ⚡ 智慧排班核心：產出班表後強制同步法規審計與月統計
 */
function apiRunAutoSchedule(token, yearMonth, confirmed) {
  const user = validateToken(token);
  if (!user || user.role !== 'Admin') return { success: false, message: "權限不足：僅店長可執行排班！" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const { allStaffList } = getCachedStaffMap();
    const availSheet = ss.getSheetByName("Availability_Raw");
    const availData = availSheet && availSheet.getLastRow() > 1 ? availSheet.getDataRange().getValues() : [];
    
    const staffPrefs = {};
    const availMap = {};
    let totalAvailCount = 0;

    for (let i = 1; i < availData.length; i++) {
      const row = availData[i];
      const ymStr = normalizeYearMonth(row[1], tz);
      const empId = String(row[2]).trim();

      if (ymStr === yearMonth) {
        staffPrefs[empId] = {
          volume: row[3] || "Standard",
          day_preference: String(row[4] || "Flexible").trim(),
          partner: String(row[5] || "").trim(),
          dislike: String(row[6] || "").trim()
        };

        try {
          const rawMatrix = row[7];
          const parsed = typeof rawMatrix === "string" ? JSON.parse(rawMatrix) : (rawMatrix || {});
          for (const [dStr, shiftArr] of Object.entries(parsed)) {
            if (Array.isArray(shiftArr)) {
              shiftArr.forEach(sId => {
                availMap[`${dStr}_${sId}_${empId}`] = true;
                totalAvailCount++;
              });
            }
          }
        } catch (e) {}
      }
    }

    const existingScheduleSheet = ss.getSheetByName("Schedule_Final");
    let existingScheduleCount = 0;

    if (existingScheduleSheet && existingScheduleSheet.getLastRow() > 1) {
      const existingScheduleData = existingScheduleSheet.getDataRange().getValues();
      for (let i = 1; i < existingScheduleData.length; i++) {
        const scheduleDate = normalizeDateStr(existingScheduleData[i][1], tz);
        if (scheduleDate && scheduleDate.startsWith(yearMonth)) {
          existingScheduleCount++;
        }
      }
    }

    const submittedEmployeeIds = new Set(
      Object.keys(staffPrefs).map(employeeId => String(employeeId).trim())
    );

    const missingEmployees = allStaffList.filter(staff => {
      const employeeId = String(staff.employee_id || "").trim();
      return employeeId && !submittedEmployeeIds.has(employeeId);
    });

    const warningLines = [];
    if (existingScheduleCount > 0) {
      warningLines.push(`該月份已有 ${existingScheduleCount} 筆班表資料，確認後會覆蓋並重新排班。`);
    }
    if (missingEmployees.length > 0) {
      const missingNames = missingEmployees.map(staff => staff.name || staff.employee_id).join("、");
      warningLines.push(`有 ${missingEmployees.length} 位員工尚未填報：${missingNames}`);
    }
    if (totalAvailCount === 0) {
      warningLines.push("該月份雖有填報紀錄，但目前沒有人選擇任何可上班時段；確認後仍會執行，可能無法排出班表。");
    }

    if (warningLines.length > 0 && confirmed !== true) {
      return {
        success: false,
        needsConfirmation: true,
        message: `即將為 ${yearMonth} 執行智慧排班：\n\n` + warningLines.join("\n\n") + "\n\n是否確認繼續？"
      };
    }
    
    const [year, month] = yearMonth.split("-").map(Number);
    const daysInMonth = new Date(year, month, 0).getDate();
    const shifts = ["SHIFT_LUNCH", "SHIFT_DINNER"];
    const nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    const precedingTailWorkDates = getPrecedingMonthTailWorkDates(ss, yearMonth, tz);
    const MAX_HOURS_CEILING = 160.0;
    const MAX_DAYS_CEILING = 20;

    const TARGET_HOURS_GOAL = {
      "More": 152.0,
      "Standard": 132.0,
      "Less": 88.0
    };

    function solveSchedule(isRelaxedMode) {
      const scheduleResult = [];
      let seq = 60000;
      const empWorkDates = {};
      const empAssignedHours = {};
      
      allStaffList.forEach(s => {
        empWorkDates[s.employee_id] = new Set(precedingTailWorkDates[s.employee_id] || []);
        empAssignedHours[s.employee_id] = 0;
      });

// ⚡ 架構升級：倒推掃描法 (徹底消滅歷史幽靈阻斷漏洞)
      function wouldViolateMaxConsecutiveDays(empId, targetDateStr) {
        const tempDates = new Set(empWorkDates[empId]);
        tempDates.add(targetDateStr);
        const sorted = Array.from(tempDates).sort();
        
        let run = 1;
        // 從最新的日期往前倒推，只要斷掉就安全，不必理會幾個月前的歷史連勤
        for (let i = sorted.length - 1; i > 0; i--) {
          const d1 = new Date(sorted[i - 1]);
          const d2 = new Date(sorted[i]);
          const diff = Math.round((d2 - d1) / (1000 * 3600 * 24));
          
          if (diff === 1) {
            run++;
            if (run > 6) return true;
          } else {
            // 只要兩天之間沒有相連，代表目前的連勤並未觸發七休一，直接安全放行
            break;
          }
        }
        return false;
      }

      function isLegallyCapped(empId, targetDateStr) {
        if (empAssignedHours[empId] + 4.0 > MAX_HOURS_CEILING) return true;
        const hasWorkedToday = empWorkDates[empId].has(targetDateStr);
        if (!hasWorkedToday) {
          const targetD = new Date(targetDateStr);
          const start28D = new Date(targetD.getTime() - 27 * 24 * 3600 * 1000);
          const start28Str = Utilities.formatDate(start28D, tz, "yyyy-MM-dd");
          let countIn28Days = 0;
          empWorkDates[empId].forEach(d => {
            if (d >= start28Str && d <= targetDateStr) countIn28Days++;
          });
          if (countIn28Days + 1 > 20) return true;
        }
        return false;
      }

      let unsatisfiedCount = 0;

      for (let d = 1; d <= daysInMonth; d++) {
        const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
        const dayAssignedInShifts = { SHIFT_LUNCH: [], SHIFT_DINNER: [] };
        const progressRatio = d / daysInMonth;
        const remainingDaysInMonth = daysInMonth - d + 1;

        shifts.forEach(shiftId => {
          const otherShiftId = shiftId === "SHIFT_LUNCH" ? "SHIFT_DINNER" : "SHIFT_LUNCH";
          const rawCandidates = allStaffList.filter(s => availMap[`${dateStr}_${shiftId}_${s.employee_id}`]);
          
          const legalCandidates = rawCandidates.filter(c =>
            !wouldViolateMaxConsecutiveDays(c.employee_id, dateStr) &&
            !isLegallyCapped(c.employee_id, dateStr)
          );

          legalCandidates.forEach(c => {
            const p = staffPrefs[c.employee_id] || { volume: "Standard", partner: "", dislike: "", day_preference: "Flexible" };
            let score = 100;
            const currentHours = empAssignedHours[c.employee_id];
            const monthlyGoal = TARGET_HOURS_GOAL[p.volume] || 132.0;

            if (!isRelaxedMode) {
              const expectedHoursAtDay = monthlyGoal * progressRatio;
              const paceDelta = currentHours - expectedHoursAtDay;

              if (paceDelta > 0) {
                score -= (paceDelta * 3.5);
              } else {
                score += (Math.abs(paceDelta) * 2.0);
              }

              let workedDaysCount = 0;
              empWorkDates[c.employee_id].forEach(dt => { if (dt.startsWith(yearMonth)) workedDaysCount++; });
              const remainingDaysBudget = MAX_DAYS_CEILING - workedDaysCount;
              const remainingHoursBudget = MAX_HOURS_CEILING - currentHours;

              if (remainingDaysBudget <= 4 && remainingDaysInMonth > 8) score -= 35;
              if (remainingHoursBudget <= 24.0 && remainingDaysInMonth > 8) score -= 40;
              if (p.volume === "More") score += 25;
              if (p.volume === "Less") score -= 25;

              const alreadyAssignedOtherShift = dayAssignedInShifts[otherShiftId].some(a => a.employee_id === c.employee_id);
              if (p.day_preference === "Both_Preferred") {
                score += alreadyAssignedOtherShift ? 50 : 15;
              } else if (p.day_preference === "Single_Preferred") {
                score += alreadyAssignedOtherShift ? -9999 : 25;
              }

              score -= (currentHours * 0.8);
            } else {
              const alreadyAssignedOtherShift = dayAssignedInShifts[otherShiftId].some(a => a.employee_id === c.employee_id);
              if (p.day_preference === "Single_Preferred" && alreadyAssignedOtherShift) {
                score -= 30;
              }
              score -= (currentHours * 0.3);
            }

            c._tempScore = score;
          });

          const assignedInShift = [];
          const coreCandidates = legalCandidates
            .filter(c => c.is_core)
            .sort((a, b) => b._tempScore - a._tempScore);
          
          if (coreCandidates.length > 0) {
            assignedInShift.push({ ...coreCandidates[0], isCoreWork: true });
          }

          const otherCandidates = legalCandidates
            .filter(c => !assignedInShift.some(a => a.employee_id === c.employee_id))
            .sort((a, b) => {
              const aPref = staffPrefs[a.employee_id] || {};
              const bPref = staffPrefs[b.employee_id] || {};
              const aHasPref = assignedInShift.some(x => x.employee_id === aPref.partner);
              const bHasPref = assignedInShift.some(x => x.employee_id === bPref.partner);
              if (aHasPref && !bHasPref) return -1;
              if (!aHasPref && bHasPref) return 1;
              return b._tempScore - a._tempScore;
            });

          for (let cand of otherCandidates) {
            if (assignedInShift.length >= 5) break;
            const p = staffPrefs[cand.employee_id] || { partner: "", dislike: "" };
            
            const conflictWithDislike = assignedInShift.some(a => {
              const aPref = staffPrefs[a.employee_id] || {};
              return (a.employee_id === p.dislike) || (aPref.dislike === cand.employee_id);
            });

            if (!conflictWithDislike || (isRelaxedMode && otherCandidates.length <= 5)) {
              assignedInShift.push({ ...cand, isCoreWork: false });
            }
          }

          dayAssignedInShifts[shiftId] = assignedInShift;
          assignedInShift.forEach(st => {
            seq++;
            empWorkDates[st.employee_id].add(dateStr);
            empAssignedHours[st.employee_id] += 4.0;
            scheduleResult.push([
              "SCH_" + seq, dateStr, shiftId, st.employee_id, st.isCoreWork, "Assigned", "Confirmed", nowStr
            ]);
          });

          const remaining = legalCandidates.filter(c => !assignedInShift.some(a => a.employee_id === c.employee_id));
          if (remaining.length > 0) {
            seq++;
            scheduleResult.push([
              "SCH_" + seq, dateStr, shiftId, remaining[0].employee_id, false, "Standby", "Confirmed", nowStr
            ]);
          }

          const hasCore = assignedInShift.some(s => s.isCoreWork);
          if (assignedInShift.length < 5 || !hasCore) {
            unsatisfiedCount++;
          }
        });
      }
      return { scheduleResult, unsatisfiedCount };
    }

    let executionResult = solveSchedule(false);
    if (executionResult.unsatisfiedCount > 0) {
      const relaxedResult = solveSchedule(true);
      if (relaxedResult.unsatisfiedCount <= executionResult.unsatisfiedCount) {
        executionResult = relaxedResult;
      }
    }

    let schedSheet = ss.getSheetByName("Schedule_Final");
    if (!schedSheet) {
      schedSheet = ss.insertSheet("Schedule_Final");
      schedSheet.appendRow(["schedule_id", "date", "shift_id", "employee_id", "is_core_at_work", "dispatch_type", "schedule_status", "created_at"]);
      schedSheet.setFrozenRows(1);
    }
    const sData = schedSheet.getDataRange().getValues();
    const keepScheduleRows = [sData[0]];
    for (let i = 1; i < sData.length; i++) {
      const dStr = normalizeDateStr(sData[i][1], tz);
      if (!dStr.startsWith(yearMonth)) keepScheduleRows.push(sData[i]);
    }
    executionResult.scheduleResult.forEach(r => keepScheduleRows.push(r));

    inPlaceReplaceSheetData(schedSheet, keepScheduleRows, keepScheduleRows[0].length);
    
    // ⚡ 核心補齊：自動排班完畢後立刻強制執行全面法規審計與月統計
    const newConflicts = syncScheduleConflictsForMonth(ss, yearMonth, tz);
    syncStaffMonthlyStatsForMonth(ss, yearMonth, tz);
    if (typeof refreshAllViewers === 'function') {
      refreshAllViewers(yearMonth);
    }
    
    const totalIssueCount = newConflicts.length;
    let feedbackMsg = `🤖 智慧排班已完成！已啟動全月工時平滑演算法，均勻分發 ${yearMonth} 正式班表，並已同步更新月工時統計。`;
    if (totalIssueCount > 0) {
      feedbackMsg += `\n\n⚠️ 實體班表異常告警（共 ${totalIssueCount} 處）：\n門市因出勤極限或法規上限仍有班次缺口，已詳實記錄於後台審計庫中，請主管審閱手動補足！`;
    }
    return { success: true, message: feedbackMsg, conflictCount: totalIssueCount };
  } catch (e) {
    return { success: false, message: "自動排班執行失敗: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

function apiGetScheduleTable(token, yearMonth) {
  const user = validateToken(token);
  if (!user || user.role !== 'Admin') return { success: false, message: "無權限！" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const { staffMap, allStaffList } = getCachedStaffMap();
  const schedSheet = ss.getSheetByName("Schedule_Final");
  const schedList = [];
  const assignedSet = new Set();
  const statMap = {};
  allStaffList.forEach(s => {
    statMap[s.employee_id] = { employee_id: s.employee_id, name: s.name, is_core: s.is_core, assignedHours: 0 };
  });
  if (schedSheet && schedSheet.getLastRow() > 1) {
    const schedData = schedSheet.getDataRange().getValues();
    for (let i = 1; i < schedData.length; i++) {
      const row = schedData[i];
      const dateStr = normalizeDateStr(row[1], tz);
      if (!dateStr.startsWith(yearMonth)) continue;
      const sId = String(row[2]).trim();
      const empId = String(row[3]).trim();
      const dispatchType = row[5];
      if (dispatchType === "Assigned") {
        assignedSet.add(`${dateStr}_${sId}_${empId}`);
        if (statMap[empId]) statMap[empId].assignedHours += 4.0;
      }
      schedList.push({
        schedule_id: row[0],
        date: dateStr,
        shift_id: sId,
        employee_id: empId,
        employee_name: staffMap[empId] ? staffMap[empId].name : empId,
        is_core_at_work: (row[4] === true || String(row[4]).toUpperCase() === "TRUE"),
        dispatch_type: dispatchType,
        status: row[6]
      });
    }
  }
  const confSheet = ss.getSheetByName("Schedule_Conflicts");
  const conflicts = [];
  if (confSheet && confSheet.getLastRow() > 1) {
    const confData = confSheet.getDataRange().getValues();
    for (let i = 1; i < confData.length; i++) {
      const ymStr = normalizeYearMonth(confData[i][1], tz);
      if (ymStr === yearMonth) {
        conflicts.push({
          conflict_id: confData[i][0],
          date: normalizeDateStr(confData[i][2], tz),
          shift_id: confData[i][3],
          reason: confData[i][4]
        });
      }
    }
  }
  const leaveSheet = ss.getSheetByName("Leave_Requests");
  const pendingLeaves = [];
  const leaveAuditHistory = [];
  if (leaveSheet && leaveSheet.getLastRow() > 1) {
    const leaveData = leaveSheet.getDataRange().getValues();
    for (let i = 1; i < leaveData.length; i++) {
      const [reqId, empId, rawD, shiftId, lType, rDetail, hasSub, subId, status] = leaveData[i];
      const dateStr = normalizeDateStr(rawD, tz);
      
      if (dateStr.startsWith(yearMonth)) {
        if (status === "Pending") {
          pendingLeaves.push({
            request_id: reqId,
            employee_id: empId,
            employee_name: staffMap[empId] ? staffMap[empId].name : empId,
            date: dateStr,
            shift_id: shiftId,
            leave_type: lType,
            reason_detail: rDetail,
            has_substitute: Boolean(hasSub),
            substitute_id: subId,
            substitute_name: staffMap[subId] ? staffMap[subId].name : "",
            recommended_candidates: []
          });
        } else {
          leaveAuditHistory.push({
            request_id: reqId,
            employee_name: staffMap[empId] ? staffMap[empId].name : empId,
            date: dateStr,
            shift_id: shiftId,
            leave_type: lType,
            reason_detail: rDetail,
            substitute_name: staffMap[subId] ? staffMap[subId].name : (hasSub ? "指定同仁" : "無指派"),
            status: status
          });
        }
      }
    }
  }
  const monthlyStats = Object.values(statMap).map(st => ({
    employee_id: st.employee_id,
    name: st.name,
    is_core: st.is_core,
    assigned_hours: st.assignedHours
  }));
  return {
    success: true,
    schedules: schedList,
    conflicts: conflicts,
    staffList: allStaffList,
    pendingLeaves: pendingLeaves,
    leaveAuditHistory: leaveAuditHistory.reverse(),
    monthlyStats: monthlyStats
  };
}

/**
 * ⚡ 批次儲存調度：徹底去除「新增缺口通知」，僅保留「班表已更新」之單一提示
 */
function apiBatchCommitScheduleModifications(token, yearMonth, updates, additions, deletions) {
  const user = validateToken(token);
  if (!user || user.role !== 'Admin') return { success: false, message: "權限不足：僅主管可儲存調度！" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const schedSheet = ss.getSheetByName("Schedule_Final");
  const { staffMap } = getCachedStaffMap();
  if (!schedSheet || schedSheet.getLastRow() <= 1) {
    return { success: false, message: "找不到正式班表，無法套用修改！" };
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sData = schedSheet.getDataRange().getValues();
    const headers = sData[0];
    const keepRows = [];
    const deleteIdSet = new Set(deletions || []);
    const updateMap = {};
    (updates || []).forEach(u => updateMap[u.schedule_id] = u);
    for (let i = 1; i < sData.length; i++) {
      const row = sData[i];
      const schedId = row[0];
      if (deleteIdSet.has(schedId)) continue;
      if (updateMap[schedId]) {
        const u = updateMap[schedId];
        const isCore = staffMap[u.employee_id] ? staffMap[u.employee_id].is_core : false;
        keepRows.push([
          row[0], row[1], row[2],
          u.employee_id, isCore, u.dispatch_type, row[6], new Date()
        ]);
      } else {
        keepRows.push(row);
      }
    }
    const nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    (additions || []).forEach(a => {
      const isCore = staffMap[a.employee_id] ? staffMap[a.employee_id].is_core : false;
      const newId = "SCH_" + Utilities.getUuid().substring(0, 8);
      keepRows.push([
        newId, a.date, a.shift_id, a.employee_id, isCore, a.dispatch_type || "Assigned", "Confirmed", nowStr
      ]);
    });
    const fullRows = [headers, ...keepRows];

    inPlaceReplaceSheetData(schedSheet, fullRows, fullRows[0].length);
    
    // 重新審計衝突庫與月統計
    const allCurrentConflicts = syncScheduleConflictsForMonth(ss, yearMonth, tz);
    syncStaffMonthlyStatsForMonth(ss, yearMonth, tz);
    if (typeof refreshAllViewers === 'function') {
      refreshAllViewers(yearMonth);
    }

    const formattedAllConflicts = allCurrentConflicts.map(c => ({
      conflict_id: c[0], date: c[2], shift_id: c[3], reason: c[4]
    }));

    // ⚡ 徹底去除缺口拼接警報，僅保留純粹的儲存成功反饋
    return {
      success: true,
      message: "🎉 班表已更新！",
      conflicts: formattedAllConflicts
    };
  } catch (err) {
    return { success: false, message: "批次儲存失敗: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

function apiUpdateScheduleItem(token, scheduleId, newEmployeeId, newDispatchType) {
  const user = validateToken(token);
  if (!user || user.role !== 'Admin') return { success: false, message: "權限不足" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const schedSheet = ss.getSheetByName("Schedule_Final");
  const { staffMap } = getCachedStaffMap();
  if (!schedSheet || schedSheet.getLastRow() <= 1) return { success: false, message: "找不到班表" };
  const data = schedSheet.getDataRange().getValues();
  let targetRow = -1;
  let targetDate = "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === scheduleId) {
      targetRow = i + 1;
      targetDate = normalizeDateStr(data[i][1], tz);
      break;
    }
  }
  if (targetRow === -1) return { success: false, message: "找不到指定的排班代碼" };
  if (!newEmployeeId) {
    schedSheet.deleteRow(targetRow);
  } else {
    const isCore = staffMap[newEmployeeId] ? staffMap[newEmployeeId].is_core : false;
    schedSheet.getRange(targetRow, 4).setValue(newEmployeeId);
    schedSheet.getRange(targetRow, 5).setValue(isCore);
    schedSheet.getRange(targetRow, 6).setValue(newDispatchType);
  }
  if (targetDate) {
    const ym = targetDate.substring(0, 7);
    syncScheduleConflictsForMonth(ss, ym, tz);
    syncStaffMonthlyStatsForMonth(ss, ym, tz);
  }
  return { success: true, message: "班表調度已成功更新！" };
}

function apiApproveLeaveRequest(token, requestId, chosenSubstituteId) {
  const user = validateToken(token);
  if (!user || user.role !== 'Admin') return { success: false, message: "權限不足" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const leaveSheet = ss.getSheetByName("Leave_Requests");
  const schedSheet = ss.getSheetByName("Schedule_Final");
  const { staffMap } = getCachedStaffMap();
  if (!leaveSheet || !schedSheet) return { success: false, message: "缺少必要表格" };
  const leaveData = leaveSheet.getDataRange().getValues();
  let targetLeaveRow = -1;
  let leaveRecord = null;
  for (let i = 1; i < leaveData.length; i++) {
    if (leaveData[i][0] === requestId) {
      targetLeaveRow = i + 1;
      const [, empId, rawD, shiftId, , , hasSub, subId] = leaveData[i];
      const dateStr = normalizeDateStr(rawD, tz);
      leaveRecord = { employee_id: empId, date: dateStr, shift_id: shiftId, final_substitute: hasSub ? subId : chosenSubstituteId };
      break;
    }
  }
  if (!leaveRecord) return { success: false, message: "找不到該申請" };
  const isCore = staffMap[leaveRecord.final_substitute] ? staffMap[leaveRecord.final_substitute].is_core : false;
  const schedData = schedSheet.getDataRange().getValues();
  let replaced = false;
  for (let i = 1; i < schedData.length; i++) {
    const [, rawD, shiftId, empId] = schedData[i];
    const dateStr = normalizeDateStr(rawD, tz);
    const sId = String(shiftId).trim();
    if (dateStr === leaveRecord.date && sId === leaveRecord.shift_id && empId === leaveRecord.employee_id) {
      if (leaveRecord.final_substitute) {
        schedSheet.getRange(i + 1, 4).setValue(leaveRecord.final_substitute);
        schedSheet.getRange(i + 1, 5).setValue(isCore);
      } else {
        schedSheet.deleteRow(i + 1);
      }
      replaced = true;
      break;
    }
  }
  if (leaveRecord.final_substitute) leaveSheet.getRange(targetLeaveRow, 8).setValue(leaveRecord.final_substitute);
  leaveSheet.getRange(targetLeaveRow, 9).setValue("Approved");
  const targetYm = leaveRecord.date.substring(0, 7);
  syncScheduleConflictsForMonth(ss, targetYm, tz);
  syncStaffMonthlyStatsForMonth(ss, targetYm, tz);
  return { success: true, message: replaced ? "請假已核准，代班人員已替換，月統計與衝突資料庫已更新！" : "已核准請假（抽單），工時統計已同步校正。" };
}

function apiRejectLeaveRequest(token, requestId) {
  const user = validateToken(token);
  if (!user || user.role !== 'Admin') return { success: false, message: "權限不足" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leaveSheet = ss.getSheetByName("Leave_Requests");
  const data = leaveSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === requestId) {
      leaveSheet.getRange(i + 1, 9).setValue("Rejected");
      return { success: true, message: "已駁回請假申請！" };
    }
  }
  return { success: false, message: "找不到該申請" };
}

function apiGetManagerSummary(token, yearMonth) {
  const user = validateToken(token);
  if (!user || user.role !== 'Admin') return { success: false, message: "無權限！" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const { staffMap, allStaffList } = getCachedStaffMap();
  const availSheet = ss.getSheetByName("Availability_Raw");
  const availData = availSheet && availSheet.getLastRow() > 1 ? availSheet.getDataRange().getValues() : [];
  const summaryMap = {};

  allStaffList.forEach(s => {
    summaryMap[s.employee_id] = {
      employee_id: s.employee_id,
      name: s.name,
      is_core: s.is_core,
      shift_count: 0,
      pref_work_volume: 'Standard',
      pref_partner_name: '',
      dislike_partner_name: '',
      day_preference: 'Flexible',
      submitted: false
    };
  });

  for (let i = 1; i < availData.length; i++) {
    const row = availData[i];
    const ymStr = normalizeYearMonth(row[1], tz);
    const empId = String(row[2]).trim();

    if (ymStr === yearMonth && summaryMap[empId]) {
      summaryMap[empId].submitted = true;
      summaryMap[empId].pref_work_volume = row[3] || 'Standard';
      summaryMap[empId].day_preference = String(row[4] || "Flexible").trim();
      let pId = String(row[5] || "").trim();
      let dpId = String(row[6] || "").trim();
      summaryMap[empId].pref_partner_name = staffMap[pId] ? staffMap[pId].name : pId;
      summaryMap[empId].dislike_partner_name = staffMap[dpId] ? staffMap[dpId].name : dpId;

      try {
        const rawJson = row[7];
        const parsed = typeof rawJson === "string" ? JSON.parse(rawJson) : (rawJson || {});
        let totalCount = 0;
        for (const shiftArr of Object.values(parsed)) {
          if (Array.isArray(shiftArr)) totalCount += shiftArr.length;
        }
        summaryMap[empId].shift_count = totalCount;
      } catch (err) {}
    }
  }

  return { success: true, summary: Object.values(summaryMap) };
}

function apiArchiveHistoricalData(token, cutoffYearMonth) {
  const user = validateToken(token);
  if (!user || user.role !== 'Admin') return { success: false, message: "權限不足！" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const availActive = ss.getSheetByName("Availability_Raw");
  let availArchive = ss.getSheetByName("Availability_Archive");
  const schedActive = ss.getSheetByName("Schedule_Final");
  let schedArchive = ss.getSheetByName("Schedule_Final_Archive");

  const standardAvailHeaders = [
    "record_id", "year_month", "employee_id", "pref_work_volume",
    "pref_day_preference", "pref_partner_id", "dislike_partner_id",
    "availability_matrix", "submitted_at"
  ];

  const standardArcAvailHeaders = [
    "record_id", "year_month", "employee_id", "pref_work_volume",
    "pref_day_preference", "pref_partner_id", "dislike_partner_id",
    "availability_matrix", "submitted_at", "archived_at"
  ];

  if (!availArchive) {
    availArchive = ss.insertSheet("Availability_Archive");
    availArchive.appendRow(standardArcAvailHeaders);
    availArchive.setFrozenRows(1);
  }

  const standardSchedHeaders = [
    "schedule_id", "date", "shift_id", "employee_id",
    "is_core_at_work", "dispatch_type", "schedule_status", "created_at"
  ];

  const standardArcSchedHeaders = [
    "schedule_id", "date", "shift_id", "employee_id",
    "is_core_at_work", "dispatch_type", "schedule_status", "created_at", "archived_at"
  ];

  if (!schedArchive) {
    schedArchive = ss.insertSheet("Schedule_Final_Archive");
    schedArchive.appendRow(standardArcSchedHeaders);
    schedArchive.setFrozenRows(1);
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);

    const aData = availActive ? availActive.getDataRange().getValues() : [];
    const aKeepActiveRows = [standardAvailHeaders];
    const aMovingRows = [];

    for (let i = 1; i < aData.length; i++) {
      const ymStr = normalizeYearMonth(aData[i][1], tz);
      const r = aData[i].slice(0, 9);
      while (r.length < 9) r.push("");

      if (ymStr < cutoffYearMonth) {
        aMovingRows.push([...r, now]);
      } else {
        aKeepActiveRows.push(r);
      }
    }

    if (aMovingRows.length > 0) {
      const oldArcData = availArchive.getDataRange().getValues();
      const movingKeySet = new Set(aMovingRows.map(r => `${normalizeYearMonth(r[1], tz)}___${String(r[2]).trim()}`));
      const finalArcRows = [standardArcAvailHeaders];

      for (let i = 1; i < oldArcData.length; i++) {
        const ym = normalizeYearMonth(oldArcData[i][1], tz);
        const emp = String(oldArcData[i][2]).trim();
        const key = `${ym}___${emp}`;
        if (!movingKeySet.has(key)) {
          const arcRow = oldArcData[i].slice(0, 10);
          while (arcRow.length < 10) arcRow.push("");
          finalArcRows.push(arcRow);
        }
      }
      aMovingRows.forEach(r => finalArcRows.push(r));
      inPlaceReplaceSheetData(availArchive, finalArcRows, 10);
    }
    inPlaceReplaceSheetData(availActive, aKeepActiveRows, 9);

    const sData = schedActive ? schedActive.getDataRange().getValues() : [];
    const sKeepActiveRows = [standardSchedHeaders];
    const sMovingRows = [];

    for (let i = 1; i < sData.length; i++) {
      const dStr = normalizeDateStr(sData[i][1], tz);
      const r = sData[i].slice(0, 8);
      while (r.length < 8) r.push("");

      if (dStr.substring(0, 7) < cutoffYearMonth) {
        sMovingRows.push([...r, now]);
      } else {
        sKeepActiveRows.push(r);
      }
    }

    if (sMovingRows.length > 0) {
      const oldSchedArcData = schedArchive.getDataRange().getValues();
      const movingSchedIdSet = new Set(sMovingRows.map(r => String(r[0]).trim()));
      const finalSchedArcRows = [standardArcSchedHeaders];

      for (let i = 1; i < oldSchedArcData.length; i++) {
        const sId = String(oldSchedArcData[i][0]).trim();
        if (!movingSchedIdSet.has(sId)) {
          const arcSchedRow = oldSchedArcData[i].slice(0, 9);
          while (arcSchedRow.length < 9) arcSchedRow.push("");
          finalSchedArcRows.push(arcSchedRow);
        }
      }
      sMovingRows.forEach(r => finalSchedArcRows.push(r));
      inPlaceReplaceSheetData(schedArchive, finalSchedArcRows, 9);
    }
    inPlaceReplaceSheetData(schedActive, sKeepActiveRows, 8);

    return {
      success: true,
      message: `歷史封存成功！已將 ${cutoffYearMonth} 以前資料安全歸檔至冷儲存區。`
    };
  } catch (e) {
    return { success: false, message: "封存失敗：" + e.message };
  } finally {
    lock.releaseLock();
  }
}

function computeSHA256(input) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  return rawHash.map(byte => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0')).join('');
}

/**
 * [API] 獲取動態節日常數
 * 從 Config_Holidays 工作表讀取資料，並轉化為前端所需的 O(1) 字典格式
 * 回傳格式範例： { "2026-09-28": { name: "教師節", isDoublePay: true, type: "法定國定假日" } }
 */
function apiGetDynamicHolidays() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config_Holidays');
  
  if (!sheet) {
    throw new Error("系統致命錯誤：找不到 Config_Holidays 工作表，請確認工作表名稱是否正確。");
  }
  
  // 抓取包含標題的所有資料
  const data = sheet.getDataRange().getValues();
  const holidaysMap = {};
  
  // 從 i = 1 開始迴圈 (跳過第一列的標題 Row)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateRaw = row[0];        // A欄: Date
    const name = row[1];           // B欄: Holiday_Name
    const isDoublePayStr = row[2]; // C欄: Is_Double_Pay
    const type = row[3];           // D欄: Holiday_Type
    
    // 防呆：如果日期或名稱是空的，直接跳過這行
    if (!dateRaw || !name) continue; 
    
    let dateStr = "";
    // 極限防呆：防止 Google Sheets 偷偷把純文字轉回 Date 物件
    if (dateRaw instanceof Date) {
      dateStr = Utilities.formatDate(dateRaw, "GMT+8", "yyyy-MM-dd");
    } else {
      dateStr = String(dateRaw).trim();
    }
    
    // 強制轉型：確保 C 欄傳上來的一定是嚴格的 Boolean true/false
    const isDoublePay = (String(isDoublePayStr).toUpperCase() === 'TRUE');
    
    // 寫入字典
    holidaysMap[dateStr] = {
      name: String(name).trim(),
      isDoublePay: isDoublePay,
      type: String(type).trim()
    };
  }
  
  return holidaysMap;
}