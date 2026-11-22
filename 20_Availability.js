/**
 * 20_Availability.js — 員工意願填報、臨時解鎖、店長填報總覽、登入後初始資料
 */

/** 解析意願點陣 JSON：{ "2026-10-01": ["SHIFT_LUNCH", ...] }，壞資料回傳 {} */
function parseAvailMatrix_(raw) {
  try {
    return typeof raw === "string" ? JSON.parse(raw) : (raw || {});
  } catch (e) {
    return {};
  }
}

/** 讀取某月份全部意願列（每人一列） */
function readAvailRowsForMonth_(yearMonth, tz) {
  const C = COL.AVAIL;
  return readRows_("AVAIL").filter(r => normalizeYearMonth(r[C.year_month], tz) === yearMonth);
}

/** 意願填報時間鎖：當月一律鎖；下月在本月最後 N+1 天起鎖 */
function isFillTimeLocked_(yearMonth) {
  const now = new Date();
  const cy = now.getFullYear(), cm = now.getMonth() + 1, cd = now.getDate();
  const currentYm = `${cy}-${String(cm).padStart(2, '0')}`;
  const { y, m } = ymParts_(yearMonth);
  const isCurrent = yearMonth === currentYm;
  const isNext = (y > cy) || (y === cy && m === cm + 1);
  const lastDay = new Date(cy, cm, 0).getDate();
  return isCurrent || (isNext && cd >= (lastDay - CONFIG.RULES.FILL_LOCK_DAYS_BEFORE_MONTH_END));
}

/** 臨時解鎖名冊（Script Properties，僅回傳未過期項目） */
function getActiveUnlockRegistry(yearMonth) {
  const raw = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_UNLOCK_PREFIX + yearMonth);
  if (!raw) return {};
  try {
    const reg = JSON.parse(raw);
    const nowMs = new Date().getTime();
    const valid = {};
    for (const [k, v] of Object.entries(reg)) {
      if (v && v.expireAt > nowMs) valid[k] = v;
    }
    return valid;
  } catch (e) {
    return {};
  }
}

function apiGetAvailabilityForMonth(token, yearMonth) {
  const user = validateToken(token);
  if (!user) return { success: false, message: "憑證無效" };
  const empId = effectiveEmpId_(user);
  const tz = getTz_();
  const C = COL.AVAIL;

  let savedPreferences = null;
  const availabilityMap = {};
  const row = readAvailRowsForMonth_(yearMonth, tz).find(r => String(r[C.employee_id]).trim() === empId);
  if (row) {
    savedPreferences = {
      work_volume: row[C.pref_work_volume] || "Standard",
      day_preference: String(row[C.pref_day_preference] || "Flexible").trim(),
      partner_id: String(row[C.pref_partner_id] || "").trim(),
      dislike_partner_id: String(row[C.dislike_partner_id] || "").trim()
    };
    for (const [dStr, shiftArr] of Object.entries(parseAvailMatrix_(row[C.availability_matrix]))) {
      if (!availabilityMap[dStr]) {
        availabilityMap[dStr] = {};
        CONFIG.SHIFTS.forEach(s => { availabilityMap[dStr][s.id] = false; });
      }
      if (Array.isArray(shiftArr)) {
        shiftArr.forEach(sId => { if (availabilityMap[dStr][sId] !== undefined) availabilityMap[dStr][sId] = true; });
      }
    }
  }

  const reg = getActiveUnlockRegistry(yearMonth);
  const myUnlockInfo = reg[empId] || reg["ALL"] || null;
  return {
    success: true,
    savedAvailabilityMap: availabilityMap,
    savedPreferences: savedPreferences,
    isTempUnlocked: Boolean(myUnlockInfo),
    myUnlockInfo: myUnlockInfo
  };
}

/** 登入後初始資料：同事清單、員工清單、解鎖名冊（意願資料由 apiGetAvailabilityForMonth 另外載入） */
function apiGetBootstrapData(token, yearMonth) {
  const user = validateToken(token);
  if (!user) return { success: false, message: "憑證無效或過期" };
  const empId = effectiveEmpId_(user);
  const { allStaffList } = getCachedStaffMap();
  return {
    success: true,
    colleagues: allStaffList
      .filter(s => s.employee_id !== empId)
      .map(s => ({ employee_id: s.employee_id, name: s.name, is_core: s.is_core })),
    allStaffList: allStaffList,
    unlockRegistry: Object.values(getActiveUnlockRegistry(yearMonth))
  };
}

function apiSetUnlockFillPermission(token, targetYearMonth, durationHours, targetEmpId) {
  const user = validateToken(token);
  if (!isAdmin_(user)) return { success: false, message: "權限不足：僅店長可執行此操作！" };
  const hours = Number(durationHours);
  if (!(hours > 0)) return { success: false, message: "解鎖時數無效！" };
  const targetVal = targetEmpId ? targetEmpId : "ALL";
  const expireAtMs = new Date().getTime() + hours * 3600 * 1000;
  const { staffMap } = getCachedStaffMap();
  const targetName = targetVal === "ALL" ? "全體同仁" : (staffMap[targetVal] ? staffMap[targetVal].name : targetVal);

  let reg;
  const err = withLock_(CONFIG.LOCK_MS.SHORT, "解鎖失敗，請重試：", () => {
    reg = getActiveUnlockRegistry(targetYearMonth); // 順便清掉過期項目
    reg[targetVal] = { target: targetVal, targetName: targetName, expireAt: expireAtMs, yearMonth: targetYearMonth };
    PropertiesService.getScriptProperties().setProperty(CONFIG.PROP_UNLOCK_PREFIX + targetYearMonth, JSON.stringify(reg));
    return null;
  });
  if (err) return err;
  return {
    success: true,
    unlockRegistry: Object.values(reg),
    message: `已成功為【${targetName}】解鎖 ${targetYearMonth} 意願填報權限，維持開放 ${durationHours} 小時！`
  };
}

function apiSubmitAvailability(token, payload) {
  const user = validateToken(token);
  if (!user) return { success: false, message: "憑證無效或過期" };
  const empId = effectiveEmpId_(user);
  const tz = getTz_();
  const ym = payload.year_month;

  if (isFillTimeLocked_(ym)) {
    const reg = getActiveUnlockRegistry(ym);
    if (!(reg["ALL"] || reg[empId])) {
      return { success: false, message: "安全鎖定：該月份處於唯讀狀態！若需補填請聯繫店長為您開放臨時權限。" };
    }
  }

  const matrixCompact = {};
  let totalActive = 0;
  for (const [dateStr, shifts] of Object.entries(payload.availability || {})) {
    const active = CONFIG.SHIFTS.map(s => s.id).filter(id => shifts[id]);
    if (active.length > 0) {
      matrixCompact[dateStr] = active;
      totalActive += active.length;
    }
  }
  if (totalActive === 0) {
    return {
      success: false,
      message: "⚠️ 防呆阻擋：您尚未勾選任何可出勤班次！系統禁止送出完全空白的整月意願（防止既有排班意願遭到意外清空）。若該月全月不便出勤，請聯繫店長備查。"
    };
  }
  const matrixJson = JSON.stringify(matrixCompact);
  if (matrixJson.length > 45000) {
    return { success: false, message: "⚠️ 儲存失敗：點陣資料體積超過系統安全閾值 (45,000 字元)，請重試或聯繫系統管理員！" };
  }

  return withLock_(CONFIG.LOCK_MS.NORMAL, "伺服器忙碌，請重試：", () => {
    const sheet = getOrCreateSheet_("AVAIL");
    const C = COL.AVAIL;
    const width = CONFIG.SHEETS.AVAIL.headers.length;
    const keep = readRows_("AVAIL")
      .filter(r => !(normalizeYearMonth(r[C.year_month], tz) === ym && String(r[C.employee_id]).trim() === empId))
      .map(r => fitRow_(r, width));
    const pref = payload.preferences || {};
    keep.push([
      `AR_${ym.replace("-", "")}_${empId}`,
      ym,
      empId,
      pref.work_volume || "Standard",
      pref.day_preference || "Flexible",
      pref.partner_id || "",
      pref.dislike_partner_id || "",
      matrixJson,
      nowStr_(tz)
    ]);
    writeAll_("AVAIL", sheet, keep);
    return { success: true };
  });
}

function apiGetManagerSummary(token, yearMonth) {
  const user = validateToken(token);
  if (!isAdmin_(user)) return { success: false, message: "無權限！" };
  const tz = getTz_();
  const { staffMap, allStaffList } = getCachedStaffMap();
  const C = COL.AVAIL;
  const nameOf = id => (staffMap[id] ? staffMap[id].name : id);
  const summaryMap = {};
  allStaffList.forEach(s => {
    summaryMap[s.employee_id] = {
      employee_id: s.employee_id, name: s.name, is_core: s.is_core, shift_count: 0,
      pref_work_volume: 'Standard', pref_partner_name: '', dislike_partner_name: '',
      day_preference: 'Flexible', submitted: false
    };
  });

  readAvailRowsForMonth_(yearMonth, tz).forEach(r => {
    const item = summaryMap[String(r[C.employee_id]).trim()];
    if (!item) return;
    item.submitted = true;
    item.pref_work_volume = r[C.pref_work_volume] || 'Standard';
    item.day_preference = String(r[C.pref_day_preference] || "Flexible").trim();
    item.pref_partner_name = nameOf(String(r[C.pref_partner_id] || "").trim());
    item.dislike_partner_name = nameOf(String(r[C.dislike_partner_id] || "").trim());
    item.shift_count = Object.values(parseAvailMatrix_(r[C.availability_matrix]))
      .reduce((n, arr) => n + (Array.isArray(arr) ? arr.length : 0), 0);
  });
  return { success: true, summary: Object.values(summaryMap) };
}
