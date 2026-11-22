/**
 * 40_Leave.js — 員工班表查詢、請假 / 代班申請、店長審核
 */

/** 員工「我的排班班表」：只回傳本人 Assigned 班次與同班夥伴 */
function apiGetMyRosterSchedule(token, yearMonth) {
  const user = validateToken(token);
  if (!user) return { success: false, message: "憑證無效或過期" };
  const me = effectiveEmpId_(user);
  const tz = getTz_();
  const { staffMap } = getCachedStaffMap();
  const S = COL.SCHED;
  const ASSIGNED = CONFIG.DISPATCH.ASSIGNED;

  const groups = {};
  readRows_("SCHED").forEach(r => {
    const date = normalizeDateStr(r[S.date], tz);
    if (!date.startsWith(yearMonth)) return;
    const key = `${date}___${String(r[S.shift_id]).trim()}`;
    const empId = String(r[S.employee_id]).trim();
    (groups[key] = groups[key] || []).push({
      employee_id: empId,
      name: staffMap[empId] ? staffMap[empId].name : empId,
      is_core: toBool_(r[S.is_core_at_work]),
      dispatch_type: r[S.dispatch_type]
    });
  });

  const list = [];
  for (const [key, staff] of Object.entries(groups)) {
    if (!staff.some(s => s.employee_id === me && s.dispatch_type === ASSIGNED)) continue;
    const [date, shiftId] = key.split("___");
    list.push({
      date: date, shift_id: shiftId, shift_name: getShiftDisplayTitle(shiftId),
      partners: staff.filter(s => s.employee_id !== me && s.dispatch_type === ASSIGNED)
    });
  }
  list.sort(compareDateShift_);
  return { success: true, myRosterSchedule: list };
}

/** 請假頁：本人今天以後的 Assigned 班次，標記是否為該班唯一核心 */
function apiGetMyShiftsForMonth(token, yearMonth) {
  const user = validateToken(token);
  if (!user) return { success: false, message: "憑證無效" };
  const me = effectiveEmpId_(user);
  const tz = getTz_();
  const todayStr = normalizeDateStr(new Date(), tz);
  const S = COL.SCHED;

  const roster = {}, coreCount = {}, mine = [];
  readRows_("SCHED").forEach(r => {
    const date = normalizeDateStr(r[S.date], tz);
    if (!date.startsWith(yearMonth) || r[S.dispatch_type] !== CONFIG.DISPATCH.ASSIGNED) return;
    const shiftId = String(r[S.shift_id]).trim();
    const empId = String(r[S.employee_id]).trim();
    const isCore = toBool_(r[S.is_core_at_work]);
    const key = `${date}_${shiftId}`;
    (roster[key] = roster[key] || []).push(empId);
    coreCount[key] = (coreCount[key] || 0) + (isCore ? 1 : 0);
    if (empId === me && date >= todayStr) mine.push({ date: date, shift_id: shiftId, is_core: isCore });
  });

  const shifts = mine.map(it => {
    const key = `${it.date}_${it.shift_id}`;
    return {
      date: it.date, shift_id: it.shift_id, shift_name: getShiftDisplayTitle(it.shift_id),
      is_sole_core: it.is_core && coreCount[key] <= 1,
      assigned_staff_in_shift: roster[key] || []
    };
  });
  shifts.sort(compareDateShift_);
  return { success: true, shifts: shifts, todayStr: todayStr };
}

function apiSubmitLeaveRequest(token, payload) {
  try {
    const user = validateToken(token);
    if (!user) return { success: false, message: "憑證無效或登入逾時，請重新整理登入！" };
    const me = effectiveEmpId_(user);
    const tz = getTz_();
    const todayStr = normalizeDateStr(new Date(), tz);
    if (!payload || !payload.date || !payload.shift_id) return { success: false, message: "請假班別參數缺失，請重新選擇！" };
    if (payload.date < todayStr) return { success: false, message: "法規防呆阻擋：該班次屬於過去歷史日期，不可提出事前請假！" };

    const S = COL.SCHED;
    const hasValidShift = readRows_("SCHED").some(r =>
      normalizeDateStr(r[S.date], tz) === payload.date && String(r[S.shift_id]).trim() === payload.shift_id &&
      r[S.employee_id] === me && r[S.dispatch_type] === CONFIG.DISPATCH.ASSIGNED);
    if (!hasValidShift) {
      return { success: false, isScheduleMutated: true, message: "⚠️ 班表即時異動告警：您選擇的班次在店長端已被調動或更換！系統已自動重新整理。" };
    }

    const hasSub = Boolean(payload.has_substitute); // 沒有自覓代班時一律不寫入代班人
    getOrCreateSheet_("LEAVE").appendRow([
      "LR_" + Utilities.getUuid().substring(0, 8), me, String(payload.date), String(payload.shift_id),
      String(payload.leave_type || "事假"), String(payload.reason_detail || ""),
      hasSub, hasSub ? String(payload.substitute_id || "") : "", CONFIG.STATUS.PENDING, new Date()
    ]);
    return { success: true, message: "請假申請已成功送出並呈報店長審核！" };
  } catch (err) {
    return { success: false, message: "伺服器處理失敗: " + err.message };
  }
}

function apiGetMyLeaveRequests(token) {
  const user = validateToken(token);
  if (!user) return { success: false, message: "憑證無效" };
  const me = effectiveEmpId_(user);
  const tz = getTz_();
  const { staffMap } = getCachedStaffMap();
  const L = COL.LEAVE;
  const list = readRows_("LEAVE")
    .filter(r => r[L.employee_id] === me)
    .map(r => ({
      request_id: r[L.request_id], date: normalizeDateStr(r[L.date], tz), shift_id: r[L.shift_id],
      leave_type: r[L.leave_type], reason_detail: r[L.reason_detail], has_substitute: Boolean(r[L.has_substitute]),
      substitute_name: staffMap[r[L.substitute_id]] ? staffMap[r[L.substitute_id]].name : "無 (由店長調派)",
      status: r[L.status]
    }));
  return { success: true, requests: list.reverse() };
}

/** 核准：有代班人 → 替換班表；無 → 刪除該班次。之後統一連動稽核與統計 */
function apiApproveLeaveRequest(token, requestId, chosenSubstituteId) {
  const user = validateToken(token);
  if (!isAdmin_(user)) return { success: false, message: "權限不足" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = getTz_();
  const leaveSheet = getSheet_("LEAVE");
  const schedSheet = getSheet_("SCHED");
  if (!leaveSheet || !schedSheet) return { success: false, message: "缺少必要表格" };

  return withLock_(CONFIG.LOCK_MS.NORMAL, "審核失敗：", () => {
    const L = COL.LEAVE, S = COL.SCHED;
    const leaveRows = leaveSheet.getDataRange().getValues();
    const li = leaveRows.findIndex((r, i) => i > 0 && r[L.request_id] === requestId);
    if (li < 0) return { success: false, message: "找不到該申請" };
    const lr = leaveRows[li];
    const rec = {
      employee_id: lr[L.employee_id],
      date: normalizeDateStr(lr[L.date], tz),
      shift_id: lr[L.shift_id],
      sub: lr[L.has_substitute] ? lr[L.substitute_id] : chosenSubstituteId
    };

    const { staffMap } = getCachedStaffMap();
    const schedRows = schedSheet.getDataRange().getValues();
    const si = schedRows.findIndex((r, i) => i > 0 &&
      normalizeDateStr(r[S.date], tz) === rec.date && String(r[S.shift_id]).trim() === rec.shift_id && r[S.employee_id] === rec.employee_id);
    if (si > 0) {
      if (rec.sub) {
        schedSheet.getRange(si + 1, S.employee_id + 1).setValue(rec.sub);
        schedSheet.getRange(si + 1, S.is_core_at_work + 1).setValue(staffMap[rec.sub] ? staffMap[rec.sub].is_core : false);
      } else {
        schedSheet.deleteRow(si + 1);
      }
    }
    if (rec.sub) leaveSheet.getRange(li + 1, L.substitute_id + 1).setValue(rec.sub);
    leaveSheet.getRange(li + 1, L.status + 1).setValue(CONFIG.STATUS.APPROVED);

    afterScheduleChanged_(ss, rec.date.substring(0, 7), tz, false);
    return { success: true, message: si > 0 ? "請假已核准，代班人員已替換，月統計與衝突資料庫已更新！" : "已核准請假（抽單），工時統計已同步校正。" };
  });
}

function apiRejectLeaveRequest(token, requestId) {
  const user = validateToken(token);
  if (!isAdmin_(user)) return { success: false, message: "權限不足" };
  const sheet = getSheet_("LEAVE");
  if (!sheet) return { success: false, message: "找不到該申請" };
  const L = COL.LEAVE;
  const rows = sheet.getDataRange().getValues();
  const i = rows.findIndex((r, idx) => idx > 0 && r[L.request_id] === requestId);
  if (i < 0) return { success: false, message: "找不到該申請" };
  sheet.getRange(i + 1, L.status + 1).setValue(CONFIG.STATUS.REJECTED);
  return { success: true, message: "已駁回請假申請！" };
}
