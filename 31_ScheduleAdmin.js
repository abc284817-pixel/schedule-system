/**
 * 31_ScheduleAdmin.js — 店長班表檢視、手動調度批次寫回、班表異動後的統一連動
 */

/**
 * 班表任何異動後的唯一連動入口：重算衝突稽核 + 月工時統計（+ 試算表看板）。
 * 新增會改動 Schedule_Final 的功能時，寫入後一律呼叫這裡，不要各自呼叫 sync*。
 * @return {Array} 本月新產生的衝突列
 */
function afterScheduleChanged_(ss, yearMonth, tz, refreshViewers) {
  const conflicts = syncScheduleConflictsForMonth(ss, yearMonth, tz);
  syncStaffMonthlyStatsForMonth(ss, yearMonth, tz);
  if (refreshViewers) refreshAllViewers(yearMonth);
  return conflicts;
}

function apiGetScheduleTable(token, yearMonth) {
  const user = validateToken(token);
  if (!isAdmin_(user)) return { success: false, message: "無權限！" };
  const tz = getTz_();
  const { staffMap, allStaffList } = getCachedStaffMap();
  const nameOf = id => (staffMap[id] ? staffMap[id].name : id);
  const S = COL.SCHED, F = COL.CONFLICTS, L = COL.LEAVE;

  const statMap = {};
  allStaffList.forEach(s => { statMap[s.employee_id] = { employee_id: s.employee_id, name: s.name, is_core: s.is_core, assigned_hours: 0 }; });

  const schedules = [];
  readRows_("SCHED").forEach(r => {
    const date = normalizeDateStr(r[S.date], tz);
    if (!date.startsWith(yearMonth)) return;
    const empId = String(r[S.employee_id]).trim();
    const type = r[S.dispatch_type];
    if (type === CONFIG.DISPATCH.ASSIGNED && statMap[empId]) statMap[empId].assigned_hours += CONFIG.RULES.SHIFT_HOURS;
    schedules.push({
      schedule_id: r[S.schedule_id], date: date, shift_id: String(r[S.shift_id]).trim(),
      employee_id: empId, employee_name: nameOf(empId), is_core_at_work: toBool_(r[S.is_core_at_work]),
      dispatch_type: type, status: r[S.schedule_status]
    });
  });

  const conflicts = readRows_("CONFLICTS")
    .filter(r => normalizeYearMonth(r[F.year_month], tz) === yearMonth)
    .map(r => ({
      conflict_id: r[F.conflict_id], date: normalizeDateStr(r[F.date], tz), shift_id: r[F.shift_id],
      reason: r[F.reason], is_legal: isLegalConflictReason_(r[F.reason])
    }));

  const pendingLeaves = [];
  const leaveAuditHistory = [];
  readRows_("LEAVE").forEach(r => {
    const date = normalizeDateStr(r[L.date], tz);
    if (!date.startsWith(yearMonth)) return;
    const empId = r[L.employee_id], subId = r[L.substitute_id], hasSub = r[L.has_substitute];
    if (r[L.status] === CONFIG.STATUS.PENDING) {
      pendingLeaves.push({
        request_id: r[L.request_id], employee_id: empId, employee_name: nameOf(empId), date: date,
        shift_id: r[L.shift_id], leave_type: r[L.leave_type], reason_detail: r[L.reason_detail],
        has_substitute: Boolean(hasSub), substitute_id: subId,
        substitute_name: staffMap[subId] ? staffMap[subId].name : "", recommended_candidates: []
      });
    } else {
      leaveAuditHistory.push({
        request_id: r[L.request_id], employee_name: nameOf(empId), date: date, shift_id: r[L.shift_id],
        leave_type: r[L.leave_type], reason_detail: r[L.reason_detail],
        substitute_name: staffMap[subId] ? staffMap[subId].name : (hasSub ? "指定同仁" : "無指派"),
        status: r[L.status]
      });
    }
  });

  return {
    success: true,
    schedules: schedules,
    conflicts: conflicts,
    staffList: allStaffList,
    pendingLeaves: pendingLeaves,
    leaveAuditHistory: leaveAuditHistory.reverse(),
    monthlyStats: Object.values(statMap)
  };
}

/** 手動調度批次寫回（前端草稿 → 一次寫入） */
function apiBatchCommitScheduleModifications(token, yearMonth, updates, additions, deletions) {
  const user = validateToken(token);
  if (!isAdmin_(user)) return { success: false, message: "權限不足：僅主管可儲存調度！" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = getTz_();
  const sheet = getSheet_("SCHED");
  const { staffMap } = getCachedStaffMap();
  if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: "找不到正式班表，無法套用修改！" };
  const isCoreOf = id => (staffMap[id] ? staffMap[id].is_core : false);
  const S = COL.SCHED;

  return withLock_(CONFIG.LOCK_MS.NORMAL, "批次儲存失敗: ", () => {
    const deleteSet = new Set(deletions || []);
    const updateMap = {};
    (updates || []).forEach(u => { updateMap[u.schedule_id] = u; });

    const rows = [];
    readRows_("SCHED").forEach(r => {
      const id = r[S.schedule_id];
      if (deleteSet.has(id)) return;
      const u = updateMap[id];
      if (!u) return rows.push(r);
      const nr = r.slice();
      nr[S.employee_id] = u.employee_id;
      nr[S.is_core_at_work] = isCoreOf(u.employee_id);
      nr[S.dispatch_type] = u.dispatch_type;
      nr[S.created_at] = new Date();
      rows.push(nr);
    });
    const now = nowStr_(tz);
    (additions || []).forEach(a => {
      rows.push(["SCH_" + Utilities.getUuid().substring(0, 8), a.date, a.shift_id, a.employee_id,
        isCoreOf(a.employee_id), a.dispatch_type || CONFIG.DISPATCH.ASSIGNED, CONFIG.STATUS.CONFIRMED, now]);
    });
    writeAll_("SCHED", sheet, rows);

    const conflicts = afterScheduleChanged_(ss, yearMonth, tz, true);
    return {
      success: true,
      message: "🎉 班表已更新！",
      conflicts: conflicts.map(c => ({ conflict_id: c[0], date: c[2], shift_id: c[3], reason: c[4] }))
    };
  });
}
