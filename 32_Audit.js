/**
 * 32_Audit.js — 班表法規 / 營運稽核（缺工、無核心、七休一、28 天 160 小時、四週例休）
 * 衝突文字一律以 CONFIG.CONFLICT[x].prefix 開頭；分類用 isLegalConflictReason_，前端不得比對中文。
 */

function isLegalConflictReason_(reason) {
  const s = String(reason || "");
  return Object.values(CONFIG.CONFLICT).some(t => t.legal && s.indexOf(t.prefix) === 0);
}

function conflictText_(type, body) {
  return CONFIG.CONFLICT[type].prefix + "：" + body;
}

/** 重算某月衝突，覆寫 Schedule_Conflicts 該月資料並寫入稽核庫。回傳新衝突列（CONFLICTS 欄位格式） */
function syncScheduleConflictsForMonth(ss, yearMonth, tz) {
  const confSheet = getOrCreateSheet_("CONFLICTS");
  const schedSheet = getSheet_("SCHED");
  if (!schedSheet || schedSheet.getLastRow() <= 1) return [];
  const R = CONFIG.RULES, S = COL.SCHED, F = COL.CONFLICTS;
  const ASSIGNED = CONFIG.DISPATCH.ASSIGNED;
  const { staffMap, allStaffList } = getCachedStaffMap();
  const nameOf = id => (staffMap[id] ? staffMap[id].name : id);

  const keepConflicts = readRows_("CONFLICTS").filter(r => normalizeYearMonth(r[F.year_month], tz) !== yearMonth);
  const sData = readRows_("SCHED");
  const { y: year, m: month, days: daysInMonth } = ymParts_(yearMonth);

  // 出勤日集合（含上月尾 28 天）
  const empWorkDatesMap = getPrecedingMonthTailWorkDates(ss, yearMonth, tz);
  allStaffList.forEach(s => { if (!empWorkDatesMap[s.employee_id]) empWorkDatesMap[s.employee_id] = new Set(); });

  // 本月每班名單（先建好每日每班空格，缺工才抓得到）
  const shiftMap = {};
  for (let d = 1; d <= daysInMonth; d++) {
    CONFIG.SHIFTS.forEach(s => { shiftMap[dateStrOf_(year, month, d) + "___" + s.id] = []; });
  }
  // 全表每日工時（160 小時跨月滾動用）
  const empDailyHoursMap = {};

  sData.forEach(r => {
    const dStr = normalizeDateStr(r[S.date], tz);
    const empId = String(r[S.employee_id]).trim();
    const type = String(r[S.dispatch_type]).trim();
    if (type === ASSIGNED) {
      if (!empDailyHoursMap[empId]) empDailyHoursMap[empId] = {};
      empDailyHoursMap[empId][dStr] = (empDailyHoursMap[empId][dStr] || 0) + R.SHIFT_HOURS;
    }
    if (!dStr.startsWith(yearMonth)) return;
    const key = dStr + "___" + String(r[S.shift_id]).trim();
    if (!shiftMap[key]) shiftMap[key] = [];
    shiftMap[key].push({ is_core: toBool_(r[S.is_core_at_work]), dispatch_type: type });
    if (type === ASSIGNED) {
      if (!empWorkDatesMap[empId]) empWorkDatesMap[empId] = new Set();
      empWorkDatesMap[empId].add(dStr);
    }
  });

  let confSeq = 50000;
  const now = nowStr_(tz);
  const fresh = [];
  const add = (date, shiftId, reason) => fresh.push(["CONF_" + (++confSeq), yearMonth, date, shiftId, reason, now]);
  const firstShift = CONFIG.SHIFTS[0].id;

  // 1. 缺工 / 無核心
  Object.entries(shiftMap).forEach(([key, list]) => {
    const [dStr, sId] = key.split("___");
    const assigned = list.filter(s => s.dispatch_type === ASSIGNED);
    if (assigned.length < R.STAFF_PER_SHIFT) {
      add(dStr, sId, conflictText_("UNDERSTAFF", `排班人數不足 ${R.STAFF_PER_SHIFT} 人 (僅 ${assigned.length} 人，缺口 ${R.STAFF_PER_SHIFT - assigned.length} 人)`));
    }
    if (assigned.length > 0 && !assigned.some(s => s.is_core)) {
      add(dStr, sId, conflictText_("NO_CORE", "排班後無核心人力出勤！"));
    }
  });

  // 2. 七休一
  Object.entries(empWorkDatesMap).forEach(([empId, dateSet]) => {
    const sorted = Array.from(dateSet).sort();
    let run = 1;
    for (let i = 1; i < sorted.length; i++) {
      const diff = Math.round((new Date(sorted[i]) - new Date(sorted[i - 1])) / 86400000);
      if (diff !== 1) { run = 1; continue; }
      run++;
      if (run > R.MAX_CONSECUTIVE_DAYS && sorted[i].startsWith(yearMonth)) {
        add(sorted[i], firstShift, conflictText_("CONSECUTIVE", `同仁 ${nameOf(empId)}(${empId}) 至本日已連續出勤達 ${run} 天，違反勞基法法定規定！`));
      }
    }
  });

  // 3. 28 天滾動工時上限（每人只記最嚴重一筆）
  Object.entries(empDailyHoursMap).forEach(([empId, dailyHours]) => {
    let worst = null;
    for (let d = 1; d <= daysInMonth; d++) {
      const currStr = dateStrOf_(year, month, d);
      if (!dailyHours[currStr]) continue;
      const currD = new Date(year, month - 1, d);
      const startD = new Date(currD.getTime() - (R.ROLLING_WINDOW_DAYS - 1) * 86400000);
      let hours = 0;
      for (let t = new Date(startD); t <= currD; t.setDate(t.getDate() + 1)) {
        hours += dailyHours[Utilities.formatDate(t, tz, "yyyy-MM-dd")] || 0;
      }
      if (hours > R.MAX_HOURS_28D && (!worst || hours > worst.hours)) {
        worst = { hours: hours, startStr: Utilities.formatDate(startD, tz, "yyyy-MM-dd"), endStr: currStr };
      }
    }
    if (worst) {
      const over = (worst.hours - R.MAX_HOURS_28D).toFixed(1);
      add(worst.endStr, firstShift, conflictText_("OVER_HOURS", `同仁 ${nameOf(empId)}(${empId}) 於滾動區間 (${worst.startStr} ~ ${worst.endStr}) 累計排班達 ${worst.hours} 小時（超標 ${over} hr），突破法定上限！`));
    }
  });

  // 4. 四週例休（28 天內出勤 > 上限天數）
  const minRest = R.ROLLING_WINDOW_DAYS - R.MAX_WORK_DAYS_28D;
  Object.entries(empWorkDatesMap).forEach(([empId, dateSet]) => {
    let worst = null;
    for (let d = 1; d <= daysInMonth; d++) {
      const currStr = dateStrOf_(year, month, d);
      if (!dateSet.has(currStr)) continue;
      const startStr = Utilities.formatDate(new Date(new Date(year, month - 1, d).getTime() - (R.ROLLING_WINDOW_DAYS - 1) * 86400000), tz, "yyyy-MM-dd");
      let count = 0;
      dateSet.forEach(w => { if (w >= startStr && w <= currStr) count++; });
      if (count > R.MAX_WORK_DAYS_28D && (!worst || count > worst.count)) {
        worst = { count: count, restDays: R.ROLLING_WINDOW_DAYS - count, startStr: startStr, endStr: currStr };
      }
    }
    if (worst) {
      add(worst.endStr, firstShift, conflictText_("REST_DAYS", `同仁 ${nameOf(empId)}(${empId}) 於滾動區間 (${worst.startStr} ~ ${worst.endStr}) 累計出勤達 ${worst.count} 天（例休僅 ${worst.restDays} 天），未滿法定 ${minRest} 天！`));
    }
  });

  writeAll_("CONFLICTS", confSheet, keepConflicts.concat(fresh));
  archiveScheduleConflictsAudit(ss, yearMonth, fresh, tz);
  return fresh;
}

/** 稽核庫：覆寫該月份稽核紀錄 */
function archiveScheduleConflictsAudit(ss, yearMonth, conflictRows, tz) {
  const sheet = getOrCreateSheet_("CONFLICTS_AUDIT");
  const A = COL.CONFLICTS_AUDIT, F = COL.CONFLICTS;
  const keep = readRows_("CONFLICTS_AUDIT").filter(r => normalizeYearMonth(r[A.year_month], tz) !== yearMonth);
  const now = nowStr_(tz);
  conflictRows.forEach(c => {
    keep.push(["AUDIT_" + Utilities.getUuid().substring(0, 8), yearMonth, c[F.date], c[F.shift_id], c[F.reason], now]);
  });
  writeAll_("CONFLICTS_AUDIT", sheet, keep);
}
