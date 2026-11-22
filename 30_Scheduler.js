/**
 * 30_Scheduler.js — 智慧排班
 * apiRunAutoSchedule：讀資料 → 確認提示 → 呼叫純演算法 → 寫入 → afterScheduleChanged_
 * solveMonthSchedule_：純函式（不碰試算表），調整排班權重只改這個函式。
 */

/** 上月最後 28 天的出勤日（跨月七休一 / 例休判斷用）：{ empId: Set<dateStr> } */
function getPrecedingMonthTailWorkDates(ss, targetYearMonth, tz) {
  const prevYm = getPreviousYearMonth(targetYearMonth);
  const tailMap = {};
  const { y, m, days } = ymParts_(prevYm);
  const cutoff = dateStrOf_(y, m, Math.max(1, days - (CONFIG.RULES.ROLLING_WINDOW_DAYS - 1)));
  const C = COL.SCHED;
  readRows_("SCHED").forEach(r => {
    const dStr = normalizeDateStr(r[C.date], tz);
    if (dStr.startsWith(prevYm) && dStr >= cutoff && r[C.dispatch_type] === CONFIG.DISPATCH.ASSIGNED) {
      const empId = String(r[C.employee_id]).trim();
      if (!tailMap[empId]) tailMap[empId] = new Set();
      tailMap[empId].add(dStr);
    }
  });
  return tailMap;
}

function apiRunAutoSchedule(token, yearMonth, confirmed) {
  const user = validateToken(token);
  if (!isAdmin_(user)) return { success: false, message: "權限不足：僅店長可執行排班！" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = getTz_();

  return withLock_(CONFIG.LOCK_MS.LONG, "自動排班執行失敗: ", () => {
    const { allStaffList } = getCachedStaffMap();
    const A = COL.AVAIL, S = COL.SCHED;

    // 1. 讀取意願
    const staffPrefs = {};
    const availMap = {};
    let totalAvailCount = 0;
    readAvailRowsForMonth_(yearMonth, tz).forEach(r => {
      const empId = String(r[A.employee_id]).trim();
      staffPrefs[empId] = {
        volume: r[A.pref_work_volume] || "Standard",
        day_preference: String(r[A.pref_day_preference] || "Flexible").trim(),
        partner: String(r[A.pref_partner_id] || "").trim(),
        dislike: String(r[A.dislike_partner_id] || "").trim()
      };
      for (const [dStr, shiftArr] of Object.entries(parseAvailMatrix_(r[A.availability_matrix]))) {
        if (!Array.isArray(shiftArr)) continue;
        shiftArr.forEach(sId => { availMap[`${dStr}_${sId}_${empId}`] = true; totalAvailCount++; });
      }
    });

    // 2. 執行前確認提示
    const schedRows = readRows_("SCHED");
    const existingCount = schedRows.filter(r => normalizeDateStr(r[S.date], tz).startsWith(yearMonth)).length;
    const missing = allStaffList.filter(s => {
      const id = String(s.employee_id || "").trim();
      return id && !staffPrefs[id];
    });
    const warnings = [];
    if (existingCount > 0) warnings.push(`該月份已有 ${existingCount} 筆班表資料，確認後會覆蓋並重新排班。`);
    if (missing.length > 0) warnings.push(`有 ${missing.length} 位員工尚未填報：${missing.map(s => s.name || s.employee_id).join("、")}`);
    if (totalAvailCount === 0) warnings.push("該月份雖有填報紀錄，但目前沒有人選擇任何可上班時段；確認後仍會執行，可能無法排出班表。");
    if (warnings.length > 0 && confirmed !== true) {
      return { success: false, needsConfirmation: true, message: `即將為 ${yearMonth} 執行智慧排班：\n\n` + warnings.join("\n\n") + "\n\n是否確認繼續？" };
    }

    // 3. 演算法（先嚴格模式，有缺口再試寬鬆模式，取缺口較少者）
    const input = {
      yearMonth, tz, allStaffList, staffPrefs, availMap,
      tailWorkDates: getPrecedingMonthTailWorkDates(ss, yearMonth, tz),
      nowStr: nowStr_(tz)
    };
    let result = solveMonthSchedule_(input, false);
    if (result.unsatisfiedCount > 0) {
      const relaxed = solveMonthSchedule_(input, true);
      if (relaxed.unsatisfiedCount <= result.unsatisfiedCount) result = relaxed;
    }

    // 4. 寫入：保留其他月份 + 本月新結果
    const sheet = getOrCreateSheet_("SCHED");
    const keep = readRows_("SCHED").filter(r => !normalizeDateStr(r[S.date], tz).startsWith(yearMonth));
    writeAll_("SCHED", sheet, keep.concat(result.scheduleResult));

    const newConflicts = afterScheduleChanged_(ss, yearMonth, tz, true);
    let msg = `🤖 智慧排班已完成！已啟動全月工時平滑演算法，均勻分發 ${yearMonth} 正式班表，並已同步更新月工時統計。`;
    if (newConflicts.length > 0) {
      msg += `\n\n⚠️ 實體班表異常告警（共 ${newConflicts.length} 處）：\n門市因出勤極限或法規上限仍有班次缺口，已詳實記錄於後台審計庫中，請主管審閱手動補足！`;
    }
    return { success: true, message: msg, conflictCount: newConflicts.length };
  });
}

/**
 * 純排班演算法。回傳 { scheduleResult: Schedule_Final 資料列[], unsatisfiedCount }
 * input: { yearMonth, tz, allStaffList, staffPrefs, availMap, tailWorkDates, nowStr }
 */
function solveMonthSchedule_(input, isRelaxedMode) {
  const { yearMonth, tz, allStaffList, staffPrefs, availMap, tailWorkDates, nowStr } = input;
  const R = CONFIG.RULES;
  const { y: year, m: month, days: daysInMonth } = ymParts_(yearMonth);
  const shiftIds = CONFIG.SHIFTS.map(s => s.id);
  const scheduleResult = [];
  let seq = 60000;
  const empWorkDates = {};
  const empAssignedHours = {};
  allStaffList.forEach(s => {
    empWorkDates[s.employee_id] = new Set(tailWorkDates[s.employee_id] || []);
    empAssignedHours[s.employee_id] = 0;
  });

  // 由最新日期往回數連續出勤天數，超過上限即違規
  function wouldViolateMaxConsecutiveDays(empId, targetDateStr) {
    const sorted = Array.from(new Set(empWorkDates[empId]).add(targetDateStr)).sort();
    let run = 1;
    for (let i = sorted.length - 1; i > 0; i--) {
      const diff = Math.round((new Date(sorted[i]) - new Date(sorted[i - 1])) / 86400000);
      if (diff !== 1) break;
      if (++run > R.MAX_CONSECUTIVE_DAYS) return true;
    }
    return false;
  }

  function isLegallyCapped(empId, targetDateStr) {
    if (empAssignedHours[empId] + R.SHIFT_HOURS > R.MAX_HOURS_28D) return true;
    if (!empWorkDates[empId].has(targetDateStr)) {
      const start = Utilities.formatDate(new Date(new Date(targetDateStr).getTime() - (R.ROLLING_WINDOW_DAYS - 1) * 86400000), tz, "yyyy-MM-dd");
      let count = 0;
      empWorkDates[empId].forEach(d => { if (d >= start && d <= targetDateStr) count++; });
      if (count + 1 > R.MAX_WORK_DAYS_28D) return true;
    }
    return false;
  }

  let unsatisfiedCount = 0;
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = dateStrOf_(year, month, d);
    const dayAssigned = {};
    shiftIds.forEach(id => { dayAssigned[id] = []; });
    const progressRatio = d / daysInMonth;
    const remainingDays = daysInMonth - d + 1;

    shiftIds.forEach(shiftId => {
      const otherShiftIds = shiftIds.filter(id => id !== shiftId);
      const workedOtherShift = empId => otherShiftIds.some(o => dayAssigned[o].some(a => a.employee_id === empId));
      const legal = allStaffList
        .filter(s => availMap[`${dateStr}_${shiftId}_${s.employee_id}`])
        .filter(c => !wouldViolateMaxConsecutiveDays(c.employee_id, dateStr) && !isLegallyCapped(c.employee_id, dateStr));

      // ---- 評分 ----
      const score = {};
      legal.forEach(c => {
        const p = staffPrefs[c.employee_id] || { volume: "Standard", partner: "", dislike: "", day_preference: "Flexible" };
        const hours = empAssignedHours[c.employee_id];
        let s = 100;
        if (!isRelaxedMode) {
          const pace = hours - (R.HOURS_GOAL[p.volume] || R.HOURS_GOAL.Standard) * progressRatio;
          s += pace > 0 ? -(pace * 3.5) : Math.abs(pace) * 2.0;
          let workedDays = 0;
          empWorkDates[c.employee_id].forEach(dt => { if (dt.startsWith(yearMonth)) workedDays++; });
          if (R.MAX_WORK_DAYS_28D - workedDays <= 4 && remainingDays > 8) s -= 35;
          if (R.MAX_HOURS_28D - hours <= 24.0 && remainingDays > 8) s -= 40;
          if (p.volume === "More") s += 25;
          if (p.volume === "Less") s -= 25;
          const both = workedOtherShift(c.employee_id);
          if (p.day_preference === "Both_Preferred") s += both ? 50 : 15;
          else if (p.day_preference === "Single_Preferred") s += both ? -9999 : 25;
          s -= hours * 0.8;
        } else {
          if (p.day_preference === "Single_Preferred" && workedOtherShift(c.employee_id)) s -= 30;
          s -= hours * 0.3;
        }
        score[c.employee_id] = s;
      });
      const byScore = (a, b) => score[b.employee_id] - score[a.employee_id];

      // ---- 先放 1 位核心，再依「偏好搭班 > 分數」補滿，避開互斥同仁 ----
      const assigned = [];
      const core = legal.filter(c => c.is_core).sort(byScore);
      if (core.length > 0) assigned.push({ ...core[0], isCoreWork: true });
      const inShift = id => assigned.some(a => a.employee_id === id);
      const others = legal
        .filter(c => !inShift(c.employee_id))
        .sort((a, b) => {
          const aHas = inShift((staffPrefs[a.employee_id] || {}).partner);
          const bHas = inShift((staffPrefs[b.employee_id] || {}).partner);
          if (aHas !== bHas) return aHas ? -1 : 1;
          return byScore(a, b);
        });
      for (const cand of others) {
        if (assigned.length >= R.STAFF_PER_SHIFT) break;
        const p = staffPrefs[cand.employee_id] || {};
        const disliked = assigned.some(a => a.employee_id === p.dislike || (staffPrefs[a.employee_id] || {}).dislike === cand.employee_id);
        if (!disliked || (isRelaxedMode && others.length <= R.STAFF_PER_SHIFT)) assigned.push({ ...cand, isCoreWork: false });
      }

      dayAssigned[shiftId] = assigned;
      assigned.forEach(st => {
        empWorkDates[st.employee_id].add(dateStr);
        empAssignedHours[st.employee_id] += R.SHIFT_HOURS;
        scheduleResult.push(["SCH_" + (++seq), dateStr, shiftId, st.employee_id, st.isCoreWork, CONFIG.DISPATCH.ASSIGNED, CONFIG.STATUS.CONFIRMED, nowStr]);
      });

      // 備援：剩餘合法人選中第一位
      const standby = legal.find(c => !inShift(c.employee_id));
      if (standby) {
        scheduleResult.push(["SCH_" + (++seq), dateStr, shiftId, standby.employee_id, false, CONFIG.DISPATCH.STANDBY, CONFIG.STATUS.CONFIRMED, nowStr]);
      }
      if (assigned.length < R.STAFF_PER_SHIFT || !assigned.some(s => s.isCoreWork)) unsatisfiedCount++;
    });
  }
  return { scheduleResult, unsatisfiedCount };
}
