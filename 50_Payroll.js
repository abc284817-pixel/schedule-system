/**
 * 50_Payroll.js — 節日字典、月工時統計（單月同步 / 全量重算）
 * total_hours 只含正班（normal + double），備援工時獨立存放於 standby_hours。
 */

/** 讀 Config_Holidays → { "2026-09-28": { name, isDoublePay, type } }（前端也會呼叫） */
function apiGetDynamicHolidays() {
  const sheet = getSheet_("HOLIDAYS");
  if (!sheet) throw new Error("系統致命錯誤：找不到 Config_Holidays 工作表，請確認工作表名稱是否正確。");
  const H = COL.HOLIDAYS;
  const tz = getTz_();
  const map = {};
  sheet.getDataRange().getValues().slice(1).forEach(r => {
    if (!r[H.date] || !r[H.holiday_name]) return;
    const dateStr = r[H.date] instanceof Date ? Utilities.formatDate(r[H.date], tz, "yyyy-MM-dd") : String(r[H.date]).trim();
    map[dateStr] = {
      name: String(r[H.holiday_name]).trim(),
      isDoublePay: String(r[H.is_double_pay]).toUpperCase() === 'TRUE',
      type: String(r[H.holiday_type]).trim()
    };
  });
  return map;
}

/**
 * 彙總班表列 → { ym: { empId: {assigned, standby, normal, double, standbyHours} } }
 * opts.ym：只算該月；opts.empSet：只算這些員工
 */
function aggregateMonthlyStats_(rows, holidays, tz, opts) {
  const S = COL.SCHED, D = CONFIG.DISPATCH, h = CONFIG.RULES.SHIFT_HOURS;
  const out = {};
  rows.forEach(r => {
    const dateStr = normalizeDateStr(r[S.date], tz);
    const empId = String(r[S.employee_id]).trim();
    if (!dateStr || !empId || String(r[S.schedule_status]).trim() !== CONFIG.STATUS.CONFIRMED) return;
    const ym = dateStr.substring(0, 7);
    if (opts.ym && ym !== opts.ym) return;
    if (opts.empSet && !opts.empSet.has(empId)) return;
    const type = String(r[S.dispatch_type]).trim();
    const byEmp = out[ym] = out[ym] || {};
    const a = byEmp[empId] = byEmp[empId] || { assigned: 0, standby: 0, normal: 0, double: 0, standbyHours: 0 };
    if (type === D.ASSIGNED) {
      a.assigned++;
      if (holidays[dateStr] && holidays[dateStr].isDoublePay === true) a.double += h; else a.normal += h;
    } else if (type === D.STANDBY) {
      a.standby++;
      a.standbyHours += h;
    }
  });
  return out;
}

/** 彙總結果 → Staff_Monthly_Stats 資料列（略過整月無班者） */
function statRowsFromAggregate_(agg, now) {
  const rows = [];
  Object.keys(agg).forEach(ym => {
    Object.keys(agg[ym]).forEach(empId => {
      const a = agg[ym][empId];
      const totalShifts = a.assigned + a.standby;
      if (totalShifts > 0) {
        rows.push([`STAT_${ym.replace("-", "")}_${empId}`, ym, empId, totalShifts, a.normal + a.double, a.normal, a.double, a.standbyHours, now]);
      }
    });
  });
  return rows;
}

function sortStatRows_(rows) {
  return rows.sort((a, b) => String(a[1]).localeCompare(String(b[1])) || String(a[2]).localeCompare(String(b[2])));
}

/** 單月同步（班表異動後由 afterScheduleChanged_ 呼叫；只算在職員工） */
function syncStaffMonthlyStatsForMonth(ss, yearMonth, tz) {
  const sheet = getOrCreateSheet_("STATS");
  const T = COL.STATS;
  const width = CONFIG.SHEETS.STATS.headers.length;
  const empSet = new Set(getCachedStaffMap().allStaffList.map(s => s.employee_id));
  const agg = aggregateMonthlyStats_(readRows_("SCHED"), apiGetDynamicHolidays(), tz, { ym: yearMonth, empSet: empSet });

  const keep = readRows_("STATS")
    .filter(r => { const ym = normalizeYearMonth(r[T.year_month], tz); return ym !== yearMonth && ym !== ""; })
    .map(r => fitRow_(r, width));
  writeAll_("STATS", sheet, sortStatRows_(keep.concat(statRowsFromAggregate_(agg, new Date()))));
}

/** [API] 全量重算（含封存區），僅限店長 */
function apiRecalculateMonthlyStats(token) {
  const user = validateToken(token);
  if (!isAdmin_(user)) return { success: false, message: "權限不足：僅店長可執行工時結算！" };
  return withLock_(CONFIG.LOCK_MS.LONG, "工時結算失敗：", () => recalculateAllMonthlyStats_());
}

/** 全量重算：Schedule_Final + Schedule_Final_Archive，全部員工（含已離職） */
function recalculateAllMonthlyStats_() {
  const sheet = getSheet_("STATS");
  if (!sheet) throw new Error("系統致命錯誤：找不到 Staff_Monthly_Stats 工作表");
  const all = readRows_("SCHED").concat(readRows_("SCHED_ARCHIVE"));
  if (all.length === 0) return { success: true, count: 0, message: "無任何班表資料" };
  const rows = statRowsFromAggregate_(aggregateMonthlyStats_(all, apiGetDynamicHolidays(), getTz_(), {}), new Date());
  if (rows.length > 0) writeAll_("STATS", sheet, sortStatRows_(rows));
  return { success: true, count: rows.length };
}
