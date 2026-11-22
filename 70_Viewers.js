/**
 * 70_Viewers.js — 試算表內的 3 張檢視看板（Viewer_Schedule / Viewer_Availability / Viewer_Leaves）
 * 只影響試算表，不影響網頁。看板 B1 改月份時由 onEdit 觸發刷新；查無資料會自動穿透到封存表。
 */
const VIEWERS = Object.freeze({
  SCHED: "Viewer_Schedule",
  AVAIL: "Viewer_Availability",
  LEAVES: "Viewer_Leaves",
  DATA_START_ROW: 4
});

function onEdit(e) {
  if (!e || !e.range) return;
  const name = e.range.getSheet().getName();
  if (name.startsWith("Viewer_") && e.range.getA1Notation() === "B1") {
    const ym = normalizeYearMonth(e.range.getValue());
    if (ym) refreshAllViewers(ym);
  }
}

/** 寫入看板：清空舊資料 → 設定 B1 → 寫入或顯示空訊息 */
function writeViewer_(sheet, ym, width, rows, emptyMsg) {
  const start = VIEWERS.DATA_START_ROW;
  if (sheet.getLastRow() >= start) sheet.getRange(start, 1, Math.max(1, sheet.getLastRow() - start + 1), width).clearContent();
  sheet.getRange("B1").setNumberFormat("@").setValue(ym);
  if (rows.length > 0) sheet.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
  else sheet.getRange("A" + start).setValue(emptyMsg);
}

/** 先查活躍表，無資料再查封存表 */
function rowsHotThenCold_(hotKey, coldKey, match) {
  const hot = readRows_(hotKey).filter(match);
  return hot.length > 0 ? hot : readRows_(coldKey).filter(match);
}

function refreshAllViewers(inputYm) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = getTz_();
  const vAvail = ss.getSheetByName(VIEWERS.AVAIL);
  const targetYm = normalizeYearMonth(inputYm || (vAvail ? vAvail.getRange("B1").getValue() : "")) || normalizeYearMonth(new Date());
  const { staffMap } = getCachedStaffMap();
  const S = COL.SCHED, A = COL.AVAIL, L = COL.LEAVE;
  const dateOf = v => (v instanceof Date ? Utilities.formatDate(v, tz, "yyyy-MM-dd") : String(v).substring(0, 10));

  // 1. 正式班表
  const vSched = ss.getSheetByName(VIEWERS.SCHED);
  if (vSched) {
    const rows = rowsHotThenCold_("SCHED", "SCHED_ARCHIVE", r => normalizeYearMonth(r[S.date]) === targetYm)
      .map(r => [dateOf(r[S.date]), r[S.shift_id], r[S.employee_id], r[S.is_core_at_work], r[S.dispatch_type], r[S.schedule_status]])
      .sort((a, b) => a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]));
    writeViewer_(vSched, targetYm, 6, rows, "該月份無正式排班紀錄 (活躍表與封存庫皆無資料)");
  }

  // 2. 意願填報（點陣展開成一班一列）
  if (vAvail) {
    const headers = ["填報月份", "員工代號", "同仁姓名", "出勤日期", "班別代碼", "可出勤", "工時傾向", "單日排班偏好", "偏好搭班同仁", "迴避同仁", "填報時間戳記"];
    vAvail.getRange(3, 1, 1, headers.length).setValues([headers]).setBackground("#16a085").setFontColor("#ffffff").setFontWeight("bold");
    const label = id => (id ? (staffMap[id] ? `${staffMap[id].name} (${id})` : id) : "無");
    const volText = { More: "高工時需求", Less: "彈性減班" };
    const dayText = { Both_Preferred: "全日班", Single_Preferred: "單一班次" };
    const rows = [];
    rowsHotThenCold_("AVAIL", "AVAIL_ARCHIVE", r => normalizeYearMonth(r[A.year_month]) === targetYm).forEach(r => {
      const empId = String(r[A.employee_id]).trim();
      const sub = r[A.submitted_at];
      const common = [
        volText[r[A.pref_work_volume]] || "標準工時",
        dayText[String(r[A.pref_day_preference] || "Flexible").trim()] || "彈性安排",
        label(String(r[A.pref_partner_id] || "").trim()),
        label(String(r[A.dislike_partner_id] || "").trim()),
        sub ? (sub instanceof Date ? Utilities.formatDate(sub, tz, "yyyy-MM-dd HH:mm:ss") : String(sub)) : ""
      ];
      const matrix = parseAvailMatrix_(r[A.availability_matrix]);
      Object.keys(matrix).sort().forEach(dStr => {
        (matrix[dStr] || []).forEach(sId => {
          rows.push([targetYm, empId, staffMap[empId] ? staffMap[empId].name : empId, dStr, sId, true].concat(common));
        });
      });
    });
    rows.sort((a, b) => a[3].localeCompare(b[3]) || a[4].localeCompare(b[4]) || a[1].localeCompare(b[1]));
    writeViewer_(vAvail, targetYm, 11, rows, "該月份無同仁申報意願紀錄 (活躍表與封存庫皆無資料)");
  }

  // 3. 請假審批
  const vLeaves = ss.getSheetByName(VIEWERS.LEAVES);
  if (vLeaves && getSheet_("LEAVE")) {
    const rows = readRows_("LEAVE")
      .filter(r => normalizeYearMonth(r[L.date]) === targetYm)
      .map(r => { const x = fitRow_(r, 10); x[L.date] = dateOf(r[L.date]); return x; });
    writeViewer_(vLeaves, targetYm, 10, rows, "該月份無請假審批紀錄");
  }
}
