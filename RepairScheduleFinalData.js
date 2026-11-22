/**
 * 專屬工具函式：一鍵標準化 Staff_Monthly_Stats
 * 1. 強制格式化 year_month 為 YYYY-MM (補齊兩位數，例如 2026-03)
 * 2. 由過去依序升冪排序到現在
 */
function cleanAndSortStaffMonthlyStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const statSheet = ss.getSheetByName("Staff_Monthly_Stats");

  if (!statSheet) throw new Error("找不到 Staff_Monthly_Stats 表格！");

  const data = statSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log("Staff_Monthly_Stats 尚無資料需整理。");
    return;
  }

  const cleanRows = [];
  const headers = ["stat_id", "year_month", "employee_id", "total_shifts", "total_hours", "assigned_hours", "standby_hours", "updated_at"];

  for (let i = 1; i < data.length; i++) {
    const [statId, rawYm, empId, totalShifts, totalHours, assignedHours, standbyHours, updatedAt] = data[i];
    
    // 統一格式化：嚴格解析並強制補 0 為 YYYY-MM
    let ymStr = "";
    if (rawYm instanceof Date) {
      ymStr = Utilities.formatDate(rawYm, tz, "yyyy-MM");
    } else {
      const match = String(rawYm).trim().match(/^(\d{4})[/-](\d{1,2})/);
      if (match) {
        ymStr = `${match[1]}-${match[2].padStart(2, '0')}`;
      }
    }

    if (ymStr && empId) {
      cleanRows.push([
        statId || ("STAT_" + Utilities.getUuid().substring(0, 8)),
        ymStr,
        String(empId).trim(),
        Number(totalShifts) || 0,
        Number(totalHours) || 0,
        Number(assignedHours) || 0,
        Number(standbyHours) || 0,
        updatedAt instanceof Date ? updatedAt : new Date()
      ]);
    }
  }

  // 嚴格升冪排序：由過去依序排到現在 (year_month 升冪，同月份再依 employee_id 升冪)
  cleanRows.sort((a, b) => {
    const ymComp = a[1].localeCompare(b[1]);
    if (ymComp !== 0) return ymComp;
    return a[2].localeCompare(b[2]);
  });

  const finalRows = [headers, ...cleanRows];

  statSheet.clear();
  if (finalRows.length > statSheet.getMaxRows()) {
    statSheet.insertRowsAfter(statSheet.getMaxRows(), finalRows.length - statSheet.getMaxRows());
  }
  statSheet.getRange(1, 1, finalRows.length, finalRows[0].length).setValues(finalRows);
  statSheet.setFrozenRows(1);
  statSheet.getRange(1, 1, 1, headers.length).setBackground("#27ae60").setFontColor("#ffffff").setFontWeight("bold");

  Logger.log(`🎉 整理完成！共標準化 ${cleanRows.length} 筆資料，year_month 已統一為 YYYY-MM 並完成升冪排序。`);
}