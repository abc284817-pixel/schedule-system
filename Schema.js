/**
 * 一鍵初始化排班系統架構：建立核心活躍表、歷史封存表與 3 張月度動態檢視看板
 * 【架構優化】：
 * 1. Availability_Raw 與 Availability_Archive 統一採「一人一月一列」聚合矩陣，零備忘浪費。
 * 2. Viewer_Availability 與 Viewer_Schedule 具備「冷熱透明穿透」能力，查歷史封存月份依然動態展開明細！
 */
function setupEntireDatabaseSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const schema = {
    "Staff_Master": [
      "employee_id", "password_salt", "password_hash",
      "name", "role", "is_core", "is_active"
    ],
    "Shift_Setting": [
      "shift_id", "shift_name", "start_time", "end_time",
      "duration_hours", "required_staff", "required_core", "enable_standby"
    ],
    "Availability_Raw": [
      "record_id", "year_month", "employee_id", "pref_work_volume",
      "pref_day_preference", "pref_partner_id", "dislike_partner_id",
      "availability_matrix", "submitted_at"
    ],
    "Availability_Archive": [
      "record_id", "year_month", "employee_id", "pref_work_volume",
      "pref_day_preference", "pref_partner_id", "dislike_partner_id",
      "availability_matrix", "submitted_at", "archived_at"
    ],
    "Schedule_Final": [
      "schedule_id", "date", "shift_id", "employee_id",
      "is_core_at_work", "dispatch_type", "schedule_status", "created_at"
    ],
    "Schedule_Final_Archive": [
      "schedule_id", "date", "shift_id", "employee_id",
      "is_core_at_work", "dispatch_type", "schedule_status", "created_at", "archived_at"
    ],
    "Schedule_Conflicts": [
      "conflict_id", "year_month", "date", "shift_id", "reason", "created_at"
    ],
    "Schedule_Conflicts_Audit": [
      "audit_id", "year_month", "date", "shift_id", "reason", "archived_at"
    ],
    "Leave_Requests": [
      "request_id", "employee_id", "date", "shift_id", "leave_type", "reason_detail", "has_substitute", "substitute_id", "status", "created_at"
    ],
    "Staff_Monthly_Stats": [
      "stat_id", "year_month", "employee_id", "total_shifts", "total_hours", "assigned_hours", "standby_hours", "updated_at"
    ],
    "Session_Tokens": [
      "token", "employee_id", "role", "last_active_at", "expires_at"
    ]
  };

  for (const [sheetName, headers] of Object.entries(schema)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    
    let headerBg = "#34495e";
    if (sheetName.includes("Archive") || sheetName.includes("Audit")) headerBg = "#1e293b";
    if (sheetName === "Schedule_Conflicts") headerBg = "#c0392b";
    if (sheetName === "Leave_Requests") headerBg = "#d35400";
    if (sheetName === "Staff_Monthly_Stats") headerBg = "#27ae60";
    if (sheetName.startsWith("Availability_")) headerBg = "#0f766e";
    
    sheet.getRange(1, 1, 1, headers.length)
         .setBackground(headerBg)
         .setFontColor("#ffffff")
         .setFontWeight("bold");
  }

  // 自動寫入預設班別需求
  const shiftSheet = ss.getSheetByName("Shift_Setting");
  if (shiftSheet) {
    const shiftData = [
      ["SHIFT_LUNCH", "午班", "10:30", "14:30", 4.0, 5, 1, true],
      ["SHIFT_DINNER", "晚班", "17:30", "21:30", 4.0, 5, 1, true]
    ];
    shiftSheet.getRange(2, 1, shiftData.length, shiftData[0].length).setValues(shiftData);
  }

  setupMonthlyViewers(ss);
  refreshAllViewers("2026-10");

  Logger.log("✅ 系統資料庫全新架構與 3 張看板初始化完成！");
}

/**
 * 建立 3 張檢視看板外觀
 */
function setupMonthlyViewers(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  // A. 正式班表看板
  let vSched = ss.getSheetByName("Viewer_Schedule");
  if (!vSched) vSched = ss.insertSheet("Viewer_Schedule", 0);
  else vSched.clear();
  vSched.getRange("A1").setValue("檢視月份 (YYYY-MM)：").setFontWeight("bold");
  vSched.getRange("B1").setNumberFormat("@").setValue("2026-10").setBackground("#fff2cc").setFontWeight("bold");
  const schedHeaders = ["出勤日期", "班別代碼", "出勤同仁代號", "核心人力覆蓋", "排定指派型態", "排班狀態"];
  vSched.getRange(3, 1, 1, schedHeaders.length).setValues([schedHeaders]).setBackground("#2980b9").setFontColor("#ffffff").setFontWeight("bold");
  vSched.setFrozenRows(3);

  // B. 意願填報看板 (維持 11 欄展示體驗)
  let vAvail = ss.getSheetByName("Viewer_Availability");
  if (!vAvail) vAvail = ss.insertSheet("Viewer_Availability", 1);
  else vAvail.clear();
  vAvail.getRange("A1").setValue("檢視月份 (YYYY-MM)：").setFontWeight("bold");
  vAvail.getRange("B1").setNumberFormat("@").setValue("2026-10").setBackground("#fff2cc").setFontWeight("bold");
  
  const availHeaders = [
    "填報月份", "員工代號", "同仁姓名", "出勤日期", "班別代碼",
    "可出勤", "工時傾向", "單日排班偏好", "偏好搭班同仁", "迴避同仁", "填報時間戳記"
  ];
  vAvail.getRange(3, 1, 1, availHeaders.length).setValues([availHeaders]).setBackground("#16a085").setFontColor("#ffffff").setFontWeight("bold");
  vAvail.setFrozenRows(3);

  // C. 請假審批看板
  let vLeaves = ss.getSheetByName("Viewer_Leaves");
  if (!vLeaves) vLeaves = ss.insertSheet("Viewer_Leaves", 2);
  else vLeaves.clear();
  vLeaves.getRange("A1").setValue("檢視月份 (YYYY-MM)：").setFontWeight("bold");
  vLeaves.getRange("B1").setNumberFormat("@").setValue("2026-10").setBackground("#fff2cc").setFontWeight("bold");
  const leaveHeaders = ["申請單號", "請假同仁", "請假日期", "請假班別", "請假假別", "詳細事由", "自覓代班", "指派代班人", "審批狀態", "送單時間"];
  vLeaves.getRange(3, 1, 1, leaveHeaders.length).setValues([leaveHeaders]).setBackground("#e67e22").setFontColor("#ffffff").setFontWeight("bold");
  vLeaves.setFrozenRows(3);
}

function normalizeToYearMonth(val) {
  if (!val) return "";
  const tz = Session.getScriptTimeZone();
  if (val instanceof Date) {
    return Utilities.formatDate(val, tz, "yyyy-MM");
  }
  const str = String(val).trim();
  const match = str.match(/^(\d{4})[-/](\d{1,2})/);
  if (match) {
    return `${match[1]}-${String(match[2]).padStart(2, '0')}`;
  }
  return str.substring(0, 7);
}

/**
 * 刷新看板資料 (核心適配：具備「冷熱資料透明穿透能力」，無論在 Raw 還是 Archive 皆動態展開明細)
 */
function refreshAllViewers(inputYm) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();

  const vAvail = ss.getSheetByName("Viewer_Availability");
  const rawB1 = inputYm || (vAvail ? vAvail.getRange("B1").getValue() : "2026-10");
  const targetYm = normalizeToYearMonth(rawB1) || "2026-10";
  const { staffMap } = getCachedStaffMap();

  // =========================================================================
  // 1. 刷新 Viewer_Schedule (具備冷熱穿透)
  // =========================================================================
  const vSched = ss.getSheetByName("Viewer_Schedule");
  const schedSheet = ss.getSheetByName("Schedule_Final");
  const schedArchiveSheet = ss.getSheetByName("Schedule_Final_Archive");
  
  if (vSched) {
    if (vSched.getLastRow() >= 4) {
      vSched.getRange(4, 1, Math.max(1, vSched.getLastRow() - 3), 6).clearContent();
    }
    vSched.getRange("B1").setNumberFormat("@").setValue(targetYm);
    
    let rows = [];

    // 先查活躍表
    if (schedSheet && schedSheet.getLastRow() > 1) {
      const sData = schedSheet.getDataRange().getValues();
      for (let i = 1; i < sData.length; i++) {
        let dStr = normalizeToYearMonth(sData[i][1]);
        if (dStr === targetYm) {
          let fullDateStr = sData[i][1] instanceof Date ? Utilities.formatDate(sData[i][1], tz, "yyyy-MM-dd") : String(sData[i][1]).substring(0, 10);
          rows.push([fullDateStr, sData[i][2], sData[i][3], sData[i][4], sData[i][5], sData[i][6]]);
        }
      }
    }

    // 活躍表無資料，穿透至歷史封存表
    if (rows.length === 0 && schedArchiveSheet && schedArchiveSheet.getLastRow() > 1) {
      const saData = schedArchiveSheet.getDataRange().getValues();
      for (let i = 1; i < saData.length; i++) {
        let dStr = normalizeToYearMonth(saData[i][1]);
        if (dStr === targetYm) {
          let fullDateStr = saData[i][1] instanceof Date ? Utilities.formatDate(saData[i][1], tz, "yyyy-MM-dd") : String(saData[i][1]).substring(0, 10);
          rows.push([fullDateStr, saData[i][2], saData[i][3], saData[i][4], saData[i][5], saData[i][6]]);
        }
      }
    }

    if (rows.length > 0) {
      rows.sort((a, b) => a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]));
      vSched.getRange(4, 1, rows.length, rows[0].length).setValues(rows);
    } else {
      vSched.getRange("A4").setValue("該月份無正式排班紀錄 (活躍表與封存庫皆無資料)");
    }
  }

  // =========================================================================
  // 2. 刷新 Viewer_Availability (具備冷熱穿透與點陣動態展開)
  // =========================================================================
  const availSheet = ss.getSheetByName("Availability_Raw");
  const availArchiveSheet = ss.getSheetByName("Availability_Archive");

  if (vAvail) {
    const availHeaders = [
      "填報月份", "員工代號", "同仁姓名", "出勤日期", "班別代碼",
      "可出勤", "工時傾向", "單日排班偏好", "偏好搭班同仁", "迴避同仁", "填報時間戳記"
    ];
    vAvail.getRange(3, 1, 1, 11).setValues([availHeaders]).setBackground("#16a085").setFontColor("#ffffff").setFontWeight("bold");
    if (vAvail.getLastRow() >= 4) {
      vAvail.getRange(4, 1, Math.max(1, vAvail.getLastRow() - 3), 11).clearContent();
    }
    vAvail.getRange("B1").setNumberFormat("@").setValue(targetYm);
    
    let sourceRows = [];

    // 先查活躍表 Availability_Raw
    if (availSheet && availSheet.getLastRow() > 1) {
      const aData = availSheet.getDataRange().getValues();
      for (let i = 1; i < aData.length; i++) {
        if (normalizeToYearMonth(aData[i][1]) === targetYm) {
          sourceRows.push(aData[i]);
        }
      }
    }

    // 活躍表無資料，自動穿透至歷史冷儲存表 Availability_Archive
    if (sourceRows.length === 0 && availArchiveSheet && availArchiveSheet.getLastRow() > 1) {
      const arcData = availArchiveSheet.getDataRange().getValues();
      for (let i = 1; i < arcData.length; i++) {
        if (normalizeToYearMonth(arcData[i][1]) === targetYm) {
          sourceRows.push(arcData[i]);
        }
      }
    }

    const expandedRows = [];
    sourceRows.forEach(row => {
      let empId = String(row[2]).trim();
      let empName = staffMap[empId] ? staffMap[empId].name : empId;
      let volText = row[3] === "More" ? "高工時需求" : (row[3] === "Less" ? "彈性減班" : "標準工時");
      let dPrefVal = String(row[4] || "Flexible").trim();
      let dPrefText = "彈性安排";
      if (dPrefVal === "Both_Preferred") dPrefText = "全日班";
      else if (dPrefVal === "Single_Preferred") dPrefText = "單一班次";
      let pId = String(row[5] || "").trim();
      let dpId = String(row[6] || "").trim();
      let pText = pId ? (staffMap[pId] ? `${staffMap[pId].name} (${pId})` : pId) : "無";
      let dpText = dpId ? (staffMap[dpId] ? `${staffMap[dpId].name} (${dpId})` : dpId) : "無";
      let matrixRaw = row[7];
      let submittedAt = row[8] ? (row[8] instanceof Date ? Utilities.formatDate(row[8], tz, "yyyy-MM-dd HH:mm:ss") : String(row[8])) : "";

      let matrix = {};
      try {
        matrix = typeof matrixRaw === "string" ? JSON.parse(matrixRaw) : (matrixRaw || {});
      } catch (e) {}

      const sortedDates = Object.keys(matrix).sort();
      sortedDates.forEach(dStr => {
        const shiftArr = matrix[dStr] || [];
        shiftArr.forEach(sId => {
          expandedRows.push([
            targetYm, empId, empName, dStr, sId, true,
            volText, dPrefText, pText, dpText, submittedAt
          ]);
        });
      });
    });

    if (expandedRows.length > 0) {
      expandedRows.sort((a, b) => a[3].localeCompare(b[3]) || a[4].localeCompare(b[4]) || a[1].localeCompare(b[1]));
      vAvail.getRange(4, 1, expandedRows.length, expandedRows[0].length).setValues(expandedRows);
    } else {
      vAvail.getRange("A4").setValue("該月份無同仁申報意願紀錄 (活躍表與封存庫皆無資料)");
    }
  }

  // =========================================================================
  // 3. 刷新 Viewer_Leaves
  // =========================================================================
  const leaveSheet = ss.getSheetByName("Leave_Requests");
  const vLeaves = ss.getSheetByName("Viewer_Leaves");
  if (leaveSheet && vLeaves) {
    if (vLeaves.getLastRow() >= 4) {
      vLeaves.getRange(4, 1, Math.max(1, vLeaves.getLastRow() - 3), 10).clearContent();
    }
    vLeaves.getRange("B1").setNumberFormat("@").setValue(targetYm);
    
    const lData = leaveSheet.getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < lData.length; i++) {
      let dStr = normalizeToYearMonth(lData[i][2]);
      if (dStr === targetYm) {
        let fullDateStr = lData[i][2] instanceof Date ? Utilities.formatDate(lData[i][2], tz, "yyyy-MM-dd") : String(lData[i][2]).substring(0, 10);
        rows.push([
          lData[i][0], lData[i][1], fullDateStr, lData[i][3], lData[i][4],
          lData[i][5], lData[i][6], lData[i][7], lData[i][8], lData[i][9]
        ]);
      }
    }
    if (rows.length > 0) {
      vLeaves.getRange(4, 1, rows.length, rows[0].length).setValues(rows);
    } else {
      vLeaves.getRange("A4").setValue("該月份無請假審批紀錄");
    }
  }
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName.startsWith("Viewer_") && e.range.getA1Notation() === "B1") {
    const rawVal = e.range.getValue();
    const ym = normalizeToYearMonth(rawVal);
    if (ym) {
      refreshAllViewers(ym);
    }
  }
}