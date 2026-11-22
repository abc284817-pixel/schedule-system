/**
 * =========================================================================
 * 💥 模擬腳本 1：生成 2024-02 至 2026-09 全量「已填報活躍資料」
 * 測試定位：
 * 1. 意願全部寫入活躍表【Availability_Raw】(一人一月一列 9 欄聚合矩陣)
 * 2. 班表全部寫入活躍表【Schedule_Final】
 * 3. 冷儲存表【Availability_Archive】與【Schedule_Final_Archive】保留空白或清空，
 *    由店長手動在系統後台執行「apiArchiveHistoricalData」進行封存驗證！
 * =========================================================================
 */
function runSimulateHistory_202402_202609() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const nowStr = Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss");

  Logger.log("🚀 開始注入 2024-02 至 2026-09 員工已填報活躍資料 (熱儲存區)...");

  const { allStaffList } = getCachedStaffMap();
  const shifts = ["SHIFT_LUNCH", "SHIFT_DINNER"];

  // 1. 產生 2024-02 至 2026-09 月份序列 (共 32 個月)
  const months = [];
  let y = 2024, m = 2;
  while (y < 2026 || (y === 2026 && m <= 9)) {
    months.push(`${y}-${String(m).padStart(2, '0')}`);
    m++;
    if (m > 12) { m = 1; y++; }
  }

  // 2. 準備標準 9 欄活躍意願表頭與資料容器
  const availRawHeaders = [
    "record_id", "year_month", "employee_id", "pref_work_volume",
    "pref_day_preference", "pref_partner_id", "dislike_partner_id",
    "availability_matrix", "submitted_at"
  ];
  const availRawRows = [availRawHeaders];

  // 3. 準備正式班表表頭與資料容器
  const schedHeaders = [
    "schedule_id", "date", "shift_id", "employee_id",
    "is_core_at_work", "dispatch_type", "schedule_status", "created_at"
  ];
  const schedRows = [schedHeaders];

  // 4. 準備員工月度工時統計表頭與資料容器
  const statsHeaders = [
    "stat_id", "year_month", "employee_id", "total_shifts", "total_hours",
    "assigned_hours", "standby_hours", "updated_at"
  ];
  const statsRows = [statsHeaders];

  let schedSeq = 100000;

  // 5. 逐月產生擬真活躍資料
  months.forEach(ymStr => {
    const [year, month] = ymStr.split("-").map(Number);
    const daysInMonth = new Date(year, month, 0).getDate();
    const monthSubmittedAt = `${ymStr}-20 18:30:00`;

    const empWorkCounts = {};
    allStaffList.forEach(s => empWorkCounts[s.employee_id] = 0);

    // A. 模擬全店 15 位同仁每人一列聚合意願矩陣
    allStaffList.forEach((staff, sIdx) => {
      const vol = (sIdx % 3 === 0) ? "More" : ((sIdx % 3 === 1) ? "Standard" : "Less");
      const dayPref = (sIdx % 2 === 0) ? "Flexible" : "Single_Preferred";
      const partner = (sIdx + 1 < allStaffList.length) ? allStaffList[sIdx + 1].employee_id : "";
      const dislike = (sIdx === 0) ? "EMP008" : "";

      const matrix = {};
      for (let d = 1; d <= daysInMonth; d++) {
        const dStr = `${year}-${String(month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
        const activeShifts = [];
        // 模擬同仁常態出勤勾選
        if ((d + sIdx) % 3 !== 0) activeShifts.push("SHIFT_LUNCH");
        if ((d + sIdx) % 4 !== 0) activeShifts.push("SHIFT_DINNER");
        if (activeShifts.length > 0) {
          matrix[dStr] = activeShifts;
        }
      }

      availRawRows.push([
        `AR_${ymStr.replace("-", "")}_${staff.employee_id}`,
        ymStr,
        staff.employee_id,
        vol,
        dayPref,
        partner,
        dislike,
        JSON.stringify(matrix),
        monthSubmittedAt
      ]);
    });

    // B. 模擬門市正式排班結果 (每日每班滿編 5 人，保證 1 核心人力出勤)
    for (let d = 1; d <= daysInMonth; d++) {
      const dStr = `${year}-${String(month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      shifts.forEach(shiftId => {
        // 核心人力挑選 1 人
        const coreStaff = allStaffList.filter(s => s.is_core);
        const chosenCore = coreStaff[(d + (shiftId === "SHIFT_LUNCH" ? 0 : 1)) % coreStaff.length];
        
        schedSeq++;
        schedRows.push([
          "SCH_" + schedSeq, dStr, shiftId, chosenCore.employee_id,
          true, "Assigned", "Confirmed", monthSubmittedAt
        ]);
        empWorkCounts[chosenCore.employee_id]++;

        // 一般同仁挑選 4 人
        const nonCore = allStaffList.filter(s => s.employee_id !== chosenCore.employee_id);
        for (let i = 0; i < 4; i++) {
          schedSeq++;
          const staff = nonCore[(d * 2 + i) % nonCore.length];
          schedRows.push([
            "SCH_" + schedSeq, dStr, shiftId, staff.employee_id,
            staff.is_core, "Assigned", "Confirmed", monthSubmittedAt
          ]);
          empWorkCounts[staff.employee_id]++;
        }
      });
    }

    // C. 建立月度工時統計基準
    allStaffList.forEach(s => {
      const assignedCount = empWorkCounts[s.employee_id] || 0;
      const assignedHours = assignedCount * 4.0;
      statsRows.push([
        `STAT_${ymStr.replace("-", "")}_${s.employee_id}`,
        ymStr,
        s.employee_id,
        assignedCount,
        assignedHours,
        assignedHours,
        0,
        now
      ]);
    });
  });

  // 6. 原地寫入【活躍資料庫工作表】
  const availRawSheet = ss.getSheetByName("Availability_Raw");
  if (availRawSheet) {
    inPlaceReplaceSheetData(availRawSheet, availRawRows, 9);
    Logger.log(`✅ 活躍意願寫入完成：共 ${months.length} 個月，嚴格佔用 ${availRawRows.length - 1} 列！`);
  }

  const schedSheet = ss.getSheetByName("Schedule_Final");
  if (schedSheet) {
    inPlaceReplaceSheetData(schedSheet, schedRows, 8);
    Logger.log(`✅ 活躍班表寫入完成：共 ${schedRows.length - 1} 筆排班記錄！`);
  }

  // 月工時統計改由正式引擎重算（9 欄：含雙薪 / 備援），避免模擬器自行寫入舊 8 欄格式

  // 7. 將冷儲存區重設為乾淨狀態 (僅保留表頭，供你手動執行封存測試)
  const availArcSheet = ss.getSheetByName("Availability_Archive");
  if (availArcSheet) {
    const arcAvailHeaders = [
      "record_id", "year_month", "employee_id", "pref_work_volume",
      "pref_day_preference", "pref_partner_id", "dislike_partner_id",
      "availability_matrix", "submitted_at", "archived_at"
    ];
    inPlaceReplaceSheetData(availArcSheet, [arcAvailHeaders], 10);
  }

  const schedArcSheet = ss.getSheetByName("Schedule_Final_Archive");
  if (schedArcSheet) {
    const arcSchedHeaders = [
      "schedule_id", "date", "shift_id", "employee_id",
      "is_core_at_work", "dispatch_type", "schedule_status", "created_at", "archived_at"
    ];
    inPlaceReplaceSheetData(schedArcSheet, [arcSchedHeaders], 9);
  }

  // 8. 刷新動態檢視看板至最新歷史月份 (2026-09)
  refreshAllViewers("2026-09");

  Logger.log("🎉 歷史活躍數據模擬完成！現在你可以登入系統後台，測試執行「歷史資料封存」，檢驗資料搬移至冷儲存區與幽靈列裁切的功能！");
  recalculateAllMonthlyStats_();
  Logger.log("✅ 月工時統計已由 recalculateAllMonthlyStats_ 重算完成");
}