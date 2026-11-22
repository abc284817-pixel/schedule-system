/**
 * 模擬情境：2026-11 缺工通報與 160 小時邊界審計測試
 * 特性：4 名核心幹部架構、方案 A 供需博弈、9 欄聚合矩陣相容、零密碼特權注入
 */
function runSimulate202611_ShortageStressTest() {
  const TARGET_YM = "2026-11";
  const DAYS_IN_MONTH = 30; // 11 月 30 天
  const TZ = Session.getScriptTimeZone();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log(`=======================================================`);
  Logger.log(`🚀 開始執行 2026-11 邊界與通報驗證模擬腳本 (4 名核心幹部架構)...`);

  // 1. 讀取 Staff_Master 並識別 4 名核心同仁
  const staffSheet = ss.getSheetByName("Staff_Master");
  if (!staffSheet) throw new Error("找不到 Staff_Master 工作表！");
  const staffData = staffSheet.getDataRange().getValues();
  if (staffData.length <= 1) throw new Error("Staff_Master 表內無員工資料！");

  const header = staffData[0];
  const colId = header.indexOf("employee_id");
  const colName = header.indexOf("name");
  const colRole = header.indexOf("role");
  const colCore = header.indexOf("is_core");
  const colActive = header.indexOf("is_active");

  const allStaff = [];
  for (let r = 1; r < staffData.length; r++) {
    const row = staffData[r];
    const isActive = colActive >= 0 ? (row[colActive] === true || String(row[colActive]).toUpperCase() === "TRUE") : true;
    const role = String(row[colRole]).trim();
    const empId = String(row[colId]).trim();

    // 只納入排班同仁 (排除虛擬管理員帳號 MGR001)
    if (isActive && empId !== "MGR001") {
      allStaff.push({
        employee_id: empId,
        name: String(row[colName]).trim(),
        role: role,
        is_core: (row[colCore] === true || String(row[colCore]).toUpperCase() === "TRUE")
      });
    }
  }

  // 鎖定蕭富駿 (EMP001)
  let targetUser = allStaff.find(s => s.employee_id === "EMP001" || s.name.includes("蕭富駿"));
  if (!targetUser) targetUser = allStaff[0];

  // 識別其餘 3 名核心與一般同仁
  const otherCores = allStaff.filter(s => s.is_core && s.employee_id !== targetUser.employee_id);
  const generalStaff = allStaff.filter(s => !s.is_core && s.employee_id !== targetUser.employee_id);
  const allCoreStaffIds = allStaff.filter(s => s.is_core).map(s => s.employee_id);

  Logger.log(`🎯 極限工時測試目標：${targetUser.name} (${targetUser.employee_id})`);
  Logger.log(`👥 全店核心幹部共 ${allCoreStaffIds.length} 名：`);
  Logger.log(`   - 主力核心：${targetUser.name} (${targetUser.employee_id})`);
  otherCores.forEach((c, idx) => Logger.log(`   - 協同核心 ${idx + 1}：${c.name} (${c.employee_id})`));
  Logger.log(`👥 一般同仁共 ${generalStaff.length} 名`);

  // 2. 隨機選取 3 個異常班次
  const allShifts = [];
  for (let day = 1; day <= DAYS_IN_MONTH; day++) {
    const dStr = `${TARGET_YM}-${String(day).padStart(2, '0')}`;
    allShifts.push({ date: dStr, shift_id: "SHIFT_LUNCH" });
    allShifts.push({ date: dStr, shift_id: "SHIFT_DINNER" });
  }

  const shuffledShifts = allShifts.slice().sort(() => Math.random() - 0.5);
  const shiftZeroCapacity = shuffledShifts[0]; // 1. 完全無人
  const shiftPartialShortage = shuffledShifts[1]; // 2. 局部缺工 (僅 2 人)
  const shiftZeroCore = shuffledShifts[2]; // 3. 零核心當值 (4 名核心全員排休)

  Logger.log(`-------------------------------------------------------`);
  Logger.log(`📌 隨機干預班次規劃：`);
  Logger.log(`  [1. 完全無人]：${shiftZeroCapacity.date} ${shiftZeroCapacity.shift_id === 'SHIFT_LUNCH' ? '午班' : '晚班'}`);
  Logger.log(`  [2. 局部缺工]：${shiftPartialShortage.date} ${shiftPartialShortage.shift_id === 'SHIFT_LUNCH' ? '午班' : '晚班'} (僅 2 人)`);
  Logger.log(`  [3. 零核心當值]：${shiftZeroCore.date} ${shiftZeroCore.shift_id === 'SHIFT_LUNCH' ? '午班' : '晚班'} (4 名核心全員排休)`);

  // 3. 清理指定月份現有資料 (Schedule_Final, Schedule_Conflicts)
  clearMonthRecords("Schedule_Final", 1, TARGET_YM, TZ);
  clearMonthRecords("Schedule_Conflicts", 1, TARGET_YM, TZ);

  // 4. 構造標準 9 欄聚合格式之 Availability_Raw (方案 A 供需博弈)
  const availSheet = ss.getSheetByName("Availability_Raw");
  if (!availSheet) throw new Error("找不到 Availability_Raw 工作表！");

  const availHeaders = [
    "record_id", "year_month", "employee_id", "pref_work_volume",
    "pref_day_preference", "pref_partner_id", "dislike_partner_id",
    "availability_matrix", "submitted_at"
  ];

  // 保留非 2026-11 月份資料
  const oldAvailData = availSheet.getDataRange().getValues();
  const keepAvailRows = [availHeaders];
  for (let i = 1; i < oldAvailData.length; i++) {
    const ym = normalizeYearMonth(oldAvailData[i][1], TZ);
    if (ym !== TARGET_YM && ym !== "") {
      keepAvailRows.push(oldAvailData[i].slice(0, 9));
    }
  }

  const submittedAtStr = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss");
  const luckyPartialStaffIds = [generalStaff[0].employee_id, generalStaff[1].employee_id];

  // 逐人建構意願矩陣
  allStaff.forEach(staff => {
    const isTarget = (staff.employee_id === targetUser.employee_id);
    const isOtherCore = otherCores.some(c => c.employee_id === staff.employee_id);

    // 方案 A 傾向設定：
    // - 蕭富駿：More + Both_Preferred (最高權重)
    // - 其餘 3 名核心：Less (主動退讓排班額度)
    // - 一般同仁：Standard 或 Less
    let prefVolume = "Standard";
    let dayPref = "Flexible";

    if (isTarget) {
      prefVolume = "More";
      dayPref = "Both_Preferred";
    } else if (isOtherCore) {
      prefVolume = "Less";
      dayPref = "Flexible";
    } else {
      prefVolume = Math.random() > 0.5 ? "Less" : "Standard";
      dayPref = "Flexible";
    }

    const prefPartner = isTarget ? (otherCores[0] ? otherCores[0].employee_id : "") : "";
    const dislikePartner = "";
    const matrix = {};

    for (let day = 1; day <= DAYS_IN_MONTH; day++) {
      const dStr = `${TARGET_YM}-${String(day).padStart(2, '0')}`;
      const activeShifts = [];

      let lunchAvail = false;
      let dinnerAvail = false;

      if (isTarget) {
        // 蕭富駿：30 天 60 班次 100% 全開！由排班引擎法規檢查器把關休假與工時
        lunchAvail = true;
        dinnerAvail = true;
      } else if (isOtherCore) {
        // 其餘 3 名核心：壓低出勤意願至約 20%~25% (僅於必要時替補)
        lunchAvail = (Math.random() < 0.22);
        dinnerAvail = (Math.random() < 0.22);
      } else {
        // 一般同仁：壓低出勤意願至約 35%~40%，使每班人數維持在 4~5 人緊繃狀態
        lunchAvail = (Math.random() < 0.38);
        dinnerAvail = (Math.random() < 0.38);
      }

      // --- 套用 3 大干預邏輯 ---
      // 干預 1：完全無人 (全店設為 0)
      if (dStr === shiftZeroCapacity.date) {
        if (shiftZeroCapacity.shift_id === "SHIFT_LUNCH") lunchAvail = false;
        if (shiftZeroCapacity.shift_id === "SHIFT_DINNER") dinnerAvail = false;
      }

      // 干預 2：局部缺工 (僅 luckyPartialStaffIds 可出勤)
      if (dStr === shiftPartialShortage.date) {
        if (shiftPartialShortage.shift_id === "SHIFT_LUNCH") {
          lunchAvail = luckyPartialStaffIds.includes(staff.employee_id);
        }
        if (shiftPartialShortage.shift_id === "SHIFT_DINNER") {
          dinnerAvail = luckyPartialStaffIds.includes(staff.employee_id);
        }
      }

      // 干預 3：零核心班次 (💥 4 名核心全員強制不可出勤；一般同仁放行)
      if (dStr === shiftZeroCore.date) {
        if (shiftZeroCore.shift_id === "SHIFT_LUNCH") {
          if (staff.is_core) lunchAvail = false;
          else lunchAvail = true;
        }
        if (shiftZeroCore.shift_id === "SHIFT_DINNER") {
          if (staff.is_core) dinnerAvail = false;
          else dinnerAvail = true;
        }
      }

      if (lunchAvail) activeShifts.push("SHIFT_LUNCH");
      if (dinnerAvail) activeShifts.push("SHIFT_DINNER");

      if (activeShifts.length > 0) {
        matrix[dStr] = activeShifts;
      }
    }

    keepAvailRows.push([
      `AR_${TARGET_YM.replace('-', '')}_${staff.employee_id}`,
      TARGET_YM,
      staff.employee_id,
      prefVolume,
      dayPref,
      prefPartner,
      dislikePartner,
      JSON.stringify(matrix),
      submittedAtStr
    ]);
  });

  // 原地覆蓋寫入 Availability_Raw (9 欄矩陣)
  if (typeof inPlaceReplaceSheetData === 'function') {
    inPlaceReplaceSheetData(availSheet, keepAvailRows, 9);
  } else {
    availSheet.clear();
    availSheet.getRange(1, 1, keepAvailRows.length, 9).setValues(keepAvailRows);
  }
  Logger.log(`✅ 成功寫入 Availability_Raw 共 ${allStaff.length} 筆同仁 9 欄聚合意願資料！`);

  // 5. 內部特權憑證頒發 (Admin 身分)
  const mockToken = "SIM_TOKEN_" + Utilities.getUuid();
  const cache = CacheService.getScriptCache();
  cache.put("TOKEN_AUTH_" + mockToken, JSON.stringify({ employee_id: "MGR001", role: "Admin" }), 21600);

  let tokenSheet = ss.getSheetByName("Session_Tokens");
  if (tokenSheet) {
    const now = new Date();
    tokenSheet.appendRow([mockToken, "MGR001", "Admin", now, new Date(now.getTime() + 21600 * 1000)]);
  }

  // 6. 自動調用刷新 Viewer 看板
  if (typeof refreshAllViewers === 'function') {
    refreshAllViewers(TARGET_YM);
    Logger.log(`📊 已完成 Viewer_Availability 看板同步刷新！`);
  }

  Logger.log(`=======================================================`);
  Logger.log(`🎉 2026-11 模擬資料注入圓滿完成！`);
  Logger.log(`  1. 請開啟前端系統切換至 2026-11。`);
  Logger.log(`  2. 點擊「一鍵智慧排班 (2026-11)」按鈕。`);
  Logger.log(`  3. 驗證通報：`);
  Logger.log(`     - 【${shiftZeroCapacity.date}】實到 0 人缺 5 人通報。`);
  Logger.log(`     - 【${shiftPartialShortage.date}】未滿 5 人缺工通報。`);
  Logger.log(`     - 【${shiftZeroCore.date}】4 名核心皆未當值之警報。`);
  Logger.log(`  4. 驗證蕭富駿月工時：自然被排班演算法推升至極限 160 hr (40 班) 且零違規！`);
  Logger.log(`=======================================================`);
}

/**
 * 輔助函式：安全清除指定月份舊紀錄
 */
function clearMonthRecords(sheetName, colIndex, yearMonth, tz) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return;

  const data = sheet.getDataRange().getValues();
  const keepRows = [data[0]];

  for (let i = 1; i < data.length; i++) {
    const val = data[i][colIndex];
    let dStr = "";
    if (val instanceof Date) {
      dStr = Utilities.formatDate(val, tz, "yyyy-MM");
    } else {
      dStr = String(val).trim().substring(0, 7);
    }
    if (dStr !== yearMonth) {
      keepRows.push(data[i]);
    }
  }

  sheet.clearContents();
  if (keepRows.length > 0) {
    sheet.getRange(1, 1, keepRows.length, keepRows[0].length).setValues(keepRows);
  }
}

/**
 * 輔助函式：月份字串正規化
 */
function normalizeYearMonth(val, tz) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, tz, "yyyy-MM");
  }
  const s = String(val).trim();
  const m = s.match(/^(\d{4})[\/-](\d{1,2})/);
  if (m) {
    return `${m[1]}-${m[2].padStart(2, '0')}`;
  }
  return s.substring(0, 7);
}