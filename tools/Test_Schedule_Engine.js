/**
 * =========================================================================
 * 【真實資料庫稽核】：直接讀取試算表中的 Schedule_Final 進行法規審查（含跨月28天例休檢驗）
 * =========================================================================
 */
function auditRealDatabaseSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const targetYearMonth = "2026-11"; // 設定你要檢驗的真實月份
  Logger.log(`🔍 開始對試算表實體資料庫進行排班合規審查 [目標月份：${targetYearMonth}]...`);
  const staffSheet = ss.getSheetByName("Staff_Master");
  const schedSheet = ss.getSheetByName("Schedule_Final");
  const availSheet = ss.getSheetByName("Availability_Raw");
  if (!staffSheet || !schedSheet) {
    throw new Error("缺少 Staff_Master 或 Schedule_Final 工作表！");
  }

  // 1. 載入員工
  const staffData = staffSheet.getDataRange().getValues();
  const staffMap = {};
  for (let i = 1; i < staffData.length; i++) {
    const [empId, , , name, , isCore, isActive] = staffData[i];
    if (isActive === true || String(isActive).toUpperCase() === "TRUE") {
      staffMap[empId] = {
        name: name,
        is_core: (isCore === true || String(isCore).toUpperCase() === "TRUE")
      };
    }
  }

  // 2. 載入員工偏好
  const staffPrefs = {};
  if (availSheet && availSheet.getLastRow() > 1) {
    const aData = availSheet.getDataRange().getValues();
    for (let i = 1; i < aData.length; i++) {
      const ym = normalizeYearMonth(aData[i][1], tz);
      const empId = String(aData[i][2]).trim();
      if (ym === targetYearMonth && !staffPrefs[empId]) {
        staffPrefs[empId] = {
          partner: String(aData[i][COL.AVAIL.pref_partner_id] || "").trim(),
          dislike: String(aData[i][COL.AVAIL.dislike_partner_id] || "").trim()
        };
      }
    }
  }

  // 3. 讀取真實排班（含前月27天出勤紀錄）
  const schedData = schedSheet.getDataRange().getValues();
  const empWorkDates = {};
  const empAssignedHours = {};
  const shiftGrouping = {};
  const prevYm = getPreviousYearMonth(targetYearMonth);
  const [y, m] = prevYm.split("-").map(Number);
  const lastDayOfPrevMonth = new Date(y, m, 0).getDate();
  const cutoffDateStr = `${y}-${String(m).padStart(2, '0')}-${String(Math.max(1, lastDayOfPrevMonth - 27)).padStart(2, '0')}`;

  Object.keys(staffMap).forEach(eId => {
    empWorkDates[eId] = new Set();
    empAssignedHours[eId] = 0;
  });

  let totalShiftRows = 0;
  for (let i = 1; i < schedData.length; i++) {
    const row = schedData[i];
    const dStr = normalizeDateStr(row[1], tz);
    const empId = String(row[3]).trim();
    const dispatchType = String(row[5]).trim();
    if (dispatchType === "Assigned") {
      // 讀取當月以及前月回溯 27 天
      if (dStr.startsWith(targetYearMonth)) {
        totalShiftRows++;
        if (empAssignedHours[empId] !== undefined) empAssignedHours[empId] += 4.0;
        const key = `${dStr}_${String(row[2]).trim()}`;
        if (!shiftGrouping[key]) shiftGrouping[key] = [];
        shiftGrouping[key].push(empId);
      }
      if (dStr.startsWith(targetYearMonth) || (dStr.startsWith(prevYm) && dStr >= cutoffDateStr)) {
        if (empWorkDates[empId]) empWorkDates[empId].add(dStr);
      }
    }
  }

  Logger.log(`📊 讀取到 ${targetYearMonth} 正式排班資料共 ${totalShiftRows} 人次。`);

  // 4. 嚴格法規審計
  let violations160Hr = 0;
  let violations22Days = 0;
  let violationsRolling28Days = 0;
  let violationsConsecutive = 0;
  let violationsDislike = 0;
  let violationsNoCore = 0;
  let violationsShortage = 0;

  const [tY, tM] = targetYearMonth.split("-").map(Number);
  const daysInMonth = new Date(tY, tM, 0).getDate();

  Object.keys(staffMap).forEach(empId => {
    if (empId === "MGR001") return;
    const name = staffMap[empId].name;
    const hours = empAssignedHours[empId] || 0;
    
    // 計算單自然月出勤天數
    let currentMonthDays = 0;
    empWorkDates[empId].forEach(d => {
      if (d.startsWith(targetYearMonth)) currentMonthDays++;
    });

    if (hours > 160.0) {
      violations160Hr++;
      Logger.log(`❌ [法規違規] 同仁 ${name} (${empId}) 總工時 ${hours} hr，已打穿四週 160 小時上限！`);
    }
    if (currentMonthDays > 22) {
      violations22Days++;
      Logger.log(`❌ [法規違規] 同仁 ${name} (${empId}) 本月出勤達 ${currentMonthDays} 天，未休滿自然月 8 休（超過 22 天）！`);
    }

    // 連續上班天數
    const sorted = Array.from(empWorkDates[empId]).sort();
    let run = 1;
    for (let i = 1; i < sorted.length; i++) {
      const diff = Math.round((new Date(sorted[i]) - new Date(sorted[i - 1])) / (1000 * 3600 * 24));
      if (diff === 1) {
        run++;
        if (run > 6 && sorted[i].startsWith(targetYearMonth)) {
          violationsConsecutive++;
          Logger.log(`❌ [法規違規] 同仁 ${name} (${empId}) 於 ${sorted[i]} 連續出勤達第 ${run} 天（違反連六不得出勤）！`);
        }
      } else {
        run = 1;
      }
    }

    // 跨月 28 天滑動視窗審查
    for (let d = 1; d <= daysInMonth; d++) {
      const currDateStr = `${tY}-${String(tM).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      if (!empWorkDates[empId].has(currDateStr)) continue;

      const currD = new Date(tY, tM - 1, d);
      const start28D = new Date(currD.getTime() - 27 * 24 * 3600 * 1000);
      const start28Str = Utilities.formatDate(start28D, tz, "yyyy-MM-dd");

      let countIn28 = 0;
      empWorkDates[empId].forEach(wDate => {
        if (wDate >= start28Str && wDate <= currDateStr) countIn28++;
      });

      if (countIn28 > 22 && start28Str < `${targetYearMonth}-01`) {
        violationsRolling28Days++;
        Logger.log(`❌ [跨月例休不足違規] 同仁 ${name} (${empId}) 於 ${start28Str} ~ ${currDateStr} 28天區間出勤達 ${countIn28} 天（例休僅 ${28 - countIn28} 天，未滿 8 天）！`);
        break;
      }
    }
  });

  // 每班編制與人際互斥
  Object.entries(shiftGrouping).forEach(([shiftKey, staffIds]) => {
    const hasCore = staffIds.some(id => staffMap[id] && staffMap[id].is_core);
    if (!hasCore) {
      violationsNoCore++;
      Logger.log(`⚠️ [營運警示] 班次 ${shiftKey} 無任何核心幹部在場！`);
    }
    if (staffIds.length < 5) {
      violationsShortage++;
      Logger.log(`⚠️ [編制缺口] 班次 ${shiftKey} 正式出勤僅 ${staffIds.length} 人（缺 ${5 - staffIds.length} 人）！`);
    }
    staffIds.forEach(idA => {
      const dislikeB = staffPrefs[idA]?.dislike;
      if (dislikeB && staffIds.includes(dislikeB)) {
        violationsDislike++;
        Logger.log(`⚠️ [人際互斥] 班次 ${shiftKey}：${staffMap[idA]?.name} 與迴避對象 ${staffMap[dislikeB]?.name} 同班出勤！`);
      }
    });
  });

  Logger.log("\n================ 實體資料庫審查總結 ================");
  Logger.log(`審查月份：${targetYearMonth}`);
  Logger.log(`1. 四週 160 小時超標：${violations160Hr} 件`);
  Logger.log(`2. 單月 22 天出勤超標：${violations22Days} 件`);
  Logger.log(`3. 跨月滾動 28 天出勤超標（例休不足 8 天）：${violationsRolling28Days} 件`);
  Logger.log(`4. 連續出勤 >6 天違規：${violationsConsecutive} 件`);
  Logger.log(`5. 無核心幹部班次：${violationsNoCore} 班`);
  Logger.log(`6. 出勤人數未達 5 人：${violationsShortage} 班`);
  Logger.log(`7. 雙向迴避衝突同班：${violationsDislike} 處`);

  if (violations160Hr === 0 && violations22Days === 0 && violationsRolling28Days === 0 && violationsConsecutive === 0) {
    Logger.log("🎉 【資料庫合規判定：完全合格】實體班表 100% 符合四週變形工時、跨月滾動例休與勞基法強制規定！");
  } else {
    Logger.log("🚨 【資料庫合規判定：存在法規違規】請根據上方日誌進行手動調度調整！");
  }
}