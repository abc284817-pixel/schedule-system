/**
 * =========================================================================
 * 💥 模擬腳本 2：2026-10 部分填報情境與智慧排班驗證 (路徑 A：零密碼依賴版)
 * 特色：
 * 1. 內部自簽管理員憑證 (Internal Elevated Session)，徹底杜絕密碼雜湊不符導致的中斷
 * 2. 模擬 15 人中前 10 位同仁已填報，5 位未填報 (催繳情境)
 * 3. 意願採標準 9 欄聚合矩陣寫入 Availability_Raw (一人一月一列)
 * 4. 執行動態工時平滑排班並刷新 3 張 Viewer 看板
 * =========================================================================
 */
function runSimulate202610_PartialFill() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const targetYm = "2026-10";
  const now = new Date();
  const nowStr = Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss");

  Logger.log(`🚀 開始模擬 ${targetYm} 部分填報情境 (10人已填 / 5人未填)...`);

  const { allStaffList } = getCachedStaffMap();
  const [year, month] = targetYm.split("-").map(Number);
  const daysInMonth = new Date(year, month, 0).getDate();

  const availHeaders = [
    "record_id", "year_month", "employee_id", "pref_work_volume",
    "pref_day_preference", "pref_partner_id", "dislike_partner_id",
    "availability_matrix", "submitted_at"
  ];

  // 1. 讀取活躍表既有資料，剔除 2026-10 舊資料
  const availSheet = ss.getSheetByName("Availability_Raw");
  const aData = availSheet ? availSheet.getDataRange().getValues() : [];
  const keepAvailRows = [availHeaders];

  for (let i = 1; i < aData.length; i++) {
    const ym = normalizeYearMonth(aData[i][1], tz);
    if (ym !== targetYm && ym !== "") {
      keepAvailRows.push(aData[i].slice(0, 9));
    }
  }

  // 2. 模擬前 10 位同仁提交意願 (後 5 位保持空白未申報)
  const submittedStaff = allStaffList.slice(0, 10);
  submittedStaff.forEach((staff, sIdx) => {
    const vol = (sIdx < 3) ? "More" : ((sIdx < 7) ? "Standard" : "Less");
    const dayPref = (sIdx % 2 === 0) ? "Flexible" : "Both_Preferred";
    const partner = (sIdx + 1 < submittedStaff.length) ? submittedStaff[sIdx + 1].employee_id : "";
    const dislike = (sIdx === 0) ? "EMP008" : "";

    const matrix = {};
    for (let d = 1; d <= daysInMonth; d++) {
      const dStr = `${year}-${String(month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      const activeShifts = [];
      if ((d + sIdx) % 4 !== 0) activeShifts.push("SHIFT_LUNCH");
      if ((d + sIdx) % 3 !== 0) activeShifts.push("SHIFT_DINNER");
      if (activeShifts.length > 0) {
        matrix[dStr] = activeShifts;
      }
    }

    keepAvailRows.push([
      `AR_202610_${staff.employee_id}`,
      targetYm,
      staff.employee_id,
      vol,
      dayPref,
      partner,
      dislike,
      JSON.stringify(matrix),
      nowStr
    ]);
  });

  // 原地覆蓋活躍意願表並裁切幽靈空白行
  if (availSheet) {
    inPlaceReplaceSheetData(availSheet, keepAvailRows, 9);
  }

  Logger.log(`✅ 意願寫入完成：共 ${submittedStaff.length} 位同仁已填報，整表該月嚴格維持 10 列！`);

  // 3. 💥 路徑 A：頒發伺服器內部模擬專用 Admin Token (繞過密碼核對)
  const internalToken = "SIM_TOKEN_" + Utilities.getUuid();
  const cache = CacheService.getScriptCache();
  cache.put("TOKEN_AUTH_" + internalToken, JSON.stringify({ employee_id: "MGR001", role: "Admin" }), 3600);

  let tokenSheet = ss.getSheetByName("Session_Tokens");
  if (!tokenSheet) {
    tokenSheet = ss.insertSheet("Session_Tokens");
    tokenSheet.appendRow(["token", "employee_id", "role", "last_active_at", "expires_at"]);
  }
  tokenSheet.appendRow([internalToken, "MGR001", "Admin", now, new Date(now.getTime() + 3600 * 1000)]);

  Logger.log("🤖 執行動態平滑排班演算法 (apiRunAutoSchedule)...");
  const schedRes = apiRunAutoSchedule(internalToken, targetYm);
  Logger.log("排班結果回傳:\n" + schedRes.message);

  // 4. 刷新 Viewer 看板
  refreshAllViewers(targetYm);
  Logger.log("🎉 2026-10 模擬與看板刷新全面完成！");
}