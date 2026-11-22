/**
 * ⚡ [獨立測試模組] 2026-12 高壓模擬意願生成器 (獨立存在，用畢可直刪)
 * 功能：針對 12 月份，生成滿編 5 人、偏好多樣、且具備高人際碰撞的假意願。
 */
function generateMockDecemberAvailability() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Availability_Raw");
  if (!sheet) throw new Error("找不到 Availability_Raw 工作表");
  
  const yearMonth = "2026-12"; 
  const { allStaffList } = getCachedStaffMap();
  const tz = Session.getScriptTimeZone();
  const nowStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
  
  const data = sheet.getDataRange().getValues();
  const keepRows = [data[0]]; // 保留第一列標題
  
  // 1. 隔離防護：無損保留歷史舊資料，僅抹除舊的 12 月模擬數據
  for (let i = 1; i < data.length; i++) {
    const ymStr = normalizeYearMonth(data[i][1], tz);
    if (ymStr !== yearMonth) {
      const r = data[i].slice(0, 9);
      while (r.length < 9) r.push("");
      keepRows.push(r);
    }
  }
  
  const allEmpIds = allStaffList.map(s => s.employee_id);
  
  // 2. 核心迴圈：生成高壓偏好矩陣
  allStaffList.forEach(emp => {
    // 參數池設定
    const volumes = ["Standard", "Standard", "More", "Less"]; 
    const dayPrefs = ["Flexible", "Flexible", "Both_Preferred", "Single_Preferred"];
    const prefWorkVolume = volumes[Math.floor(Math.random() * volumes.length)];
    const prefDayPref = dayPrefs[Math.floor(Math.random() * dayPrefs.length)];
    
    let partner = "";
    let dislike = "";
    
    // 🚨 變更點：30% 機率生成偏好搭班
    if (Math.random() < 0.30) {
       partner = allEmpIds[Math.floor(Math.random() * allEmpIds.length)];
       if (partner === emp.employee_id) partner = ""; // 禁止自戀
    }
    
    // 🚨 變更點：15% 機率生成迴避搭班
    if (Math.random() < 0.15) {
       dislike = allEmpIds[Math.floor(Math.random() * allEmpIds.length)];
       // 邏輯防呆：禁止自我厭惡，且不能與偏好對象重疊 (邏輯悖論)
       if (dislike === emp.employee_id || dislike === partner) dislike = "";
    }
    
    const matrixCompact = {};
    // 核心人員排班存活率 90%，一般人員 75%
    const survivalRate = emp.is_core ? 0.90 : 0.75; 
    
    for (let d = 1; d <= 31; d++) {
      const dateStr = yearMonth + "-" + String(d).padStart(2, '0');
      const activeShifts = [];
      
      if (Math.random() < survivalRate) activeShifts.push("SHIFT_LUNCH");
      if (Math.random() < survivalRate) activeShifts.push("SHIFT_DINNER");
      
      // 極限防呆：非減班意願者，絕不允許當天出現雙空班，強制塞入一班以防缺工
      if (activeShifts.length === 0 && prefWorkVolume !== "Less") {
        activeShifts.push(Math.random() > 0.5 ? "SHIFT_LUNCH" : "SHIFT_DINNER");
      }
      
      if (activeShifts.length > 0) {
        matrixCompact[dateStr] = activeShifts;
      }
    }
    
    const recId = "AR_202612_" + emp.employee_id;
    keepRows.push([
      recId, 
      yearMonth, 
      emp.employee_id, 
      prefWorkVolume, 
      prefDayPref, 
      partner, 
      dislike, 
      JSON.stringify(matrixCompact), 
      nowStr
    ]);
  });
  
  // 3. 實體寫回資料庫
  inPlaceReplaceSheetData(sheet, keepRows, 9);
}