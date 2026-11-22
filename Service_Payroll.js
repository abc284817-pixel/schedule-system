/**
 * [API] 全量重算月工時統計（僅限店長）
 * 前端呼叫入口：驗證登入憑證與 Admin 角色，並以 ScriptLock 防止與排班 / 調度同時寫入。
 */
function apiRecalculateMonthlyStats(token) {
  const user = validateToken(token);
  if (!user || user.role !== 'Admin') {
    return { success: false, message: "權限不足：僅店長可執行工時結算！" };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    return recalculateAllMonthlyStats_();
  } catch (e) {
    return { success: false, message: "工時結算失敗：" + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * [全新計薪引擎 2.0] 全量對帳重建與雙薪分流
 * 1. 跨表讀取 (Schedule_Final + Schedule_Final_Archive)
 * 2. 嚴格比對底層原生英文 Enum ("Assigned", "Standby")
 * 3. 精準隔離 Standby Hours，確保 total_hours 不被污染
 *
 * 函式名稱以底線結尾 = Apps Script 私有函式，前端無法透過 google.script.run 直接呼叫。
 */
function recalculateAllMonthlyStats_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statsSheet = ss.getSheetByName('Staff_Monthly_Stats');
  
  if (!statsSheet) throw new Error("系統致命錯誤：找不到 Staff_Monthly_Stats 工作表");

  // 1. 抓取動態節日字典
  const holidaysMap = apiGetDynamicHolidays(); 
  
  // 2. 建立跨表 (Hot + Cold) 資料陣列
  let allScheduleData = [];
  
  // 讀取熱區 (當前班表)
  const hotSheet = ss.getSheetByName('Schedule_Final');
  if (hotSheet && hotSheet.getLastRow() > 1) {
    allScheduleData = allScheduleData.concat(hotSheet.getDataRange().getValues().slice(1));
  }
  
  // 讀取冷區 (歷史封存區)
  const coldSheet = ss.getSheetByName('Schedule_Final_Archive');
  if (coldSheet && coldSheet.getLastRow() > 1) {
    allScheduleData = allScheduleData.concat(coldSheet.getDataRange().getValues().slice(1));
  }
  
  if (allScheduleData.length === 0) return { success: true, count: 0, message: "無任何班表資料" };
  
  // 3. 準備聚合容器
  const aggregation = {};

  allScheduleData.forEach(row => {
    // [0]id, [1]date, [2]shift_id, [3]emp_id, [4]is_core, [5]dispatch_type, [6]status
    const dateObj = row[1];
    const empId = String(row[3]).trim();
    // 嚴格比對底層原生英文 Enum，避免中文殘留造成統計脫軌 (參照 Node-003)
    const dispatchType = String(row[5]).trim(); 
    const status = String(row[6]).trim();
    
    // 防呆：略過空白無效列，且只計算狀態為 'Confirmed' 者
    if (!dateObj || !empId || status !== 'Confirmed') return;

    const dateStr = Utilities.formatDate(new Date(dateObj), "GMT+8", "yyyy-MM-dd");
    const ymStr = dateStr.substring(0, 7);
    const key = `${ymStr}_${empId}`;
    
    if (!aggregation[key]) {
      aggregation[key] = { assigned_shifts: 0, standby_shifts: 0, normal_hours: 0, double_hours: 0, standby_hours: 0 };
    }
    
    const hours = 4.0;
    
    // 4. 水管分流：嚴格隔離備援與正班
    if (dispatchType === 'Standby') {
      aggregation[key].standby_shifts += 1;
      aggregation[key].standby_hours += hours; 
      // ⚠️ 備援工時獨立存放，絕對不干涉一般與雙倍薪資池
      
    } else if (dispatchType === 'Assigned') {
      aggregation[key].assigned_shifts += 1;
      
      const isDouble = (holidaysMap[dateStr] && holidaysMap[dateStr].isDoublePay === true);
      if (isDouble) {
        aggregation[key].double_hours += hours;
      } else {
        aggregation[key].normal_hours += hours;
      }
    }
  });

  // 5. 轉換陣列與複合主鍵校準 (參照 Node-001)
  const now = new Date();
  const rowsToWrite = [];
  
  for (const key in aggregation) {
    const [ym, empId] = key.split('_');
    const data = aggregation[key];
    
    const totalShifts = data.assigned_shifts + data.standby_shifts;
    
    // 🚨 致命修正：total_hours 僅包含正班時數，絕不包含 standby_hours
    const totalHours = data.normal_hours + data.double_hours; 
    
    // 排除雙零（整個月既無正班也無備援）的無效資料
    if (totalShifts > 0) {
      rowsToWrite.push([
        `STAT_${ym.replace("-", "")}_${empId}`, // stat_id
        ym,                                     // year_month
        empId,                                  // employee_id
        totalShifts,                            // total_shifts
        totalHours,                             // total_hours (僅正班)
        data.normal_hours,                      // normal_hours 
        data.double_hours,                      // double_pay_hours 
        data.standby_hours,                     // standby_hours (獨立統計)
        now                                     // updated_at
      ]);
    }
  }
  
  // 6. 寫回資料庫 (先清空髒資料，再記憶體雙重升冪排序覆寫)
  if (rowsToWrite.length > 0) {
    // 排序：月份優先，員工次之
    rowsToWrite.sort((a, b) => a[1].localeCompare(b[1]) || a[2].localeCompare(b[2]));
    
    const oldMaxRows = statsSheet.getMaxRows();
    if (oldMaxRows > 1) {
      statsSheet.getRange(2, 1, oldMaxRows - 1, 9).clearContent();
    }
    statsSheet.getRange(2, 1, rowsToWrite.length, 9).setValues(rowsToWrite);
  }
  
  return { success: true, count: rowsToWrite.length };
}