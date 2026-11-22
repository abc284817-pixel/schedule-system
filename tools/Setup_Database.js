/**
 * tools/Setup_Database.js — 一次性建置 / 種子資料（只在 Apps Script 編輯器手動執行，網頁不會呼叫）
 * ⚠️ setupEntireDatabaseSystem() 會清空所有工作表，正式環境請勿執行。
 * 表頭一律取自 00_Config.js 的 CONFIG.SHEETS，不在此重複定義。
 */
function setupEntireDatabaseSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(CONFIG.SHEETS).forEach(key => {
    if (key === "HOLIDAYS") return; // 節日表由人工維護，不清空
    const spec = CONFIG.SHEETS[key];
    let sheet = ss.getSheetByName(spec.name);
    if (!sheet) sheet = ss.insertSheet(spec.name); else sheet.clear();
    sheet.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers])
      .setBackground(spec.color).setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  });
  seedShiftSetting();
  setupMonthlyViewers(ss);
  refreshAllViewers(normalizeYearMonth(new Date()));
  Logger.log("✅ 系統資料庫與 3 張看板初始化完成！");
}

/** 建立 3 張檢視看板外觀 */
function setupMonthlyViewers(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const ym = normalizeYearMonth(new Date());
  const specs = [
    [VIEWERS.SCHED, 0, "#2980b9", ["出勤日期", "班別代碼", "出勤同仁代號", "核心人力覆蓋", "排定指派型態", "排班狀態"]],
    [VIEWERS.AVAIL, 1, "#16a085", ["填報月份", "員工代號", "同仁姓名", "出勤日期", "班別代碼", "可出勤", "工時傾向", "單日排班偏好", "偏好搭班同仁", "迴避同仁", "填報時間戳記"]],
    [VIEWERS.LEAVES, 2, "#e67e22", ["申請單號", "請假同仁", "請假日期", "請假班別", "請假假別", "詳細事由", "自覓代班", "指派代班人", "審批狀態", "送單時間"]]
  ];
  specs.forEach(([name, idx, color, headers]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name, idx); else sh.clear();
    sh.getRange("A1").setValue("檢視月份 (YYYY-MM)：").setFontWeight("bold");
    sh.getRange("B1").setNumberFormat("@").setValue(ym).setBackground("#fff2cc").setFontWeight("bold");
    sh.getRange(3, 1, 1, headers.length).setValues([headers]).setBackground(color).setFontColor("#ffffff").setFontWeight("bold");
    sh.setFrozenRows(3);
  });
}

/** 班別設定：每班 5 人、1 位核心（目前排班引擎讀的是 CONFIG.RULES，此表僅供人工參考） */
function seedShiftSetting() {
  const sheet = getOrCreateSheet_("SHIFT_SETTING");
  const rows = [
    ["SHIFT_LUNCH", "午班", "10:30", "14:30", CONFIG.RULES.SHIFT_HOURS, CONFIG.RULES.STAFF_PER_SHIFT, 1, true],
    ["SHIFT_DINNER", "晚班", "17:30", "21:30", CONFIG.RULES.SHIFT_HOURS, CONFIG.RULES.STAFF_PER_SHIFT, 1, true]
  ];
  writeAll_("SHIFT_SETTING", sheet, rows);
}

/** 依 CONFIG.SHEETS.STAFF 欄位順序組一列員工資料 */
function buildStaffRow_(id, name, role, isCore, password) {
  const C = COL.STAFF;
  const row = new Array(CONFIG.SHEETS.STAFF.headers.length).fill("");
  const salt = Utilities.getUuid().substring(0, 8);
  row[C.employee_id] = id;
  row[C.password_salt] = salt;
  row[C.password_hash] = computeSHA256(password + salt);
  row[C.name] = name;
  row[C.role] = role;
  row[C.is_core] = isCore;
  row[C.is_active] = true;
  return row;
}

/** 15 位模擬員工 + MGR001 店長帳號（預設密碼 123456，僅供測試） */
function seedStaffMaster() {
  const pw = "123456";
  const staff = [
    ["EMP001", "蕭富駿", true], ["EMP002", "陳雅婷", true], ["EMP003", "林冠宇", true], ["EMP004", "張瑋倫", true],
    ["EMP005", "黃詩涵", false], ["EMP006", "吳柏毅", false], ["EMP007", "蔡佩君", false], ["EMP008", "劉承翰", false],
    ["EMP009", "鄭宇翔", false], ["EMP010", "謝佳蓉", false], ["EMP011", "賴俊豪", false], ["EMP012", "許家豪", false],
    ["EMP013", "楊千又", false], ["EMP014", "孫詠潔", false], ["EMP015", "趙哲偉", false]
  ];
  const rows = staff.map(([id, name, core]) => buildStaffRow_(id, name, "Staff", core, pw));
  rows.push(buildStaffRow_("MGR001", "蕭富駿 (主管端)", CONFIG.ROLE_ADMIN, true, pw));
  const sheet = getOrCreateSheet_("STAFF");
  writeAll_("STAFF", sheet, rows);
  CacheService.getScriptCache().remove(CONFIG.CACHE.STAFF);
  Logger.log(`✅ 已建立 ${rows.length} 位員工種子資料，預設密碼 ${pw}`);
}
