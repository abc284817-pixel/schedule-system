/**
 * 一鍵更新門市班別設定與人力需求檔 (Shift_Setting)
 * 規格：每班 5 人出勤，其中 1 人為核心人力
 */
function seedShiftSetting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Shift_Setting");

  const headers = [
    "shift_id", "shift_name", "start_time", "end_time", 
    "duration_hours", "required_staff", "required_core", "enable_standby"
  ];

  if (!sheet) {
    sheet = ss.insertSheet("Shift_Setting");
  } else {
    sheet.clear();
  }

  // 規格：required_staff = 5, required_core = 1
  const shiftData = [
    headers,
    ["SHIFT_LUNCH", "午班", "10:30", "14:30", 4.0, 5, 1, true],
    ["SHIFT_DINNER", "晚班", "17:30", "21:30", 4.0, 5, 1, true]
  ];

  sheet.getRange(1, 1, shiftData.length, headers.length).setValues(shiftData);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
       .setBackground("#34495e")
       .setFontColor("#ffffff")
       .setFontWeight("bold");

  Logger.log("✅ Shift_Setting 已成功更新為：每班 5 人出勤（含 1 位核心人力）！");
}

/**
 * 初始化 15 位門市員工帳號資料 (預設密碼 123456)
 */
function seedStaffMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Staff_Master");
  
  const headers = [
    "employee_id", "password_salt", "password_hash", 
    "name", "role", "is_core", "target_hours_pref", "is_active"
  ];

  if (!sheet) {
    sheet = ss.insertSheet("Staff_Master");
  } else {
    sheet.clear();
  }
  sheet.appendRow(headers);

  const defaultPassword = "123456";

  const staffConfig = [
    { id: "EMP001", name: "蕭富駿", role: "Admin", isCore: true, pref: "Standard" },
    { id: "EMP002", name: "陳雅婷", role: "Staff", isCore: true, pref: "Standard" },
    { id: "EMP003", name: "林冠宇", role: "Staff", isCore: true, pref: "More" },
    { id: "EMP004", name: "張瑋倫", role: "Staff", isCore: true, pref: "More" },
    { id: "EMP005", name: "黃詩涵", role: "Staff", isCore: false, pref: "Standard" },
    { id: "EMP006", name: "吳柏毅", role: "Staff", isCore: false, pref: "More" },
    { id: "EMP007", name: "蔡佩君", role: "Staff", isCore: false, pref: "Standard" },
    { id: "EMP008", name: "劉承翰", role: "Staff", isCore: false, pref: "Less" },
    { id: "EMP009", name: "鄭宇翔", role: "Staff", isCore: false, pref: "More" },
    { id: "EMP010", name: "謝佳蓉", role: "Staff", isCore: false, pref: "Standard" },
    { id: "EMP011", name: "賴俊豪", role: "Staff", isCore: false, pref: "Less" },
    { id: "EMP012", name: "許家豪", role: "Staff", isCore: false, pref: "More" },
    { id: "EMP013", name: "楊千又", role: "Staff", isCore: false, pref: "Standard" },
    { id: "EMP014", name: "孫詠潔", role: "Staff", isCore: false, pref: "Less" },
    { id: "EMP015", name: "趙哲偉", role: "Staff", isCore: false, pref: "Standard" }
  ];

  const rows = staffConfig.map(staff => {
    const salt = Utilities.getUuid().substring(0, 8);
    const hash = computeSHA256(defaultPassword + salt);
    return [
      staff.id,
      salt,
      hash,
      staff.name,
      staff.role,
      staff.isCore,
      staff.pref,
      true
    ];
  });

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setBackground("#34495e").setFontColor("#ffffff").setFontWeight("bold");

  Logger.log(`✅ 成功建立 ${rows.length} 位員工種子資料！預設密碼全數為 ${defaultPassword}`);
}

/**
 * SHA-256 加密工具
 */
function computeSHA256(input) {
  const rawHash = Maps.newDirectionFinder ? Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8) : [];
  return rawHash.map(byte => {
    const v = (byte < 0 ? byte + 256 : byte).toString(16);
    return v.length === 1 ? "0" + v : v;
  }).join("");
}

/**
 * 建立店長雙帳號體系：
 * 1. 將 EMP001 的權限設定為一般員工 (Staff)，僅具備員工視角
 * 2. 新增專屬店長帳號 MGR001 (Admin)，具備完整管理權限
 */
function setupDualManagerAccount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffSheet = ss.getSheetByName("Staff_Master");
  if (!staffSheet) throw new Error("找不到 Staff_Master 表格！");

  const data = staffSheet.getDataRange().getValues();
  let emp001Found = false;
  let mgr001Found = false;

  for (let i = 1; i < data.length; i++) {
    // 1. 將 EMP001 改為一般同仁 Staff 模式
    if (data[i][0] === "EMP001") {
      staffSheet.getRange(i + 1, 5).setValue("Staff");
      emp001Found = true;
    }
    if (data[i][0] === "MGR001") {
      mgr001Found = true;
    }
  }

  // 2. 新增 MGR001 專屬店長管理帳號 (預設密碼 123456)
  if (!mgr001Found) {
    const salt = Utilities.getUuid().substring(0, 8);
    const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, "123456" + salt, Utilities.Charset.UTF_8);
    const hash = rawHash.map(byte => {
      const v = (byte < 0 ? byte + 256 : byte).toString(16);
      return v.length === 1 ? "0" + v : v;
    }).join("");

    staffSheet.appendRow([
      "MGR001",
      salt,
      hash,
      "蕭富駿 (主管端)",
      "Admin",
      true,
      "Standard",
      true,
      new Date()
    ]);
  }

  Logger.log("🎉 主管雙帳號配置完成！\n- EMP001：員工模式 (Staff)\n- MGR001：店長模式 (Admin，密碼 123456)");
}