/**
 * 00_Config.js — 全系統唯一設定來源（Single Source of Truth）
 * 規則：工作表名稱、欄位順序、班別、法規數值只能在這裡定義；其他檔案一律引用 CONFIG / COL。
 * 改欄位 = 只改 SHEETS[x].headers，欄位索引 COL 會自動跟著變。
 */
const CONFIG = Object.freeze({
  SHEETS: {
    STAFF: { name: "Staff_Master", color: "#34495e", headers: ["employee_id", "password_salt", "password_hash", "name", "role", "is_core", "is_active"] },
    SHIFT_SETTING: { name: "Shift_Setting", color: "#34495e", headers: ["shift_id", "shift_name", "start_time", "end_time", "duration_hours", "required_staff", "required_core", "enable_standby"] },
    AVAIL: { name: "Availability_Raw", color: "#0f766e", headers: ["record_id", "year_month", "employee_id", "pref_work_volume", "pref_day_preference", "pref_partner_id", "dislike_partner_id", "availability_matrix", "submitted_at"] },
    AVAIL_ARCHIVE: { name: "Availability_Archive", color: "#1e293b", headers: ["record_id", "year_month", "employee_id", "pref_work_volume", "pref_day_preference", "pref_partner_id", "dislike_partner_id", "availability_matrix", "submitted_at", "archived_at"] },
    SCHED: { name: "Schedule_Final", color: "#34495e", headers: ["schedule_id", "date", "shift_id", "employee_id", "is_core_at_work", "dispatch_type", "schedule_status", "created_at"] },
    SCHED_ARCHIVE: { name: "Schedule_Final_Archive", color: "#1e293b", headers: ["schedule_id", "date", "shift_id", "employee_id", "is_core_at_work", "dispatch_type", "schedule_status", "created_at", "archived_at"] },
    CONFLICTS: { name: "Schedule_Conflicts", color: "#c0392b", headers: ["conflict_id", "year_month", "date", "shift_id", "reason", "created_at"] },
    CONFLICTS_AUDIT: { name: "Schedule_Conflicts_Audit", color: "#1e293b", headers: ["audit_id", "year_month", "date", "shift_id", "reason", "archived_at"] },
    LEAVE: { name: "Leave_Requests", color: "#d35400", headers: ["request_id", "employee_id", "date", "shift_id", "leave_type", "reason_detail", "has_substitute", "substitute_id", "status", "created_at"] },
    STATS: { name: "Staff_Monthly_Stats", color: "#27ae60", headers: ["stat_id", "year_month", "employee_id", "total_shifts", "total_hours", "normal_hours", "double_pay_hours", "standby_hours", "updated_at"] },
    TOKENS: { name: "Session_Tokens", color: "#34495e", headers: ["token", "employee_id", "role", "last_active_at", "expires_at"] },
    HOLIDAYS: { name: "Config_Holidays", color: "#34495e", headers: ["date", "holiday_name", "is_double_pay", "holiday_type"] }
  },

  // 班別：順序即顯示順序（午班在前）
  SHIFTS: [
    { id: "SHIFT_LUNCH", name: "午班" },
    { id: "SHIFT_DINNER", name: "晚班" }
  ],

  RULES: {
    STAFF_PER_SHIFT: 5,          // 每班正式人數
    SHIFT_HOURS: 4.0,            // 每班工時
    MAX_HOURS_28D: 160.0,        // 28 天滾動工時上限
    MAX_WORK_DAYS_28D: 20,       // 28 天內最多出勤天數（例休 ≥ 8 天）
    MAX_CONSECUTIVE_DAYS: 6,     // 七休一
    ROLLING_WINDOW_DAYS: 28,
    HOURS_GOAL: { More: 152.0, Standard: 132.0, Less: 88.0 },
    FILL_LOCK_DAYS_BEFORE_MONTH_END: 2, // 下月意願於本月最後 N+1 天起鎖定
    TOKEN_TTL_HOURS: 12,
    HISTORY_START_YM: "2024-02"   // 前端月份選單起點
  },

  // 代理帳號：店長帳號 MGR001 以 EMP001 身分填意願 / 看自己的班
  PROXY_ACCOUNTS: { MGR001: "EMP001" },

  DISPATCH: { ASSIGNED: "Assigned", STANDBY: "Standby" },
  STATUS: { CONFIRMED: "Confirmed", PENDING: "Pending", APPROVED: "Approved", REJECTED: "Rejected" },
  ROLE_ADMIN: "Admin",

  // 衝突原因前綴：產生文字與分類都用這裡，前端不再比對中文字串
  CONFLICT: {
    UNDERSTAFF: { prefix: "出勤編制不足", legal: false },
    NO_CORE: { prefix: "重大營運警示", legal: false },
    CONSECUTIVE: { prefix: "七休一違規", legal: true },
    OVER_HOURS: { prefix: "工時超標違規", legal: true },
    REST_DAYS: { prefix: "四週例休不足", legal: true }
  },

  CACHE: { STAFF: "CACHE_STAFF_MAP_REAL_7COL_FINAL", TOKEN_PREFIX: "TOKEN_AUTH_", TTL_SEC: 21600 },
  PROP_UNLOCK_PREFIX: "UNLOCK_REG_",
  LOCK_MS: { SHORT: 10000, NORMAL: 15000, ARCHIVE: 25000, LONG: 30000 }
});

/** 欄位索引：COL.SCHED.employee_id === 3。由 headers 自動產生，禁止手寫數字索引。 */
const COL = (function () {
  const out = {};
  Object.keys(CONFIG.SHEETS).forEach(key => {
    const map = {};
    CONFIG.SHEETS[key].headers.forEach((h, i) => { map[h] = i; });
    out[key] = Object.freeze(map);
  });
  return Object.freeze(out);
})();

/** 給前端的設定（透過 js_constants.html 注入，前後端共用同一組數值） */
function getClientConfig() {
  return {
    SHIFTS: CONFIG.SHIFTS,
    RULES: CONFIG.RULES,
    DISPATCH: CONFIG.DISPATCH
  };
}
