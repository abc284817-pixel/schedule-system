# AGENTS.md — schedule-system 修改規範（所有 AI 模型共用）

> 任何模型修改本專案前先讀本檔。目標：**不影響使用者體驗與 UI**、**低耦合**、**最少 token**、**不讓系統變肥**。

## 0. 最高原則（依優先順序）

1. **UI / 使用體驗不變**：除非需求就是改畫面，否則不動 `view_*.html` 的結構、class、文字。
2. **只讀需要的檔**：先查 §2 讀檔地圖，用 `grep -n` 定位後局部讀取，不要整包讀。
3. **最小改動**：用精準取代，不重寫整檔；不留死碼、不加「相容用」別名、不加 emoji 長註解。
4. **不加依賴**：不新增框架、建置流程、npm 套件（`dev/` 測試工具除外）。
5. **改完必驗證**：§7 的指令全部通過才算完成。

## 1. 架構

```
瀏覽器 (Vue3 + Tailwind CDN) ──google.script.run──▶ api*()（Apps Script V8）──▶ Google 試算表
```

後端（根目錄 `*.js`，Apps Script 所有檔案共用同一個全域命名空間）：

| 層 | 檔案 | 職責 |
|---|---|---|
| 設定 | `00_Config.js` | **唯一設定來源**：工作表名稱/欄位、班別、法規數值、衝突類型、代理帳號。`COL` 由表頭自動產生 |
| 工具 | `01_Util.js` | 日期正規化、讀寫表（`getSheet_` `getOrCreateSheet_` `readRows_` `writeAll_`）、`withLock_`、`isAdmin_`、`effectiveEmpId_` |
| 領域 | `10_Auth` `20_Availability` `30_Scheduler` `31_ScheduleAdmin` `32_Audit` `40_Leave` `50_Payroll` `60_Archive` `70_Viewers` | 各自的 `api*` 與私有函式 |
| 連動 | `31_ScheduleAdmin.js` → `afterScheduleChanged_()` | 班表任何異動後**唯一**的連動入口（稽核 + 月統計 + 看板） |
| 工具腳本 | `tools/` | 只在 Apps Script 編輯器手動執行（建表、種子、模擬、稽核測試），網頁不會呼叫 |
| 本機測試 | `dev/` | Node 回歸比對工具，`.claspignore` 已排除，不會推上 GAS |

前端（`index.html` 以 `<?!= include('x') ?>` 組裝；`include` 走樣板引擎，**支援巢狀 include 與 scriptlet**）：

| 層 | 檔案 | 職責 |
|---|---|---|
| 設定 | `js_constants.html` | `APP_CONFIG`（由後端 `getClientConfig()` 注入，前後端共用數值）、`ShiftUtils` |
| 純工具 | `js_calendar_utils.html` | `CalendarUtils`：`today` `addMonths` `monthOption` `normalizeDateStr` `normalizeYearMonth` `buildMonthCalendarGrid` |
| 共用狀態 | `js_core.html` | token、currentUser、`myEmployeeId`、`colleagueList`、提示框、時鐘 |
| 功能模組 | `js_availability` `js_my_roster` `js_leave` `js_manager_summary` `js_schedule_admin` | `useXxxModule(core)`，**只能依賴 core** |
| 組裝 | `js_main.html` | 建立 Vue app、登入/登出、跨模組流程、模板用別名（`handleXxx`） |
| 模板 | `view_*.html` | 純模板；`view_schedule_admin` 拆成 本體 / `_calendar` / `_table` 三檔 |

## 2. 讀檔地圖（依任務只讀這些）

| 任務 | 讀 |
|---|---|
| 改欄位、法規數值、班別、衝突文字前綴 | `00_Config.js` |
| 排班權重 / 演算法 | `30_Scheduler.js` 的 `solveMonthSchedule_`（純函式） |
| 法規稽核 / 衝突 | `32_Audit.js` |
| 工時 / 雙薪 / 節日 | `50_Payroll.js` |
| 登入 / 權限 | `10_Auth.js`（前端：`js_main.html` 的 `handleLogin`） |
| 意願填報 / 解鎖 / 填報總覽 | `20_Availability.js` + `js_availability.html` 或 `js_manager_summary.html` + 對應 view |
| 請假 / 我的班表 | `40_Leave.js` + `js_leave.html` / `js_my_roster.html` + 對應 view |
| 店長班表頁 | `31_ScheduleAdmin.js` + `js_schedule_admin.html` + 需要的那個 `view_schedule_admin*.html` |
| 封存 | `60_Archive.js`；試算表看板：`70_Viewers.js` |
| **不要讀** | `tools/`（除非任務就是它）、`dev/`（直接執行即可）、`README.md` |

## 3. API 對照表（前端呼叫 → 後端）

| 前端（呼叫處） | 後端 API | 檔案 | 權限 |
|---|---|---|---|
| `handleLogin` (js_main) | `apiLogin(id, pw)` | 10_Auth | 公開 |
| `loadBootstrapData` (js_main) | `apiGetBootstrapData(token, ym)` | 20_Availability | 登入 |
| `loadAvailabilityOnly` (avail) | `apiGetAvailabilityForMonth(token, ym)` | 20_Availability | 登入 |
| `submitAvailability` (avail) | `apiSubmitAvailability(token, payload)` | 20_Availability | 登入 |
| `executeUnlockMonth` (schedAdmin) | `apiSetUnlockFillPermission(token, ym, hours, empId)` | 20_Availability | Admin |
| `fetchManagerSummary` (summary) | `apiGetManagerSummary(token, ym)` | 20_Availability | Admin |
| `runAutoScheduleRequest` (js_main) | `apiRunAutoSchedule(token, ym, confirmed)` | 30_Scheduler | Admin |
| `fetchScheduleData` (schedAdmin) | `apiGetScheduleTable(token, ym)` | 31_ScheduleAdmin | Admin |
| `commitAllChanges` (schedAdmin) | `apiBatchCommitScheduleModifications(token, ym, upd, add, del)` | 31_ScheduleAdmin | Admin |
| `loadRosterForMonth` (js_main) | `apiGetMyRosterSchedule(token, ym)` | 40_Leave | 登入 |
| `fetchMyShiftsForMonth` (leave) | `apiGetMyShiftsForMonth(token, ym)` | 40_Leave | 登入 |
| `submitLeave` (leave) | `apiSubmitLeaveRequest(token, payload)` | 40_Leave | 登入 |
| `fetchMyLeaves` (leave) | `apiGetMyLeaveRequests(token)` | 40_Leave | 登入 |
| `approveLeave` / `rejectLeave` (schedAdmin) | `apiApproveLeaveRequest` / `apiRejectLeaveRequest` | 40_Leave | Admin |
| `handleRecalculateStats` (schedAdmin) | `apiRecalculateMonthlyStats(token)` | 50_Payroll | Admin |
| `fetchDynamicHolidays` (js_main) | `apiGetDynamicHolidays()` | 50_Payroll | 公開 |
| `executeArchiveFromSchedule` (schedAdmin) | `apiArchiveHistoricalData(token, cutoffYm)` | 60_Archive | Admin |

新增 / 刪除 API 時同步更新此表。

## 4. 後端寫法規範

- **欄位**：一律 `COL.<表>.<欄名>`（例：`r[COL.SCHED.employee_id]`）。**禁止** `row[3]` 這類數字索引。
- **工作表**：用 `getSheet_("SCHED")` / `getOrCreateSheet_` / `readRows_`（不含表頭）/ `writeAll_`（自動帶 CONFIG 表頭）。禁止散寫表名字串與 `appendRow(表頭)`。
- **數值**：用 `CONFIG.RULES.*`、`CONFIG.DISPATCH.*`、`CONFIG.STATUS.*`；不得寫 5、4.0、160、"Assigned" 等字面值。
- **權限**：開頭 `const user = validateToken(token);`，管理功能用 `if (!isAdmin_(user)) return { success:false, message:"權限不足" }`。
- **員工身分**：一律 `effectiveEmpId_(user)`；**禁止**寫死 `"MGR001"` / `"EMP001"`（代理帳號在 `CONFIG.PROXY_ACCOUNTS`）。
- **讀-改-寫**：包在 `withLock_(CONFIG.LOCK_MS.X, "錯誤前綴：", () => {...})`，callback 回傳 API 結果。
- **班表異動後**：只呼叫 `afterScheduleChanged_(ss, ym, tz, refreshViewers)`，不要自己串 `sync*`。
- **回傳格式**：`{ success: boolean, message?: string, ...資料 }`；可預期的錯誤用回傳，不要 throw。
- **命名**：`api*` = 前端可呼叫；結尾 `_` = 私有（`google.script.run` 無法呼叫）。全專案（含 `tools/`）**函式與頂層 const 名稱必須唯一**——同名時後載入者會靜默覆蓋。
- **頂層程式碼**：只允許函式宣告與 `const` 常數；不得在載入時呼叫其他檔案的函式。
- **日期**：`normalizeDateStr` / `normalizeYearMonth` / `ymParts_` / `dateStrOf_`；時區用 `getTz_()`。
- **演算法**：保持純函式（輸入資料 → 輸出列），不得在裡面呼叫 `SpreadsheetApp`。
- **衝突文字**：以 `conflictText_(類型, 內容)` 產生；前端只看 `is_legal` 欄位。

## 5. 前端寫法規範

- 模組簽名固定 `function useXxxModule(core)`；**模組之間不可互相引用**，需要協調的流程寫在 `js_main.html`。
- 多模組共用的資料放 `js_core.html`（例：`colleagueList`、`myEmployeeId`）。
- 各模組 `return` 的名稱**全域唯一**（`js_main` 會展開合併；重複時 console 會出現「架構警告」）。
- 數值與班別用 `APP_CONFIG.RULES` / `ShiftUtils`；年月計算用 `CalendarUtils`；不要再各自 `new Date()` 拆年月。
- `google.script.run` 一律同時寫 `withSuccessHandler` 與 `withFailureHandler`；錯誤顯示用 `core.triggerStickyAlert(msg, "error")`。
- 邏輯放 `js_*.html`，`view_*.html` 只綁定；不得在模板或 JS 用中文字串比對來判斷邏輯。
- 單一 view 檔超過約 20KB 時，依區塊拆成 `view_xxx_part.html` 並以 `<?!= include('view_xxx_part') ?>` 引入（純搬移，不改內容）。
- 不改 `index.html` 的 CDN 來源與版本。

## 6. 常見修改 SOP

- **改法規 / 每班人數 / 工時目標**：只改 `CONFIG.RULES`（前端經 `APP_CONFIG` 自動同步；文字說明若寫死數字再另行調整 view）。
- **表格新增欄位**：在 `CONFIG.SHEETS.X.headers` **最後面**追加（不可插在中間，既有試算表資料不會自動搬移），之後用 `COL.X.新欄位` 讀寫，並在試算表手動補上表頭。
- **新增 API**：放進對應領域檔 → 依 §4 開頭寫權限檢查 → 前端在對應模組呼叫 → 更新 §3 表 → 在 `dev/backend_diff.js` 的 `runScenario` 補一行 `call(...)`。
- **新增分頁**：`view_x.html` + `js_x.html`（`useXModule(core)`）→ `index.html` 加 include（view 與 js 各一處）與導覽按鈕（桌機列 + 手機抽屜兩處）→ `js_main.html` 建立模組並展開到 return。
- **新增衝突類型**：`CONFIG.CONFLICT` 加一項（`prefix`、`legal`）→ 在 `32_Audit.js` 用 `add(..., conflictText_("新類型", ...))`。
- **新增班別**：`CONFIG.SHIFTS` 加一項即可讓後端運作；但 view 內有 `'SHIFT_LUNCH'` 字面值決定顏色/名稱（`view_availability`、`view_my_roster`、`view_leave`、`view_schedule_admin*`），需逐一補上並跑 §7 UI 比對。
- **刪除功能**：同步刪 API、前端呼叫、§3 表列；不要留「以防萬一」的舊函式。

## 7. 驗證（改完必做，於 repo 根目錄）

```bash
# 1) 語法
for f in *.js tools/*.js; do node --check "$f" || echo "FAIL $f"; done
# 2) 全域重名（應無輸出）
grep -h "^function \|^const " *.js tools/*.js | sed -E 's/^(function|const) ([A-Za-z0-9_$]+).*/\2/' | sort | uniq -d
# 3) 後端行為回歸：目前修改 vs 上一個 commit（預期 issues=0；刻意變更只應出現在你改的 API）
node dev/backend_diff.js HEAD
# 4) 有動 html / js_*.html 時：前端畫面回歸（需 Chrome；首次先 npm i --prefix dev）
node dev/ui_diff.js HEAD
```

情境沒覆蓋到你的新功能時，先在 `dev/backend_diff.js` 補 `call(...)` 或在 `dev/ui_diff.js` 補操作步驟再比對。

## 8. 部署（clasp）

- `.claspignore` 已排除 `dev/` 與 `*.md`；`tools/` 會推上去（供編輯器手動執行）。
- `appsscript.json`（時區、V8、網頁應用程式權限）已在 repo，修改專案設定時同步更新此檔。**不要在 repo 內直接 `clasp pull`**，會把遠端已刪除的舊檔（如 `Api.js`）拉回來造成函式重複。
- `clasp push` 會以本機檔案完整取代遠端專案，再到「部署 → 管理部署 → 編輯 → 新版本」發布（網址不變）。

## 9. 已知技術債（非需求不要順手改）

- `Session_Tokens` 只增不減，快取失效時會全表掃描。
- `inPlaceReplaceSheetData` 每次覆寫整張表；資料量大時變慢（定期封存可緩解）。
- `Shift_Setting` 表未被程式讀取，以 `CONFIG` 為準。
- 意願鎖定規則前端（顯示）與後端（把關）各一份，共用 `FILL_LOCK_DAYS_BEFORE_MONTH_END`。
- `apiGetDynamicHolidays` 不驗 token（節日為公開資料）；密碼為 SHA-256 加鹽。

## 10. 省 token 守則

- 先 §2 決定檔案 → `grep -n` 找函式 → 只讀該段。
- 一次修改只碰必要的檔；回報只寫「改了什麼 + 驗證結果」。
- 不重貼整檔、不複製既有函式改名、不新增說明性 md（規範只維護本檔）。
