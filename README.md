# 餐飲門市智慧排班系統

以 Google Apps Script 建置的門市排班網頁系統，後端資料存放在 Google 試算表，前端使用 Vue 3 + Tailwind CSS。

## 功能

**員工端**
- 登入（員工代碼 + 密碼，SHA-256 加鹽驗證）
- 下月出勤意願填報：午班 / 晚班可出勤日期、工時傾向、單日偏好、偏好 / 迴避搭班同仁
- 我的排班班表（清單 / 月曆檢視）
- 請假 / 代班申請

**店長端**
- 全店意願填報總覽
- 一鍵智慧排班（每班 5 人、至少 1 位核心人力，含 Standby 備援）
- 手動調度（草稿 → 批次寫回）、請假審核、臨時解鎖意願填報
- 法規稽核：七休一、四週 28 天滾動例休（≥ 8 天）、28 天滾動 160 小時上限、缺工與無核心人力警示
- 月工時統計（正常 / 雙倍薪 / 備援工時，節日讀取 `Config_Holidays`）
- 歷史資料封存（Archive 冷儲存）

## 檔案結構

架構、檔案職責、API 對照與修改規範請見 **[AGENTS.md](AGENTS.md)**。

| 類型 | 檔案 |
|---|---|
| 設定 / 工具 | `00_Config.js`（唯一設定來源）、`01_Util.js` |
| 後端領域 | `10_Auth` `20_Availability` `30_Scheduler` `31_ScheduleAdmin` `32_Audit` `40_Leave` `50_Payroll` `60_Archive` `70_Viewers`（`.js`） |
| 一次性腳本 | `tools/`（建表、種子資料、模擬、稽核測試；只在 Apps Script 編輯器手動執行） |
| 前端頁面 | `index.html`、`view_*.html`、`css.html` |
| 前端邏輯 | `js_*.html`（Vue 模組） |
| 本機回歸測試 | `dev/`（Node，不會推上 Apps Script） |

> Apps Script 裡的 `.gs` 檔在這裡存成 `.js`，這是 `clasp` 的慣例，推回 Apps Script 時會自動轉回 `.gs`。

## 與 Apps Script 同步（clasp）

本 repo 已附 `.clasp.json`（指向原專案的 scriptId）與 `.claspignore`。目前**尚未包含 `appsscript.json`**：請到 Apps Script 編輯器「專案設定 → 在編輯器中顯示 appsscript.json」，把內容存成根目錄的 `appsscript.json`。

> ⚠️ 不要在 repo 內直接 `clasp pull`：會把遠端已移除的舊檔（例如 `Api.js`）拉回來，造成函式重複定義。

```bash
npm install -g @google/clasp
clasp login
clasp push        # 以本機檔案完整取代遠端專案
```

需要先在 https://script.google.com/home/usersettings 開啟「Google Apps Script API」。

## 注意

- `tools/Setup_Database.js` 內的員工資料與預設密碼為模擬資料。
- `tools/Setup_Database.js` 的 `setupEntireDatabaseSystem()` 會**清空所有工作表**，正式環境請勿執行。
