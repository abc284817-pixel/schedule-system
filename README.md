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

| 類型 | 檔案 |
|---|---|
| 後端主程式 | `Api.js`（API、排班演算法、稽核、封存） |
| 計薪 | `Service_Payroll.js` |
| 資料庫建置 / 工具 | `Schema.js`、`SeedData.js`、`RepairScheduleFinalData.js` |
| 測試 / 模擬 | `Simulate_*.js`、`Test_Schedule_Engine.js` |
| 前端頁面 | `index.html`、`view_*.html`、`css.html` |
| 前端邏輯 | `js_*.html`（Vue 模組） |

> Apps Script 裡的 `.gs` 檔在這裡存成 `.js`，這是 `clasp` 的慣例，推回 Apps Script 時會自動轉回 `.gs`。

## 與 Apps Script 同步（clasp）

本 repo 已附 `.clasp.json`（指向原專案的 scriptId）。目前**尚未包含 `appsscript.json`**（專案設定檔），第一次同步前請先拉下來：

```bash
npm install -g @google/clasp
clasp login
clasp pull        # 取得 appsscript.json 與最新程式碼
clasp push        # 把本機修改推回 Apps Script
```

需要先在 https://script.google.com/home/usersettings 開啟「Google Apps Script API」。

## 注意

- `SeedData.js` 內的員工資料與預設密碼為模擬資料。
- `Schema.js` 的 `setupEntireDatabaseSystem()` 會**清空所有工作表**，正式環境請勿執行。
