/**
 * 60_Archive.js — 歷史資料封存（意願 / 班表 移至 *_Archive 冷儲存）
 */
function apiArchiveHistoricalData(token, cutoffYearMonth) {
  const user = validateToken(token);
  if (!isAdmin_(user)) return { success: false, message: "權限不足！" };
  const tz = getTz_();
  const now = new Date();
  getOrCreateSheet_("AVAIL_ARCHIVE");
  getOrCreateSheet_("SCHED_ARCHIVE");

  return withLock_(CONFIG.LOCK_MS.ARCHIVE, "封存失敗：", () => {
    const A = COL.AVAIL, S = COL.SCHED;
    // 意願：同一人同月份在封存區只保留最新一筆
    moveToArchive_("AVAIL", "AVAIL_ARCHIVE", r => normalizeYearMonth(r[A.year_month], tz) < cutoffYearMonth,
      r => `${normalizeYearMonth(r[A.year_month], tz)}___${String(r[A.employee_id]).trim()}`, now);
    // 班表：以 schedule_id 去重
    moveToArchive_("SCHED", "SCHED_ARCHIVE", r => normalizeDateStr(r[S.date], tz).substring(0, 7) < cutoffYearMonth,
      r => String(r[S.schedule_id]).trim(), now);
    return { success: true, message: `歷史封存成功！已將 ${cutoffYearMonth} 以前資料安全歸檔至冷儲存區。` };
  });
}

/**
 * 將 activeKey 表中 shouldMove 的列搬到 archiveKey 表（尾端加 archived_at），
 * 封存區內 keyOf 相同的舊列會被新列取代。
 */
function moveToArchive_(activeKey, archiveKey, shouldMove, keyOf, now) {
  const activeW = CONFIG.SHEETS[activeKey].headers.length;
  const archW = CONFIG.SHEETS[archiveKey].headers.length;
  const keep = [], moving = [];
  readRows_(activeKey).forEach(r => {
    const row = fitRow_(r, activeW);
    if (shouldMove(row)) moving.push(row.concat([now])); else keep.push(row);
  });
  if (moving.length > 0) {
    const movingKeys = new Set(moving.map(keyOf));
    const arc = readRows_(archiveKey).filter(r => !movingKeys.has(keyOf(r))).map(r => fitRow_(r, archW));
    writeAll_(archiveKey, getSheet_(archiveKey), arc.concat(moving));
  }
  const activeSheet = getSheet_(activeKey);
  if (activeSheet) writeAll_(activeKey, activeSheet, keep);
}
