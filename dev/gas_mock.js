// dev/gas_mock.js — 最小 Apps Script 模擬（試算表含日期自動轉型、快取、屬性、鎖、時區）。只給 dev/ 測試用，不會推上 GAS。
const crypto = require('crypto');
const FIXED_NOW = Date.UTC(2026, 10, 22, 13, 50, 0); // 2026-11-22 21:50 Asia/Taipei

function makeEnv() {
  const RealDate = Date;
  let tick = 0;
  class FakeDate extends RealDate {
    constructor(...a) { if (a.length === 0) super(FIXED_NOW + (tick++) * 0); else super(...a); }
    static now() { return FIXED_NOW; }
  }
  const isDate = v => v instanceof RealDate;

  // Sheets auto-convert date-like strings into Date (spreadsheet tz = Asia/Taipei = process TZ)
  function coerce(v) {
    if (typeof v !== 'string') return v;
    let m = v.match(/^(\d{4})-(\d{1,2})$/);
    if (m) return new FakeDate(+m[1], +m[2] - 1, 1);
    m = v.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return new FakeDate(+m[1], +m[2] - 1, +m[3]);
    m = v.match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})$/);
    if (m) return new FakeDate(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);
    return v;
  }
  const clone = v => (isDate(v) ? new FakeDate(v.getTime()) : v);

  class Range {
    constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
    setValues(vals) {
      if (vals.length !== this.nr || vals[0].length !== this.nc) throw new Error(`dim mismatch ${vals.length}x${vals[0].length} vs ${this.nr}x${this.nc}`);
      this.sh._ensure(this.r + this.nr - 1, this.c + this.nc - 1, true);
      vals.forEach((row, i) => row.forEach((v, j) => { this.sh.g[this.r - 1 + i][this.c - 1 + j] = coerce(clone(v)); }));
      return this;
    }
    setValue(v) { return this.setValues([[v]]); }
    getValue() { return clone(this.sh.g[this.r - 1][this.c - 1]); }
    getValues() { const o = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(clone(this.sh.g[this.r - 1 + i][this.c - 1 + j])); o.push(row); } return o; }
    clearContent() { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) if (this.sh.g[this.r - 1 + i]) this.sh.g[this.r - 1 + i][this.c - 1 + j] = ""; return this; }
    getA1Notation() { return String.fromCharCode(64 + this.c) + this.r; }
    getSheet() { return this.sh; }
    setBackground() { return this; } setFontColor() { return this; } setFontWeight() { return this; } setNumberFormat() { return this; }
  }
  class Sheet {
    constructor(name) { this.name = name; this.maxR = 1000; this.maxC = 26; this.g = []; this.frozen = 0; this._ensure(this.maxR, this.maxC); }
    _ensure(r, c, strict) {
      if (strict && (r > this.maxR || c > this.maxC)) throw new Error(`out of bounds ${r},${c} max ${this.maxR},${this.maxC} in ${this.name}`);
      while (this.g.length < this.maxR) this.g.push([]);
      this.g.forEach(row => { while (row.length < this.maxC) row.push(""); });
    }
    getName() { return this.name; }
    getMaxRows() { return this.maxR; } getMaxColumns() { return this.maxC; }
    getLastRow() { for (let i = this.maxR - 1; i >= 0; i--) if (this.g[i].some(v => v !== "" && v !== null)) return i + 1; return 0; }
    _lastCol() { let m = 0; this.g.forEach(row => row.forEach((v, j) => { if (v !== "" && v !== null && j + 1 > m) m = j + 1; })); return m; }
    getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this._lastCol())); }
    getRange(a, b, c, d) {
      if (typeof a === 'string') { const m = a.match(/^([A-Z])(\d+)$/); return new Range(this, +m[2], m[1].charCodeAt(0) - 64, 1, 1); }
      return new Range(this, a, b, c || 1, d || 1);
    }
    insertColumnsAfter(after, n) { this.maxC += n; this._ensure(this.maxR, this.maxC); }
    insertRowsAfter(after, n) { this.maxR += n; this._ensure(this.maxR, this.maxC); }
    deleteRows(start, n) { this.g.splice(start - 1, n); this.maxR -= n; }
    deleteRow(r) { this.deleteRows(r, 1); }
    appendRow(row) { const r = this.getLastRow() + 1; if (r > this.maxR) this.insertRowsAfter(this.maxR, 1); if (row.length > this.maxC) this.insertColumnsAfter(this.maxC, row.length - this.maxC); this.getRange(r, 1, 1, row.length).setValues([row]); }
    setFrozenRows(n) { this.frozen = n; } getFrozenRows() { return this.frozen; }
    clear() { this.g = []; this._ensure(this.maxR, this.maxC); }
    dump() { const lr = this.getLastRow(), lc = this._lastCol(); return this.g.slice(0, lr).map(r => r.slice(0, lc)); }
  }
  const sheets = {};
  const ss = {
    getSheetByName: n => sheets[n] || null,
    insertSheet: n => { sheets[n] = new Sheet(n); return sheets[n]; },
    _sheets: sheets
  };
  const cache = {}, props = {};
  let uuid = 0;
  const tzName = tz => (tz === 'GMT+8' ? 'Etc/GMT-8' : tz);
  function formatDate(d, tz, fmt) {
    const p = {};
    new Intl.DateTimeFormat('en-CA', { timeZone: tzName(tz), year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })
      .formatToParts(new RealDate(d.getTime())).forEach(x => { p[x.type] = x.value; });
    return fmt.replace('yyyy', p.year).replace('MM', p.month).replace('dd', p.day).replace('HH', p.hour).replace('mm', p.minute).replace('ss', p.second);
  }
  const env = {
    Date: FakeDate, console,
    SpreadsheetApp: { getActiveSpreadsheet: () => ss },
    CacheService: { getScriptCache: () => ({ get: k => (k in cache ? cache[k] : null), put: (k, v) => { cache[k] = v; }, remove: k => { delete cache[k]; } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = v; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Session: { getScriptTimeZone: () => 'Asia/Taipei' },
    Logger: { log() {} },
    HtmlService: {}, Maps: { newDirectionFinder() {} },
    Utilities: {
      formatDate,
      getUuid: () => { uuid++; return uuid.toString(16).padStart(8, '0') + '-aaaa-bbbb-cccc-dddddddddddd'; },
      computeDigest: (alg, input) => Array.from(crypto.createHash('sha256').update(input, 'utf8').digest()).map(b => (b > 127 ? b - 256 : b)),
      DigestAlgorithm: { SHA_256: 'SHA_256' }, Charset: { UTF_8: 'UTF_8' }
    }
  };
  return { env, ss, Sheet, RealDate };
}
module.exports = { makeEnv };
