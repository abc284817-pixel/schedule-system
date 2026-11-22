// dev/backend_diff.js — 後端行為回歸比對：同一套情境，分別跑 <REF> 版本與目前工作目錄版本，逐一比對 API 回傳與所有工作表。
// 用法（在 repo 根目錄）：TZ=Asia/Taipei node dev/backend_diff.js [REF=HEAD] [種子數=3]
// 預期結果：issues=0。若是「刻意」改變行為，差異應只出現在你改的地方。
process.env.TZ = 'Asia/Taipei';
const fs = require('fs'), vm = require('vm'), path = require('path'), crypto = require('crypto'), { execSync } = require('child_process');
const { makeEnv } = require('./gas_mock');
const REPO = path.resolve(__dirname, '..');
const REF = process.argv[2] || 'HEAD';
const SEEDS = Number(process.argv[3] || 3);

/** 讀取 GAS 程式碼：REF=null 代表工作目錄；只取根目錄 .js（tools/、dev/ 不載入） */
function loadSources(ref) {
  if (!ref) return fs.readdirSync(REPO).filter(f => f.endsWith('.js')).sort().map(f => fs.readFileSync(path.join(REPO, f), 'utf8'));
  return execSync(`git -C "${REPO}" ls-tree --name-only ${ref}`, { encoding: 'utf8' }).split('\n')
    .filter(f => f.endsWith('.js')).sort().map(f => execSync(`git -C "${REPO}" show ${ref}:${f}`, { encoding: 'utf8', maxBuffer: 1 << 26 }));
}

function runScenario(sources, seed) {
  const { env, ss } = makeEnv();
  const ctx = vm.createContext(env);
  vm.runInContext(sources.join('\n;\n'), ctx);
  const G = (name, ...a) => vm.runInContext(name, ctx)(...a);
  // ---------- seed ----------
  function rng(s) { return () => { s |= 0; s = s + 0x6D2B79F5 | 0; let t = Math.imul(s ^ s >>> 15, 1 | s); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }
  const R = rng(seed);
  const put = (name, rows) => { const sh = ss.insertSheet(name); sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows); sh.setFrozenRows(1); };
  const sha = s => crypto.createHash('sha256').update(s, 'utf8').digest('hex');
  const staff = [["employee_id", "password_salt", "password_hash", "name", "role", "is_core", "is_active"]];
  const ids = [];
  for (let i = 1; i <= 15; i++) {
    const id = 'EMP' + String(i).padStart(3, '0'); ids.push(id);
    staff.push([id, 'salt' + i, sha('123456salt' + i), '員工' + i, 'Staff', i <= 4, i === 15 && seed % 2 === 0 ? false : true]);
  }
  staff.push(['MGR001', 'saltm', sha('123456saltm'), '店長', 'Admin', true, true]);
  put('Staff_Master', staff);
  put('Config_Holidays', [["Date", "Holiday_Name", "Is_Double_Pay", "Holiday_Type"], ["2026-12-25", "行憲紀念日", "TRUE", "國定"], ["2026-11-12", "國父誕辰", "FALSE", "紀念日"], ["2026-12-31", "跨年", "TRUE", "節慶"]]);
  put('Availability_Raw', [["record_id", "year_month", "employee_id", "pref_work_volume", "pref_day_preference", "pref_partner_id", "dislike_partner_id", "availability_matrix", "submitted_at"]]);
  ['Viewer_Schedule', 'Viewer_Availability', 'Viewer_Leaves'].forEach(n => { const sh = ss.insertSheet(n); sh.getRange("B1").setValue("2026-10"); });
  put('Staff_Monthly_Stats', [["stat_id", "year_month", "employee_id", "total_shifts", "total_hours", "normal_hours", "double_pay_hours", "standby_hours", "updated_at"]]);

  const vols = ['More', 'Standard', 'Less'], dprefs = ['Flexible', 'Both_Preferred', 'Single_Preferred'];
  function availFor(ym, density) {
    const [y, m] = ym.split('-').map(Number); const days = new Date(y, m, 0).getDate();
    const av = {};
    for (let d = 1; d <= days; d++) {
      const ds = `${ym}-${String(d).padStart(2, '0')}`;
      av[ds] = { SHIFT_LUNCH: R() < density, SHIFT_DINNER: R() < density };
    }
    return av;
  }
  const calls = [];
  const call = (label, fn, ...a) => { let r; try { r = G(fn, ...a); } catch (e) { r = { THROWN: e.message }; } calls.push([label, r]); return r; };

  // ---------- scenario ----------
  const tokens = {};
  for (const id of ids.concat(['MGR001'])) { const r = call('login ' + id, 'apiLogin', id, '123456'); if (r && r.token) tokens[id] = r.token; }
  call('login bad', 'apiLogin', 'EMP001', 'wrong');
  const adm = tokens.MGR001;

  // submit availability (Nov via unlock, Dec, Jan)
  call('unlock ALL 2026-11', 'apiSetUnlockFillPermission', adm, '2026-11', 5, '');
  call('unlock bad hours', 'apiSetUnlockFillPermission', adm, '2026-11', 0, '');
  call('unlock non-admin', 'apiSetUnlockFillPermission', tokens.EMP002, '2026-11', 5, '');
  for (const ym of ['2026-11', '2026-12', '2027-01']) {
    ids.forEach((id, i) => {
      if (!tokens[id]) return;
      if (ym === '2026-12' && i % 7 === 3) return; // some missing
      const pick = a => a[Math.floor(R() * a.length)];
      const partner = R() < 0.4 ? pick(ids) : '';
      const dislike = R() < 0.3 ? pick(ids) : '';
      call(`submit ${ym} ${id}`, 'apiSubmitAvailability', tokens[id], { year_month: ym, preferences: { work_volume: pick(vols), day_preference: pick(dprefs), partner_id: partner, dislike_partner_id: dislike }, availability: availFor(ym, 0.35 + R() * 0.5) });
    });
  }
  call('submit empty', 'apiSubmitAvailability', tokens.EMP002, { year_month: '2027-01', preferences: {}, availability: availFor('2027-01', 0) });
  call('submit locked oct', 'apiSubmitAvailability', tokens.EMP002, { year_month: '2026-10', preferences: {}, availability: availFor('2026-10', 0.5) });
  call('resubmit dec EMP002', 'apiSubmitAvailability', tokens.EMP002, { year_month: '2026-12', preferences: { work_volume: 'More', day_preference: 'Both_Preferred', partner_id: 'EMP003', dislike_partner_id: 'EMP009' }, availability: availFor('2026-12', 0.9) });
  call('mgr submit dec (proxy EMP001)', 'apiSubmitAvailability', adm, { year_month: '2026-12', preferences: { work_volume: 'Standard' }, availability: availFor('2026-12', 0.6) });

  call('bootstrap EMP005', 'apiGetBootstrapData', tokens.EMP005, '2026-12');
  call('bootstrap MGR', 'apiGetBootstrapData', adm, '2026-11');
  call('avail EMP005 dec', 'apiGetAvailabilityForMonth', tokens.EMP005, '2026-12');
  call('avail MGR nov', 'apiGetAvailabilityForMonth', adm, '2026-11');
  call('summary dec', 'apiGetManagerSummary', adm, '2026-12');
  call('summary non-admin', 'apiGetManagerSummary', tokens.EMP002, '2026-12');

  call('auto nov ask', 'apiRunAutoSchedule', adm, '2026-11', false);
  call('auto nov', 'apiRunAutoSchedule', adm, '2026-11', true);
  call('auto dec ask', 'apiRunAutoSchedule', adm, '2026-12', false);
  call('auto dec', 'apiRunAutoSchedule', adm, '2026-12', true);
  call('auto jan', 'apiRunAutoSchedule', adm, '2027-01', true);
  call('auto dec again (overwrite)', 'apiRunAutoSchedule', adm, '2026-12', true);
  call('table dec', 'apiGetScheduleTable', adm, '2026-12');
  call('table nov', 'apiGetScheduleTable', adm, '2026-11');

  for (const id of ['EMP001', 'EMP003', 'EMP008', 'MGR001']) {
    const t = tokens[id];
    call('roster ' + id, 'apiGetMyRosterSchedule', t, '2026-12');
    call('roster nov ' + id, 'apiGetMyRosterSchedule', t, '2026-11');
    call('shifts ' + id, 'apiGetMyShiftsForMonth', t, '2026-12');
    call('shifts nov ' + id, 'apiGetMyShiftsForMonth', t, '2026-11');
  }
  // leaves
  const leaveIds = [];
  for (const id of ['EMP001', 'EMP003', 'EMP008', 'EMP002']) {
    const sh = G('apiGetMyShiftsForMonth', tokens[id], '2026-12').shifts || [];
    sh.slice(0, 2).forEach((s, k) => {
      const r = call(`leave ${id} ${k}`, 'apiSubmitLeaveRequest', tokens[id], { date: s.date, shift_id: s.shift_id, leave_type: '事假', reason_detail: 'x', has_substitute: k === 0, substitute_id: k === 0 ? 'EMP010' : 'EMP011' });
    });
  }
  call('leave past', 'apiSubmitLeaveRequest', tokens.EMP003, { date: '2026-11-01', shift_id: 'SHIFT_LUNCH' });
  call('leave invalid shift', 'apiSubmitLeaveRequest', tokens.EMP003, { date: '2026-12-31', shift_id: 'SHIFT_XX' });
  call('leave missing', 'apiSubmitLeaveRequest', tokens.EMP003, null);
  for (const id of ['EMP001', 'EMP003', 'EMP008', 'EMP002']) {
    const r = call('myleaves ' + id, 'apiGetMyLeaveRequests', tokens[id]);
    (r.requests || []).forEach(q => leaveIds.push(q.request_id));
  }
  call('table dec w/ leaves', 'apiGetScheduleTable', adm, '2026-12');
  leaveIds.forEach((rid, i) => {
    if (i % 3 === 0) call('approve ' + i, 'apiApproveLeaveRequest', adm, rid, i % 2 ? 'EMP012' : '');
    else if (i % 3 === 1) call('approve sub ' + i, 'apiApproveLeaveRequest', adm, rid, 'EMP013');
    else call('reject ' + i, 'apiRejectLeaveRequest', adm, rid);
  });
  call('approve missing', 'apiApproveLeaveRequest', adm, 'LR_nope', '');
  call('reject missing', 'apiRejectLeaveRequest', adm, 'LR_nope');
  call('approve non-admin', 'apiApproveLeaveRequest', tokens.EMP002, leaveIds[0], '');

  // batch commit
  const tbl = G('apiGetScheduleTable', adm, '2026-12');
  const sch = tbl.schedules;
  call('batch', 'apiBatchCommitScheduleModifications', adm, '2026-12',
    [{ schedule_id: sch[0].schedule_id, employee_id: 'EMP014', dispatch_type: 'Assigned' }, { schedule_id: sch[5].schedule_id, employee_id: sch[5].employee_id, dispatch_type: 'Standby' }],
    [{ date: '2026-12-05', shift_id: 'SHIFT_DINNER', employee_id: 'EMP002', dispatch_type: 'Assigned' }, { date: '2026-12-06', shift_id: 'SHIFT_LUNCH', employee_id: 'EMP004' }],
    [sch[10].schedule_id, sch[11].schedule_id]);
  call('table dec after batch', 'apiGetScheduleTable', adm, '2026-12');
  { // stress: force legal violations via manual additions
    const adds = [];
    for (let d = 1; d <= 31; d++) for (const sid of ['SHIFT_LUNCH', 'SHIFT_DINNER']) {
      adds.push({ date: `2026-12-${String(d).padStart(2, '0')}`, shift_id: sid, employee_id: 'EMP006', dispatch_type: 'Assigned' });
      if (d % 3) adds.push({ date: `2026-12-${String(d).padStart(2, '0')}`, shift_id: sid, employee_id: 'EMP007', dispatch_type: d % 2 ? 'Assigned' : 'Standby' });
    }
    call('batch stress', 'apiBatchCommitScheduleModifications', adm, '2026-12', [], adds, []);
    call('table dec stress', 'apiGetScheduleTable', adm, '2026-12');
  }
  call('recalc', 'apiRecalculateMonthlyStats', adm);
  call('recalc non-admin', 'apiRecalculateMonthlyStats', tokens.EMP002);
  call('holidays', 'apiGetDynamicHolidays');
  call('archive', 'apiArchiveHistoricalData', adm, '2026-12');
  call('table nov after archive', 'apiGetScheduleTable', adm, '2026-11');
  call('recalc after archive', 'apiRecalculateMonthlyStats', adm);
  call('viewers nov', 'refreshAllViewers', '2026-11');
  call('viewers dec', 'refreshAllViewers', '2026-12');
  call('validate bad token', 'validateToken', 'nope');


  const sheets = {};
  Object.keys(ss._sheets).sort().forEach(n => { sheets[n] = ss._sheets[n].dump(); });
  return JSON.parse(JSON.stringify({ calls, sheets }));
}

function diff(a, b, seed) {
  let issues = 0;
  a.calls.forEach(([label, ra], i) => {
    const [lb, rb] = b.calls[i] || [];
    const sa = JSON.stringify(ra), sb = JSON.stringify(rb);
    if (label !== lb || sa !== sb) {
      if (issues++ < (process.env.SHOW || 8)) console.log(`\n[seed ${seed}] CALL #${i} ${label}\n  ref: ${sa.slice(0, 500)}\n  now: ${String(sb).slice(0, 500)}`);
    }
  });
  new Set([...Object.keys(a.sheets), ...Object.keys(b.sheets)]).forEach(n => {
    const A = a.sheets[n] || [], B = b.sheets[n] || [];
    if (JSON.stringify(A) === JSON.stringify(B)) return;
    issues++;
    const i = A.findIndex((r, k) => JSON.stringify(r) !== JSON.stringify(B[k]));
    console.log(`\n[seed ${seed}] SHEET ${n} rows ref=${A.length} now=${B.length}, first diff row ${i}\n  ref: ${JSON.stringify(A[i])}\n  now: ${JSON.stringify(B[i])}`);
  });
  return issues;
}

const refSrc = loadSources(REF), nowSrc = loadSources(null);
let total = 0;
for (let seed = 1; seed <= SEEDS; seed++) {
  const n = diff(runScenario(refSrc, seed), runScenario(nowSrc, seed), seed);
  console.log(`seed ${seed}: issues=${n}`);
  total += n;
}
console.log(total === 0 ? `✅ 與 ${REF} 行為完全一致` : `⚠️ 共 ${total} 處差異（請確認是否皆為刻意變更）`);
process.exit(total === 0 ? 0 : 1);
