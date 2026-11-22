// dev/ui_diff.js — 前端畫面回歸比對：用真的 Chrome 渲染 <REF> 版與目前工作目錄版（各接各自的後端模擬），
// 走一遍登入 / 各分頁 / 切月份 / 智慧排班 / 員工請假，逐步比對 #app 的 DOM。
// 需求：Chrome（預設 /usr/bin/google-chrome，可用 CHROME=路徑 覆寫）與 `npm i --prefix dev`。
// 用法：node dev/ui_diff.js [REF=HEAD]     預期：全部 SAME；DIFF 只應出現在你刻意改的畫面。
process.env.TZ = 'Asia/Taipei';
const fs = require('fs'), vm = require('vm'), path = require('path'), crypto = require('crypto'), { execSync } = require('child_process');
const { chromium } = require('playwright-core');
const { makeEnv } = require('./gas_mock');
const REPO = path.resolve(__dirname, '..');
const REF = process.argv[2] || 'HEAD';
const VUE = fs.readFileSync(require.resolve('vue/dist/vue.global.prod.js'), 'utf8');

function readFile(which, f) {
  return which === 'old' ? execSync(`git -C "${REPO}" show ${REF}:${f}`, { encoding: 'utf8', maxBuffer: 1 << 26 }) : fs.readFileSync(path.join(REPO, f), 'utf8');
}
function listJs(which) {
  if (which !== 'old') return fs.readdirSync(REPO).filter(f => f.endsWith('.js')).sort();
  return execSync(`git -C "${REPO}" ls-tree --name-only ${REF}`, { encoding: 'utf8' }).split('\n').filter(f => f.endsWith('.js')).sort();
}
const ctxCache = {};
function makeBackend(which) {
  const files = listJs(which).map(f => readFile(which, f));
  const { env, ss } = makeEnv();
  const ctx = vm.createContext(env);
  vm.runInContext(files.join('\n;\n'), ctx);
  const G = (name, ...a) => vm.runInContext(name, ctx)(...a);
  ctxCache[which] = ctx;
  // seed (deterministic)
  let s = 7; const R = () => { s = (s * 1103515245 + 12345) % 2147483648; return s / 2147483648; };
  const put = (name, rows) => { const sh = ss.insertSheet(name); sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows); };
  const sha = x => crypto.createHash('sha256').update(x, 'utf8').digest('hex');
  const staff = [["employee_id", "password_salt", "password_hash", "name", "role", "is_core", "is_active"]];
  const ids = [];
  for (let i = 1; i <= 15; i++) { const id = 'EMP' + String(i).padStart(3, '0'); ids.push(id); staff.push([id, 's' + i, sha('123456s' + i), '員工' + i, 'Staff', i <= 4, true]); }
  staff.push(['MGR001', 'sm', sha('123456sm'), '店長', 'Admin', true, true]);
  put('Staff_Master', staff);
  put('Config_Holidays', [["Date", "Holiday_Name", "Is_Double_Pay", "Holiday_Type"], ["2026-12-25", "行憲紀念日", "TRUE", "國定"], ["2026-11-12", "國父誕辰", "FALSE", "紀念日"]]);
  const tok = {};
  ids.concat(['MGR001']).forEach(id => { tok[id] = G('apiLogin', id, '123456').token; });
  G('apiSetUnlockFillPermission', tok.MGR001, '2026-11', 3, '');
  for (const ym of ['2026-11', '2026-12']) ids.forEach((id, i) => {
    if (ym === '2026-12' && i === 9) return;
    const [y, m] = ym.split('-').map(Number), days = new Date(y, m, 0).getDate(), av = {};
    for (let d = 1; d <= days; d++) av[`${ym}-${String(d).padStart(2, '0')}`] = { SHIFT_LUNCH: R() < 0.6, SHIFT_DINNER: R() < 0.6 };
    G('apiSubmitAvailability', tok[id], { year_month: ym, preferences: { work_volume: ['More', 'Standard', 'Less'][i % 3], day_preference: 'Flexible', partner_id: i === 2 ? 'EMP004' : '', dislike_partner_id: i === 5 ? 'EMP007' : '' }, availability: av });
  });
  G('apiRunAutoSchedule', tok.MGR001, '2026-11', true);
  // many manual extra shifts for EMP003 in Nov → legal conflicts in Nov
  const adds = []; for (let d = 1; d <= 30; d++) adds.push({ date: `2026-11-${String(d).padStart(2, '0')}`, shift_id: 'SHIFT_DINNER', employee_id: 'EMP003', dispatch_type: 'Assigned' });
  G('apiBatchCommitScheduleModifications', tok.MGR001, '2026-11', [], adds, []);
  const sh = G('apiGetMyShiftsForMonth', tok.EMP002, '2026-11').shifts;
  if (sh[0]) G('apiSubmitLeaveRequest', tok.EMP002, { date: sh[0].date, shift_id: sh[0].shift_id, leave_type: '病假', reason_detail: '發燒', has_substitute: false });
  return G;
}
function assemble(which) {
  const cfg = vm.runInContext("typeof getClientConfig === 'function' ? JSON.stringify(getClientConfig()) : '{}'", ctxCache[which]);
  const expand = txt => txt.replace(/<\?!=\s*include\('([^']+)'\)\s*\?>/g, (_, f) => expand(readFile(which, f + '.html')))
    .replace(/<\?!=\s*JSON\.stringify\(getClientConfig\(\)\)\s*\?>/g, cfg);
  return expand(readFile(which, 'index.html'))
    .replace('<script src="https://cdn.tailwindcss.com"></script>', '')
    .replace('<script src="https://unpkg.com/vue@3/dist/vue.global.prod.js"></script>', '<script src="https://vue.local/vue.js"></script>');
}

const INIT = `
  window.__pending = 0; window.__errors = [];
  window.addEventListener('error', e => window.__errors.push(String(e.message)));
  window.google = { script: { get run() {
    let ok = () => {}, fail = () => {};
    const p = new Proxy({}, { get(_, name) {
      if (name === 'withSuccessHandler') return f => { ok = f; return p; };
      if (name === 'withFailureHandler') return f => { fail = f; return p; };
      return (...args) => { window.__pending++; window.__gas(name, JSON.stringify(args)).then(r => { const v = JSON.parse(r);
        setTimeout(() => { try { v.err !== undefined ? fail(new Error(v.err)) : ok(v.res); } finally { window.__pending--; } }, 5); }); };
    } });
    return p; } } };
  window.confirm = () => true;`;

async function run(which, browser) {
  const G = makeBackend(which);
  const html = assemble(which);
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 }, timezoneId: 'Asia/Taipei' });
  await page.clock.setFixedTime(new Date('2026-11-22T21:50:00+08:00'));
  const consoleMsgs = [];
  page.on('console', m => { if (m.type() === 'warning' || m.type() === 'error') consoleMsgs.push(m.text()); });
  page.on('pageerror', e => consoleMsgs.push('PAGEERROR ' + e.message));
  await page.exposeFunction('__gas', (name, argsJson) => {
    try { return JSON.stringify({ res: G(name, ...JSON.parse(argsJson)) }); } catch (e) { return JSON.stringify({ err: e.message }); }
  });
  await page.addInitScript(INIT);
  await page.route('**/*', r => {
    const u = r.request().url();
    if (u.startsWith('https://vue.local/')) return r.fulfill({ contentType: 'application/javascript', body: VUE });
    if (u.startsWith('https://app.local/')) return r.fulfill({ contentType: 'text/html', body: html });
    return r.abort();
  });
  await page.goto('https://app.local/');
  const idle = async () => { await page.waitForFunction(() => window.__pending === 0); await page.waitForTimeout(60); await page.waitForFunction(() => window.__pending === 0); };
  const snaps = [];
  const snap = async label => { await idle(); const h = await page.evaluate(() => document.getElementById('app').innerHTML.replace(/<!--[\s\S]*?-->/g, '')); snaps.push([label, h]); };
  const nav = async text => { await page.locator('main > div.hidden.md\\:flex button', { hasText: text }).first().click(); await snap('tab ' + text); };
  const pick = async (value, nth = 0) => { const sel = page.locator(`select:visible:has(option[value="${value}"])`).nth(nth); await sel.selectOption(value); };
  const login = async id => {
    await page.locator('input[placeholder="例：EMP001"]').fill(id);
    await page.locator('input[placeholder="請輸入密碼"]').fill('123456');
    await page.locator('button', { hasText: '登入' }).last().click();
    await snap('login ' + id);
  };

  await snap('start');
  await login('MGR001');
  await nav('我的排班班表');
  await page.locator('button', { hasText: '月曆' }).first().click().catch(() => {}); await snap('roster calendar');
  await nav('請假/代班申請');
  await nav('全店填報總覽');
  await nav('正式班表預覽與手動調度');
  await pick('2026-11'); await snap('sched month nov');
  await page.locator('button', { hasText: /清單/ }).first().click(); await snap('sched table view');
  await nav('個人意願填報'); // ← old: calendar was hijacked by schedule tab
  await page.locator('[class*="grid-cols-7"] > div:not(.border-dashed)').nth(4).click().catch(() => {}); await snap('avail toggle a day');
  await nav('全店填報總覽');
  await page.locator('button', { hasText: '智慧排班' }).first().click(); await snap('auto schedule confirm modal');
  await page.locator('button', { hasText: /確認/ }).last().click(); await snap('auto schedule done');
  await nav('正式班表預覽與手動調度');
  await page.locator('button', { hasText: '登出' }).first().click(); await snap('logout');
  await login('EMP003');
  await nav('我的排班班表');
  await pick('2026-11'); await snap('emp roster nov');
  await nav('請假/代班申請');
  await pick('2026-12'); await snap('emp leave month');
  const errs = await page.evaluate(() => window.__errors);
  await page.close();
  return { snaps, errs, consoleMsgs };
}

(async () => {
  const browser = await chromium.launch({ executablePath: process.env.CHROME || '/usr/bin/google-chrome', args: ['--no-sandbox'] });
  const o = await run('old', browser), n = await run('new', browser);
  await browser.close();
  let diffs = 0;
  o.snaps.forEach(([label, h], i) => {
    const [ln, hn] = n.snaps[i] || [];
    const same = h === hn;
    if (!same) diffs++;
    let at = 0; while (!same && at < h.length && h[at] === hn[at]) at++;
    console.log(`${same ? 'SAME' : 'DIFF'}  ${label}  (${h.length} chars)` + (same ? '' : `\n   old…${h.slice(Math.max(0, at - 150), at + 250)}\n   new…${(hn || '').slice(Math.max(0, at - 150), at + 250)}`));
  });
  console.log('ref errors', o.errs, o.consoleMsgs.slice(0, 5));
  console.log('now errors', n.errs, n.consoleMsgs.slice(0, 5));
  console.log(`steps=${o.snaps.length} diffs=${diffs}（ref=${REF}）`);
  process.exit(diffs === 0 && !n.errs.length ? 0 : 1);
})();
