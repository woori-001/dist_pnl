const { chromium } = require('playwright');
const fs = require('fs'), path = require('path'), http = require('http');
const ROOT = process.env.APP_ROOT || require('path').resolve(__dirname, '..'), D = (process.env.XLSX_DIR || '') + '/';
const OUT = (process.env.OUT_DIR || '/tmp/dist_pnl_shots') + '/'; fs.mkdirSync(OUT, { recursive: true });
const srv = http.createServer((q, s) => { const p = path.join(ROOT, decodeURIComponent(q.url.split('?')[0]) === '/' ? 'index.html' : decodeURIComponent(q.url.split('?')[0]));
  fs.readFile(p, (e, b) => { if (e) { s.writeHead(404); s.end(); return; } s.writeHead(200, { 'content-type': p.endsWith('.js') ? 'text/javascript' : p.endsWith('.css') ? 'text/css' : 'text/html; charset=utf-8' }); s.end(b); }); }).listen(8765);
const runs = [], linesBy = {};
{ const old = (process.env.LEGACY_JSON ? require(process.env.LEGACY_JSON)['08'] : { summary: { segments: {} }, lines: [] }); runs.push({ id: 'runL', period: '2026-08-01', status: '가정산', seg_status: { '유통이관': '가정산', 'SKB웨이브넷': '가정산', 'SKY우리정보': '가정산', 'SK매직': '가정산' }, note: null, files: [], summary: old.summary, rules: {}, created_by: 'old@x.com', created_at: '2026-10-26T08:00:00Z' }); linesBy['runL'] = old.lines; }
(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  const ctx = await b.newContext({ viewport: { width: 1280, height: 900 }, locale: 'ko-KR', acceptDownloads: true });
  await ctx.route('https://cdn.jsdelivr.net/**', r => r.request().url().includes('pretendard') ? r.fulfill({ contentType: 'text/css', body: '' }) : r.fulfill({ contentType: 'text/javascript', body: fs.readFileSync(process.env.SB_ESM) }));
  await ctx.route('https://cdn.sheetjs.com/**', r => r.fulfill({ contentType: 'text/javascript', body: fs.readFileSync((process.env.NODE_MODS || 'node_modules') + '/xlsx/dist/xlsx.full.min.js') }));
  await ctx.route('https://cdnjs.cloudflare.com/**', r => r.fulfill({ contentType: 'text/javascript', body: fs.readFileSync(r.request().url().includes('jszip') ? (process.env.NODE_MODS || 'node_modules') + '/jszip/dist/jszip.min.js' : (process.env.NODE_MODS || 'node_modules') + '/html2canvas/dist/html2canvas.min.js') }));
  await ctx.route('https://fonts.**/**', r => r.fulfill({ contentType: 'text/css', body: '' }));
  await ctx.route('https://omdyyvqkxxigilxjysbz.supabase.co/**', async r => {
    const u = new URL(r.request().url()), m = r.request().method();
    const json = (o, st = 200) => r.fulfill({ status: st, contentType: 'application/json', body: JSON.stringify(o) });
    if (u.pathname.endsWith('/dist_pnl_members')) return json([{ email: 'test@x.com' }]);
    if (u.pathname.endsWith('/rpc/dist_pnl_save_run2')) {
      const p = JSON.parse(r.request().postData()); const id = 'run' + (runs.length + 1);
      for (const [k, v] of Object.entries(p.p_seg_status)) if (v === '확정') runs.forEach(x => { if (x.period === p.p_period && x.seg_status[k] === '확정') x.seg_status[k] = '가정산'; });
      runs.unshift({ id, period: p.p_period, status: '가정산', seg_status: p.p_seg_status, note: p.p_note, files: p.p_files, summary: p.p_summary, rules: p.p_rules, created_by: 'test@x.com', created_at: new Date().toISOString() });
      linesBy[id] = p.p_lines; return json(id);
    }
    if (u.pathname.endsWith('/rpc/dist_pnl_set_seg_status')) { const p = JSON.parse(r.request().postData()); const run = runs.find(x => x.id === p.p_run);
      if (p.p_status === '확정') runs.forEach(x => { if (x.period === run.period && x.seg_status[p.p_seg] === '확정') x.seg_status[p.p_seg] = '가정산'; });
      run.seg_status[p.p_seg] = p.p_status; return json(null); }
    if (u.pathname.endsWith('/rpc/dist_pnl_save_run')) {
      const p = JSON.parse(r.request().postData()); const id = 'run' + (runs.length + 1);
      if (p.p_status === '확정') runs.forEach(x => { if (x.period === p.p_period) x.status = '가정산'; });
      runs.unshift({ id, period: p.p_period, status: p.p_status, note: p.p_note, files: p.p_files, summary: p.p_summary, rules: p.p_rules, created_by: 'test@x.com', created_at: new Date().toISOString() });
      linesBy[id] = p.p_lines; return json(id);
    }
    if (u.pathname.endsWith('/rpc/dist_pnl_set_status')) { const p = JSON.parse(r.request().postData()); const run = runs.find(x => x.id === p.p_run);
      if (p.p_status === '확정') runs.forEach(x => { if (x.period === run.period) x.status = '가정산'; }); run.status = p.p_status; return json(null); }
    if (u.pathname.endsWith('/dist_pnl_runs') && m === 'GET') return json([...runs].sort((a, b) => b.period.localeCompare(a.period)));
    if (u.pathname.endsWith('/dist_pnl_lines')) return json(linesBy[u.searchParams.get('run_id').slice(3)] || []);
    return json({});
  });
  const page = await ctx.newPage();
  const errs = []; page.on('pageerror', e => errs.push(e.message)); page.on('console', m => { if (m.type() === 'error') errs.push(m.text()); });
  const exp = Math.floor(Date.now() / 1000) + 3600;
  const jwt = 'x.' + Buffer.from(JSON.stringify({ sub: 'u1', email: 'test@x.com', exp, role: 'authenticated' })).toString('base64url') + '.y';
  await page.addInitScript(([jwt, exp]) => localStorage.setItem('sb-omdyyvqkxxigilxjysbz-auth-token', JSON.stringify({ access_token: jwt, refresh_token: 'r', expires_at: exp, expires_in: 3600, token_type: 'bearer', user: { id: 'u1', email: 'test@x.com', aud: 'authenticated' } })), [jwt, exp]);
  await page.goto('http://localhost:8765/');
  await page.waitForSelector('#viewApp:not([hidden])', { timeout: 10000 });
  const files = fs.readdirSync(D);
  console.log('initial rail:', (await page.innerText('#railInfo')).replace(/\n/g,' | '));
  for (const m of ['9', '7']) {
    await page.click(`.cm-m[data-p="2026-0${m}"]`); await page.waitForTimeout(300);
    await page.setInputFiles('#fileInput', files.filter(f => f.includes(`-${m}_`) && f.includes('_2026102')).map(f => D + f));
    await page.waitForSelector('#fileList table');
    await page.click('#btnCalc');
    await page.waitForSelector('#calcEval .ev-tbl');
    console.log(m, 'rail:', (await page.innerText('#railInfo')).replace(/\n/g,' | '));
    console.log(' eval:', (await page.innerText('#calcEval .ev-tbl')).split('\n').slice(1,3).join(' | '));
    if (m === '9') {
      console.log(' auto:', (await page.innerText('#calcGifts')).split('\n').slice(0,3).join(' | '));
      const cks = await page.$$('#calcGifts .gsel'); await cks[0].click(); await cks[3].click({ modifiers: ['Shift'] });
      console.log(' selN', await page.innerText('#giftSelN'));
      await page.$eval('#giftBox', e => e.scrollIntoView()); const y0 = await page.$eval('.gsel', e => Math.round(e.getBoundingClientRect().top));
      await (await page.$('#calcGifts .gift-bar')).screenshot({ path: OUT + 'giftbar.png' }); await (await page.$('#giftBox')).screenshot({ path: OUT + 'giftbox.png' });
      await page.click('#calcGifts [data-giftsel=ded]'); await page.waitForTimeout(300);
      console.log(' after sel ded:', await page.innerText('#giftStats'), 'row moved px', (await page.$eval('.gsel', e => Math.round(e.getBoundingClientRect().top))) - y0);
      await page.click('#calcGifts [data-giftall=ded]'); await page.waitForTimeout(200);
      await page.screenshot({ path: OUT + 'calc.png', fullPage: true });
      await page.click('#btnToggleFiles'); await page.waitForTimeout(200); await page.screenshot({ path: OUT + 'files.png' }); await page.click('#btnToggleFiles'); }
    await page.click('#btnSave'); await page.waitForTimeout(800);
    console.log(' save:', await page.innerText('#toast'), '| rail:', (await page.innerText('#railInfo')).replace(/\n/g,' | '));
  }
  await page.click('.cm-m[data-p="2026-09"]'); await page.waitForTimeout(600);
  console.log('saved 9:', (await page.innerText('#railInfo')).replace(/\n/g,' | '));
  console.log('after confirm:', (await page.innerText('#railInfo')).replace(/\n/g,' | '), await page.$eval('.cm-m[data-p="2026-09"] i', e => e.className));
  await page.screenshot({ path: OUT + 'history.png', fullPage: true });
  await page.click('.cm-m[data-p="2026-08"]'); await page.waitForTimeout(800);
  console.log('legacy 8:', (await page.innerText('#railInfo')).replace(/\n/g,' | '));
  await page.screenshot({ path: OUT + 'legacy.png', fullPage: true });
  await page.click('.cm-m[data-p="2026-06"]'); await page.waitForTimeout(300);
  console.log('empty 6:', (await page.innerText('#railInfo')).replace(/\n/g,' | '), 'upload visible', await page.isVisible('#uploadCard'));
  await page.setViewportSize({ width: 390, height: 800 }); await page.waitForTimeout(300);
  await page.screenshot({ path: OUT + 'mobile.png' });
  console.log('scrollW', await page.evaluate(() => [document.documentElement.scrollWidth, ...[...document.querySelectorAll('body *')].filter(e => e.getBoundingClientRect().right > 392 && getComputedStyle(e).position !== 'fixed').slice(0, 8).map(e => e.tagName + '.' + e.className + ' ' + Math.round(e.getBoundingClientRect().right))]));
  console.log('errors', errs);
  await b.close(); srv.close();
})();
