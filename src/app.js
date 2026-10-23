import { createClient } from 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.45.4/+esm';
import { SUPABASE_URL, SUPABASE_KEY } from './config.js';
import { RULES, SEGMENTS } from './rules.js';
import { parseWorkbook, compute, summarize, byUpper } from './calc.js';

const sb = createClient(SUPABASE_URL, SUPABASE_KEY);
const $ = id => document.getElementById(id);
const XLSX = window.XLSX;

const state = { files: [], result: null, runs: [], selPeriod: null, selRun: null, year: new Date().getFullYear(), segStatus: null };
const ST = ['가정산', '확정'];
const segSt = (run, k) => (run.seg_status && run.seg_status[k]) || run.status || '가정산';
const blankSt = () => Object.fromEntries(SEGMENTS.map(k => [k, '가정산']));
const SEG_COLOR = { '유통이관': '#3d6ef2', 'SKB웨이브넷': '#8b5cf6', 'SKY우리정보': '#06b6d4', 'SK매직': '#f59e0b' };

// ---------- 공통 ----------
const won = v => {
  const n = Math.round(v || 0);
  return `<span class="${n < 0 ? 'neg' : ''}">${n === 0 ? '-' : n.toLocaleString('ko-KR')}</span>`;
};
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const ym = d => d.slice(0, 7);
const monthLabel = p => `${p.slice(0, 4)}년 ${Number(p.slice(5, 7))}월`;
function toast(msg) { const t = $('toast'); t.textContent = msg; t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => t.hidden = true, 2600); }
function ask(msg) {
  return new Promise(res => {
    $('modalMsg').textContent = msg; $('modal').hidden = false;
    const done = v => { $('modal').hidden = true; $('modalYes').onclick = $('modalNo').onclick = null; res(v); };
    $('modalYes').onclick = () => done(true); $('modalNo').onclick = () => done(false);
  });
}
function show(view) {
  for (const v of ['viewLogin', 'viewDenied', 'viewApp']) $(v).hidden = v !== view;
  $('nav').hidden = view !== 'viewApp';
  $('bmCard').hidden = view !== 'viewApp' || location.hash.includes('admin-import');
  $('ruleCard').hidden = view !== 'viewApp';
  if (view === 'viewLogin') setTitle('유통 손익매출 계산기', '로그인하면 계산·저장할 수 있어요');
  if (view === 'viewApp') setTitle(...PAGE[currentTab]);
}
const PAGE = {
  calc: ['손익매출 계산하기', '어드민 가져오기 버튼이나 엑셀로 대상월 매출을 계산해요 · 매출 = 상부(B) − 전체(A) − 자점운영비(코드)'],
  history: ['월별 내역', '달력에서 달을 누르면 그 달의 가정산·확정 내역이 열려요'],
};
let currentTab = 'calc';
function setTitle(t, sub) { $('pageTitle').textContent = t; $('pageSub').textContent = sub; }

// ---------- 로그인 ----------
async function boot() {
  const { data: { session } } = await sb.auth.getSession();
  await onSession(session);
  sb.auth.onAuthStateChange((_e, s) => { if (!s) onSession(null); });
}
async function onSession(session) {
  if (!session) { $('who').hidden = true; show('viewLogin'); return; }
  $('who').hidden = false; $('whoEmail').textContent = session.user.email; $('whoEmail').title = session.user.email; $('ruleVer').textContent = RULES.version;
  const { data, error } = await sb.from('dist_pnl_members').select('email').limit(1);
  if (error || !data?.length) { show('viewDenied'); return; }
  show('viewApp');
  loadRuns();
  applyAdmin();
}
$('formLogin').addEventListener('submit', async e => {
  e.preventDefault(); $('loginErr').textContent = '';
  const { data, error } = await sb.auth.signInWithPassword({ email: $('loginEmail').value.trim(), password: $('loginPw').value });
  if (error) { $('loginErr').textContent = '로그인 실패: 이메일 또는 비밀번호를 확인하세요'; return; }
  onSession(data.session);
});
$('btnLogout').onclick = async () => { await sb.auth.signOut(); onSession(null); };

// ---------- 탭 ----------
document.querySelectorAll('#nav button').forEach(b => b.onclick = () => {
  currentTab = b.dataset.tab;
  document.querySelectorAll('#nav button').forEach(x => x.classList.toggle('on', x === b));
  document.querySelectorAll('.pane').forEach(p => p.hidden = p.id !== 'pane-' + b.dataset.tab);
  setTitle(...PAGE[currentTab]);
  if (b.dataset.tab === 'history') loadRuns();
});

// ---------- 파일 올리기 ----------
const prev = new Date(); prev.setDate(1); prev.setMonth(prev.getMonth() - 1);
$('period').value = `${prev.getFullYear()}-${String(prev.getMonth() + 1).padStart(2, '0')}`;
const drop = $('drop');
drop.onclick = () => $('fileInput').click();
drop.onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') $('fileInput').click(); };
drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
drop.ondragleave = () => drop.classList.remove('over');
drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); addFiles(e.dataTransfer.files); };
$('fileInput').onchange = e => { addFiles(e.target.files); e.target.value = ''; };
$('btnClear').onclick = () => { state.files = []; state.result = null; renderFiles(); $('calcResult').hidden = true; };
function compactUpload(on, info = {}) {
  $('uploadCard').classList.toggle('compact', on);
  $('uploadBody').hidden = on;
  $('btnToggleFiles').hidden = !on; $('btnNew').hidden = !on; $('btnSave').hidden = !on; $('btnExportCalc').hidden = !on;
  $('btnToggleFiles').textContent = '파일 목록 보기';
  $('uploadTitle').textContent = on ? `${monthLabel($('period').value)} · 계산에 쓴 파일 ${state.files.length}개` : '엑셀로 계산하기';
  $('uploadSub').innerHTML = on ? info.html : '어드민 가져오기 버튼을 쓰면 이 단계는 자동이에요. 엑셀을 직접 받았다면 여기에 올리세요.';
}
$('btnToggleFiles').onclick = () => { const open = $('uploadBody').hidden; $('uploadBody').hidden = !open; $('btnToggleFiles').textContent = open ? '파일 목록 접기' : '파일 목록 보기'; };
$('btnNew').onclick = () => { state.files = []; state.result = null; renderFiles(); $('calcResult').hidden = true; compactUpload(false); };

async function addFiles(fileList) {
  const list = [...fileList]; // FileList는 input 초기화 시 비워지므로 먼저 복사
  for (const f of list) {
    if (state.files.some(x => x.fileName === f.name && x.size === f.size)) { toast(`${f.name}: 이미 올린 파일입니다`); continue; }
    try {
      const parsed = parseWorkbook(XLSX, new Uint8Array(await f.arrayBuffer()), f.name);
      parsed.size = f.size; state.files.push(parsed);
    } catch (err) { toast(err.message); }
  }
  state.result = null; $('calcResult').hidden = true; compactUpload(false); renderFiles();
}
function renderFiles() {
  $('btnCalc').disabled = $('btnClear').disabled = !state.files.length;
  if (!state.files.length) { $('fileList').innerHTML = ''; return; }
  const rows = state.files.map((f, i) => {
    const k = {}; f.rows.forEach(r => k[r.kind] = (k[r.kind] || 0) + 1);
    const sum = f.rows.reduce((s, r) => s + r.amount, 0);
    const ok = f.fileTotal === null || f.fileTotal === sum;
    return `<tr><td><span class="pill ${f.side === '상부' ? 'up' : 'dn'}">${f.side}정산</span></td><td>${f.source === 'admin' ? '<span class="pill ok">어드민</span> ' : ''}${esc(f.fileName)}</td>
      <td>${Object.entries(k).map(([a, b]) => `${a} ${b.toLocaleString()}`).join(' · ')}</td>
      <td class="n">${won(sum)}</td><td>${ok ? '<span class="pill ok">합계 일치</span>' : '<span class="pill warn">합계 불일치</span>'}</td>
      <td><button class="link danger" data-rm="${i}">빼기</button></td></tr>`;
  }).join('');
  $('fileList').innerHTML = `<div class="tbl-wrap"><table class="files"><tr><th>구분</th><th>파일</th><th>정산종류별 건수</th><th>정산금액 합계</th><th>점검</th><th></th></tr>${rows}</table></div>`;
  $('fileList').querySelectorAll('[data-rm]').forEach(b => b.onclick = () => { state.files.splice(+b.dataset.rm, 1); state.result = null; $('calcResult').hidden = true; renderFiles(); });
}

// ---------- 어드민에서 바로 가져오기 (북마클릿 ↔ 이 창 postMessage) ----------
const ADMIN_ORIGIN = 'https://workline.kr';
const CALC_URL = location.origin + location.pathname.replace(/index\.html$/, '');
let pendingAdmin = null, importDone = false;
function adminWait(title, msg, opts = {}) {
  $('adminWait').hidden = false;
  $('adminWait').classList.toggle('err', !!opts.error);
  $('adminSpin').hidden = !!opts.done || !!opts.error;
  if (title) $('adminWaitTitle').textContent = title;
  if (msg !== undefined) $('adminWaitMsg').textContent = msg;
}
let waitTimer = null;
function enterImportMode() {
  if (!location.hash.includes('admin-import')) return;
  importDone = false;
  $('bmCard').hidden = true; $('uploadCard').hidden = true; $('calcResult').hidden = true;
  adminWait('어드민에서 데이터를 받는 중이에요', '어드민 창에서 엑셀 10개(옛날 데이터 포함)를 만들고 있어요. 1~2분 걸려요. 어드민 창을 닫지 마세요.');
  clearTimeout(waitTimer);
  waitTimer = setTimeout(() => adminWait('아직 데이터가 오지 않았어요', '어드민 창의 작은 창에 오류가 떠 있는지 확인해 주세요. 엑셀을 직접 올려도 됩니다.', { error: true }) || ($('uploadCard').hidden = false), 4 * 60 * 1000);
}
function announceReady() {
  enterImportMode();
  if (window.opener && location.hash.includes('admin-import')) {
    try { window.opener.postMessage({ type: 'dist-pnl-calc-ready' }, ADMIN_ORIGIN); } catch { /* 어드민 창이 닫힘 */ }
  }
}
window.addEventListener('message', e => {
  if (e.origin !== ADMIN_ORIGIN || !e.data) return; // 어드민 창에서 온 것만 받음
  if (e.data.type === 'dist-pnl-progress') {
    if (importDone) return; // 계산이 끝난 뒤 오는 안내는 무시
    clearTimeout(waitTimer);
    if (e.data.error) { adminWait('어드민에서 가져오지 못했어요', String(e.data.text || ''), { error: true }); $('uploadCard').hidden = false; }
    else adminWait('어드민에서 데이터를 받는 중이에요', String(e.data.text || ''));
    return;
  }
  if (e.data.type !== 'dist-pnl-admin-files') return;
  clearTimeout(waitTimer); importDone = true;
  const { period, files } = e.data;
  if (!/^\d{4}-\d{2}$/.test(period) || !Array.isArray(files)) return;
  const parsed = [], errors = [];
  for (const f of files) {
    try {
      // 직접 올린 엑셀과 똑같은 방식으로 읽음 (상부/전체는 시트 이름으로 다시 판별)
      const p = parseWorkbook(XLSX, new Uint8Array(f.buffer), String(f.name || '어드민 엑셀'));
      p.source = 'admin'; p.db = f.old ? 'old' : ''; parsed.push(p);
    } catch (err) { errors.push(err.message); }
  }
  pendingAdmin = { period, pulledAt: e.data.pulledAt, files: parsed, errors };
  if ($('viewApp').hidden) adminWait('데이터를 다 받았어요', '로그인하면 바로 계산돼요.', { done: true });
  applyAdmin();
});
window.addEventListener('hashchange', announceReady);
function applyAdmin() {
  if (!pendingAdmin || $('viewApp').hidden) return;
  const { period, files, pulledAt, errors } = pendingAdmin; pendingAdmin = null;
  if (errors.length) errors.forEach(m => toast(m));
  $('adminWait').hidden = true; $('uploadCard').hidden = false;
  document.querySelector('#nav button[data-tab=calc]').click();
  state.files = files; $('period').value = period;
  renderFiles();
  $('btnCalc').click();
  $('uploadSub').insertAdjacentHTML('afterbegin', `<b>어드민에서 가져옴</b> (${new Date(pulledAt || Date.now()).toLocaleString('ko-KR')}) · `);
  history.replaceState(null, '', location.pathname);
  toast('어드민 데이터로 계산했습니다');
}
async function setupBookmarklet() {
  try {
    let code = await fetch('src/admin-pull.js', { cache: 'no-cache' }).then(r => r.text());
    code = code.replace(/^\/\*[\s\S]*?\*\/\s*/, '').replace('__CALC_URL__', CALC_URL);
    $('bmLink').href = 'javascript:' + encodeURIComponent(code);
  } catch { $('bmLink').removeAttribute('href'); }
}
$('bmLink').onclick = e => { e.preventDefault(); toast('클릭하지 말고, 이 버튼을 즐겨찾기(북마크) 막대로 끌어다 놓으세요'); };

// ---------- 계산 ----------
$('btnCalc').onclick = () => {
  const period = $('period').value;
  if (!period) { toast('대상월을 고르세요'); return; }
  const { lines, checks } = compute(state.files, period);
  const summary = summarize(lines);
  state.result = { period, lines, checks, summary };
  state.segStatus = blankSt();
  const warn = [...checks.warnings];
  if (checks.outOfPeriod) warn.push(`대상월이 아닌 행 ${checks.outOfPeriod.toLocaleString()}건은 계산에서 뺐어요 (월을 잘못 골랐거나 다른 달 파일이 섞였는지 확인)`);
  $('calcChecks').innerHTML = warn.length ? `<div class="warns">${warn.map(w => `<div class="notice warn">⚠ ${esc(w)}</div>`).join('')}</div>` : '';
  $('calcChecks').hidden = !warn.length;
  const allOk = checks.files.every(f => f.totalOk);
  compactUpload(true, { html: `${warn.length ? '' : '<span class="okline">✓ 점검 통과</span> · '}${allOk ? '파일 합계 모두 일치' : '<b class="neg">파일 합계 불일치 있음</b>'} · 계산에 쓴 행 ${checks.usedRows.toLocaleString()}건 · 타부서 코드 ${checks.otherDept.toLocaleString()}건 제외` });
  $('calcSummary').innerHTML = renderSummary(summary, monthLabel(period), state.segStatus);
  $('calcUpper').innerHTML = renderUpper(byUpper(lines));
  $('calcResult').hidden = false; $('btnSave').textContent = '이 달 결과 저장';
};

$('btnSave').onclick = async () => {
  const r = state.result; if (!r) return;
  const st = { ...state.segStatus };
  const periodDate = r.period + '-01';
  const existing = state.runs.filter(x => ym(x.period) === r.period);
  const clash = SEGMENTS.filter(k => st[k] === '확정' && existing.some(x => segSt(x, k) === '확정'));
  if (clash.length && !await ask(`${monthLabel(r.period)} ${clash.join(', ')}은(는) 이미 확정된 회차가 있어요. 이번 계산을 확정으로 저장하면 이전 확정은 가정산으로 바뀌어요. 계속할까요?`)) return;
  $('btnSave').disabled = true; $('btnSave').textContent = '저장 중…';
  const files = state.files.map(f => ({ name: f.fileName, side: f.side, rows: f.rows.length, total: f.fileTotal }));
  const rules = { version: RULES.version, checks: { usedRows: r.checks.usedRows, outOfPeriod: r.checks.outOfPeriod, otherDept: r.checks.otherDept, warnings: r.checks.warnings } };
  const { error } = await sb.rpc('dist_pnl_save_run2', { p_period: periodDate, p_seg_status: st, p_note: null,
    p_files: files, p_summary: r.summary, p_rules: rules, p_lines: r.lines });
  $('btnSave').disabled = false;
  if (error) { $('btnSave').textContent = '이 달 결과 저장'; toast('저장 실패: ' + error.message); return; }
  $('btnSave').textContent = '✓ 저장됨';
  const nConf = SEGMENTS.filter(k => st[k] === '확정').length;
  toast(`${monthLabel(r.period)} 저장 완료 · ${nConf === 4 ? '전체 확정' : nConf ? `확정 ${nConf}/4` : '가정산'} — 월별 내역에서 볼 수 있어요`); loadRuns();
};
$('btnExportCalc').onclick = () => { const r = state.result; if (r) exportXlsx(r.period, r.summary, r.lines, state.segStatus); };

// ---------- 요약 화면 ----------
const LINES = [ // 엑셀 내보내기용 (전체 항목)
  ['sec', '전체정산 · 우리가 준 돈'],
  ['', 'a. 수수료 (사은품 차감전, 기타 차감 반영)', 'D_A'], ['', 'b. 사은품 - 현금', 'D_B'], ['', 'c. 사은품 - 물류', 'D_C'], ['', 'd. 사은품 - 쿠폰', 'D_D'], ['', 'f. 환수', 'D_F'],
  ['b', '최종 정산금액 (A) = a − (b+c+d+f)', 'A'],
  ['sec', '상부정산 · 우리가 받은 돈'],
  ['', 'a. 수수료 (a-1·a-2 제외)', 'U_A'], ['', 'a-1. 관리수수료', 'U_A1'], ['', 'a-2. 정산관련', 'U_A2'], ['', 'b. 사은품 - 현금', 'U_B'], ['', 'c. 사은품 - 물류', 'U_C'], ['', 'd. 사은품 - 쿠폰', 'U_D'], ['', 'f. 환수', 'U_F'],
  ['b', '최종 정산금액 (B) = (a+a1+a2) − (b+c+d+f)', 'B'],
  ['sec', '매출'],
  ['', '(−) 자점운영비 (코드 · 자점 일반 건)', 'JF', -1, 'code'],
  ['key', '매출 = B − A − 자점운영비', 'REV'],
  ['ref', '참고: 자점운영비 (A 중 자점 지급분)', 'JJ', 1, 'code'],
];
const SEG_NOTE = { '유통이관': '이관처 · 자점 제외', 'SKB웨이브넷': '코드 · 자점 포함', 'SKY우리정보': '코드 · 자점 포함', 'SK매직': '코드 · 자점 포함' };
let detailSeg = '유통이관';
function renderSummary(summary, title, segStatus = blankSt()) {
  const s = summary.segments;
  const kv = v => `${Math.round(v || 0).toLocaleString('ko-KR')}<small>원</small>`;
  const kpi = SEGMENTS.map(k => `<div class="kpi kpi-btn${k === detailSeg ? ' hot' : ''}" data-seg="${k}" role="tab" tabindex="0" aria-selected="${k === detailSeg}">
      <div class="kpi-top"><div class="kpi-label">${k}</div>
        <div class="stt" role="group" aria-label="${k} 정산 상태">${ST.map(x => `<button type="button" data-st="${x}" class="${segStatus[k] === x ? 'on ' + x : ''}" aria-pressed="${segStatus[k] === x}">${x}</button>`).join('')}</div></div>
      <div class="kpi-value">${kv(s[k].REV)}</div>
      <div class="kpi-note">${SEG_NOTE[k]}</div></div>`).join('');
  const t = s[detailSeg];
  // 상부·전체를 같은 항목으로 나란히 (차감 항목은 −로 표시)
  const rows = [
    ['수수료', '일반', t.U_A, t.D_A0], ['', '관리', t.U_A1, t.D_A1], ['', '정산관련', t.U_A2, t.D_A2],
    ['사은품', '현금', -t.U_B, -t.D_B], ['', '물류', -t.U_C, -t.D_C], ['', '쿠폰', -t.U_D, -t.D_D],
  ];
  const grpSpan = { '수수료': 3, '사은품': 3 };
  const body = rows.map(([g, item, up, dn]) => `<tr>${g ? `<th class="grp" rowspan="${grpSpan[g]}">${g}</th>` : ''}<td class="item">${item}</td>
      <td class="n">${won(up)}</td><td class="n">${won(dn)}</td><td class="n diff">${won(up - dn)}</td></tr>`).join('')
    + `<tr><th class="grp" colspan="2">환수</th><td class="n">${won(-t.U_F)}</td><td class="n">${won(-t.D_F)}</td><td class="n diff">${won(-t.U_F + t.D_F)}</td></tr>`
    + `<tr><th class="grp" colspan="2">자점운영비</th><td class="n">${won(-t.JF)}</td><td class="n">-</td><td class="n diff">${won(-t.JF)}</td></tr>`
    + `<tr class="total"><th colspan="2">합계</th><td class="n">${won(t.B - (t.JF || 0))}</td><td class="n">${won(t.A)}</td><td class="n rev">${won(t.REV)}</td></tr>`;
  const jj = detailSeg === '유통이관' ? '' : `<p class="muted small" style="margin:0">자점운영비 = 자점 영업팀 '일반' 건의 자점 마진(접수일 기준 금액, 7.유통사업부·0.영업정산 제외) · 참고: 전체정산 중 자점 협력점 지급분 ${won(t.JJ)}원</p>`;
  return `<div class="sum-split"><div class="kpi-col" role="tablist" aria-label="구분 선택">${kpi}</div>
    <div class="panel sum-panel"><div class="panel-head"><div><h2>${esc(title)} · ${detailSeg} <span class="pill ${segStatus[detailSeg]}">${segStatus[detailSeg]}</span></h2></div><button type="button" class="small-btn js-snap" data-name="${esc(title)}_${detailSeg}">이미지 저장</button></div>
    <div class="tbl-wrap"><table class="cmp-tbl"><thead><tr><th colspan="2">구분 / 항목</th><th class="n">상부정산</th><th class="n">전체정산</th><th class="n">합계(상부−전체)</th></tr></thead><tbody>${body}</tbody></table></div>${jj}</div></div>`;
}
function rerenderSummary(host) {
  if (host.id === 'calcSummary' && state.result) host.innerHTML = renderSummary(state.result.summary, monthLabel(state.result.period), state.segStatus);
  else if (host.id === 'runBody' && state.detail) host.innerHTML = renderSummary(state.detail.summary, state.detail.title, state.detail.segStatus) + renderUpper(byUpper(state.detail.lines));
}
document.addEventListener('click', async e => {
  const tile = e.target.closest('[data-seg]'); if (!tile) return;
  const host = tile.closest('#calcSummary, #runBody'); if (!host) return;
  const seg = tile.dataset.seg;
  const stBtn = e.target.closest('[data-st]');
  if (stBtn) {
    const to = stBtn.dataset.st;
    if (host.id === 'calcSummary') { state.segStatus[seg] = to; detailSeg = seg; rerenderSummary(host); $('btnSave').textContent = '이 달 결과 저장'; return; }
    // 저장된 회차: 바로 DB에 반영
    const d = state.detail; if (!d || d.segStatus[seg] === to) return;
    const others = state.runs.filter(r => ym(r.period) === d.period && r.id !== d.runId && segSt(r, seg) === '확정');
    if (to === '확정' && others.length && !await ask(`${monthLabel(d.period)} ${seg}은(는) 다른 회차가 확정되어 있어요. 이 회차를 확정하면 그 회차는 가정산으로 바뀌어요. 계속할까요?`)) return;
    const { error } = await sb.rpc('dist_pnl_set_seg_status', { p_run: d.runId, p_seg: seg, p_status: to });
    if (error) { toast('변경 실패: ' + error.message); return; }
    detailSeg = seg; toast(`${seg} → ${to}`); await loadRuns(); return;
  }
  detailSeg = seg; rerenderSummary(host);
});
// ---------- 이미지 저장 (구분 타일 + 결과표) ----------
let h2cLoading = null;
function loadH2C() {
  if (window.html2canvas) return Promise.resolve(window.html2canvas);
  return h2cLoading ||= new Promise((ok, no) => {
    const sc = document.createElement('script');
    sc.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js';
    sc.onload = () => ok(window.html2canvas); sc.onerror = () => { h2cLoading = null; no(new Error('이미지 도구를 불러오지 못했어요')); };
    document.head.appendChild(sc);
  });
}
document.addEventListener('click', async e => {
  const b = e.target.closest('.js-snap'); if (!b) return;
  e.stopPropagation();
  const target = b.closest('.sum-split'); if (!target) return;
  b.disabled = true; const old = b.textContent; b.textContent = '만드는 중…';
  try {
    const h2c = await loadH2C();
    document.body.classList.add('snapping');
    const canvas = await h2c(target, { backgroundColor: getComputedStyle(document.body).backgroundColor, scale: 2, useCORS: true, logging: false });
    const blob = await new Promise(ok => canvas.toBlob(ok, 'image/png'));
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `유통사업부_손익매출_${b.dataset.name.replace(/\s+/g, '')}.png`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    toast('이미지를 저장했어요');
  } catch (err) { toast(err.message || '이미지 저장 실패'); }
  finally { document.body.classList.remove('snapping'); b.disabled = false; b.textContent = old; }
}, true);
document.addEventListener('keydown', e => {
  if ((e.key === 'Enter' || e.key === ' ') && e.target.matches && e.target.matches('.kpi-btn[data-seg]')) { e.preventDefault(); e.target.click(); }
});
function renderUpper(rows) {
  const tr = rows.map(r => `<tr><td>${r.segment}</td><td>${esc(r.upper_name)}</td><td class="n">${r.cntUp.toLocaleString()}</td><td class="n">${won(r.B)}</td>
    <td class="n">${r.cntDn.toLocaleString()}</td><td class="n">${won(r.A)}</td><td class="n">${won(-r.JF)}</td><td class="n"><b>${won(r.REV)}</b></td></tr>`).join('');
  return `<div class="panel"><div class="panel-head"><div><h2>상부점별 내역</h2><p>수기 계산과 대조할 때 쓰세요</p></div></div><div class="tbl-wrap"><table class="grid">
    <tr><th>구분</th><th>상부점</th><th class="n">상부 건수</th><th class="n">상부 B</th><th class="n">전체 건수</th><th class="n">전체 A</th><th class="n">자점운영비</th><th class="n">매출</th></tr>${tr}</table></div></div>`;
}

// ---------- 월별 내역 ----------
async function loadRuns() {
  const { data, error } = await sb.from('dist_pnl_runs')
    .select('id, period, status, seg_status, note, files, summary, rules, created_by, created_at, confirmed_by, confirmed_at')
    .order('period', { ascending: false }).order('created_at', { ascending: false });
  if (error) { toast('불러오기 실패: ' + error.message); return; }
  state.runs = data; $('runCount').textContent = `${data.length}건`; renderMonths();
}
function repRun(period, seg) {
  const list = state.runs.filter(r => ym(r.period) === period);
  if (seg) return list.find(r => segSt(r, seg) === '확정') || list[0];
  return list.find(r => r.status === '확정') || list.find(r => SEGMENTS.some(k => segSt(r, k) === '확정')) || list[0];
}
const revOf = (run, k) => { const t = run.summary.segments[k]; return t.B - t.A - (t.JF || 0); }; // 저장 시점과 무관하게 B − A − 자점운영비
const man = v => { const n = Math.round((v || 0) / 1e4); return `${n.toLocaleString('ko-KR')}<small>만</small>`; };
function renderMonths() {
  const y = state.year;
  $('yLabel').textContent = `${y}`;
  const now = new Date(), nowKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  const cells = []; let nAll = 0, nPart = 0, nNone = 0;
  for (let m = 1; m <= 12; m++) {
    const p = `${y}-${String(m).padStart(2, '0')}`;
    const r = repRun(p);
    const sts = r ? SEGMENTS.map(k => { const rr = repRun(p, k); return segSt(rr, k); }) : [];
    const nConf = sts.filter(x => x === '확정').length;
    if (r) nConf === 4 ? nAll++ : nConf ? nPart++ : nNone++;
    const cls = ['mcell', r ? '' : 'empty', state.selPeriod === p ? 'sel' : '', p === nowKey ? 'now' : ''].filter(Boolean).join(' ');
    const tip = r ? SEGMENTS.map((k, i) => `${k} ${Math.round(revOf(repRun(p, k), k)).toLocaleString('ko-KR')}원 (${sts[i]})`).join('\n') : '기록 없음';
    const rev = r ? revOf(repRun(p, '유통이관'), '유통이관') : 0;
    const label = !r ? (p === nowKey ? '이번 달' : '기록 없음') : nConf === 4 ? '전체 확정' : nConf ? `확정 ${nConf}/4` : '가정산';
    cells.push(`<button class="${cls}" data-p="${p}" title="${esc(`${m}월\n${tip}`)}">
      ${r ? `<span class="mdots">${SEGMENTS.map((k, i) => `<i class="${sts[i]}" title="${k} ${sts[i]}"></i>`).join('')}</span>` : ''}
      <span class="mnum">${m}<small>월</small></span>
      <span class="mcap">${label}</span>
      ${r ? `<span class="mval${rev < 0 ? ' neg' : ''}">${man(rev)}</span><span class="mcap">유통이관 매출</span>` : ''}
    </button>`);
  }
  $('calGrid').innerHTML = cells.join('');
  $('yStat').innerHTML = `<span><i style="background:var(--ok)"></i>전체 확정 ${nAll}개월</span><span><i style="background:linear-gradient(90deg,var(--ok) 50%,#f59e0b 50%)"></i>일부 확정 ${nPart}개월</span><span><i style="background:#f59e0b"></i>가정산 ${nNone}개월</span>`;
  $('calGrid').querySelectorAll('.mcell').forEach(b => b.onclick = () => { state.selPeriod = b.dataset.p; state.selRun = null; state.scrollDetail = true; renderMonths(); });
  if (state.selPeriod && state.selPeriod.startsWith(String(y))) renderDetail(); else $('monthDetail').innerHTML = '';
}
$('yPrev').onclick = () => { state.year--; renderMonths(); };
$('yNext').onclick = () => { state.year++; renderMonths(); };
async function renderDetail() {
  const p = state.selPeriod; if (!p) return;
  const list = state.runs.filter(r => ym(r.period) === p);
  if (!list.length) { $('monthDetail').innerHTML = `<div class="panel"><div class="panel-head"><div><h2>${monthLabel(p)}</h2><p>저장된 회차가 없어요. 계산하기에서 계산한 뒤 저장하면 여기에 나와요.</p></div></div></div>`; return; }
  const run = list.find(r => r.id === state.selRun) || repRun(p);
  state.selRun = run.id;
  const stCell = r => { const n = SEGMENTS.filter(k => segSt(r, k) === '확정').length; return `<span class="pill ${n === 4 ? '확정' : '가정산'}">${n === 4 ? '전체 확정' : n ? `확정 ${n}/4` : '가정산'}</span>`; };
  const runRows = list.map(r => `<tr class="click ${r.id === run.id ? 'sel' : ''}" data-id="${r.id}"><td>${stCell(r)}</td>
    <td class="muted small">${new Date(r.created_at).toLocaleString('ko-KR')}</td><td class="small">${esc(r.created_by || '')}</td>
    ${SEGMENTS.map(k => `<td class="n">${won(revOf(r, k))} <span class="sdot ${segSt(r, k)}" title="${segSt(r, k)}"></span></td>`).join('')}<td class="small">${esc(r.note || '')}</td></tr>`).join('');
  $('monthDetail').innerHTML = `<div class="panel"><div class="panel-head"><div><h2>${monthLabel(p)} 회차</h2><p>회차를 누르면 그 회차 숫자로 바뀌어요 · 구분별 가정산/확정은 아래 타일에서 바꿔요</p></div>
      <div class="row gap wrap">
        <button id="btnExportRun">엑셀 다운로드</button><button id="btnDelRun" class="danger">회차 삭제</button></div></div>
      <div class="tbl-wrap"><table class="grid"><tr><th>상태</th><th>저장 시각</th><th>저장자</th>${SEGMENTS.map(k => `<th class="n">${k}</th>`).join('')}<th>메모</th></tr>${runRows}</table></div>
      <p class="muted small" style="margin:0">사용 파일: ${(run.files || []).map(f => `${esc(f.name)} (${f.side}, ${Number(f.rows).toLocaleString()}행)`).join(' · ')}${run.confirmed_at ? ` · 확정: ${esc(run.confirmed_by)} ${new Date(run.confirmed_at).toLocaleString('ko-KR')}` : ''}</p>
    </div><div id="runBody" class="stack-lg"><div class="panel muted">불러오는 중…</div></div>`;
  $('monthDetail').querySelectorAll('tr.click').forEach(tr => tr.onclick = () => { state.selRun = tr.dataset.id; renderDetail(); });
  if (state.scrollDetail) { state.scrollDetail = false; $('monthDetail').scrollIntoView({ behavior: 'smooth', block: 'start' }); }
  $('btnDelRun').onclick = async () => {
    if (!await ask(`${monthLabel(p)} 회차(${new Date(run.created_at).toLocaleString('ko-KR')} 저장)를 삭제할까요? 되돌릴 수 없습니다.`)) return;
    const { error } = await sb.from('dist_pnl_runs').delete().eq('id', run.id);
    if (error) { toast('삭제 실패: ' + error.message); return; }
    toast('삭제했습니다'); state.selRun = null; await loadRuns();
  };
  const { data: lines, error } = await sb.from('dist_pnl_lines').select('*').eq('run_id', run.id);
  if (error) { $('runBody').innerHTML = `<div class="panel err">${esc(error.message)}</div>`; return; }
  const summary = summarize(lines);
  const segStatus = Object.fromEntries(SEGMENTS.map(k => [k, segSt(run, k)]));
  state.detail = { summary, lines, title: monthLabel(p), segStatus, runId: run.id, period: p };
  $('runBody').innerHTML = renderSummary(summary, state.detail.title, segStatus) + renderUpper(byUpper(lines));
  $('btnExportRun').onclick = () => exportXlsx(p, summary, lines, segStatus);
}

// ---------- 엑셀 내보내기 ----------
function exportXlsx(period, summary, lines, segStatus = blankSt()) {
  const wb = XLSX.utils.book_new();
  const s = summary.segments;
  const aoa = [[`유통사업부 손익매출 ${monthLabel(period)}`], ['단위: 원 · 매출 = B − A − 자점운영비(코드)'], [], ['항목', ...SEGMENTS], ['정산 상태', ...SEGMENTS.map(k => segStatus[k])]];
  for (const [cls, lab, key, sign = 1, only] of LINES) {
    if (cls === 'sec') { aoa.push([`[${lab}]`]); continue; }
    const vals = SEGMENTS.map(k => (only === 'code' && k === '유통이관') ? null : sign * s[k][key]);
    aoa.push([lab, ...vals]);
  }
  const ws1 = XLSX.utils.aoa_to_sheet(aoa); ws1['!cols'] = [{ wch: 42 }, ...SEGMENTS.map(() => ({ wch: 16 }))];
  numFmt(ws1); XLSX.utils.book_append_sheet(wb, ws1, '요약');
  const up = byUpper(lines).map(r => ({ 구분: r.segment, 상부점: r.upper_name, '상부 건수': r.cntUp, '상부 B': r.B, '전체 건수': r.cntDn, '전체 A': r.A, '자점운영비': r.JF, '매출': r.REV }));
  const ws2 = XLSX.utils.json_to_sheet(up); ws2['!cols'] = [{ wch: 14 }, { wch: 34 }, { wch: 10 }, { wch: 16 }, { wch: 10 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];
  numFmt(ws2); XLSX.utils.book_append_sheet(wb, ws2, '상부점별');
  const raw = lines.map(l => ({ 구분: l.segment, 상부점: l.upper_name, 측: l.side, 항목: l.item, 자점구분: l.partner_type, 포함: l.included ? 'Y' : 'N', 건수: l.cnt, 정산금액: l.amount, 사은품현금: l.gift_cash, 사은품물류: l.gift_goods, 사은품쿠폰: l.gift_coupon, 자점운영비: l.jj_fee || 0 }));
  const ws3 = XLSX.utils.json_to_sheet(raw); numFmt(ws3); XLSX.utils.book_append_sheet(wb, ws3, '집계데이터');
  XLSX.writeFile(wb, `유통사업부_손익매출_${period}.xlsx`);
}
function numFmt(ws) {
  for (const a in ws) if (a[0] !== '!' && ws[a].t === 'n') ws[a].z = '#,##0;-#,##0;"-"';
}

setupBookmarklet();
announceReady();
boot();
