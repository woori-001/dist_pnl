import { createClient } from 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.45.4/+esm';
import { SUPABASE_URL, SUPABASE_KEY } from './config.js';
import { RULES, SEGMENTS, LEGACY_SEGMENTS, DIST_VIEWS, CODE_VIEWS, STATUS_KEYS, statusKeyOf } from './rules.js';
import { parseWorkbook, compute, summarize, byUpper } from './calc.js';
import { computeEval, evalForSave, EVAL_COLS, EVAL_ROWS } from './eval.js';

const sb = createClient(SUPABASE_URL, SUPABASE_KEY);
const $ = id => document.getElementById(id);
const XLSX = window.XLSX;

const prevMonth = (() => { const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - 1); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`; })();
// period = 왼쪽 달력에서 고른 달 · mode: 'upload'(파일 올리기) / 'calc'(방금 계산한 결과) / 'saved'(저장된 결과)
const state = { files: [], result: null, runs: [], period: prevMonth, mode: 'upload', selRun: null, year: +prevMonth.slice(0, 4), segStatus: null, view: null, showFiles: false, pulledAt: null };
const ST = ['가정산', '확정'];
const segSt = (run, k) => (run.seg_status && run.seg_status[k]) || run.status || '가정산';
const blankSt = () => Object.fromEntries(STATUS_KEYS.map(k => [k, '가정산']));
// 이전 기준(10/27 전) 회차는 구분 이름이 달라요 — 회차의 seg_status 키로 판별
const isLegacyRun = run => !!run && !!run.seg_status && '유통이관' in run.seg_status;
const segsOf = run => isLegacyRun(run) ? LEGACY_SEGMENTS : STATUS_KEYS; // 상태 단위
// 상태 단위별 매출: 유통사업부 = 코드매출 + 외부유통매출 (= 유통매출 + 직영영업매출)
const valOf = (run, k) => k === '유통사업부' ? revOf(run, '코드매출') + revOf(run, '외부유통매출') : revOf(run, k);
const totalOf = run => isLegacyRun(run) ? LEGACY_SEGMENTS.reduce((a, k) => a + revOf(run, k), 0) : valOf(run, '유통사업부');

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
  $('bmCard').hidden = view !== 'viewApp' || location.hash.includes('admin-import');
  $('ruleCard').hidden = view !== 'viewApp';
  $('rail').hidden = view !== 'viewApp';
  if (view === 'viewLogin') setTitle('유통 손익매출 계산기', '로그인하면 계산·저장할 수 있어요');
  if (view === 'viewApp') setTitle(...PAGE.calc);
}
const PAGE = {
  calc: ['유통사업부 코드매출·외부유통매출', ''],
  history: ['월별 내역', '달력에서 달을 누르면 그 달의 가정산·확정 내역이 열려요'],
};
function setTitle(t, sub) { $('pageTitle').textContent = t; $('pageSub').textContent = sub; $('pageSub').hidden = !sub; }

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
  await loadRuns();
  if (pendingAdmin) applyAdmin(); else openMonth(state.period);
}
$('formLogin').addEventListener('submit', async e => {
  e.preventDefault(); $('loginErr').textContent = '';
  const { data, error } = await sb.auth.signInWithPassword({ email: $('loginEmail').value.trim(), password: $('loginPw').value });
  if (error) { $('loginErr').textContent = '로그인 실패: 이메일 또는 비밀번호를 확인하세요'; return; }
  onSession(data.session);
});
$('btnLogout').onclick = async () => { await sb.auth.signOut(); onSession(null); };

// ---------- 파일 올리기 ----------
const drop = $('drop');
drop.onclick = () => $('fileInput').click();
drop.onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') $('fileInput').click(); };
drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
drop.ondragleave = () => drop.classList.remove('over');
drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); addFiles(e.dataTransfer.files); };
$('fileInput').onchange = e => { addFiles(e.target.files); e.target.value = ''; };
$('btnClear').onclick = () => { state.files = []; state.result = null; renderFiles(); refresh(); };
$('btnToggleFiles').onclick = () => { state.showFiles = !state.showFiles; refresh(); };
$('btnNew').onclick = async () => {
  if (state.result && state.mode === 'calc' && !state.resultSaved && !await ask('저장하지 않은 계산 결과가 있어요. 지우고 새로 계산할까요?')) return;
  state.files = []; state.result = null; state.mode = 'upload'; state.showFiles = false; renderFiles(); refresh();
};

async function addFiles(fileList) {
  const list = [...fileList]; // FileList는 input 초기화 시 비워지므로 먼저 복사
  for (const f of list) {
    if (state.files.some(x => x.fileName === f.name && x.size === f.size)) { toast(`${f.name}: 이미 올린 파일입니다`); continue; }
    try {
      const parsed = parseWorkbook(XLSX, new Uint8Array(await f.arrayBuffer()), f.name);
      parsed.size = f.size; state.files.push(parsed);
    } catch (err) { toast(err.message); }
  }
  state.result = null; state.mode = 'upload'; renderFiles(); refresh();
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
  $('fileList').querySelectorAll('[data-rm]').forEach(b => b.onclick = () => { state.files.splice(+b.dataset.rm, 1); state.result = null; state.mode = 'upload'; renderFiles(); refresh(); });
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
  $('bmCard').hidden = true; $('rail').hidden = true; $('viewApp').classList.add('importing');
  adminWait('어드민에서 데이터를 받는 중이에요', '어드민 창에서 엑셀 10개(옛날 데이터 포함)를 만들고 있어요. 1~2분 걸려요. 어드민 창을 닫지 마세요.');
  clearTimeout(waitTimer);
  waitTimer = setTimeout(() => adminWait('아직 데이터가 오지 않았어요', '어드민 창의 작은 창에 오류가 떠 있는지 확인해 주세요. 엑셀을 직접 올려도 됩니다.', { error: true }) || $('viewApp').classList.remove('importing'), 4 * 60 * 1000);
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
    if (e.data.error) { adminWait('어드민에서 가져오지 못했어요', String(e.data.text || ''), { error: true }); $('viewApp').classList.remove('importing'); }
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
  $('adminWait').hidden = true; $('viewApp').classList.remove('importing');
  state.files = files; state.period = period; state.year = +period.slice(0, 4); state.pulledAt = pulledAt || new Date().toISOString();
  renderFiles();
  $('btnCalc').click();
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
  const period = state.period;
  if (!period) { toast('왼쪽 달력에서 달을 고르세요'); return; }
  const { lines, checks } = compute(state.files, period);
  const summary = summarize(lines);
  state.result = { period, lines, checks, summary, pulledAt: state.pulledAt }; state.pulledAt = null; state.resultSaved = false;
  state.segStatus = blankSt();
  const warn = [...checks.warnings];
  if (checks.outOfPeriod) warn.push(`대상월이 아닌 행 ${checks.outOfPeriod.toLocaleString()}건은 계산에서 뺐어요 (월을 잘못 골랐거나 다른 달 파일이 섞였는지 확인)`);
  state.result.warn = warn;
  state.giftChoices = {}; state.result.eval = computeEval(state.files, period, state.giftChoices);
  state.mode = 'calc'; state.showFiles = false; refresh();
};

$('btnSave').onclick = async () => {
  const r = state.result; if (!r) return;
  const st = { ...state.segStatus };
  const periodDate = r.period + '-01';
  const existing = state.runs.filter(x => ym(x.period) === r.period && !isLegacyRun(x));
  const clash = STATUS_KEYS.filter(k => st[k] === '확정' && existing.some(x => segSt(x, k) === '확정'));
  if (clash.length && !await ask(`${monthLabel(r.period)} ${clash.join(', ')}은(는) 이미 확정된 회차가 있어요. 이번 계산을 확정으로 저장하면 이전 확정은 가정산으로 바뀌어요. 계속할까요?`)) return;
  $('btnSave').disabled = true;
  const files = state.files.map(f => ({ name: f.fileName, side: f.side, rows: f.rows.length, total: f.fileTotal }));
  const rules = { version: RULES.version, checks: { usedRows: r.checks.usedRows, outOfPeriod: r.checks.outOfPeriod, otherDept: r.checks.otherDept, warnings: r.checks.warnings } };
  const { error } = await sb.rpc('dist_pnl_save_run2', { p_period: periodDate, p_seg_status: st, p_note: null,
    p_files: files, p_summary: { ...r.summary, eval: r.eval ? evalForSave(r.eval) : null }, p_rules: rules, p_lines: r.lines });
  $('btnSave').disabled = false;
  if (error) { toast('저장 실패: ' + error.message); return; }
  state.resultSaved = true;
  toast(`${monthLabel(r.period)} 가정산으로 저장했어요 — 달력에서 이 달을 누르면 다시 볼 수 있어요`);
  await loadRuns(); state.selRun = null; openMonth(r.period);
};

// ---------- 요약 화면 ----------
const LINES = [ // 엑셀 내보내기용 (전체 항목)
  ['sec', '전체정산 · 우리가 준 돈'],
  ['', 'a. 수수료 (사은품 차감전, 기타 차감 반영)', 'D_A'], ['', 'b. 사은품 - 현금', 'D_B'], ['', 'c. 사은품 - 물류', 'D_C'], ['', 'd. 사은품 - 쿠폰', 'D_D'], ['', 'f. 환수', 'D_F'],
  ['b', '최종 정산금액 (A) = a − (b+c+d+f)', 'A'],
  ['sec', '상부정산 · 우리가 받은 돈'],
  ['', 'a. 수수료 (a-1·a-2 제외)', 'U_A'], ['', 'a-1. 관리수수료', 'U_A1'], ['', 'a-2. 정산관련', 'U_A2'], ['', 'b. 사은품 - 현금', 'U_B'], ['', 'c. 사은품 - 물류', 'U_C'], ['', 'd. 사은품 - 쿠폰', 'U_D'], ['', 'f. 환수', 'U_F'],
  ['b', '최종 정산금액 (B) = (a+a1+a2) − (b+c+d+f)', 'B'],
  ['sec', '매출'],
  ['', '(−) 자점운영비 (코드 상부 · 자점 일반 건) — 유통사업부 4분류는 별도 표기, 매출에 미반영', 'JF', -1],
  ['', '(−) 직접지급사은품 (웨이브넷 지급분)', 'GD', -1],
  ['key', '매출 = B − A − 자점운영비 − 직접지급사은품', 'REV'],
  ['ref', '참고: 전체정산 중 자점 협력점 지급분', 'JJ', 1],
];
const SEG_NOTE = { '코드매출': '협력점·자점 / 코드', '외부유통매출': '협력점·자점 / 이관', '유통매출': '협력점 / 코드·이관', '직영영업매출': '자점 / 코드·이관',
  'SKB웨이브넷': '코드 · 자점 포함', 'SKY우리정보': '코드 · 자점 포함', 'SK매직': '코드 · 자점 포함' };
let detailSeg = SEGMENTS[0];
const SHOW_VIEWS = ['코드매출', '외부유통매출']; // 화면에 보이는 유통사업부 항목 (유통매출·직영영업매출·SKB·SKY·매직은 계산만 하고 숨김)
function renderSummary(summary, title, segStatus = blankSt(), opts = {}) {
  const s = summary.segments;
  const keys = Object.keys(s), legacy = !!summary.legacy;
  const allowed = legacy ? keys : SHOW_VIEWS;
  if (!allowed.includes(detailSeg)) detailSeg = allowed[0];
  const kv = v => `${Math.round(v || 0).toLocaleString('ko-KR')}<small>원</small>`;
  const man1 = v => `${Math.round(v || 0).toLocaleString('ko-KR')}`;
  const stt = k => (legacy || opts.toggles === false) ? '' : `<div class="stt" role="group" aria-label="${k} 정산 상태">${ST.map(x => `<button type="button" data-st="${x}" class="${segStatus[k] === x ? 'on ' + x : ''}" aria-pressed="${segStatus[k] === x}">${x}</button>`).join('')}</div>`;
  const jfNote = k => s[k].JF ? ` · 자점운영비 −${man1(s[k].JF)}` : '';
  const tile = k => `<div class="kpi kpi-btn${k === detailSeg ? ' hot' : ''}" data-seg="${k}" ${legacy ? '' : `data-stk="${k}"`} role="tab" tabindex="0" aria-selected="${k === detailSeg}">
      <div class="kpi-top"><div class="kpi-label">${k}</div>${stt(k)}</div>
      <div class="kpi-value">${kv(s[k].REV)}</div>
      <div class="kpi-note">${legacy ? '이전 기준' : (SEG_NOTE[k] || '')}${jfNote(k)}</div></div>`;
  let left;
  const tot = summary.total || {};
  if (legacy) left = keys.map(tile).join('');
  else {
    const same = Math.round(tot.REV) === Math.round(tot.REV2);
    const cells = tot.cells || {};
    const PARTS = { '코드매출': ['협력점|코드', '자점|코드'], '외부유통매출': ['협력점|이관', '자점|이관'], '유통매출': ['협력점|코드', '협력점|이관'], '직영영업매출': ['자점|코드', '자점|이관'] };
    const sub = k => `<button type="button" class="kpi-sub${k === detailSeg ? ' hot' : ''}" data-seg="${k}" role="tab" aria-selected="${k === detailSeg}">
        <span class="sh"><span class="sl">${k}</span><span class="sv">${man1(s[k].REV)}</span></span>
        ${PARTS[k].map(c => `<span class="sp"><span>${c.replace('|', ' > ')}</span><span>${man1(cells[c])}</span></span>`).join('')}</button>`;
    const sumLine = (lab, v, ok) => `<div class="sub-sum"><span>${lab} 합계</span><b>${man1(v)}</b>${ok === undefined ? '' : ok ? '<i>✓</i>' : '<i class="bad">≠</i>'}</div>`;
    left = `<div class="kpi-grp${DIST_VIEWS.includes(detailSeg) ? ' on' : ''}" data-stk="유통사업부">
        <div class="kpi-top"><div class="kpi-label">유통사업부</div>${stt('유통사업부')}</div>
        <div class="kpi-value">${kv(tot.REV)}</div>
        <div class="kpi-subs">${SHOW_VIEWS.map(sub).join('')}</div>
      </div>`;
  }
  const t = s[detailSeg];
  const stKey = legacy ? null : statusKeyOf(detailSeg);
  // 상부·전체를 같은 항목으로 나란히 (차감 항목은 −로 표시)
  const rows = [
    ['수수료', '일반', t.U_A, t.D_A0], ['', '관리', t.U_A1, t.D_A1], ['', '정산관련', t.U_A2, t.D_A2],
    ['사은품', '현금', -t.U_B, -t.D_B], ['', '물류', -t.U_C, -t.D_C], ['', '쿠폰', -t.U_D, -t.D_D],
  ];
  const grpSpan = { '수수료': 3, '사은품': 3 };
  const body = rows.map(([g, item, up, dn]) => `<tr>${g ? `<th class="grp" rowspan="${grpSpan[g]}">${g}</th>` : ''}<td class="item">${item}</td>
      <td class="n">${won(up)}</td><td class="n">${won(dn)}</td><td class="n diff">${won(up - dn)}</td></tr>`).join('')
    + `<tr><th class="grp" colspan="2">환수</th><td class="n">${won(-t.U_F)}</td><td class="n">${won(-t.D_F)}</td><td class="n diff">${won(-t.U_F + t.D_F)}</td></tr>`
    + (t.JF_OUT ? '' : `<tr><th class="grp" colspan="2">자점운영비</th><td class="n">${won(-t.JF)}</td><td class="n">-</td><td class="n diff">${won(-t.JF)}</td></tr>`)
    + `<tr><th class="grp" colspan="2">직접지급사은품</th><td class="n">-</td><td class="n">${won(t.GD)}</td><td class="n diff">${won(-t.GD)}</td></tr>`
    + `<tr class="total"><th colspan="2">합계</th><td class="n">${won(t.B - (t.JF_OUT ? 0 : (t.JF || 0)))}</td><td class="n">${won(t.A + t.GD)}</td><td class="n rev">${won(t.REV)}</td></tr>`;
  const jfDef = `코드 상부 아래 자점 '일반' 건의 자점 마진(접수일 기준 금액, 7.유통사업부·0.영업정산 제외)`;
  const jj = t.JF_OUT && t.JF
    ? `<div class="jf-out"><div><b>자점운영비</b><span>${won(t.JF)}원</span></div><p>※ 위 표 금액에는 반영되지 않았어요 (별도 표기) · ${jfDef}</p></div>`
    : !t.JF || t.JF_OUT ? '' : `<p class="muted small" style="margin:0">자점운영비 = ${jfDef}</p>`;
  const name = DIST_VIEWS.includes(detailSeg) ? `유통사업부 ${detailSeg}` : detailSeg;
  const head = legacy ? `<div class="notice warn">이전 기준(유통이관·SKB·SKY·매직)으로 저장된 회차예요. 유통사업부 4분류로 보려면 이 달을 다시 계산해서 저장하세요.</div>` : '';
  return `${head}<div class="sum-split"><div class="kpi-col" role="tablist" aria-label="구분 선택">${left}</div>
    <div class="panel sum-panel"><div class="panel-head"><div><h2>${esc(title)} · ${name}${stKey ? ` <span class="pill ${segStatus[stKey]}">${segStatus[stKey]}</span>` : ''}</h2></div></div>
    <div class="tbl-wrap"><table class="cmp-tbl"><thead><tr><th colspan="2">구분 / 항목</th><th class="n">상부정산</th><th class="n">전체정산</th><th class="n">합계(상부−전체)</th></tr></thead><tbody>${body}</tbody></table></div>${jj}</div></div>`;
}
function rerenderSummary() {
  const v = state.view; $('calcSummary').innerHTML = v ? renderSummary(v.summary, v.title, v.segStatus, { toggles: false }) : '';
}
document.addEventListener('click', e => {
  if (!e.target.closest('#calcSummary')) return;
  const tile = e.target.closest('[data-seg]');
  if (tile) { detailSeg = tile.dataset.seg; rerenderSummary(); }
});
// ---------- 평가손익 입력값 ----------
const PICK = [['ded', '차감 금액'], ['paid', '지급 금액'], ['none', '넣지 않음']];
function renderEval(ev, title, live) {
  const n = v => v == null ? '-' : Math.round(v).toLocaleString('ko-KR');
  const cell = v => `<td class="n"><button type="button" class="copy-n" data-copy="${Math.round(v || 0)}" title="눌러서 복사">${won(v)}</button></td>`;
  const head = EVAL_COLS.map(c => `<th class="n${c[0] === 'expect' ? ' ref' : ''}" title="${esc(c[2])}">${c[1]}</th>`).join('');
  const row = (lab, t, cls = '') => `<tr class="${cls}"><th>${lab}</th>${EVAL_COLS.map(c => c[0] === 'expect' ? `<td class="n ref">${won(t[c[0]])}</td>` : cell(t[c[0]])).join('')}</tr>`;
  const notes = [];
  if (ev.ratesMissing && ev.ratesMissing.length) notes.push(`기대관리 단가 확인 중 — ${ev.ratesMissing.join(', ')} (금액 0으로 계산, 단가가 정해지면 rules.js에 넣으면 돼요)`);
  if (ev.expectHold) notes.push(`기대관리 보류 ${ev.expectHold.toLocaleString()}건 — 약정갱신·업셀링/전환 상품 (넣을지 확인 중)`);
  const pend = live ? ev.giftPending : ev.gift && ev.gift.pending;
  if (pend) notes.push(`사은품 확인 ${pend}건이 아직 '넣지 않음'이에요 — 아래 목록에서 넣을 금액을 골라 주세요`);
  const exp = (ev.expectRows || []).map(e => `<tr><td>${e.carrier}</td><td>${e.type}</td><td class="n">${e.cnt.toLocaleString()}</td><td class="n">${e.rate == null ? '<span class="muted">확인 중</span>' : n(e.rate)}</td><td class="n">${e.months}</td><td class="n">${e.amount == null ? '<span class="muted">-</span>' : won(e.amount)}</td></tr>`).join('');
  let gifts = '';
  if (live && ev.giftList.length) {
    const sum = (k) => ev.giftList.reduce((s, g) => s + (g[k] || 0), 0);
    gifts = `<details class="panel ev-more" ${ev.giftPending ? 'open' : ''}><summary><b>사은품 확인 ${ev.giftList.length}건</b><span class="muted small"> · 차감_사은품과 사은품 지급(현금+지급완료일)이 안 맞는 건 · 고른 금액이 사은품에 들어가요</span></summary>
      <div class="row gap wrap ev-bulk"><span class="small muted">한꺼번에:</span>${PICK.map(([k, l]) => `<button type="button" class="small-btn" data-giftall="${k}">모두 ${l}</button>`).join('')}
        <span class="small muted">차감 합 ${n(sum('ded'))} · 지급 합 ${n(sum('paid'))}</span></div>
      <div class="tbl-wrap"><table class="grid ev-gift"><tr><th>구분</th><th>내용</th><th>협력점</th><th>고객</th><th>개통일</th><th>상품</th><th>지급처</th><th>지급완료</th><th class="n">차감</th><th class="n">지급</th><th>넣을 금액</th></tr>
      ${ev.giftList.map((g, i) => `<tr><td>${g.seg === '코드매출' ? '코드' : '외부'}</td><td class="small">${esc(g.why)}</td><td class="small">${esc(g.partner)}</td><td class="small">${esc(g.customer)}</td><td class="small">${esc(g.openDate)}</td><td class="small">${esc(g.product)}</td><td class="small">${esc(g.payer)}</td><td class="small">${esc(g.done || '-')}</td><td class="n">${n(g.ded)}</td><td class="n">${g.paid == null ? '-' : n(g.paid)}</td>
        <td><div class="pick" role="radiogroup">${PICK.filter(([k]) => k !== 'paid' || g.paid != null).map(([k, l]) => `<label><input type="radio" name="gp${i}" data-gift="${esc(g.key)}" value="${k}" ${g.pick === k ? 'checked' : ''}>${l}</label>`).join('')}</div></td></tr>`).join('')}</table></div></details>`;
  } else if (!live && ev.gift && ev.gift.checked) {
    const pk = ev.gift.picks || {};
    gifts = `<p class="muted small" style="margin:0">사은품 확인 ${ev.gift.checked}건 — 차감 금액 ${pk.ded || 0} · 지급 금액 ${pk.paid || 0} · 넣지 않음 ${pk.none || 0}건으로 저장됨</p>`;
  }
  return `<div class="panel ev-panel"><div class="panel-head"><div><h2>${esc(title)} · 평가손익 입력값 <span class="pill">유통사업부</span></h2><p>경영지원 '배분 전·배분 후 손익'의 유통사업부 › 코드매출 / 외부유통매출 칸에 넣을 숫자예요 · 숫자를 누르면 복사돼요</p></div></div>
    ${notes.length ? `<div class="warns">${notes.map(x => `<div class="notice warn">⚠ ${esc(x)}</div>`).join('')}</div>` : ''}
    <div class="tbl-wrap"><table class="grid ev-tbl"><tr><th>구분</th>${head}</tr>${EVAL_ROWS.map(k => row(k, ev.rows[k])).join('')}${row('유통사업부 합계', ev.total, 'total')}</table></div>
    <p class="muted small" style="margin:0">협력점이 자점인 건은 빼고(0.영업정산은 포함), 사은품과 기대관리만 자점을 포함해요 · 발생수수료 = 상부정산 일반·추가(관리수수료 제외, 정산관련 포함) + 기대관리 · 사은품 = 현금+물류+쿠폰(일반 건, 유치자지급 제외)</p>
    <details class="ev-more"><summary><b>기대관리 내역</b><span class="muted small"> · 상부정산 코드 · 일반 · 해지 제외 · 자점 포함 · 건수 × 월 단가 × ${RULES.expectMonths}개월</span></summary>
      <div class="tbl-wrap"><table class="grid"><tr><th>통신사</th><th>유형</th><th class="n">건수</th><th class="n">월 단가</th><th class="n">개월</th><th class="n">기대관리</th></tr>${exp || '<tr><td colspan="6" class="muted">해당 없음</td></tr>'}</table></div></details>
  </div>${gifts}`;
}
function renderEvalCalc() {
  const r = state.result; if (!r || !r.eval || state.mode !== 'calc') return;
  $('calcEval').innerHTML = renderEval(r.eval, monthLabel(r.period), true);
}
function recomputeEval() {
  const r = state.result; if (!r) return;
  const keepOpen = [...$('calcEval').querySelectorAll('details')].map(d => d.open);
  r.eval = computeEval(state.files, r.period, state.giftChoices); renderEvalCalc(); state.resultSaved = false; renderRail();
  [...$('calcEval').querySelectorAll('details')].forEach((d, i) => { if (keepOpen[i] !== undefined) d.open = keepOpen[i]; });
}
document.addEventListener('change', e => {
  const el = e.target.closest('input[data-gift]'); if (!el) return;
  state.giftChoices[el.dataset.gift] = el.value; recomputeEval();
});
document.addEventListener('click', async e => {
  const all = e.target.closest('[data-giftall]');
  if (all && state.result && state.result.eval) {
    for (const g of state.result.eval.giftList) state.giftChoices[g.key] = (all.dataset.giftall === 'paid' && g.paid == null) ? 'none' : all.dataset.giftall;
    recomputeEval(); return;
  }
  const c = e.target.closest('.copy-n'); if (!c) return;
  try { await navigator.clipboard.writeText(c.dataset.copy); toast(`${Number(c.dataset.copy).toLocaleString('ko-KR')} 복사했어요`); }
  catch { toast('복사하지 못했어요 — 숫자를 직접 선택해 복사하세요'); }
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
function loadScript(src, globalName) {
  if (window[globalName]) return Promise.resolve(window[globalName]);
  return new Promise((ok, no) => {
    const sc = document.createElement('script'); sc.src = src;
    sc.onload = () => ok(window[globalName]); sc.onerror = () => no(new Error('도구를 불러오지 못했어요'));
    document.head.appendChild(sc);
  });
}
async function captureSplit(el) {
  const h2c = await loadH2C();
  document.body.classList.add('snapping');
  try {
    const canvas = await h2c(el, { backgroundColor: getComputedStyle(document.body).backgroundColor, scale: 2, useCORS: true, logging: false });
    return await new Promise(ok => canvas.toBlob(ok, 'image/png'));
  } finally { document.body.classList.remove('snapping'); }
}
function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}
const fileTag = t => String(t).replace(/\s+/g, '').replace(/[\\/:*?"<>|()]/g, '');
document.addEventListener('click', async e => {
  const b = e.target.closest('.js-snap'); if (!b) return;
  e.stopPropagation();
  const target = b.closest('.sum-panel'); if (!target) return;
  b.disabled = true; const old = b.textContent; b.textContent = '만드는 중…';
  try {
    downloadBlob(await captureSplit(target), `유통사업부_손익매출_${fileTag(b.dataset.name)}.png`);
    toast('이미지를 저장했어요');
  } catch (err) { toast(err.message || '이미지 저장 실패'); }
  finally { b.disabled = false; b.textContent = old; }
}, true);
/** 4개 구분 이미지를 한 번에 → ZIP 파일 하나로 저장 */
async function snapAll(hostId, title, btn) {
  const host = $(hostId); if (!host || !host.querySelector('.sum-panel')) return;
  const keep = detailSeg; const old = btn.textContent; btn.disabled = true;
  try {
    const JSZip = await loadScript('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js', 'JSZip');
    const zip = new JSZip();
    const keys = [...host.querySelectorAll('.kpi-col [data-seg]')].map(x => x.dataset.seg);
    for (const [i, seg] of keys.entries()) {
      btn.textContent = `만드는 중 ${i + 1}/${keys.length}`;
      detailSeg = seg; rerenderSummary(host);
      await new Promise(r => requestAnimationFrame(() => setTimeout(r, 60)));
      zip.file(`${i + 1}_${DIST_VIEWS.includes(seg) ? '유통사업부_' : ''}${seg}.png`, await captureSplit(host.querySelector('.sum-panel')));
    }
    downloadBlob(await zip.generateAsync({ type: 'blob' }), `유통사업부_손익매출_${fileTag(title)}_이미지.zip`);
    toast(`${keys.length}개 구분 이미지를 ZIP으로 저장했어요`);
  } catch (err) { toast(err.message || '이미지 저장 실패'); }
  finally { detailSeg = keep; rerenderSummary(host); btn.disabled = false; btn.textContent = old; }
}
document.addEventListener('keydown', e => {
  if ((e.key === 'Enter' || e.key === ' ') && e.target.matches && e.target.matches('div.kpi-btn[data-seg]')) { e.preventDefault(); e.target.click(); }
});
function renderUpper(rows) {
  const tr = rows.map(r => `<tr><td>${r.segment}</td><td>${esc(r.upper_name)}</td><td>${r.partner_type}</td><td class="n">${r.cntUp.toLocaleString()}</td><td class="n">${won(r.B)}</td>
    <td class="n">${r.cntDn.toLocaleString()}</td><td class="n">${won(r.A)}</td><td class="n">${won(-r.JF)}</td><td class="n">${won(-r.GD)}</td><td class="n"><b>${won(r.REV)}</b></td></tr>`).join('');
  return `<div class="panel"><div class="panel-head"><div><h2>상부점별 내역</h2><p>상부 구분 · 상부점 · 하부(협력점/자점)별 — 수기 계산과 대조할 때 쓰세요</p></div></div><div class="tbl-wrap"><table class="grid">
    <tr><th>상부</th><th>상부점</th><th>하부</th><th class="n">상부 건수</th><th class="n">상부 B</th><th class="n">전체 건수</th><th class="n">전체 A</th><th class="n">자점운영비</th><th class="n">직접지급사은품</th><th class="n">매출</th></tr>${tr}</table></div></div>`;
}

// ---------- 달력 · 저장된 결과 ----------
async function loadRuns() {
  const { data, error } = await sb.from('dist_pnl_runs')
    .select('id, period, status, seg_status, note, files, summary, rules, created_by, created_at, confirmed_by, confirmed_at')
    .order('period', { ascending: false }).order('created_at', { ascending: false });
  if (error) { toast('불러오기 실패: ' + error.message); return; }
  state.runs = data; $('runCount').textContent = `${data.length}건`; renderMonths();
}
function repRun(period, seg) {
  const all = state.runs.filter(r => ym(r.period) === period);
  const cur = all.filter(r => !isLegacyRun(r));
  const list = cur.length ? cur : all; // 새 기준 회차가 있으면 그것만, 없으면 이전 기준 회차
  if (seg) return list.find(r => segSt(r, seg) === '확정') || list[0];
  return list.find(r => r.status === '확정') || list.find(r => segsOf(r).some(k => segSt(r, k) === '확정')) || list[0];
}
const revOf = (run, k) => { const t = run.summary.segments[k]; const jf = DIST_VIEWS.includes(k) ? 0 : (t.JF || 0); return t.B - t.A - jf - (t.GD ?? ((t.D_B || 0) + (t.D_C || 0) + (t.D_D || 0))); }; // B − A − 자점운영비(유통사업부는 별도라 미차감) − 직접지급사은품
const monthSt = run => !run ? null : isLegacyRun(run) ? '이전' : segSt(run, '유통사업부');
function renderMonths() {
  const y = state.year;
  $('yLabel').textContent = `${y}년`;
  const nowKey = (() => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`; })();
  const cells = [];
  for (let m = 1; m <= 12; m++) {
    const p = `${y}-${String(m).padStart(2, '0')}`;
    const r = repRun(p), st = monthSt(r);
    const calc = state.result && state.result.period === p && !state.resultSaved;
    const future = p > nowKey;
    const dot = calc ? 'calc' : st === '확정' ? '확정' : st ? '가정산' : '';
    const tip = calc ? '계산 중 (저장 안 함)' : !r ? (future ? '' : '저장된 결과 없음') : st === '이전' ? '이전 기준으로 저장됨' : `${st} · ${Math.round(valOf(r, '유통사업부')).toLocaleString('ko-KR')}원`;
    cells.push(`<button type="button" class="cm-m${p === state.period ? ' on' : ''}${p === nowKey ? ' now' : ''}" data-p="${p}" ${future ? 'disabled' : ''} title="${esc(`${m}월${tip ? ' · ' + tip : ''}`)}">${m}월${dot ? `<i class="${dot}"></i>` : ''}</button>`);
  }
  $('calGrid').innerHTML = cells.join('');
  $('calGrid').querySelectorAll('.cm-m').forEach(b => b.onclick = () => { state.selRun = null; openMonth(b.dataset.p); });
}
$('yPrev').onclick = () => { state.year--; renderMonths(); };
$('yNext').onclick = () => { state.year++; renderMonths(); };

/** 달 고르기: 방금 계산한 달이면 그 결과, 저장된 결과가 있으면 그것, 없으면 파일 올리기 */
async function openMonth(p) {
  state.period = p; state.year = +p.slice(0, 4); state.showFiles = false;
  if (state.result && state.result.period === p && !state.resultSaved) { state.mode = 'calc'; refresh(); return; }
  const list = state.runs.filter(r => ym(r.period) === p);
  if (!list.length) { state.mode = 'upload'; state.view = null; refresh(); return; }
  const run = list.find(r => r.id === state.selRun) || repRun(p);
  state.selRun = run.id; state.mode = 'saved'; state.view = { loading: true, run, title: monthLabel(p) }; refresh();
  const { data: lines, error } = await sb.from('dist_pnl_lines').select('*').eq('run_id', run.id);
  if (state.selRun !== run.id || state.mode !== 'saved') return; // 그 사이 다른 달을 눌렀으면 무시
  if (error) { state.view = { error: error.message, run, title: monthLabel(p) }; refresh(); return; }
  const summary = summarize(lines);
  state.view = { run, lines, summary, title: monthLabel(p), segStatus: Object.fromEntries(segsOf(run).map(k => [k, segSt(run, k)])), eval: run.summary && run.summary.eval };
  refresh();
}

/** 화면 전체 다시 그리기 (왼쪽 정보 + 오른쪽 결과) */
function refresh() {
  renderMonths(); renderRail();
  const m = state.mode, r = state.result;
  const showUpload = m === 'upload' || state.showFiles;
  $('uploadCard').hidden = !showUpload;
  $('uploadTitle').textContent = m === 'upload' ? `${monthLabel(state.period)} · 엑셀로 계산하기` : '계산에 쓴 파일';
  $('uploadSub').textContent = m === 'upload' ? '어드민 가져오기 버튼을 쓰면 자동으로 계산돼요. 엑셀을 직접 받았다면 여기에 올리세요.' : '파일을 빼거나 더 올린 뒤 다시 계산할 수 있어요.';
  if (m === 'calc' && r) {
    state.view = { summary: r.summary, title: monthLabel(r.period), segStatus: state.segStatus };
    $('calcChecks').innerHTML = r.warn.length ? `<div class="warns">${r.warn.map(w => `<div class="notice warn">⚠ ${esc(w)}</div>`).join('')}</div>` : '';
    renderEvalCalc(); rerenderSummary();
  } else if (m === 'saved' && state.view) {
    const v = state.view;
    $('calcChecks').innerHTML = v.error ? `<div class="notice warn">⚠ ${esc(v.error)}</div>` : '';
    $('calcEval').innerHTML = v.loading ? '<div class="panel muted">불러오는 중…</div>' : v.eval ? renderEval(v.eval, v.title, false)
      : (v.summary && !v.summary.legacy ? '<div class="notice warn">⚠ 평가손익 입력값이 생기기 전에 저장한 결과예요. 이 달을 다시 계산해서 저장하면 나와요.</div>' : '');
    if (v.summary) rerenderSummary(); else $('calcSummary').innerHTML = '';
  } else { $('calcChecks').innerHTML = ''; $('calcEval').innerHTML = ''; $('calcSummary').innerHTML = ''; }
}

/** 왼쪽 버튼 · 정보 카드 */
function renderRail() {
  const m = state.mode, r = state.result;
  $('btnToggleFiles').hidden = m === 'upload';
  $('btnToggleFiles').textContent = state.showFiles ? '파일 목록 접기' : '파일 목록 보기';
  $('btnToggleFiles').disabled = m === 'saved';
  $('btnSave').hidden = m !== 'calc';
  $('btnSave').disabled = state.resultSaved;
  $('btnNew').hidden = m === 'upload' && !state.files.length;
  if (m === 'saved') $('btnToggleFiles').hidden = true;
  let html = '';
  if (m === 'upload') {
    html = `<h3>${monthLabel(state.period)}</h3><p class="muted small">저장된 결과가 없어요.<br>어드민에서 <b>유통손익 가져오기</b>를 누르거나 오른쪽에 엑셀을 올려 계산하세요.</p>`;
  } else if (m === 'calc' && r) {
    const c = r.checks, allOk = c.files.every(f => f.totalOk);
    html = `<h3>${monthLabel(r.period)}</h3><div class="ri-sub">계산에 쓴 파일 ${state.files.length}개</div>
      ${r.pulledAt ? `<div class="ri-line">어드민에서 가져옴<br><span class="muted">${new Date(r.pulledAt).toLocaleString('ko-KR')}</span></div>` : ''}
      <div class="ri-checks">${r.warn.length ? `<span class="neg">⚠ 확인할 것 ${r.warn.length}개</span>` : '<span class="okline">✓ 점검 통과</span>'}<span>${allOk ? '파일 합계 모두 일치' : '<b class="neg">파일 합계 불일치</b>'}</span><span>계산에 쓴 행 ${c.usedRows.toLocaleString()}건</span></div>
      <div class="ri-state ${state.resultSaved ? 'ok' : 'warn'}">${state.resultSaved ? '✓ 저장됨' : '아직 저장 안 함 — 💾 버튼으로 저장'}</div>`;
  } else if (m === 'saved' && state.view) {
    const v = state.view, run = v.run, list = state.runs.filter(x => ym(x.period) === state.period);
    const st = monthSt(run), legacy = st === '이전';
    html = `<h3>${monthLabel(state.period)}</h3><div class="ri-sub">저장된 결과${list.length > 1 ? ` · ${list.length}회 저장` : ''}</div>
      <div class="ri-big"><span>${won(legacy ? totalOf(run) : valOf(run, '유통사업부'))}<small>원</small></span><span class="muted small">${legacy ? '이전 기준 합계' : '유통사업부 매출'}</span></div>
      ${legacy ? '<div class="ri-state warn">이전 기준으로 저장된 결과예요 — 다시 계산해서 저장하세요</div>'
        : `<div class="stt big" role="group" aria-label="정산 상태">${ST.map(x => `<button type="button" data-runst="${x}" class="${st === x ? 'on ' + x : ''}" aria-pressed="${st === x}">${x}</button>`).join('')}</div>`}
      <div class="ri-line muted">${new Date(run.created_at).toLocaleString('ko-KR')} 저장<br>${esc(run.created_by || '')}${run.confirmed_at ? `<br>확정 ${new Date(run.confirmed_at).toLocaleString('ko-KR')}` : ''}</div>
      ${list.length > 1 ? `<div class="ri-runs"><div class="small muted">저장 기록</div>${list.map(x => `<button type="button" class="ri-run${x.id === run.id ? ' on' : ''}" data-run="${x.id}"><span>${new Date(x.created_at).toLocaleString('ko-KR', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' })}</span><i class="${monthSt(x) === '확정' ? '확정' : '가정산'}"></i></button>`).join('')}</div>` : ''}
      <button type="button" class="link danger small" id="btnDelRun">이 저장 기록 삭제</button>`;
  }
  $('railInfo').innerHTML = html;
  const del = $('btnDelRun');
  if (del) del.onclick = async () => {
    const run = state.view.run;
    if (!await ask(`${monthLabel(state.period)} ${new Date(run.created_at).toLocaleString('ko-KR')}에 저장한 결과를 삭제할까요? 되돌릴 수 없습니다.`)) return;
    const { error } = await sb.from('dist_pnl_runs').delete().eq('id', run.id);
    if (error) { toast('삭제 실패: ' + error.message); return; }
    toast('삭제했습니다'); state.selRun = null; await loadRuns(); openMonth(state.period);
  };
  $('railInfo').querySelectorAll('[data-run]').forEach(b => b.onclick = () => { state.selRun = b.dataset.run; openMonth(state.period); });
  $('railInfo').querySelectorAll('[data-runst]').forEach(b => b.onclick = async () => {
    const run = state.view.run, to = b.dataset.runst; if (monthSt(run) === to) return;
    const others = state.runs.filter(x => ym(x.period) === state.period && x.id !== run.id && segSt(x, '유통사업부') === '확정');
    if (to === '확정' && others.length && !await ask(`${monthLabel(state.period)}에 이미 확정된 저장 기록이 있어요. 이 기록을 확정하면 그 기록은 가정산으로 바뀌어요. 계속할까요?`)) return;
    const { error } = await sb.rpc('dist_pnl_set_seg_status', { p_run: run.id, p_seg: '유통사업부', p_status: to });
    if (error) { toast('변경 실패: ' + error.message); return; }
    toast(`${monthLabel(state.period)} → ${to}`); state.selRun = run.id; await loadRuns(); openMonth(state.period);
  });
}

// ---------- 엑셀 내보내기 ----------
function exportXlsx(period, summary, lines, segStatus = blankSt(), ev = null) {
  const wb = XLSX.utils.book_new();
  const s = summary.segments;
  const KS = Object.keys(s);
  const aoa = [[`유통사업부 손익매출 ${monthLabel(period)}`], ['단위: 원 · 매출 = B − A − 자점운영비 − 직접지급사은품 · 유통사업부: 코드매출+외부유통매출 = 유통매출+직영영업매출'], [], ['항목', ...KS.map(k => DIST_VIEWS.includes(k) ? `유통사업부 ${k}` : k)], ['정산 상태', ...KS.map(k => summary.legacy ? '이전 기준' : segStatus[statusKeyOf(k)])]];
  for (const [cls, lab, key, sign = 1] of LINES) {
    if (cls === 'sec') { aoa.push([`[${lab}]`]); continue; }
    aoa.push([lab, ...KS.map(k => sign * s[k][key])]);
  }
  if (!summary.legacy) {
    aoa.push([], ['[유통사업부 검산]'], ['코드매출 + 외부유통매출', summary.total.REV], ['유통매출 + 직영영업매출', summary.total.REV2], ['자점운영비 (별도 표기 · 유통사업부 매출에 미반영)', summary.total.JF]);
  }
  const ws1 = XLSX.utils.aoa_to_sheet(aoa); ws1['!cols'] = [{ wch: 42 }, ...KS.map(() => ({ wch: 16 }))];
  numFmt(ws1); XLSX.utils.book_append_sheet(wb, ws1, '요약');
  const up = byUpper(lines).map(r => ({ 상부: r.segment, 상부점: r.upper_name, 하부: r.partner_type, '상부 건수': r.cntUp, '상부 B': r.B, '전체 건수': r.cntDn, '전체 A': r.A, '자점운영비': r.JF, '직접지급사은품': r.GD, '매출': r.REV }));
  const ws2 = XLSX.utils.json_to_sheet(up); ws2['!cols'] = [{ wch: 8 }, { wch: 34 }, { wch: 8 }, { wch: 10 }, { wch: 16 }, { wch: 10 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];
  numFmt(ws2); XLSX.utils.book_append_sheet(wb, ws2, '상부점별');
  const raw = lines.map(l => ({ 상부구분: l.segment, 상부점: l.upper_name, 측: l.side, 항목: l.item, 하부구분: l.partner_type, 포함: l.included ? 'Y' : 'N', 건수: l.cnt, 정산금액: l.amount, 사은품현금: l.gift_cash, 사은품물류: l.gift_goods, 사은품쿠폰: l.gift_coupon, 직접지급사은품: l.direct_gift ?? 0, 자점운영비: l.jj_fee || 0 }));
  const ws3 = XLSX.utils.json_to_sheet(raw); numFmt(ws3); XLSX.utils.book_append_sheet(wb, ws3, '집계데이터');
  if (ev) {
    const cols = EVAL_COLS;
    const aoaE = [[`평가손익 입력값 · 유통사업부 ${monthLabel(period)}`], ['단위: 원 · 협력점 자점 제외(사은품·기대관리는 자점 포함) · 기대관리는 발생수수료에 포함'], [],
      ['구분', ...cols.map(c => c[1])],
      ...EVAL_ROWS.map(k => [k, ...cols.map(c => ev.rows[k][c[0]])]), ['유통사업부 합계', ...cols.map(c => ev.total[c[0]])], [],
      ...cols.map(c => [c[1], c[2]])];
    if (ev.ratesMissing && ev.ratesMissing.length) aoaE.push([], ['기대관리 단가 확인 중 (금액 0으로 계산)', ev.ratesMissing.join(', ')]);
    const wsE = XLSX.utils.aoa_to_sheet(aoaE); wsE['!cols'] = [{ wch: 18 }, ...cols.map(() => ({ wch: 16 }))]; numFmt(wsE);
    XLSX.utils.book_append_sheet(wb, wsE, '평가손익');
    const wsX = XLSX.utils.json_to_sheet((ev.expectRows || []).map(e => ({ 통신사: e.carrier, 유형: e.type, 건수: e.cnt, '월 단가': e.rate ?? '확인 중', 개월: e.months, 기대관리: e.amount ?? 0 })));
    numFmt(wsX); XLSX.utils.book_append_sheet(wb, wsX, '기대관리');
    if (ev.giftList) {
      const wsG = XLSX.utils.json_to_sheet(ev.giftList.map(g => ({ 구분: g.seg, 내용: g.why, 협력점: g.partner, 고객명: g.customer, 개통일: g.openDate, 상품명: g.product, 지급처: g.payer, 지급완료: g.done, 차감금액: g.ded, 지급금액: g.paid, 선택: { ded: '차감 금액', paid: '지급 금액', none: '넣지 않음' }[g.pick] })));
      numFmt(wsG); XLSX.utils.book_append_sheet(wb, wsG, '사은품확인');
    }
  }
  XLSX.writeFile(wb, `유통사업부_손익매출_${period}.xlsx`);
}
function numFmt(ws) {
  for (const a in ws) if (a[0] !== '!' && ws[a].t === 'n') ws[a].z = '#,##0;-#,##0;"-"';
}

setupBookmarklet();
announceReady();
boot();
