import { createClient } from 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.45.4/+esm';
import { SUPABASE_URL, SUPABASE_KEY } from './config.js';
import { RULES, SEGMENTS } from './rules.js';
import { parseWorkbook, compute, summarize, byUpper } from './calc.js';

const sb = createClient(SUPABASE_URL, SUPABASE_KEY);
const $ = id => document.getElementById(id);
const XLSX = window.XLSX;

const state = { files: [], result: null, runs: [], selPeriod: null, selRun: null };

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
function show(view) { for (const v of ['viewLogin', 'viewDenied', 'viewApp']) $(v).hidden = v !== view; }

// ---------- 로그인 ----------
async function boot() {
  const { data: { session } } = await sb.auth.getSession();
  await onSession(session);
  sb.auth.onAuthStateChange((_e, s) => { if (!s) onSession(null); });
}
async function onSession(session) {
  if (!session) { $('who').hidden = true; show('viewLogin'); return; }
  $('who').hidden = false; $('whoEmail').textContent = session.user.email;
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
document.querySelectorAll('.tab').forEach(b => b.onclick = () => {
  document.querySelectorAll('.tab').forEach(x => x.classList.toggle('active', x === b));
  document.querySelectorAll('.pane').forEach(p => p.hidden = p.id !== 'pane-' + b.dataset.tab);
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

async function addFiles(fileList) {
  const list = [...fileList]; // FileList는 input 초기화 시 비워지므로 먼저 복사
  for (const f of list) {
    if (state.files.some(x => x.fileName === f.name && x.size === f.size)) { toast(`${f.name}: 이미 올린 파일입니다`); continue; }
    try {
      const parsed = parseWorkbook(XLSX, new Uint8Array(await f.arrayBuffer()), f.name);
      parsed.size = f.size; state.files.push(parsed);
    } catch (err) { toast(err.message); }
  }
  state.result = null; $('calcResult').hidden = true; renderFiles();
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
  $('fileList').innerHTML = `<table class="files"><tr><th>구분</th><th>파일</th><th>정산종류별 건수</th><th>정산금액 합계</th><th>점검</th><th></th></tr>${rows}</table>`;
  $('fileList').querySelectorAll('[data-rm]').forEach(b => b.onclick = () => { state.files.splice(+b.dataset.rm, 1); state.result = null; $('calcResult').hidden = true; renderFiles(); });
}

// ---------- 어드민에서 바로 가져오기 (북마클릿 ↔ 이 창 postMessage) ----------
const ADMIN_ORIGIN = 'https://workline.kr';
const CALC_URL = location.origin + location.pathname.replace(/index\.html$/, '');
let pendingAdmin = null;
function announceReady() {
  if (window.opener && location.hash.includes('admin-import')) {
    try { window.opener.postMessage({ type: 'dist-pnl-calc-ready' }, ADMIN_ORIGIN); } catch { /* 어드민 창이 닫힘 */ }
  }
}
window.addEventListener('message', e => {
  if (e.origin !== ADMIN_ORIGIN || !e.data || e.data.type !== 'dist-pnl-admin-files') return; // 어드민 창에서 온 것만 받음
  const { period, files } = e.data;
  if (!/^\d{4}-\d{2}$/.test(period) || !Array.isArray(files)) return;
  const parsed = [], errors = [];
  for (const f of files) {
    try {
      // 직접 올린 엑셀과 똑같은 방식으로 읽음 (상부/전체는 시트 이름으로 다시 판별)
      const p = parseWorkbook(XLSX, new Uint8Array(f.buffer), String(f.name || '어드민 엑셀'));
      p.source = 'admin'; parsed.push(p);
    } catch (err) { errors.push(err.message); }
  }
  pendingAdmin = { period, pulledAt: e.data.pulledAt, files: parsed, errors };
  applyAdmin();
});
window.addEventListener('hashchange', announceReady);
function applyAdmin() {
  if (!pendingAdmin || $('viewApp').hidden) return;
  const { period, files, pulledAt, errors } = pendingAdmin; pendingAdmin = null;
  if (errors.length) errors.forEach(m => toast(m));
  document.querySelector('.tab[data-tab=calc]').click();
  state.files = files; $('period').value = period;
  renderFiles();
  $('btnCalc').click();
  $('calcChecks').insertAdjacentHTML('afterbegin', `<div class="notice ok">어드민에서 직접 가져온 ${monthLabel(period)} 데이터입니다 (${new Date(pulledAt || Date.now()).toLocaleString('ko-KR')})</div>`);
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
  const warn = [...checks.warnings];
  if (checks.outOfPeriod) warn.push(`대상월(${monthLabel(period)})이 아닌 행 ${checks.outOfPeriod.toLocaleString()}건은 계산에서 뺐습니다 — 월을 잘못 골랐거나 다른 달 파일이 섞였는지 확인하세요`);
  $('calcChecks').innerHTML = (warn.length ? warn.map(w => `<div class="notice warn">⚠ ${esc(w)}</div>`).join('')
    : `<div class="notice ok">✓ 파일 점검 통과 — 계산에 쓴 행 ${checks.usedRows.toLocaleString()}건 (타부서 코드 ${checks.otherDept.toLocaleString()}건 제외)</div>`);
  $('calcSummary').innerHTML = renderSummary(summary, monthLabel(period));
  $('calcUpper').innerHTML = renderUpper(byUpper(lines));
  $('calcResult').hidden = false; $('saveMsg').textContent = '';
  $('calcResult').scrollIntoView({ behavior: 'smooth' });
};

$('btnSave').onclick = async () => {
  const r = state.result; if (!r) return;
  const status = document.querySelector('input[name=status]:checked').value;
  const periodDate = r.period + '-01';
  const existing = state.runs.filter(x => ym(x.period) === r.period);
  if (status === '확정' && existing.some(x => x.status === '확정')) {
    if (!await ask(`${monthLabel(r.period)}에는 이미 확정 회차가 있습니다. 이번 계산을 확정으로 저장하면 이전 확정은 가정산으로 바뀝니다. 계속할까요?`)) return;
  }
  $('btnSave').disabled = true; $('saveMsg').textContent = '저장 중…';
  const files = state.files.map(f => ({ name: f.fileName, side: f.side, rows: f.rows.length, total: f.fileTotal }));
  const rules = { version: RULES.version, checks: { usedRows: r.checks.usedRows, outOfPeriod: r.checks.outOfPeriod, otherDept: r.checks.otherDept, warnings: r.checks.warnings } };
  const { error } = await sb.rpc('dist_pnl_save_run', { p_period: periodDate, p_status: status, p_note: $('note').value.trim() || null,
    p_files: files, p_summary: r.summary, p_rules: rules, p_lines: r.lines });
  $('btnSave').disabled = false;
  if (error) { $('saveMsg').innerHTML = `<span class="err">저장 실패: ${esc(error.message)}</span>`; return; }
  $('saveMsg').innerHTML = `<span class="pill ${status}">${status}</span> ${monthLabel(r.period)} 저장 완료 — 월별 내역 탭에서 볼 수 있어요`;
  toast('저장했습니다'); loadRuns();
};
$('btnExportCalc').onclick = () => { const r = state.result; if (r) exportXlsx(r.period, r.summary, r.lines, '계산중'); };

// ---------- 요약 화면 ----------
const SEG_TITLE = {
  '유통이관': '① 유통이관 (상부=이관처 · 자점 협력점 제외)',
  'SKB웨이브넷': '② 코드 – SKB웨이브넷 (자점 포함)',
  'SKY우리정보': '③ 코드 – SKY우리정보 (자점 포함)',
  'SK매직': '④ 코드 – SK매직 (자점 포함)',
};
const LINES = [
  ['sec', '전체정산 · 우리가 준 돈'],
  ['', 'a. 수수료 (사은품 차감전, 기타 차감 반영)', 'D_A'], ['', 'b. 사은품 - 현금', 'D_B'], ['', 'c. 사은품 - 물류', 'D_C'], ['', 'd. 사은품 - 쿠폰', 'D_D'], ['', 'f. 환수', 'D_F'],
  ['b', '최종 정산금액 (A) = a − (b+c+d+f)', 'A'],
  ['sec', '상부정산 · 우리가 받은 돈'],
  ['', 'a. 수수료 (a-1·a-2 제외)', 'U_A'], ['', 'a-1. 관리수수료', 'U_A1'], ['', 'a-2. 정산관련', 'U_A2'], ['', 'b. 사은품 - 현금', 'U_B'], ['', 'c. 사은품 - 물류', 'U_C'], ['', 'd. 사은품 - 쿠폰', 'U_D'], ['', 'f. 환수', 'U_F'],
  ['b', '최종 정산금액 (B) = (a+a1+a2) − (b+c+d+f)', 'B'],
  ['sec', '매출'],
  ['', 'B − A', 'BA'], ['', '(−) 직접지급사은품 (전체 b+c+d)', 'GIFT', -1],
  ['key', '매출 = B − A − 직접지급사은품', 'REV'],
  ['ref', '참고: 자점운영비 (A 중 자점 지급분)', 'JJ', 1, 'code'],
];
function renderSummary(summary, title) {
  const s = summary.segments;
  const kpi = SEGMENTS.map(k => `<div class="kpi"><div class="k">${k}</div><div class="v">${won(s[k].REV)}</div></div>`).join('')
    + `<div class="kpi total"><div class="k">유통사업부 합계</div><div class="v">${won(summary.total.REV)}</div></div>`;
  const cards = SEGMENTS.map(k => {
    const t = s[k];
    const rows = LINES.filter(l => !(l[4] === 'code' && k === '유통이관')).map(([cls, lab, key, sign = 1]) =>
      cls === 'sec' ? `<tr class="sec"><td colspan="2">${lab}</td></tr>`
        : `<tr class="${cls}"><td>${lab}</td><td class="n">${won(sign * t[key])}</td></tr>`).join('');
    return `<div class="card"><h3>${SEG_TITLE[k]}</h3><table class="sum">${rows}</table></div>`;
  }).join('');
  return `<div class="card"><h3>${esc(title)} 매출 요약</h3><div class="kpis">${kpi}</div><p class="muted small">단위: 원 · 매출 = 상부(B) − 전체(A) − 직접지급사은품</p></div><div class="segs">${cards}</div>`;
}
function renderUpper(rows) {
  const tr = rows.map(r => `<tr><td>${r.segment}</td><td>${esc(r.upper_name)}</td><td class="n">${r.cntUp.toLocaleString()}</td><td class="n">${won(r.B)}</td>
    <td class="n">${r.cntDn.toLocaleString()}</td><td class="n">${won(r.A)}</td><td class="n">${won(r.GIFT)}</td><td class="n"><b>${won(r.REV)}</b></td></tr>`).join('');
  return `<div class="card"><h3>상부점별 내역 (수기 대조용)</h3><div class="scroll"><table class="grid">
    <tr><th>구분</th><th>상부점</th><th class="n">상부 건수</th><th class="n">상부 B</th><th class="n">전체 건수</th><th class="n">전체 A</th><th class="n">직접지급사은품</th><th class="n">매출</th></tr>${tr}</table></div></div>`;
}

// ---------- 월별 내역 ----------
$('btnReload').onclick = () => loadRuns();
async function loadRuns() {
  const { data, error } = await sb.from('dist_pnl_runs')
    .select('id, period, status, note, files, summary, rules, created_by, created_at, confirmed_by, confirmed_at')
    .order('period', { ascending: false }).order('created_at', { ascending: false });
  if (error) { toast('불러오기 실패: ' + error.message); return; }
  state.runs = data; renderMonths();
}
function repRun(period) {
  const list = state.runs.filter(r => ym(r.period) === period);
  return list.find(r => r.status === '확정') || list[0];
}
function renderMonths() {
  const periods = [...new Set(state.runs.map(r => ym(r.period)))];
  if (!periods.length) { $('tblMonths').innerHTML = '<tr><td class="muted">저장된 회차가 없습니다. 계산하기 탭에서 계산 후 저장하세요.</td></tr>'; $('monthDetail').innerHTML = ''; return; }
  const head = `<tr><th>월</th><th>상태</th>${SEGMENTS.map(s => `<th class="n">${s}</th>`).join('')}<th class="n">합계</th><th>회차</th><th>저장</th></tr>`;
  const tot = Object.fromEntries([...SEGMENTS, 'total'].map(k => [k, 0]));
  const body = periods.map(p => {
    const r = repRun(p), s = r.summary.segments, n = state.runs.filter(x => ym(x.period) === p).length;
    SEGMENTS.forEach(k => tot[k] += s[k].REV); tot.total += r.summary.total.REV;
    return `<tr class="click ${state.selPeriod === p ? 'sel' : ''}" data-p="${p}"><td>${monthLabel(p)}</td><td><span class="pill ${r.status}">${r.status}</span></td>
      ${SEGMENTS.map(k => `<td class="n">${won(s[k].REV)}</td>`).join('')}<td class="n"><b>${won(r.summary.total.REV)}</b></td>
      <td>${n}건</td><td class="muted small">${new Date(r.created_at).toLocaleString('ko-KR')}</td></tr>`;
  }).join('');
  const foot = `<tr class="tot"><td>표시 합계</td><td></td>${SEGMENTS.map(k => `<td class="n">${won(tot[k])}</td>`).join('')}<td class="n">${won(tot.total)}</td><td></td><td></td></tr>`;
  $('tblMonths').innerHTML = head + body + foot;
  $('tblMonths').querySelectorAll('tr.click').forEach(tr => tr.onclick = () => { state.selPeriod = tr.dataset.p; state.selRun = null; renderMonths(); });
  if (state.selPeriod) renderDetail();
}
async function renderDetail() {
  const p = state.selPeriod; if (!p) return;
  const list = state.runs.filter(r => ym(r.period) === p);
  if (!list.length) { $('monthDetail').innerHTML = ''; return; }
  const run = list.find(r => r.id === state.selRun) || repRun(p);
  state.selRun = run.id;
  const runRows = list.map(r => `<tr class="click ${r.id === run.id ? 'sel' : ''}" data-id="${r.id}"><td><span class="pill ${r.status}">${r.status}</span></td>
    <td class="muted small">${new Date(r.created_at).toLocaleString('ko-KR')}</td><td class="small">${esc(r.created_by || '')}</td>
    <td class="n">${won(r.summary.total.REV)}</td><td class="small">${esc(r.note || '')}</td></tr>`).join('');
  $('monthDetail').innerHTML = `<div class="card"><div class="row between wrap gap"><h3>${monthLabel(p)} 회차</h3>
      <div class="row gap wrap">
        ${run.status === '확정' ? '<button id="btnUnconfirm" class="ghost">가정산으로 되돌리기</button>' : '<button id="btnConfirm" class="primary">이 회차를 확정으로</button>'}
        <button id="btnExportRun" class="ghost">엑셀 다운로드</button><button id="btnDelRun" class="ghost danger">회차 삭제</button></div></div>
      <div class="scroll"><table class="grid"><tr><th>상태</th><th>저장 시각</th><th>저장자</th><th class="n">합계 매출</th><th>메모</th></tr>${runRows}</table></div>
      <p class="muted small">사용 파일: ${(run.files || []).map(f => `${esc(f.name)} (${f.side}, ${Number(f.rows).toLocaleString()}행)`).join(' · ')}${run.confirmed_at ? ` · 확정: ${esc(run.confirmed_by)} ${new Date(run.confirmed_at).toLocaleString('ko-KR')}` : ''}</p>
    </div><div id="runBody"><div class="card muted">불러오는 중…</div></div>`;
  $('monthDetail').querySelectorAll('tr.click').forEach(tr => tr.onclick = () => { state.selRun = tr.dataset.id; renderDetail(); });
  const setStatus = async st => {
    if (st === '확정' && list.some(x => x.status === '확정' && x.id !== run.id) && !await ask('이미 확정된 회차가 있습니다. 이 회차를 확정하면 기존 확정은 가정산으로 바뀝니다. 계속할까요?')) return;
    const { error } = await sb.rpc('dist_pnl_set_status', { p_run: run.id, p_status: st });
    if (error) { toast('변경 실패: ' + error.message); return; }
    toast(`${st}(으)로 변경했습니다`); await loadRuns();
  };
  if ($('btnConfirm')) $('btnConfirm').onclick = () => setStatus('확정');
  if ($('btnUnconfirm')) $('btnUnconfirm').onclick = () => setStatus('가정산');
  $('btnDelRun').onclick = async () => {
    if (!await ask(`${monthLabel(p)} 회차(${run.status}, ${new Date(run.created_at).toLocaleString('ko-KR')})를 삭제할까요? 되돌릴 수 없습니다.`)) return;
    const { error } = await sb.from('dist_pnl_runs').delete().eq('id', run.id);
    if (error) { toast('삭제 실패: ' + error.message); return; }
    toast('삭제했습니다'); state.selRun = null; await loadRuns();
  };
  const { data: lines, error } = await sb.from('dist_pnl_lines').select('*').eq('run_id', run.id);
  if (error) { $('runBody').innerHTML = `<div class="card err">${esc(error.message)}</div>`; return; }
  const summary = summarize(lines);
  $('runBody').innerHTML = renderSummary(summary, `${monthLabel(p)} (${run.status})`) + renderUpper(byUpper(lines));
  $('btnExportRun').onclick = () => exportXlsx(p, summary, lines, run.status);
}

// ---------- 엑셀 내보내기 ----------
function exportXlsx(period, summary, lines, status) {
  const wb = XLSX.utils.book_new();
  const s = summary.segments;
  const aoa = [[`유통사업부 손익매출 ${monthLabel(period)} (${status})`], ['단위: 원'], [], ['항목', ...SEGMENTS, '합계']];
  for (const [cls, lab, key, sign = 1, only] of LINES) {
    if (cls === 'sec') { aoa.push([`[${lab}]`]); continue; }
    const vals = SEGMENTS.map(k => (only === 'code' && k === '유통이관') ? null : sign * s[k][key]);
    aoa.push([lab, ...vals, vals.reduce((a, b) => a + (b || 0), 0)]);
  }
  const ws1 = XLSX.utils.aoa_to_sheet(aoa); ws1['!cols'] = [{ wch: 42 }, ...SEGMENTS.map(() => ({ wch: 16 })), { wch: 16 }];
  numFmt(ws1); XLSX.utils.book_append_sheet(wb, ws1, '요약');
  const up = byUpper(lines).map(r => ({ 구분: r.segment, 상부점: r.upper_name, '상부 건수': r.cntUp, '상부 B': r.B, '전체 건수': r.cntDn, '전체 A': r.A, 직접지급사은품: r.GIFT, 매출: r.REV }));
  const ws2 = XLSX.utils.json_to_sheet(up); ws2['!cols'] = [{ wch: 14 }, { wch: 34 }, { wch: 10 }, { wch: 16 }, { wch: 10 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];
  numFmt(ws2); XLSX.utils.book_append_sheet(wb, ws2, '상부점별');
  const raw = lines.map(l => ({ 구분: l.segment, 상부점: l.upper_name, 측: l.side, 항목: l.item, 자점구분: l.partner_type, 포함: l.included ? 'Y' : 'N', 건수: l.cnt, 정산금액: l.amount, 사은품현금: l.gift_cash, 사은품물류: l.gift_goods, 사은품쿠폰: l.gift_coupon }));
  const ws3 = XLSX.utils.json_to_sheet(raw); numFmt(ws3); XLSX.utils.book_append_sheet(wb, ws3, '집계데이터');
  XLSX.writeFile(wb, `유통사업부_손익매출_${period}_${status}.xlsx`);
}
function numFmt(ws) {
  for (const a in ws) if (a[0] !== '!' && ws[a].t === 'n') ws[a].z = '#,##0;-#,##0;"-"';
}

// ---------- 기준 ----------
$('rulesBox').innerHTML = `
<h3>대상</h3><ul>
<li><b>유통이관</b>: 상부점 이름이 숫자로 시작하지 않는 이관처. 협력점이 자점(숫자로 시작)인 건은 제외 — 단 <code>0.영업정산</code>은 정산 입력용이라 포함</li>
<li><b>코드</b>: 상부점에 ${RULES.codeSegments.map(c => `<code>${c.match}</code> → ${c.segment}`).join(', ')} 가 들어간 숫자 코드. 협력점은 자점·외부 모두 포함</li>
<li>그 밖의 숫자 코드(KT·SKT, 우신·대성 SKY 등)는 타부서라 제외</li></ul>
<h3>월 귀속</h3><ul><li>일반·추가 = 개통일이 대상월인 건 / 환수 = 정산예정일이 대상월인 건 (파일에 다른 달 행이 있어도 자동으로 걸러냄)</li></ul>
<h3>항목</h3><ul>
<li>사은품: b.현금 = 차감내역-사은품, c.물류 = 차감내역-물류, d.쿠폰 = 차감내역-쿠폰 (일반·추가 건)</li>
<li>a.수수료 = 정산금액 + b + c + d (선납·할인탭·미비서류·기타 차감은 수수료에 반영)</li>
<li>상부 a-1.관리수수료 = 상품명에 <code>${RULES.mgmtProductKeyword}</code>, a-2.정산관련 = 상품명에 <code>${RULES.settleProductKeyword}</code> (a에서는 제외)</li>
<li>f.환수 = 환수 건 정산금액(차감액을 양수로 표시)</li></ul>
<h3>산식</h3><ul>
<li>전체 A = a − (b+c+d+f) = 데이터 정산금액 합계</li>
<li>상부 B = (a+a1+a2) − (b+c+d+f) = 데이터 정산금액 합계</li>
<li><b>매출 = B − A − 직접지급사은품(전체 b+c+d)</b></li>
<li>자점운영비 = 코드 매출 중 자점 협력점에 지급한 전체정산 금액 (참고 표시, 매출에서 이미 차감됨)</li></ul>
<p class="muted small">기준 버전 ${RULES.version} · 기준 변경은 <code>src/rules.js</code></p>`;

setupBookmarklet();
announceReady();
boot();
