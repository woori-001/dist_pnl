// 손익매출 계산 엔진 — 브라우저와 Node 양쪽에서 동작 (XLSX 객체는 밖에서 넘겨받음)
import { RULES, SEGMENTS, VIEW_FILTER, LEGACY_SEGMENTS } from './rules.js';

const SIDE_BY_SHEET = { '상부정산관리': '상부', '전체정산관리': '전체' };
const NEED = ['정산종류', '협력점', '상부점', '상품명', '개통일', '정산예정일', '정산금액',
  '약정금액', '수납', '번들수수료', '동판수수료'];
const DEDUCT = ['사은품', '물류', '쿠폰', '선납', '할인탭', '미비서류', '기타'];
const OPTIONAL = ['접수일', '가입서비스', '상품옵션', '셋트유형', '사은품지급처', '지급완료'];

const num = v => (v === null || v === undefined || v === '') ? 0 : Math.round(Number(String(v).replace(/,/g, '')) || 0);
const str = v => (v === null || v === undefined) ? '' : String(v).trim();

/** 어드민에서 받은 엑셀 1개를 읽어 필요한 칸만 뽑는다 */
export function parseWorkbook(XLSX, data, fileName) {
  const wb = XLSX.read(data, { type: 'array' });
  const sheetName = wb.SheetNames.find(n => SIDE_BY_SHEET[n]);
  if (!sheetName) throw new Error(`${fileName}: '상부정산관리' 또는 '전체정산관리' 시트가 없습니다`);
  const side = SIDE_BY_SHEET[sheetName];
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: null });
  const h0 = aoa[0] || [], h1 = aoa[1] || [];
  const idx = {};
  h0.forEach((h, i) => { if (h !== null && idx[str(h)] === undefined) idx[str(h)] = i; });
  const dStart = idx['차감내역'];
  for (const n of NEED) if (idx[n] === undefined) throw new Error(`${fileName}: '${n}' 칸을 찾지 못했습니다`);
  if (dStart === undefined) throw new Error(`${fileName}: '차감내역' 칸을 찾지 못했습니다`);
  const didx = {};
  for (let i = dStart; i < dStart + 10 && i < h1.length; i++) if (h1[i]) didx[str(h1[i])] = i;
  for (const n of DEDUCT) if (didx[n] === undefined) throw new Error(`${fileName}: 차감내역 '${n}' 칸을 찾지 못했습니다`);

  const rows = []; let fileTotal = null;
  for (let r = 2; r < aoa.length; r++) {
    const a = aoa[r]; if (!a) continue;
    const kind = str(a[idx['정산종류']]);
    if (!kind) {
      if (str(a[0]).includes('합계')) fileTotal = num(a[idx['정산금액']]);
      continue;
    }
    const d = {}; for (const n of DEDUCT) d[n] = num(a[didx[n]]);
    rows.push({
      kind, partner: str(a[idx['협력점']]), upper: str(a[idx['상부점']]), product: str(a[idx['상품명']]),
      openDate: str(a[idx['개통일']]).slice(0, 10), dueDate: str(a[idx['정산예정일']]).slice(0, 10),
      amount: num(a[idx['정산금액']]), base: num(a[idx['약정금액']]), receipt: num(a[idx['수납']]),
      bundle: num(a[idx['번들수수료']]), dongpan: num(a[idx['동판수수료']]), d,
      recvDate: idx['접수일'] === undefined ? '' : str(a[idx['접수일']]).slice(0, 10),
      svc: idx['가입서비스'] === undefined ? '' : str(a[idx['가입서비스']]),
      opt: idx['상품옵션'] === undefined ? '' : str(a[idx['상품옵션']]),
      set: idx['셋트유형'] === undefined ? '' : str(a[idx['셋트유형']]),
      giftPayer: idx['사은품지급처'] === undefined ? '' : str(a[idx['사은품지급처']]),
      giftDone: idx['지급완료'] === undefined ? '' : str(a[idx['지급완료']]),
    });
  }
  return { fileName, side, rows, fileTotal };
}

/** 상부 구분: '코드' (숫자로 시작) / '이관' (그 외 + 0.기타) / null (상부점 비어 있음) */
export function segmentOf(upper) {
  if (!upper) return null;
  if (!RULES.codePattern.test(upper)) return '이관';
  if (RULES.transferCodePrefixes.some(p => upper.trim().startsWith(p))) return '이관';
  return '코드';
}
export function isJajeom(partner) {
  return RULES.jajeomPattern.test(partner) && !RULES.jajeomExceptPrefixes.some(p => partner.trim().startsWith(p));
}
/** 자점운영비 항목 판별 (첨부 '자점운영비(수동반영) 관리' 표 기준) */
export function jajeomFeeCategory(row) {
  const p = row.product || '';
  const type = p.includes('_') ? p.slice(p.indexOf('_') + 1) : '';
  // 인터넷: 상품유형이 '인터넷'으로 시작 (SK-biz 소호 인터넷처럼 앞에 구분 표시가 붙은 것도 포함, 인터넷전화 제외)
  const netType = type.replace(/^\([^)]*\)/, '');
  if (netType.startsWith('인터넷') && !netType.startsWith('인터넷전화')) return (row.set || '').startsWith('단독') ? '인터넷 단독' : '인터넷 번들';
  if (p.startsWith('모바일')) {
    const svc = row.svc || '';
    const fam = (row.opt || '').includes('가족동판');
    if (p.includes('SKYLIFE') && svc.includes('(유심)')) return '유심 상품 SKYLIFE';
    if (svc.includes('(유심)')) return fam ? '유심 가족동판' : '유심';
    if (svc.includes('(기기)')) return fam ? '기기 가족동판' : '기기';
    if (svc.includes('(사전)')) return fam ? '사전예약 가족동판' : '사전예약';
  }
  return null;
}
export function jajeomFeeOf(row) {
  const cat = jajeomFeeCategory(row); if (!cat) return 0;
  const date = row.recvDate || row.openDate || '';
  const hit = RULES.jajeomFees.filter(f => f.name === cat && f.from <= date).sort((a, b) => b.from.localeCompare(a.from))[0];
  return hit ? hit.amount : 0;
}
function itemOf(row) {
  if (row.kind === '환수') return '환수';
  if (row.product.includes(RULES.mgmtProductKeyword)) return '관리수수료';
  if (row.product.includes(RULES.settleProductKeyword)) return '정산관련';
  return '수수료';
}

/** 파싱된 파일들 + 대상월(YYYY-MM) → 집계행 + 점검결과 */
export function compute(parsed, period) {
  const map = new Map();
  const checks = { files: [], outOfPeriod: 0, otherDept: 0, formulaMismatch: 0, usedRows: 0, warnings: [] };
  const seen = new Set();
  const kindFiles = {}; // 같은 정산종류가 여러 파일에서 계산되면 중복 가능성 경고
  for (const f of parsed) {
    const sumAll = f.rows.reduce((s, r) => s + r.amount, 0);
    const kinds = {}; f.rows.forEach(r => kinds[r.kind] = (kinds[r.kind] || 0) + 1);
    checks.files.push({ name: f.fileName, side: f.side, rows: f.rows.length, kinds,
      fileTotal: f.fileTotal, sumAll, totalOk: f.fileTotal === null || f.fileTotal === sumAll });
    if (f.fileTotal !== null && f.fileTotal !== sumAll) checks.warnings.push(`${f.fileName}: 파일 합계(${f.fileTotal})와 행 합계(${sumAll})가 다릅니다`);
    for (const r of f.rows) {
      // 일반 = 개통일 기준, 추가·환수 = 정산예정일 기준 (RULES.openDateKinds)
      const dateKey = RULES.openDateKinds.includes(r.kind) ? r.openDate : r.dueDate;
      if (!dateKey.startsWith(period)) { checks.outOfPeriod++; continue; }
      const calc = r.base - DEDUCT.reduce((s, k) => s + r.d[k], 0) + r.receipt + r.bundle + r.dongpan;
      if (calc !== r.amount) checks.formulaMismatch++;
      const seg = segmentOf(r.upper);
      if (!seg) { checks.otherDept++; continue; } // 상부점이 비어 있는 행
      const jj = isJajeom(r.partner);
      const included = true;
      const item = itemOf(r);
      const gift = r.kind === '환수' ? { c: 0, g: 0, k: 0 } : { c: r.d['사은품'], g: r.d['물류'], k: r.d['쿠폰'] };
      const key = [seg, r.upper, f.side, item, jj ? '자점' : '협력점', included].join('|');
      let l = map.get(key);
      if (!l) { l = { segment: seg, upper_name: r.upper, side: f.side, item, partner_type: jj ? '자점' : '협력점', included,
        cnt: 0, amount: 0, gift_cash: 0, gift_goods: 0, gift_coupon: 0, jj_fee: 0, direct_gift: 0 }; map.set(key, l); }
      l.cnt++; l.amount += r.amount; l.gift_cash += gift.c; l.gift_goods += gift.g; l.gift_coupon += gift.k;
      // 직접지급사은품: 회사가 고객에게 준 것만 (유치자지급 = 협력점이 직접 지급 → 제외)
      const partnerPaid = (r.giftPayer || '').includes(RULES.partnerPaidGiftLabel);
      if (!partnerPaid) l.direct_gift += gift.c + gift.g + gift.k;
      if (partnerPaid && (gift.c + gift.g + gift.k) && r.giftDone) checks.partnerPaidDone = (checks.partnerPaidDone || 0) + 1;
      // 자점운영비: 코드 상부 · 전체정산 · 자점(7.유통사업부·0.영업정산 제외) · 일반 건
      if (seg === '코드' && f.side === '전체' && jj && r.kind === '일반'
          && !RULES.jajeomFeeExcludePrefixes.some(p => r.partner.trim().startsWith(p))) {
        const fee = jajeomFeeOf(r); l.jj_fee += fee; if (fee) checks.jajeomFeeRows = (checks.jajeomFeeRows || 0) + 1;
      }
      checks.usedRows++;
      const kf = f.side + '|' + r.kind + '|' + (f.db || ''); (kindFiles[kf] = kindFiles[kf] || new Set()).add(f.fileName);
    }
    const sk = f.side + JSON.stringify(kinds) + sumAll;
    if (f.rows.length && seen.has(sk)) checks.warnings.push(`${f.fileName}: 같은 내용의 파일이 두 번 올라온 것 같습니다`);
    seen.add(sk);
  }
  for (const [k, set] of Object.entries(kindFiles)) if (set.size > 1) {
    const [side, kind] = k.split('|');
    checks.warnings.push(`${side}정산 '${kind}' 건이 파일 ${set.size}개에서 계산됐습니다 — 같은 건이 두 번 들어갔을 수 있어요 (${[...set].join(', ')})`);
  }
  const sides = new Set(parsed.map(f => f.side));
  for (const s of ['상부', '전체']) if (!sides.has(s)) checks.warnings.push(`${s}정산 파일이 없습니다`);
  const hasRefund = s => parsed.some(f => f.side === s && f.rows.some(r => r.kind === '환수'));
  for (const s of ['상부', '전체']) if (sides.has(s) && !hasRefund(s)) checks.warnings.push(`${s}정산 환수 건이 하나도 없습니다 (환수 파일 확인)`);
  if (checks.partnerPaidDone) checks.warnings.push(`사은품지급처가 '유치자지급'인데 지급완료일이 있는 건 ${checks.partnerPaidDone}건 — 확인이 필요해요 (직접지급사은품에서는 제외됨)`);
  if (checks.formulaMismatch) checks.warnings.push(`정산금액 산식이 맞지 않는 행 ${checks.formulaMismatch}건`);
  return { lines: [...map.values()], checks };
}

/** 저장된 집계행이 이전 기준(유통이관·SKB·SKY·매직)인지 */
export const isLegacyLines = lines => lines.some(l => LEGACY_SEGMENTS.includes(l.segment));

function addLine(t, l) {
  const gifts = l.gift_cash + l.gift_goods + l.gift_coupon;
  t.JF += l.jj_fee || 0;
  if (l.side === '전체') {
    t.GD += (l.direct_gift == null ? gifts : l.direct_gift);
    if (l.item === '환수') t.D_F += -l.amount;
    else {
      const k = l.item === '관리수수료' ? 'D_A1' : l.item === '정산관련' ? 'D_A2' : 'D_A0';
      t[k] += l.amount + gifts; t.D_A += l.amount + gifts; t.D_B += l.gift_cash; t.D_C += l.gift_goods; t.D_D += l.gift_coupon;
    }
    if (l.partner_type === '자점') t.JJ += l.amount;
  } else {
    if (l.item === '환수') t.U_F += -l.amount;
    else {
      const k = l.item === '관리수수료' ? 'U_A1' : l.item === '정산관련' ? 'U_A2' : 'U_A';
      t[k] += l.amount + gifts; t.U_B += l.gift_cash; t.U_C += l.gift_goods; t.U_D += l.gift_coupon;
    }
  }
}
function finish(t) {
  t.A = t.D_A - (t.D_B + t.D_C + t.D_D + t.D_F);
  t.B = (t.U_A + t.U_A1 + t.U_A2) - (t.U_B + t.U_C + t.U_D + t.U_F);
  t.BA = t.B - t.A;
  t.GIFT = t.D_B + t.D_C + t.D_D;
  // 매출 = 상부 최종(B) − 전체 최종(A) − 자점운영비 − 직접지급사은품(웨이브넷 지급분)
  t.REV = t.BA - t.JF - t.GD;
}

/** 집계행 → 화면 요약 숫자 (저장된 회차도 같은 함수로 다시 계산)
 *  4분류: 코드매출·외부유통매출(상부 기준) / 유통매출·직영영업매출(하부 기준). 두 묶음의 합계는 항상 같음 */
export function summarize(lines) {
  const seg = {};
  if (isLegacyLines(lines)) { // 이전 기준으로 저장된 회차
    for (const s of LEGACY_SEGMENTS) seg[s] = blank();
    for (const l of lines) if (l.included && seg[l.segment]) addLine(seg[l.segment], l);
    for (const s of LEGACY_SEGMENTS) finish(seg[s]);
    const rev = LEGACY_SEGMENTS.reduce((a, k) => a + seg[k].REV, 0);
    return { segments: seg, total: { REV: rev, REV2: rev }, legacy: true };
  }
  for (const s of SEGMENTS) seg[s] = blank();
  for (const l of lines) {
    if (!l.included) continue;
    for (const s of SEGMENTS) if (VIEW_FILTER[s](l)) addLine(seg[s], l);
  }
  for (const s of SEGMENTS) finish(seg[s]);
  const [a, b, c, d] = ['코드매출', '외부유통매출', '유통매출', '직영영업매출'].map(k => seg[k]);
  // 하부 × 상부 4칸 (협력점>코드, 자점>코드, 협력점>이관, 자점>이관)
  const cells = {};
  for (const pt of ['협력점', '자점']) for (const up of ['코드', '이관']) {
    const t = blank();
    for (const l of lines) if (l.included && l.segment === up && (l.partner_type === '자점' ? '자점' : '협력점') === pt) addLine(t, l);
    finish(t); cells[`${pt}|${up}`] = t.REV;
  }
  return { segments: seg, total: { REV: a.REV + b.REV, REV2: c.REV + d.REV, JF: a.JF + b.JF, cells } };
}
function blank() {
  return { JF: 0, GD: 0, D_A: 0, D_A0: 0, D_A1: 0, D_A2: 0, D_B: 0, D_C: 0, D_D: 0, D_F: 0, A: 0, U_A: 0, U_A1: 0, U_A2: 0, U_B: 0, U_C: 0, U_D: 0, U_F: 0, B: 0, BA: 0, GIFT: 0, REV: 0, JJ: 0 };
}

/** 상부점별 내역 (점검용) — 상부 구분 · 상부점 · 하부(협력점/자점)별 */
export function byUpper(lines) {
  const legacy = isLegacyLines(lines);
  const order = legacy ? LEGACY_SEGMENTS : ['코드', '이관'];
  const m = new Map();
  for (const l of lines) {
    if (!l.included) continue;
    const pt = l.partner_type === '자점' ? '자점' : '협력점';
    const k = l.segment + '|' + l.upper_name + '|' + pt;
    let r = m.get(k); if (!r) { r = { segment: l.segment, upper_name: l.upper_name, partner_type: pt, B: 0, A: 0, GD: 0, JF: 0, cntUp: 0, cntDn: 0 }; m.set(k, r); }
    r.JF += l.jj_fee || 0;
    if (l.side === '상부') { r.B += l.amount; r.cntUp += l.cnt; }
    else { r.A += l.amount; r.cntDn += l.cnt; r.GD += (l.direct_gift == null ? l.gift_cash + l.gift_goods + l.gift_coupon : l.direct_gift); }
  }
  return [...m.values()].map(r => ({ ...r, REV: r.B - r.A - r.JF - r.GD }))
    .sort((a, b) => order.indexOf(a.segment) - order.indexOf(b.segment) || a.upper_name.localeCompare(b.upper_name, 'ko') || (a.partner_type === '자점') - (b.partner_type === '자점'));
}
