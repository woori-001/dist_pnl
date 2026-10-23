// 손익매출 계산 엔진 — 브라우저와 Node 양쪽에서 동작 (XLSX 객체는 밖에서 넘겨받음)
import { RULES, SEGMENTS } from './rules.js';

const SIDE_BY_SHEET = { '상부정산관리': '상부', '전체정산관리': '전체' };
const NEED = ['정산종류', '협력점', '상부점', '상품명', '개통일', '정산예정일', '정산금액',
  '약정금액', '수납', '번들수수료', '동판수수료'];
const DEDUCT = ['사은품', '물류', '쿠폰', '선납', '할인탭', '미비서류', '기타'];

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
    });
  }
  return { fileName, side, rows, fileTotal };
}

export function segmentOf(upper) {
  if (!RULES.codePattern.test(upper)) return RULES.transferSegment;
  if (RULES.transferCodePrefixes.some(p => upper.trim().startsWith(p))) return RULES.transferSegment;
  const hit = RULES.codeSegments.find(c => upper.includes(c.match));
  return hit ? hit.segment : null; // null = 타부서 코드
}
export function isJajeom(partner) {
  return RULES.jajeomPattern.test(partner) && !RULES.jajeomExceptPrefixes.some(p => partner.trim().startsWith(p));
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
      if (!seg) { checks.otherDept++; continue; }
      const jj = isJajeom(r.partner);
      const included = !(seg === RULES.transferSegment && jj);
      const item = itemOf(r);
      const gift = r.kind === '환수' ? { c: 0, g: 0, k: 0 } : { c: r.d['사은품'], g: r.d['물류'], k: r.d['쿠폰'] };
      const key = [seg, r.upper, f.side, item, jj ? '자점' : '외부', included].join('|');
      let l = map.get(key);
      if (!l) { l = { segment: seg, upper_name: r.upper, side: f.side, item, partner_type: jj ? '자점' : '외부', included,
        cnt: 0, amount: 0, gift_cash: 0, gift_goods: 0, gift_coupon: 0 }; map.set(key, l); }
      l.cnt++; l.amount += r.amount; l.gift_cash += gift.c; l.gift_goods += gift.g; l.gift_coupon += gift.k;
      checks.usedRows++;
      const kf = f.side + '|' + r.kind; (kindFiles[kf] = kindFiles[kf] || new Set()).add(f.fileName);
    }
    const sk = f.side + JSON.stringify(kinds) + sumAll;
    if (seen.has(sk)) checks.warnings.push(`${f.fileName}: 같은 내용의 파일이 두 번 올라온 것 같습니다`);
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
  if (checks.formulaMismatch) checks.warnings.push(`정산금액 산식이 맞지 않는 행 ${checks.formulaMismatch}건`);
  return { lines: [...map.values()], checks };
}

/** 집계행 → 화면 요약 숫자 (저장된 회차도 같은 함수로 다시 계산) */
export function summarize(lines) {
  const seg = {};
  for (const s of SEGMENTS) seg[s] = blank();
  for (const l of lines) {
    if (!l.included) continue;
    const t = seg[l.segment]; if (!t) continue;
    const gifts = l.gift_cash + l.gift_goods + l.gift_coupon;
    if (l.side === '전체') {
      if (l.item === '환수') t.D_F += -l.amount;
      else { t.D_A += l.amount + gifts; t.D_B += l.gift_cash; t.D_C += l.gift_goods; t.D_D += l.gift_coupon; }
      if (l.partner_type === '자점') t.JJ += l.amount;
    } else {
      if (l.item === '환수') t.U_F += -l.amount;
      else {
        const k = l.item === '관리수수료' ? 'U_A1' : l.item === '정산관련' ? 'U_A2' : 'U_A';
        t[k] += l.amount + gifts; t.U_B += l.gift_cash; t.U_C += l.gift_goods; t.U_D += l.gift_coupon;
      }
    }
  }
  const total = { REV: 0 };
  for (const s of SEGMENTS) {
    const t = seg[s];
    t.A = t.D_A - (t.D_B + t.D_C + t.D_D + t.D_F);
    t.B = (t.U_A + t.U_A1 + t.U_A2) - (t.U_B + t.U_C + t.U_D + t.U_F);
    t.BA = t.B - t.A;
    t.GIFT = t.D_B + t.D_C + t.D_D;
    t.REV = t.BA; // 매출 = 상부 최종(B) − 전체 최종(A). 사은품은 A에서 이미 차감되어 있음
    total.REV += t.REV;
  }
  return { segments: seg, total };
}
function blank() {
  return { D_A: 0, D_B: 0, D_C: 0, D_D: 0, D_F: 0, A: 0, U_A: 0, U_A1: 0, U_A2: 0, U_B: 0, U_C: 0, U_D: 0, U_F: 0, B: 0, BA: 0, GIFT: 0, REV: 0, JJ: 0 };
}

/** 상부점별 내역 (점검용) */
export function byUpper(lines) {
  const m = new Map();
  for (const l of lines) {
    if (!l.included) continue;
    const k = l.segment + '|' + l.upper_name;
    let r = m.get(k); if (!r) { r = { segment: l.segment, upper_name: l.upper_name, B: 0, A: 0, GIFT: 0, cntUp: 0, cntDn: 0 }; m.set(k, r); }
    if (l.side === '상부') { r.B += l.amount; r.cntUp += l.cnt; }
    else { r.A += l.amount; r.cntDn += l.cnt; r.GIFT += l.gift_cash + l.gift_goods + l.gift_coupon; }
  }
  return [...m.values()].map(r => ({ ...r, REV: r.B - r.A }))
    .sort((a, b) => SEGMENTS.indexOf(a.segment) - SEGMENTS.indexOf(b.segment) || b.REV - a.REV);
}
