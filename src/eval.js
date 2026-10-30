// 평가손익(경영지원 '배분 전·배분 후 손익') 유통사업부 입력값 — 코드매출 / 외부유통매출
//  공통: 코드 = 숫자로 시작하는 상부(0.기타 제외), 외부유통 = 그 외. 협력점이 자점인 건은 제외(사은품·기대관리만 자점 포함)
//  월 기준은 계산기와 같음: 일반 = 개통일, 추가·환수 = 정산예정일, 옛날 데이터 포함
import { RULES } from './rules.js';
import { segmentOf, isJajeom, itemOf } from './calc.js';

export const EVAL_ROWS = ['코드매출', '외부유통매출'];
export const EVAL_COLS = [
  ['issue', '발생수수료', '상부정산 일반·추가 정산금액(관리수수료 제외, 정산관련 포함) + 기대관리'],
  ['expect', '기대관리(참고)', '건수 × 월 단가 × 28개월 · 발생수수료에 포함됨'],
  ['mgmt', '관리수수료', "상부정산 일반·추가 중 상품명 '_관리'"],
  ['refundUp', '환수(차감)', '상부정산 환수 — 우리에게 내려온 것'],
  ['refundDn', '환수(대체)', '전체정산 환수 — 우리가 하부점에 내린 것'],
  ['gift', '사은품', '전체정산 일반 · 현금(차감↔지급 확인)+물류+쿠폰 · 자점 포함'],
  ['partner', '협력점정산', '전체정산 일반·추가 정산금액(사은품 차감 후)'],
];
const segName = up => up === '코드' ? '코드매출' : '외부유통매출';
const blank = () => ({ issue: 0, issueBase: 0, expect: 0, mgmt: 0, refundUp: 0, refundDn: 0, gift: 0, giftCash: 0, giftGoods: 0, giftCoupon: 0, partner: 0 });

/** 기대관리 분류: 상품명 → { carrier, type } 또는 { hold } / null */
export function expectKind(product) {
  const p = (product || '').replace(/\s+/g, '');
  if (/정산관련|추가지급|_관리$/.test(p)) return null;
  if (RULES.expectHoldKeywords.some(k => p.includes(k))) return { hold: true };
  const mob = p.match(/^모바일_(.+)$/);
  if (mob) return RULES.expectFees[mob[1]] ? { carrier: mob[1], type: '유심' } : null;
  let carrier = p.split('_')[0], rest = p.slice(p.indexOf('_') + 1);
  if (carrier === '유선기타') { const m = rest.match(/^\((SK|KT)-biz\)/); if (!m) return null; carrier = m[1] === 'SK' ? 'SKB' : 'KT'; rest = rest.replace(/^\([^)]*\)/, ''); }
  if (!RULES.expectFees[carrier]) return null;
  // 코드실적 계산기와 같은 순서: CCTV·인터넷전화·일반전화를 먼저 거르고, TV를 인터넷보다 먼저
  if (/CCTV/i.test(rest) || /인터넷전화/.test(rest) || /일반전화|^전화$/.test(rest)) return null;
  if (/TV/.test(rest)) return { carrier, type: 'TV' };
  if (/인터넷/.test(rest)) return { carrier, type: '인터넷' };
  return null;
}

/** 사은품 현금: 차감_사은품 ↔ 사은품 쪽 현금(지급완료일 있는 것) 대조 */
function giftCheck(r) {
  const ded = r.d['사은품'], paid = r.giftDone && r.giftCash > 0 ? r.giftCash : 0;
  if ((r.giftPayer || '').includes(RULES.partnerPaidGiftLabel)) return { auto: 0, why: '유치자지급' };
  if (!ded && !paid) return { auto: 0 };
  if (ded === paid) return { auto: ded };
  return { ded, paid, why: !paid ? '차감만 있음 (지급 안 됨)' : !ded ? '지급만 있음 (차감 없음)' : '금액 다름' };
}

/** parsed 파일들 + 대상월 + 사은품 선택({key: 'ded'|'paid'|'none'}) → 평가손익 입력값 */
export function computeEval(parsed, period, choices = {}) {
  const rows = { 코드매출: blank(), 외부유통매출: blank() };
  const expect = {};      // carrier|type → { carrier, type, cnt, seg }
  let expectHold = 0, partnerPaidDone = 0;
  const giftList = [];    // 확인이 필요한 사은품 건 (화면 전용)
  for (const f of parsed) for (const r of f.rows) {
    const dateKey = RULES.openDateKinds.includes(r.kind) ? r.openDate : r.dueDate;
    if (!dateKey.startsWith(period)) continue;
    const up = segmentOf(r.upper); if (!up) continue;
    if (up === '코드' && RULES.evalCodeUppers && !RULES.evalCodeUppers.some(c => r.upper.includes(c))) continue; // 다른 사업부 코드
    const jj = isJajeom(r.partner) && !RULES.evalNotJajeomPrefixes.some(p => r.partner.trim().startsWith(p));
    const t = rows[segName(up)], item = itemOf(r);
    if (f.side === '상부') {
      if (r.kind === '환수') { if (!jj) t.refundUp += -r.amount; continue; }
      if (!jj) { if (item === '관리수수료') t.mgmt += r.amount; else t.issueBase += r.amount; }
      // 기대관리: 코드 · 일반 · 해지 제외 · 자점 포함
      if (up === '코드' && r.kind === '일반' && !(r.status || '').includes('해지')) {
        const k = expectKind(r.product);
        if (k && k.hold) expectHold++;
        else if (k) { const key = k.carrier + '|' + k.type; (expect[key] ||= { carrier: k.carrier, type: k.type, cnt: 0 }).cnt++; }
      }
    } else {
      if (r.kind === '환수') { if (!jj) t.refundDn += -r.amount; continue; }
      if (!jj) t.partner += r.amount;
      // 사은품: 일반 건(개통일 기준) · 자점 포함. 추가 행의 사은품 칸은 원 건의 복사라 쓰지 않음(차감 보정만 확인 목록에)
      if (r.kind === '일반') {
        t.giftGoods += r.d['물류']; t.giftCoupon += r.d['쿠폰'];
        const g = giftCheck(r);
        if (g.why === '유치자지급' && r.giftDone && r.giftCash) partnerPaidDone++;
        if (g.why && g.why !== '유치자지급') {
          const key = [f.side, r.no, r.customer, r.openDate, r.product].join('|');
          const pick = choices[key] || 'none';
          t.giftCash += pick === 'ded' ? g.ded : pick === 'paid' ? g.paid : 0;
          giftList.push({ key, seg: segName(up), why: g.why, ded: g.ded, paid: g.paid, pick, partner: r.partner, customer: r.customer, openDate: r.openDate, product: r.product, payer: r.giftPayer, done: r.giftDone });
        } else t.giftCash += g.auto || 0;
      } else if (r.d['사은품'] || r.d['물류'] || r.d['쿠폰']) {
        const key = [f.side, r.no, r.customer, r.dueDate, r.product, r.kind].join('|');
        const amt = r.d['사은품'] + r.d['물류'] + r.d['쿠폰'];
        const pick = choices[key] || 'none';
        if (pick === 'ded') t.giftCash += amt;
        giftList.push({ key, seg: segName(up), why: `${r.kind} 행 사은품 차감 보정`, ded: amt, paid: null, pick, partner: r.partner, customer: r.customer, openDate: r.openDate, product: r.product, payer: r.giftPayer, done: r.giftDone });
      }
    }
  }
  // 기대관리 금액 (코드매출에만)
  const expectRows = Object.values(expect).map(e => {
    const rate = RULES.expectFees[e.carrier][e.type];
    return { ...e, rate, months: RULES.expectMonths, amount: rate == null ? null : e.cnt * rate * RULES.expectMonths };
  }).sort((a, b) => a.carrier.localeCompare(b.carrier) || a.type.localeCompare(b.type));
  rows['코드매출'].expect = expectRows.reduce((s, e) => s + (e.amount || 0), 0);
  for (const k of EVAL_ROWS) {
    const t = rows[k];
    t.gift = t.giftCash + t.giftGoods + t.giftCoupon;
    t.issue = t.issueBase + t.expect;
  }
  const total = blank();
  for (const k of Object.keys(total)) total[k] = rows['코드매출'][k] + rows['외부유통매출'][k];
  const pending = giftList.filter(g => g.pick === 'none').length;
  const ratesMissing = expectRows.filter(e => e.amount == null);
  return { rows, total, expectRows, expectHold, ratesMissing: ratesMissing.map(e => `${e.carrier} ${e.type} ${e.cnt}건`), giftList, giftPending: pending, partnerPaidDone };
}

/** 저장용 (고객명 등 개인정보 없이 숫자만) */
export function evalForSave(ev) {
  return { rows: ev.rows, total: ev.total, expectRows: ev.expectRows, expectHold: ev.expectHold, ratesMissing: ev.ratesMissing,
    gift: { checked: ev.giftList.length, pending: ev.giftPending, picks: ev.giftList.reduce((a, g) => (a[g.pick] = (a[g.pick] || 0) + 1, a), {}) } };
}
