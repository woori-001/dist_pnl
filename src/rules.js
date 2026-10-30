// 계산 기준 — 기준이 바뀌면 이 파일만 고치면 됩니다.
export const RULES = {
  version: '2026-10-27',

  // 상부점 이름이 "숫자." 으로 시작하면 코드(모든 코드: KT-우리·SKT-우성·대성·우신·SKB·SKY·매직 …), 아니면 이관
  codePattern: /^\s*\d+\./,
  // 숫자로 시작하지만 이관으로 보는 상부점 (0.기타)
  transferCodePrefixes: ['0.'],

  // 따로 보는 코드 3개 (상부점 이름에 포함된 글자로 판별) — 협력점·자점 모두, 자점운영비 차감
  codeSegments: [
    { segment: 'SKB웨이브넷', match: 'SKB-웨이브넷' },
    { segment: 'SKY우리정보', match: 'SKY-우리' },
    { segment: 'SK매직',      match: 'SK매직' },
  ],

  // 월 귀속: 여기 있는 정산종류는 개통일, 나머지(추가·환수 등)는 정산예정일이 대상월인 건
  openDateKinds: ['일반'],

  // 협력점이 "숫자." 으로 시작하면 자점 (0.영업정산·7.유통사업부 포함)
  jajeomPattern: /^\s*\d+\./,
  jajeomExceptPrefixes: [],

  // 자점운영비(자점 마진): 코드 상부 아래 자점 협력점의 '일반' 건에서 차감, 시행일은 접수일 기준
  //  - 대상 협력점: 숫자로 시작하는 자점 중 아래 접두어 제외 (7.유통사업부 = 유통 자체, 0.영업정산 = 정산 입력용)
  jajeomFeeExcludePrefixes: ['7.', '0.'],
  //  - 항목별 금액 (from = 시행일, 접수일이 시행일 이후인 가장 최근 금액 적용)
  jajeomFees: [
    { name: '인터넷 단독',      from: '2026-01-01', amount: 100000 },
    { name: '인터넷 번들',      from: '2026-01-01', amount: 120000 },
    { name: '인터넷 번들',      from: '2026-10-01', amount: 150000 },
    { name: '유심',             from: '2026-01-01', amount: 100000 },
    { name: '기기',             from: '2026-01-01', amount: 50000 },
    { name: '사전예약',         from: '2026-01-01', amount: 50000 },
    { name: '유심 가족동판',    from: '2026-01-01', amount: 50000 },
    { name: '기기 가족동판',    from: '2026-01-01', amount: 50000 },
    { name: '사전예약 가족동판', from: '2026-01-01', amount: 50000 },
    { name: '유심 상품 SKYLIFE', from: '2026-01-01', amount: 50000 },
  ],

  // 직접지급사은품: 사은품지급처가 아래가 아닌 건만 회사 지출로 봄 (유치자지급 = 협력점이 직접 지급)
  partnerPaidGiftLabel: '유치자지급',

  // ── 평가손익(경영지원 '배분 전·후 손익') 입력값 ──
  // 기대관리수수료 = 건수 × 월 단가 × 개월 (코드실적 계산기의 '예상관리'와 같은 방식)
  //   대상: 상부정산 · 코드 상부(숫자 시작, 0.기타 제외) · 일반 · 해지 제외 · 정산관련/추가지급/_관리 제외 · 자점 포함
  //   통신사: SKB·SKYLIFE만 (KT·SKT 제외)
  expectMonths: 28,
  expectFees: {
    SKYLIFE: { 인터넷: 950,  TV: 400, 유심: 1400 },
    SKB:     { 인터넷: 1350, TV: 800, 유심: 1500 },
    // KT·SKT는 기대관리를 넣지 않음 (실장님 확인 2026-10-30), SK매직은 유지수수료 없음
  },
  // 평가손익에서 '자점 제외'를 할 때 자점으로 보지 않는 협력점 (0.영업정산 = 관리·정산관련 입력용) — 확인 중
  evalNotJajeomPrefixes: ['0.'],
  // 평가손익 코드매출에 넣을 코드 상부 (null = 숫자로 시작하는 모든 코드 — 유통은 전 통신사, 우신·우성 사업부와 겹쳐도 됨)
  evalCodeUppers: null,
  // 기대관리 건수에서 보류(확인 중)로 따로 세는 상품명 글자
  expectHoldKeywords: ['약정갱신', '업셀링', '전환'],

  // 상부정산 항목 구분 (상품명 기준)
  mgmtProductKeyword: '_관리',     // a-1 관리수수료
  settleProductKeyword: '정산관련', // a-2 정산관련
};

// 유통사업부 4분류: 상부 기준(코드/이관) 두 개와 하부 기준(협력점/자점) 두 개 — 코드매출+외부유통매출 = 유통매출+직영영업매출
//   코드 = 숫자로 시작하는 모든 상부(KT·SKT·SKB·SKY·SK매직 …), 자점 = 숫자로 시작하는 모든 협력점(0.영업정산·1.시흥영업팀 …)
export const DIST_VIEWS = ['코드매출', '외부유통매출', '유통매출', '직영영업매출'];
export const CODE_VIEWS = RULES.codeSegments.map(c => c.segment);
export const SEGMENTS = [...DIST_VIEWS, ...CODE_VIEWS];         // 화면·엑셀에 나오는 구분 7개
export const STATUS_KEYS = ['유통사업부', ...CODE_VIEWS];          // 가정산/확정을 정하는 단위 4개
export const statusKeyOf = view => DIST_VIEWS.includes(view) ? '유통사업부' : view;
const codeMatch = seg => RULES.codeSegments.find(c => c.segment === seg).match;
export const VIEW_FILTER = {
  '코드매출': l => l.segment === '코드',
  '외부유통매출': l => l.segment === '이관',
  '유통매출': l => l.partner_type !== '자점',
  '직영영업매출': l => l.partner_type === '자점',
  ...Object.fromEntries(CODE_VIEWS.map(k => [k, l => l.segment === '코드' && (l.upper_name || '').includes(codeMatch(k))])),
};
// 2026-10-27 이전 저장 회차의 구분 (유통 범위만, 코드는 SKB·SKY·매직만)
export const LEGACY_SEGMENTS = ['유통이관', 'SKB웨이브넷', 'SKY우리정보', 'SK매직'];
