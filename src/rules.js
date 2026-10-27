// 계산 기준 — 기준이 바뀌면 이 파일만 고치면 됩니다.
export const RULES = {
  version: '2026-10-27',

  // 상부점 이름이 "숫자." 으로 시작하면 코드(모든 코드: KT-우리·SKT-우성·대성·우신·SKB·SKY·매직 …), 아니면 이관
  codePattern: /^\s*\d+\./,
  // 숫자로 시작하지만 이관으로 보는 상부점 (0.기타)
  transferCodePrefixes: ['0.'],

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

  // 상부정산 항목 구분 (상품명 기준)
  mgmtProductKeyword: '_관리',     // a-1 관리수수료
  settleProductKeyword: '정산관련', // a-2 정산관련
};

// 4분류: 상부 기준(코드/이관) 두 개와 하부 기준(협력점/자점) 두 개. 코드매출+외부유통매출 = 유통매출+직영영업매출
export const SEGMENTS = ['코드매출', '외부유통매출', '유통매출', '직영영업매출'];
export const VIEW_FILTER = {
  '코드매출': l => l.segment === '코드',
  '외부유통매출': l => l.segment === '이관',
  '유통매출': l => l.partner_type !== '자점',
  '직영영업매출': l => l.partner_type === '자점',
};
// 2026-10-27 이전 저장 회차의 구분 (유통 범위만, 코드는 SKB·SKY·매직만)
export const LEGACY_SEGMENTS = ['유통이관', 'SKB웨이브넷', 'SKY우리정보', 'SK매직'];
