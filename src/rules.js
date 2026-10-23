// 계산 기준 — 기준이 바뀌면 이 파일만 고치면 됩니다.
export const RULES = {
  version: '2026-10-23d',

  // 상부점 이름이 "숫자." 으로 시작하면 우리 코드점, 아니면 이관처
  codePattern: /^\s*\d+\./,

  // 유통사업부가 운영하는 자체 코드 (상부점 이름에 포함된 글자로 판별)
  codeSegments: [
    { segment: 'SKB웨이브넷', match: 'SKB-웨이브넷' },
    { segment: 'SKY우리정보', match: 'SKY-우리' },
    { segment: 'SK매직',      match: 'SK매직' },
  ],
  transferSegment: '유통이관',
  // 숫자로 시작하지만 유통이관으로 보는 상부점 (0.기타)
  transferCodePrefixes: ['0.'],

  // 월 귀속: 여기 있는 정산종류는 개통일, 나머지(추가·환수 등)는 정산예정일이 대상월인 건
  openDateKinds: ['일반'],

  // 협력점이 "숫자." 으로 시작하면 자점. 단 0.영업정산 같은 정산 입력용은 자점으로 보지 않음
  jajeomPattern: /^\s*\d+\./,
  jajeomExceptPrefixes: ['0.'],

  // 상부정산 항목 구분 (상품명 기준)
  mgmtProductKeyword: '_관리',     // a-1 관리수수료
  settleProductKeyword: '정산관련', // a-2 정산관련
};

export const SEGMENTS = ['유통이관', 'SKB웨이브넷', 'SKY우리정보', 'SK매직'];
