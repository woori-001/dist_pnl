# 유통 손익매출 계산기

유통사업부 월별 손익매출(유통이관 / SKB웨이브넷 / SKY우리정보 / SK매직)을 어드민 정산 엑셀로 계산하고,
월별 결과를 **가정산 / 확정** 상태로 Supabase에 남기는 웹앱입니다.

- 엑셀은 **브라우저 안에서만** 읽고 계산합니다. 고객명·연락처 등 원자료는 서버로 보내지 않습니다.
- Supabase에는 집계 숫자(구분 × 상부점 × 상부/전체 × 항목 × 자점여부)만 저장합니다.
- 빌드 과정 없는 정적 사이트라 GitHub Pages에 그대로 올리면 됩니다.

## 매월 사용 순서 — 어드민에서 바로 가져오기 (권장)

1. (처음 한 번) 계산기 화면의 **⭳ 유통손익 가져오기** 버튼을 브라우저 즐겨찾기 막대로 끌어다 놓기
2. 어드민(workline.kr)에 로그인 → 즐겨찾기 **유통손익 가져오기** 클릭
3. 대상월 고르고 **가져오기** → 어드민 엑셀 다운로드 기능으로 6개 파일(전체/상부 × 일반[개통일] / 추가·환수[정산예정일])을 만들고, 계산기 창이 열리며 자동 계산
4. 결과 확인 → 가정산/확정 저장

- 즐겨찾기(북마클릿) 원본: `src/admin-pull.js`. 어드민 보안정책(CSP)상 외부 스크립트를 불러올 수 없어 코드 전체가 즐겨찾기에 들어갑니다. **코드를 바꾸면 즐겨찾기를 다시 끌어다 놓아야** 합니다.
- 어드민 데이터는 읽기만 하며(엑셀 작업 생성·다운로드), 파일은 브라우저 창 사이(postMessage, 출처 검증)로만 전달됩니다.
- 어드민 API: `/api/exact/calc/export`, `/api/exact/calc/upper/export` (chargeType G/A/B, dateChk 2=개통일 / 4=정산예정) → `/api/orders/export-jobs/{id}` → `/download`

## 매월 사용 순서 — 엑셀 직접 올리기

1. 어드민(workline.kr) 정산관리에서 4개 파일 다운로드
   - 전체정산관리 · 상부정산관리 각각
   - 일반: **개통일** 기준 대상월 1일~말일 (정산종류 '일반'만)
   - 추가·환수: **정산예정일** 기준 대상월 1일~말일 (정산종류 '추가'·'환수'만)
   - 같은 건이 두 파일에 겹치면 계산기가 경고를 띄웁니다
2. 계산기에서 대상월 선택 → 4개 파일 끌어다 놓기 → **계산하기**
3. 점검 메시지 확인(파일 합계 일치, 다른 달 행 여부) → 상부점별 내역으로 수기 대조
4. **가정산**으로 저장 → 본사 월정산·관리수수료가 다 들어온 뒤 다시 계산해 **확정**으로 저장
   (같은 달 확정은 1건만 유지, 새로 확정하면 이전 확정은 가정산으로 바뀜)

## 계산 기준

| 항목 | 기준 |
|---|---|
| 유통이관 | 상부점 이름이 숫자로 시작하지 않는 이관처 + `0.기타`. 협력점은 자점 포함 전체 (2026-10-26부터) |
| 코드 | 상부점에 `SKB-웨이브넷` / `SKY-우리` / `SK매직`. 자점·외부 협력점 모두 포함 |
| 사은품 | b.현금=차감-사은품, c.물류=차감-물류, d.쿠폰=차감-쿠폰 (일반·추가 건) |
| 상부 a-1 / a-2 | 상품명에 `_관리` → 관리수수료, `정산관련` → 정산관련 |
| 전체 A | a − (b+c+d+f) = 정산금액 합계 |
| 상부 B | (a+a1+a2) − (b+c+d+f) = 정산금액 합계 |
| **매출** | **B − A − 자점운영비 − 직접지급사은품** (A = 협력점 지급분, 사은품 = 회사가 고객에게 준 현금·물류·쿠폰, 자점운영비는 코드만) |
| 자점운영비 | 코드 전체정산 중 자점(7.유통사업부·0.영업정산 제외) 협력점의 일반 건. 접수일 기준 금액은 `src/rules.js`의 `jajeomFees` |
| 자점운영비 | 코드 매출 중 자점 협력점에 지급한 전체정산 금액(참고 표시) |

기준을 바꿀 때는 `src/rules.js`만 고치면 됩니다.

## 파일 구성

```
index.html                 화면
src/app.js                 화면 동작, Supabase 저장/조회, 엑셀 내보내기, 어드민 데이터 받기
src/admin-pull.js          어드민에서 실행하는 즐겨찾기(북마클릿) 원본
src/calc.js                계산 엔진 (브라우저·Node 공용)
src/rules.js               계산 기준
src/config.js              Supabase URL / publishable 키 (공개 키라 커밋해도 됨)
src/styles.css             스타일
supabase/migrations/       테이블·RLS·함수 SQL
test/verify.mjs            Node로 계산 결과 확인: node test/verify.mjs 2026-09 파일들.xlsx
```

## Supabase (woori-001's Project)

| 테이블 | 내용 |
|---|---|
| `dist_pnl_members` | 사용 가능 계정(이메일). 여기 있는 계정만 조회/저장 가능 |
| `dist_pnl_runs` | 월별 계산 회차 (period, status 가정산/확정, summary, 메모, 사용 파일) |
| `dist_pnl_lines` | 회차별 집계 행 |

- RLS: 로그인(authenticated) + `dist_pnl_members` 명단에 있는 이메일만 읽기/쓰기. anon 접근 불가.
- 저장은 `dist_pnl_save_run`, 상태 변경은 `dist_pnl_set_status` 함수(RLS 그대로 적용되는 invoker 함수).
- 사용자 추가: Supabase SQL Editor에서
  ```sql
  insert into public.dist_pnl_members (email, name) values ('someone@010king.com', '이름');
  ```
  (해당 이메일로 Supabase Auth 계정이 있어야 로그인 가능)

## GitHub Pages로 배포

1. GitHub에서 새 저장소(Private 권장) 생성 후 이 폴더를 push
2. Settings → Pages → Branch: `main` / root → Save
3. 몇 분 뒤 `https://<계정>.github.io/<저장소>/` 로 접속
   (Private 저장소의 Pages는 유료 플랜 필요. 무료면 Public 저장소 또는 Netlify/Vercel/Cloudflare Pages 사용 — 코드에 비밀값이 없어 Public이어도 데이터는 RLS로 보호됨)
