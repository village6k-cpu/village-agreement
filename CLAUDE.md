# 개고생2.0 (my-gas-project)

빌리지 렌탈샵 메인 관리 시스템. Google Apps Script + clasp.

## 배포
- `clasp push` → GAS 편집기에서 웹앱 새 버전 배포
- 웹앱 URL: `https://script.google.com/macros/s/AKfycbwX2V0SqRf23DCwaVojlc5YFXKTfMNLBt68edpGmCx8j0i9hkYdP_bXHKEGIcde2iS5EA/exec`
- GitHub Pages: `https://village6k-cpu.github.io/village-agreement/` (약관 동의 페이지, docs/index.html)

## 파일 구조
- **Code.js** — 메인 로직 (onChange 트리거, 세금계산서/현금영수증 발행, 알림톡, 입금매칭)
- **agreement.js** — 약관 동의 API (doGet/doPost, GitHub Pages 연동)
- **sheetProtection.js** — 시트 보호 설정
- **Sidebar.html** — 사이드바 UI
- **docs/index.html** — 약관 동의 모바일 페이지 (GitHub Pages)

## 거래내역 시트 컬럼 (COL_*)
A=날짜, B=예약자명, C=입금자명, D=거래ID, E=예약자연락처, F=상호, G=사업자번호, H=금액, I=결제수단, J=증빙유형, K=발행상태, L=입금상태, M=계약서링크, N=비고, O=관리키, P=약관동의일시

## 핵심 로직
- I열 "카드결제/현금" → K열 자동 "발행완료" + L열 "입금완료"
- I열 "계좌이체(VAT별도)" → K열 "발행완료" + L열 "미입금" (금액 /1.1)
- K열 "발행요청" → 팝빌 API로 실제 발행 실행
- 계약서 링크/금액은 5분 트리거로 자동 동기화
- 약관 동의: JSONP API → GitHub Pages에서 호출

## 외부 연동
- 팝빌 API (세금계산서, 현금영수증, 알림톡, 계좌조회)
- Google Drive (계약서 파일)
- GitHub Pages (약관 동의 페이지)

## 주의사항
- GAS는 같은 프로젝트 내 모든 .js 파일이 전역 스코프 공유
- doGet/doPost는 프로젝트당 하나만 가능 → agreement.js에 정의
- .clasp.json의 scriptId로 GAS 프로젝트 식별
