# 취업/이직 미니 진단

상담 예약자에게 발송하는 취업/이직 미니 진단 페이지.

## 구성
- `index.html`: 진단 페이지 (단일 파일)
- `apps-script/Code.gs`: 응답을 구글 시트에 쌓는 Apps Script

## 배포
1. Vercel에서 이 저장소를 Import (Framework: Other, 빌드 설정 없음)
2. 응답 시트에 `apps-script/Code.gs`를 웹 앱으로 배포 (실행: 나, 액세스: 모든 사용자)
3. 웹 앱 URL을 `index.html`의 `SHEET_ENDPOINT`에 입력

## 시연
- 주소 끝에 `?demo=1`을 붙이면 시연 메뉴가 보임
- 현재는 예시 고객 데이터로 동작함. 실제 고객 정보 연동은 추후 작업
