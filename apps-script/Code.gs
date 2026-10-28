/**
 * 취업/이직 미니 진단 → 응답 시트 저장
 *
 * 설치 방법
 * 1. "[취업정보회사] 미니 진단 응답" 시트에서 확장 프로그램 > Apps Script 를 연다.
 * 2. 기본 코드를 지우고 이 파일 내용을 그대로 붙여 넣고 저장한다.
 * 3. 배포 > 새 배포 > 유형 "웹 앱"
 *    - 다음 사용자 인증 정보로 실행: 나
 *    - 액세스 권한이 있는 사용자: 모든 사용자
 * 4. 배포 후 나오는 "웹 앱 URL"을 진단 페이지의 SHEET_ENDPOINT 값에 넣는다.
 *
 * 동작
 * - 진단 페이지가 결과지를 열 때 응답 한 건을 JSON으로 보낸다.
 * - 첫 행의 열 이름에 맞춰 한 줄을 추가한다. 시트에 없는 항목이 오면 맨 오른쪽에 열을 새로 만든다.
 * - 같은 연락처가 다시 제출하면 새 줄로 쌓인다(최신 응답이 아래).
 */

const SHEET_INDEX = 0; // 응답을 쌓을 탭 (첫 번째 탭)

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[SHEET_INDEX];

    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);

    // 시트에 없는 항목은 열을 새로 추가
    Object.keys(data).forEach(function (key) {
      if (headers.indexOf(key) === -1) {
        headers.push(key);
        sheet.getRange(1, headers.length).setValue(key);
      }
    });

    const row = headers.map(function (h) {
      return data[h] === undefined || data[h] === null ? '' : data[h];
    });
    sheet.appendRow(row);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 브라우저에서 URL을 열었을 때 연결 확인용
function doGet() {
  return json({ ok: true, message: '미니 진단 응답 수집 중' });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
