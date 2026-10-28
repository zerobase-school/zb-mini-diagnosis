/**
 * 취업/이직 미니 진단 → 응답 시트 저장 + PDF 결과지 발행
 *
 * 설치 방법
 * 1. "[취업정보회사] 미니 진단 응답" 시트에서 확장 프로그램 > Apps Script 를 연다.
 * 2. 기본 코드를 지우고 이 파일 내용을 그대로 붙여 넣고 저장한다.
 * 3. 함수 선택 칸에서 setupPdfTrigger 를 골라 실행한다 (처음 한 번만, 권한 승인 필요).
 *    1분마다 PDF가 없는 줄을 찾아 결과지를 만드는 자동 실행이 등록된다.
 * 4. 배포 > 배포 관리 > 연필 아이콘 > 버전: 새 버전 > 배포
 *    - 다음 사용자 인증 정보로 실행: 나
 *
 * 동작
 * - 링크에 연락처와 이름이 있으면 [ZB] SD 2.0 시트 DB RAW 탭에서 앞서 남긴 응답을 찾아 진단 페이지에 보여준다.
 * - 진단 페이지가 결과지를 열 때 응답 한 건을 JSON으로 보낸다.
 * - 응답은 바로 시트에 쌓이고, PDF 칸에는 "생성 중"이 들어간다.
 * - 1분마다 도는 자동 실행(makePendingPdfs)이 PDF 결과지를 만들어 드라이브 "미니 진단 결과지" 폴더에 저장하고 링크로 바꾼다.
 * - 시트 첫 행의 열 이름에 맞춰 한 줄을 추가한다. PDF 링크는 "연락처" 바로 옆 열에 들어간다.
 * - 시트에 없는 항목이 오면 맨 오른쪽에 열을 새로 만든다.
 */

const SHEET_ID = '1YS72FDuRNgBQVu9uD3x_ZZbTTTndEd4Vef1BUha7k7A'; // [취업정보회사] 미니 진단 응답
const SHEET_INDEX = 0;                 // 응답을 쌓을 탭 (첫 번째 탭)
const PDF_COL = '결과지 PDF';          // PDF 링크 열 이름
const PDF_FOLDER = '미니 진단 결과지';  // PDF를 저장할 폴더 이름 (시트와 같은 위치에 생성)
const PDF_PENDING = '생성 중';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[SHEET_INDEX];

    // 1) PDF는 자동 실행이 1분 안에 만들어 링크로 바꿈
    //    결과지 상세 내용(_report)은 시트 열 대신 스크립트 저장소에 잠시 보관
    const detail = data._report; delete data._report;
    data['응답 ID'] = Utilities.getUuid().slice(0, 8);
    if (detail) PropertiesService.getScriptProperties().setProperty('r_' + data['응답 ID'], detail);
    data[PDF_COL] = PDF_PENDING;

    // 2) 열 준비: 결과지 PDF는 연락처 바로 옆
    let headers = getHeaders(sheet);
    if (headers.indexOf(PDF_COL) === -1) {
      const phoneIdx = headers.indexOf('연락처');
      if (phoneIdx > -1) {
        sheet.insertColumnAfter(phoneIdx + 1);
        sheet.getRange(1, phoneIdx + 2).setValue(PDF_COL);
      } else {
        sheet.getRange(1, headers.length + 1).setValue(PDF_COL);
      }
      headers = getHeaders(sheet);
    }
    Object.keys(data).forEach(function (key) {
      if (headers.indexOf(key) === -1) {
        headers.push(key);
        sheet.getRange(1, headers.length).setValue(key);
      }
    });

    // 3) 한 줄 추가
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

function getHeaders(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const h = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  // 완전히 빈 시트면 빈 칸 하나만 나오므로 정리
  return h.length === 1 && h[0] === '' ? [] : h;
}

/* ---------- PDF 결과지 ---------- */

// 자동 실행: PDF 칸이 "생성 중"이거나 실패한 줄을 찾아 PDF를 만들고 링크로 바꾼다
function makePendingPdfs() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[SHEET_INDEX];
    const headers = getHeaders(sheet);
    const pdfIdx = headers.indexOf(PDF_COL);
    if (pdfIdx === -1 || sheet.getLastRow() < 2) return;
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    const formulas = sheet.getRange(2, pdfIdx + 1, values.length, 1).getFormulas();
    let made = 0;
    for (let i = 0; i < values.length && made < 20; i++) {
      const cell = String(values[i][pdfIdx]);
      if (formulas[i][0] || !(cell === PDF_PENDING || cell.indexOf('PDF 생성 실패') === 0)) continue;
      const data = {};
      headers.forEach(function (h, j) { const v = values[i][j]; data[h] = v instanceof Date ? Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd HH:mm') : v; });
      const target = sheet.getRange(i + 2, pdfIdx + 1);
      const key = 'r_' + data['응답 ID'], props = PropertiesService.getScriptProperties();
      let detail = null;
      try { detail = JSON.parse(props.getProperty(key) || 'null'); } catch (e) {}
      try {
        target.setFormula('=HYPERLINK("' + makePdf(data, detail) + '","PDF 보기")');
        props.deleteProperty(key);
      } catch (err) {
        target.setValue('PDF 생성 실패: ' + String(err));
      }
      made++;
    }
  } finally {
    lock.releaseLock();
  }
}

// 편집기에서 한 번만 실행: 1분마다 makePendingPdfs 자동 실행 등록 (이미 있으면 다시 만들지 않음)
function setupPdfTrigger() {
  const exists = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'makePendingPdfs'; });
  if (!exists) ScriptApp.newTrigger('makePendingPdfs').timeBased().everyMinutes(1).create();
  makePendingPdfs(); // 지금 쌓여 있는 줄도 바로 처리
  Logger.log('PDF 자동 생성 등록 완료');
}


function makePdf(d, detail) {
  const html = detail ? sheetHtml(d, detail) : reportHtml(d);
  const name = '미니진단_' + (d['이름'] || '이름없음') + '_' + String(d['연락처'] || '').replace(/\D/g, '').slice(-4) + '_' +
    Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmm') + '.pdf';
  const blob = htmlToPdf(html).setName(name);
  const file = getFolder().createFile(blob);
  try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return file.getUrl();
}

// HTML을 PDF로 변환
// Apps Script 편집기에서 서비스 > Drive API 를 추가해 두면 구글 문서를 거쳐 변환해 폰트(Noto Sans KR)가 정확히 나옴
// 추가하지 않았으면 기본 변환기를 사용
function htmlToPdf(html) {
  const src = Utilities.newBlob(html, 'text/html', 'report.html');
  if (typeof Drive !== 'undefined') {
    try {
      const doc = Drive.Files.create({ name: 'tmp_minidiag_' + Date.now(), mimeType: 'application/vnd.google-apps.document' }, src);
      const pdf = DriveApp.getFileById(doc.id).getAs('application/pdf');
      DriveApp.getFileById(doc.id).setTrashed(true);
      return pdf;
    } catch (e) {}
  }
  return src.getAs('application/pdf');
}

function getFolder() {
  // 시트가 있는 폴더를 찾고, 권한 문제로 못 찾으면 내 드라이브 첫 화면에 폴더를 만든다
  let base;
  try {
    const parents = DriveApp.getFileById(SHEET_ID).getParents();
    base = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  } catch (e) {
    base = DriveApp.getRootFolder();
  }
  const it = base.getFoldersByName(PDF_FOLDER);
  return it.hasNext() ? it.next() : base.createFolder(PDF_FOLDER);
}

function esc(v) {
  return String(v === undefined || v === null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function verdict(s) {
  s = Number(s) || 0;
  return s < 45 ? '보완 필요' : s < 75 ? '표준' : '우수';
}

/* ---------- 연차별 합격 사례와 제로베이스 전략 (PDF 전용) ----------
 * 출처: 슬랙 #zb_pm 설명회 합격 사례 정리, #zb_student_success 합격 공유 (2026년)
 * 고객 이름은 넣지 않음. 사례를 바꾸거나 추가할 때는 이 목록만 고치면 됨
 * [결과, 출발점, 제로베이스에서 한 일, 프로그램(내부 참고용, 결과지에는 표시 안 함)] */
const STAGE_CASES = {
  new: [
    ['LIG넥스원 최종 합격', '신입 취준생', '100문항 진단으로 목표 기업을 방산, 반도체 대기업으로 좁히고 지원 집중', '브라이트'],
    ['현대자동차, CJ제일제당, 삼일회계법인 인적성 합격', '신입 취준생', '컨설팅 2, 3회차에서 타깃 기업 비교와 적합도 진단 후 단기, 중기, 장기 실행 루틴 설계', '노바'],
    ['현대트랜시스 파워트레인 채용연계형 인턴', '첫 취업', '현대차 직무 특강과 멘토링으로 서류, 면접 집중 준비. 결제 후 2개월 만에 합격', '노바'],
    ['포스코DX 신입 공채 합격', '장교 전역 (군 경력 4년 6개월)', '멘토링 1회차에 경험과 전략 총괄 설계, 3회차에 경험 기반 지원 기업 정리', '프레스티지'],
    ['한국항공우주산업(KAI), 대한송유관공사 동시 최종 합격', '신입 취준생', '멘토링 12회차 동안 리포트 기반으로 기업 지원 병행', '']
  ],
  jr: [
    ['SK하이닉스 경력 서류 합격', '공정 엔지니어 1년 7개월', '컨설팅 1회차에 업무 로그 속 경험 정리, 멘토링 3회차에 실제 합격 자소서로 피드백', '노바'],
    ['토스 CX 최종 합격', '전환형 인턴 경력', '직무 경험을 CX 역량으로 재구성해 플랫폼 기업으로 확장 지원', ''],
    ['DL이앤씨 합격', '저연차 경력', '진행 중 서류, 면접 합격을 이어가며 목표 기업군 단계적 상향', '']
  ],
  mid: [
    ['신한투자증권 마케팅데이터 분석 이직', '은행 마케팅 7년차', '컨설팅 1회차에 경력 기반 Plan A(데이터 마케팅) 설정, 2회차에 대기업과 대형 금융사로 타깃 확정', '브라이트'],
    ['KT&G 향료 연구 이직', '연구직 5년 경력', '중도 포기 직전에서 목표를 다시 잡고 기업별 지원 전략 재정비', '노바'],
    ['현대홈쇼핑 CX 기획 이직', '총 경력 3년 6개월', '컨설팅 2, 3회차에 기획, CX 직무 추천과 서류 전략, 5회차 포트폴리오 피드백', '프리미엄'],
    ['CJ올리브영 클라우드 엔지니어 면접 합격', '금융권 3년 8개월', '채용 공고 분석 피드백으로 강점, 약점, 키워드 배치 순서 재구성 후 서류 합격', '노바'],
    ['쏘카 이직', '스타트업 5년 6개월', '모빌리티 산업을 1순위로 두고 산업 적합도 기준으로 지원 우선순위 설정', ''],
    ['자동차 기업 PM 연봉 1,500만 원 상향 이직', 'PM 5년차', '경력 성과를 수치 중심으로 재정리해 처우 협상까지 연결', '']
  ],
  sr: [
    ['JYP엔터테인먼트 최종 합격', '8년차', '컨설팅 단계에서 서류를 처음부터 다시 쓰고, 멘토링으로 사전과제와 1, 2차 면접 대비', ''],
    ['LH 법무 이직', '경력 6년 6개월', '경력을 공공기관 직무 요건에 맞춰 재구성', ''],
    ['한화자산운용 미들오피스 이직', '금융권 경력직', '멘토링 10회차를 모두 진행하며 희망 금융권 포지션으로 이동', '브라이트'],
    ['펑타이코리아 경력직 마케터', '경력직 마케터', '컨설팅 2회차 산업 VALUE CHAIN 분석에서 적합도 5점 기업을 우선 준비', '프리미엄']
  ]
};
STAGE_CASES.new2 = STAGE_CASES.new;

// 연차별 제로베이스만의 전략
const STAGE_STRATEGY = {
  new: ['100문항 진단으로 스펙 위치를 확인하고 목표 기업군을 대기업, 중견, 공기업 중 확정', '주차별 대기업 공채 마감 캘린더로 지원 일정을 놓치지 않게 관리', '타깃 기업 비교와 적합도 진단 후 단기, 중기, 장기 실행 루틴 설계', '기업별 인적성(GSAT, SKCT, HMAT 등) 모의고사와 스터디로 필기 대비'],
  jr: ['업무 로그 속 경험을 문제, 행동, 결과로 정리해 경력과 직무 역량 연결', '중고신입과 경력직 중 유리한 포지셔닝 판단', '같은 연차 실제 합격 자소서와 경력기술서로 1:1 비교 피드백', '1년차 경험도 수치 성과로 바꾸는 성과 수치화 작업'],
  mid: ['경력 기반 Plan A, B, C 설정 후 대기업, 대형 금융사 등 목표 도메인 확정', '산업 VALUE CHAIN 분석으로 기업별 적합도를 점수화해 우선순위 결정', '마스터 경력기술서, 포트폴리오 완성 프로그램(노바, 프레스티지 공통)', '상향 이직 시 처우 협상 포인트 정리'],
  sr: ['리더급, 비공개 포지션은 헤드헌팅으로 연결', '경력기술서를 처음부터 다시 설계하고 사전과제, 다차 면접 대비', '다음 10년 기준 커리어 Plan A, B, C 재설계', '연봉, 직급 처우 협상 전략']
};
STAGE_STRATEGY.new2 = STAGE_STRATEGY.new;

// 가장 약한 축별 제로베이스 처방
const AXIS_STRATEGY = {
  '업무 구조화': '합격 경력기술서의 구조화 패턴과 비교하고, 마스터 경력기술서 프로그램으로 경험을 문제, 행동, 결과 구조로 재작성',
  '목표 구체화': 'Plan A, B, C 기법으로 연차와 직무 기준 목표를 세우고, 목표 기업별 최근 합격자 조건과 비교해 지원 우선순위 확정',
  '취업/이직 전략': 'Pivot 프레임워크로 지금 강점을 축으로 옮길 경로를 설계하고, 지원 결과 기록으로 탈락 원인을 추적',
  '실전 준비도': '최신 합격 서류와 내 서류를 비교하고, 기업별 면접 복기 자료로 탈락 패턴과 예상 질문을 대비'
};

/* 검사지형 결과지 (InBody 결과지처럼 한 장에 수치와 막대를 촘촘히) */
function sheetHtml(d, r) {
  const INK = '#111', INK2 = '#444', INK3 = '#888', LINE = '#c4c4c4', HEAD = '#e3e3e3', SEG = '#ededed', NAVY = '#14206b', RED = '#b3263e';
  const F = "font-family:'Noto Sans KR','Noto Sans CJK KR','Apple SD Gothic Neo','Malgun Gothic',sans-serif;";
  const T = 'width:100%;border-collapse:collapse;';
  const w = r.axes[0];
  const cases = STAGE_CASES[r.stageKey] || [];
  const byName = {}; r.axes.forEach(function (a) { byName[a.name] = a; });
  const ORDER = ['업무 구조화', '목표 구체화', '취업/이직 전략', '실전 준비도'].filter(function (n) { return byName[n]; });

  // 회색 머리띠 (한글 제목 + 영문 부제)
  const band = function (ko, en) {
    return '<table style="' + T + 'margin-top:9px"><tr><td style="background:' + HEAD + ';padding:3px 7px;font-size:11.5px;font-weight:700;color:' + INK + '">' + ko +
      ' <span style="font-size:8.5px;font-weight:400;color:' + INK2 + ';font-style:italic">' + en + '</span></td></tr></table>';
  };
  // 0~100 눈금 막대 (20칸, 보완 필요/표준/우수 구간 음영)
  const scaleHead = function () {
    let cells = '';
    for (let i = 0; i < 20; i++) {
      const v = i * 5, lab = (v % 10 === 0) ? v : '';
      cells += '<td style="width:5%;font-size:7.5px;color:' + INK3 + ';text-align:left;padding:0;border-left:' + (v % 10 === 0 ? '0.5pt solid #cfcfcf' : '0') + '">' + lab + '</td>';
    }
    return cells;
  };
  const segHead = '<td colspan="9" style="font-size:8px;text-align:center;border:0.5pt solid #cfcfcf;background:#fff;padding:1px 0">보완 필요</td>' +
    '<td colspan="6" style="font-size:8px;text-align:center;border:0.5pt solid #cfcfcf;background:' + SEG + ';padding:1px 0">표준</td>' +
    '<td colspan="5" style="font-size:8px;text-align:center;border:0.5pt solid #cfcfcf;background:#fff;padding:1px 0">우수</td>';
  const barCells = function (score) {
    let cells = '';
    const s = Math.max(0, Math.min(100, Number(score) || 0));
    for (let i = 0; i < 20; i++) {
      const v = i * 5, fill = v < s, std = v >= 45 && v < 75;
      cells += '<td style="width:5%;height:7px;padding:0;font-size:1px;line-height:1px;background:' +
        (fill ? (s < 45 ? RED : INK) : (std ? SEG : '#fff')) + '"></td>';
    }
    return '<td colspan="20" style="padding:4px 0;border-bottom:0.5pt solid #e2e2e2"><table style="' + T + '"><tr>' + cells + '</tr></table></td>';
  };

  // ---------- 상단 ----------
  const top =
    '<table style="' + T + '"><tr>' +
    '<td style="font-size:30px;font-weight:900;letter-spacing:-1.5px;color:' + INK + '">ZEROBASE</td>' +
    '<td style="text-align:right;vertical-align:bottom;font-size:10px;color:' + INK2 + '">[취업/이직 미니 진단] [취업정보회사]</td>' +
    '<td style="text-align:right;vertical-align:bottom;width:150px;font-size:16px;font-weight:900;color:' + NAVY + '">제로베이스</td></tr></table>' +
    '<table style="' + T + 'margin-top:4px;border-top:1pt solid ' + INK + ';border-bottom:0.5pt solid ' + LINE + '"><tr>' +
    ['이름', '구분', '연차', '신청 DB', '검사일시'].map(function (h) { return '<td style="font-size:8.5px;color:' + INK2 + ';padding:3px 5px 0">' + h + '</td>'; }).join('') + '</tr><tr>' +
    [d['이름'], r.track, r.stage, r.db || '-', d['제출일시']].map(function (v) { return '<td style="font-size:11px;font-weight:700;padding:0 5px 4px">' + esc(v) + '</td>'; }).join('') +
    '</tr></table>';

  // ---------- 왼쪽 ----------
  // 1) 취업/이직 공식 분석 표
  const formula = band('취업/이직 공식 분석', 'Career Formula Analysis') +
    '<table style="' + T + 'border:0.5pt solid ' + LINE + '"><tr>' +
    '<td style="width:24%;border:0.5pt solid ' + LINE + '"></td>' +
    ['측정값', '기준 범위', '판정', '유형 글자'].map(function (h) { return '<td style="border:0.5pt solid ' + LINE + ';background:#f6f6f6;font-size:9px;text-align:center;padding:2px">' + h + '</td>'; }).join('') + '</tr>' +
    ORDER.map(function (n) {
      const a = byName[n];
      return '<tr><td style="border:0.5pt solid ' + LINE + ';padding:3px 6px;font-size:11px;font-weight:700">' + esc(n) + '</td>' +
        '<td style="border:0.5pt solid ' + LINE + ';text-align:center;font-size:13px;font-weight:700;color:' + (a.score < 45 ? RED : INK) + '">' + a.score + '</td>' +
        '<td style="border:0.5pt solid ' + LINE + ';text-align:center;font-size:9.5px;color:' + INK2 + '">(45.0 ~ 74.0)</td>' +
        '<td style="border:0.5pt solid ' + LINE + ';text-align:center;font-size:10px">' + esc(a.verdict) + '</td>' +
        '<td style="border:0.5pt solid ' + LINE + ';text-align:center;font-size:10px">' + esc(a.side) + '</td></tr>';
    }).join('') + '</table>';

  // 2) 축별 수준 분석 (눈금 막대)
  const levels = band('축별 수준 분석', 'Formula Level Analysis') +
    '<table style="' + T + '"><tr><td style="width:22%"></td>' + segHead + '<td style="width:9%"></td></tr>' +
    '<tr><td></td>' + scaleHead() + '<td></td></tr>' +
    ORDER.map(function (n) {
      const a = byName[n];
      return '<tr><td style="font-size:10.5px;font-weight:700;padding:4px 4px 4px 0;border-bottom:0.5pt solid #e2e2e2">' + esc(n) + ' <span style="font-size:8px;font-weight:400;color:' + INK3 + '">(점)</span></td>' +
        barCells(a.score) + '<td style="font-size:10.5px;font-weight:700;padding-left:5px;border-bottom:0.5pt solid #e2e2e2;color:' + (a.score < 45 ? RED : INK) + '">' + a.score + '</td></tr>';
    }).join('') + '</table>';

  // 3) 문항별 응답 분석 (0~3점)
  const qrows = ORDER.map(function (n) {
    const a = byName[n];
    return a.qa.map(function (x, i) {
      const s = x.s === null ? 0 : x.s;
      let cells = '';
      for (let k = 0; k < 3; k++) cells += '<td style="height:6px;padding:0;font-size:1px;line-height:1px;border-right:2px solid #fff;background:' + (k < s ? (s <= 1 ? RED : INK) : '#e4e4e4') + '"></td>';
      cells = '<td colspan="3" style="width:13%;padding:0 4px;border-bottom:0.5pt solid #e6e6e6' + (i === 2 ? ';border-bottom-color:' + LINE : '') + '"><table style="' + T + '"><tr>' + cells + '</tr></table></td>';
      return '<tr>' +
        (i === 0 ? '<td rowspan="3" style="width:16%;font-size:10px;font-weight:700;border-bottom:0.5pt solid ' + LINE + ';vertical-align:middle">' + esc(n) + '</td>' : '') +
        '<td style="font-size:8.5px;line-height:1.3;color:' + INK2 + ';padding:2px 4px;border-bottom:0.5pt solid #e6e6e6' + (i === 2 ? ';border-bottom-color:' + LINE : '') + '">' + esc(x.q) + '</td>' +
        cells + '<td style="width:5%;font-size:9px;text-align:right;padding-right:3px;color:' + (x.s !== null && x.s <= 1 ? RED : INK) + ';font-weight:700">' + (x.s === null ? '-' : x.s) + '</td>' +
        '<td style="width:24%;font-size:8.5px;line-height:1.3;font-weight:700;padding:2px 4px;border-bottom:0.5pt solid #e6e6e6' + (i === 2 ? ';border-bottom-color:' + LINE : '') + '">' + esc(x.a) + '</td></tr>';
    }).join('');
  }).join('');
  const qa = band('문항별 응답 분석', 'Item Response Analysis') +
    '<table style="' + T + '"><tr><td></td><td style="font-size:8px;color:' + INK3 + '">문항</td><td colspan="3" style="font-size:8px;color:' + INK3 + ';text-align:center">점수 (0~3)</td><td></td><td style="font-size:8px;color:' + INK3 + '">응답</td></tr>' + qrows + '</table>';

  // 4) 같은 연차 합격 사례
  const caseBlock = band('같은 연차 합격 사례 (' + esc(r.stage) + ')', 'Placement History') +
    '<table style="' + T + '">' + (cases.length ? cases.slice(0, 3).map(function (c) {
      return '<tr><td style="font-size:9.5px;font-weight:700;padding:3px 4px 3px 0;border-bottom:0.5pt solid #e6e6e6;width:38%;vertical-align:top">' + esc(c[0]) + '</td>' +
        '<td style="font-size:9px;color:' + INK2 + ';padding:3px 4px;border-bottom:0.5pt solid #e6e6e6;width:20%;vertical-align:top">' + esc(c[1]) + '</td>' +
        '<td style="font-size:9px;padding:3px 0 3px 4px;border-bottom:0.5pt solid #e6e6e6;vertical-align:top">' + esc(c[2]) + '</td></tr>';
    }).join('') : '<tr><td style="font-size:9.5px;color:' + INK2 + '">해당 연차 사례 없음</td></tr>') + '</table>';

  const left = top.replace(/^/, '') && (formula + levels + qa + caseBlock);

  // ---------- 오른쪽 ----------
  const rbox = function (ko, en, body) {
    return '<div style="margin-top:9px"><div style="font-size:11px;font-weight:700;border-bottom:0.75pt solid ' + INK + ';padding-bottom:2px">' + ko +
      ' <span style="font-size:8px;font-weight:400;color:' + INK2 + ';font-style:italic">' + en + '</span></div><div style="padding-top:4px">' + body + '</div></div>';
  };
  const scoreBox = rbox('진단 점수', 'Diagnosis Score',
    '<div style="text-align:center;padding:6px 0 2px"><span style="font-size:34px;font-weight:900">' + r.total + '</span><span style="font-size:13px"> / 100 점</span></div>' +
    '<div style="font-size:8.5px;color:' + INK2 + '">* 취업/이직 공식 4개 축 점수의 평균. 45점 미만 보완 필요, 75점 이상 우수. 현재 판정: <b>' + esc(r.verdict) + '</b></div>');
  const typeBox = rbox('취업/이직 유형', 'Career Type',
    '<table style="' + T + 'margin:3px 0"><tr>' + String(r.code).split('').map(function (ch) {
      return '<td style="border:0.75pt solid ' + INK + ';text-align:center;font-size:18px;font-weight:900;padding:3px 0">' + esc(ch) + '</td>';
    }).join('<td style="width:4px"></td>') + '</tr></table>' +
    '<div style="font-size:12px;font-weight:700;text-align:center">' + esc(r.typeName) + '</div>' +
    '<table style="' + T + 'margin-top:5px">' + ORDER.map(function (n) {
      const a = byName[n];
      return '<tr><td style="font-size:9px;padding:1px 0">' + esc(n) + '</td><td style="font-size:9px;text-align:right;padding:1px 0">' + esc(a.side) + ' <span style="color:' + INK3 + '">' + esc(a.strength) + '</span></td></tr>';
    }).join('') + '</table>');
  const planBox = rbox('추천 프로그램', 'Recommended Program',
    '<table style="' + T + '">' +
    '<tr><td style="font-size:9.5px;padding:2px 0">프로그램</td><td style="text-align:right;font-size:12px;font-weight:900;color:' + NAVY + '">' + esc(r.plan.name) + '</td></tr>' +
    '<tr><td style="font-size:9.5px;padding:2px 0">월 납부</td><td style="text-align:right;font-size:11px;font-weight:700">' + esc(r.plan.monthly) + '</td></tr>' +
    '<tr><td style="font-size:9.5px;padding:2px 0">총액</td><td style="text-align:right;font-size:10px">' + esc(r.plan.total) + ' <span style="font-size:8px;color:' + INK3 + '">(VAT 별도, 무이자 12개월)</span></td></tr></table>' +
    '<div style="font-size:9px;color:' + INK2 + ';margin-top:3px">' + esc(r.plan.why) + '</div>');
  const concernBox = rbox('고민과 추천 서비스', 'Concern',
    r.concern ? '<div style="font-size:10px;font-weight:700">' + esc(r.concern.title) + '</div><div style="font-size:9.5px;margin-top:2px">→ ' + esc(r.concern.service) + '</div><div style="font-size:8.5px;color:' + INK2 + '">' + esc(r.concern.desc) + '</div>' : '<div style="font-size:9.5px">미선택</div>');
  const stratBox = rbox('제로베이스 처방', 'Prescription',
    '<div style="font-size:9.5px;font-weight:700">' + esc(w.name) + ' (' + w.score + '점)</div><div style="font-size:9px;margin-top:2px">' + esc(AXIS_STRATEGY[w.name] || '') + '</div>' +
    '<div style="font-size:9.5px;font-weight:700;margin-top:5px">' + esc(r.stage) + ' 공통 전략</div>' +
    (STAGE_STRATEGY[r.stageKey] || []).map(function (x) { return '<div style="font-size:9px;margin-top:2px">- ' + esc(x) + '</div>'; }).join(''));
  const right = scoreBox + typeBox + planBox + concernBox + stratBox;

  // ---------- 2쪽: 상담 진행 가이드 ----------
  const lowQs = [];
  r.axes.forEach(function (a) { a.qa.forEach(function (x) { if (x.s !== null && x.s <= 1) lowQs.push([a.name, x.q, x.a]); }); });
  const steps = [
    ['오프닝', '유형 ' + esc(r.code) + ' ' + esc(r.typeName) + ', 종합 ' + r.total + '점을 함께 보며 시작'],
    ['고민 확인', r.concern ? '"' + esc(r.concern.title) + '"의 구체적 상황과 지금까지 시도한 방법 질문' : '현재 가장 막히는 지점 질문'],
    ['핵심 진단', esc(w.name) + ' ' + w.score + '점이 최저. ' + esc(w.flag)],
    ['해결 방향', esc((w.locked || []).join(', '))],
    ['사례 제시', cases.length ? cases.slice(0, 2).map(function (c) { return esc(c[1]) + ' → ' + esc(c[0]); }).join(' / ') : '-'],
    ['프로그램 제안', esc(r.plan.name) + ' 월 ' + esc(r.plan.monthly) + '. ' + esc(r.plan.why)]
  ];
  const page2 =
    '<div style="page-break-before:always"></div>' +
    '<table style="' + T + 'border-bottom:2px solid ' + INK + '"><tr><td style="font-size:16px;font-weight:900">상담 진행 가이드</td><td style="text-align:right;font-size:9px;color:' + INK2 + '">' + esc(d['이름']) + '님  |  세일즈팀 내부용</td></tr></table>' +
    band('진행 순서', 'Consulting Flow') +
    '<table style="' + T + '">' + steps.map(function (s, i) {
      return '<tr><td style="width:20px;font-size:10px;font-weight:700;padding:4px 0;border-bottom:0.5pt solid #e2e2e2;vertical-align:top">' + (i + 1) + '</td>' +
        '<td style="width:80px;font-size:10px;font-weight:700;padding:4px 0;border-bottom:0.5pt solid #e2e2e2;vertical-align:top">' + s[0] + '</td>' +
        '<td style="font-size:10px;padding:4px 0;border-bottom:0.5pt solid #e2e2e2">' + s[1] + '</td></tr>';
    }).join('') + '</table>' +
    band('확인이 필요한 응답', 'Low Score Items (1점 이하)') +
    '<table style="' + T + '">' + (lowQs.length ? lowQs.map(function (x) {
      return '<tr><td style="width:90px;font-size:9.5px;color:' + INK2 + ';padding:3px 0;border-bottom:0.5pt solid #e6e6e6">' + esc(x[0]) + '</td><td style="font-size:9.5px;padding:3px 4px;border-bottom:0.5pt solid #e6e6e6">' + esc(x[1]) + '</td><td style="width:150px;font-size:9.5px;font-weight:700;color:' + RED + ';padding:3px 0;border-bottom:0.5pt solid #e6e6e6">' + esc(x[2]) + '</td></tr>';
    }).join('') : '<tr><td style="font-size:9.5px">없음</td></tr>') + '</table>' +
    band('축별 상담 의제', 'Agenda by Axis') +
    '<table style="' + T + '">' + r.axes.map(function (a) {
      return '<tr><td style="width:90px;font-size:10px;font-weight:700;padding:4px 0;border-bottom:0.5pt solid #e2e2e2;vertical-align:top">' + esc(a.name) + '<div style="font-size:9px;font-weight:400;color:' + (a.score < 45 ? RED : INK2) + '">' + a.score + '점 ' + esc(a.verdict) + '</div></td>' +
        '<td style="font-size:9.5px;padding:4px 0 4px 6px;border-bottom:0.5pt solid #e2e2e2"><b>진단</b> ' + esc(a.flag) + '<br><b>함께 볼 부분</b> ' + esc(a.point) + '<br><b>상담 의제</b> ' + esc((a.locked || []).join(', ')) + '</td></tr>';
    }).join('') + '</table>' +
    band('추천 프로그램 제공 내용', esc(r.plan.name)) +
    '<div style="font-size:9.5px;padding-top:3px">' + (r.plan.items || []).map(esc).join('  |  ') + '</div>';

  return '<!doctype html><html><head><meta charset="utf-8"><style>' +
    '@page{size:A4;margin:12mm}body{' + F + 'color:' + INK + ';margin:0;line-height:1.45}td,div,span,b{' + F + '}' +
    '</style></head><body>' +
    top +
    '<table style="' + T + 'margin-top:2px"><tr>' +
    '<td style="width:66%;vertical-align:top;padding-right:12px;border-right:0.5pt solid #d5d5d5">' + left + '</td>' +
    '<td style="width:34%;vertical-align:top;padding-left:12px">' + right + '</td></tr></table>' +
    '<div style="border-top:0.5pt solid #d5d5d5;margin-top:10px;padding-top:4px;font-size:8px;color:' + INK3 + '">본 결과지는 12개 자가 응답 문항으로 산출한 참고용 진단임. 고객 화면에는 유형까지만 공개. 응답 ID ' + esc(d['응답 ID'] || '-') + '  |  Copyright ZEROBASE</div>' +
    '</body></html>';
}

/* 상세 결과지 (세일즈팀 상담 준비용) — 컨설팅 리포트 톤 */
function richHtml(d, r) {
  const NAVY = '#14206b', NAVY2 = '#2f3a9e', GRAY = '#f4f5f8', INK = '#16181f', INK2 = '#4b5060', INK3 = '#8a8f9e', RULE = '#d7d9e2', RED = '#b3263e';
  const F = "font-family:'Noto Sans KR','Noto Sans CJK KR','Apple SD Gothic Neo','Malgun Gothic',sans-serif;";
  const T = 'width:100%;border-collapse:collapse;';
  const w = r.axes[0];
  const cases = STAGE_CASES[r.stageKey] || [];

  const sec = function (no, title, msg) {
    return '<table style="' + T + 'margin:26px 0 10px"><tr>' +
      '<td style="width:34px;font-size:15px;font-weight:700;color:' + NAVY2 + ';border-bottom:1.5px solid ' + NAVY + ';padding:0 0 6px">' + no + '</td>' +
      '<td style="font-size:14px;font-weight:700;color:' + NAVY + ';border-bottom:1.5px solid ' + NAVY + ';padding:0 0 6px">' + title + '</td></tr></table>' +
      (msg ? '<div style="font-size:12.5px;font-weight:700;color:' + INK + ';margin:0 0 10px">' + msg + '</div>' : '');
  };
  const th = function (t, wd) { return '<td style="padding:7px 8px;background:' + NAVY + ';color:#fff;font-size:10.5px;font-weight:700;' + (wd ? 'width:' + wd + ';' : '') + '">' + t + '</td>'; };
  const td = function (t, st) { return '<td style="padding:7px 8px;border-bottom:1px solid ' + RULE + ';vertical-align:top;font-size:11px;' + (st || '') + '">' + t + '</td>'; };
  const bullets = function (arr) {
    return '<table style="' + T + '">' + arr.map(function (x) {
      return '<tr><td style="width:14px;vertical-align:top;color:' + NAVY2 + ';font-size:11px;padding:2px 0">■</td><td style="font-size:11.5px;padding:2px 0;color:' + INK + '">' + x + '</td></tr>';
    }).join('') + '</table>';
  };
  const bar = function (score) {
    const s = Math.max(0, Math.min(100, Number(score) || 0)), color = s < 45 ? RED : NAVY;
    return '<table style="' + T + '"><tr>' +
      (s > 0 ? '<td style="width:' + s + '%;height:8px;background:' + color + '"></td>' : '') +
      (s < 100 ? '<td style="height:8px;background:#e6e8ef"></td>' : '') + '</tr></table>';
  };

  // 핵심 메시지 (명사형)
  const keyMsg = esc(w.name) + ' 보완이 최우선 과제이며, ' + esc(r.plan.name) + ' 프로그램 제안이 적합한 상황';

  // 표지 헤더
  const header =
    '<table style="' + T + '"><tr>' +
    '<td style="font-size:9.5px;letter-spacing:1.5px;color:' + INK3 + '">ZEROBASE CAREER AGENCY  |  INTERNAL</td>' +
    '<td style="text-align:right;font-size:9.5px;color:' + INK3 + '">세일즈팀 상담 준비용</td></tr></table>' +
    '<div style="height:3px;background:' + NAVY + ';margin:8px 0 14px"></div>' +
    '<div style="font-size:22px;font-weight:700;color:' + INK + '">취업/이직 미니 진단 결과지</div>' +
    '<div style="font-size:11.5px;color:' + INK2 + ';margin-top:4px">' + esc(d['이름']) + '님  |  ' + esc([r.track, r.stage].filter(Boolean).join(', ')) + (r.db ? '  |  ' + esc(r.db) : '') + '</div>' +
    '<table style="' + T + 'margin-top:14px;border-top:1px solid ' + RULE + ';border-bottom:1px solid ' + RULE + '"><tr>' +
    [['진단 유형', esc(r.code) + '  ' + esc(r.typeName)], ['종합 점수', r.total + '점 (' + esc(r.verdict) + ')'], ['추천 프로그램', esc(r.plan.name)], ['제출일시', esc(d['제출일시'])]].map(function (x, i) {
      return '<td style="padding:10px 10px;' + (i ? 'border-left:1px solid ' + RULE + ';' : '') + 'vertical-align:top">' +
        '<div style="font-size:9.5px;color:' + INK3 + '">' + x[0] + '</div><div style="font-size:12.5px;font-weight:700;color:' + (i === 2 ? NAVY2 : INK) + ';margin-top:2px">' + x[1] + '</div></td>';
    }).join('') + '</tr></table>';

  // 핵심 요약 박스
  const exec =
    '<table style="' + T + 'margin-top:16px"><tr><td style="width:4px;background:' + NAVY + '"></td><td style="background:' + GRAY + ';padding:14px 16px">' +
    '<div style="font-size:9.5px;letter-spacing:1px;color:' + NAVY2 + ';font-weight:700">KEY MESSAGE</div>' +
    '<div style="font-size:14px;font-weight:700;color:' + INK + ';margin:4px 0 8px">' + keyMsg + '</div>' +
    bullets([
      (r.concern ? '고객 고민: ' + esc(r.concern.title) + ' → ' + esc(r.concern.service) : '고객 고민: 미선택'),
      '약점: ' + esc(w.name) + ' ' + w.score + '점 (' + esc(w.verdict) + '). ' + esc(w.flag),
      '강점: ' + esc((r.strengths || []).join(', ') || '두드러진 강점 없음'),
      '제안: ' + esc(r.plan.name) + ' 월 ' + esc(r.plan.monthly) + ' (무이자 12개월, 총 ' + esc(r.plan.total) + ', VAT 별도). ' + esc(r.plan.why)
    ]) + '</td></tr></table>';

  // 01 점수
  const scoreTable = '<table style="' + T + '"><tr>' + th('평가 축', '110px') + th('점수 분포') + th('점수', '52px') + th('판정', '64px') + th('성향', '150px') + '</tr>' +
    r.axes.map(function (a) {
      return '<tr>' + td('<b>' + esc(a.name) + '</b>') + td(bar(a.score), 'vertical-align:middle;') +
        td('<b style="color:' + (a.score < 45 ? RED : INK) + '">' + a.score + '</b>', 'text-align:right;') +
        td(esc(a.verdict), 'color:' + (a.score < 45 ? RED : INK2) + ';') + td(esc(a.side) + ', ' + esc(a.strength), 'color:' + INK2 + ';') + '</tr>';
    }).join('') + '</table>' +
    '<div style="font-size:9.5px;color:' + INK3 + ';margin-top:5px">기준: 45점 미만 보완 필요, 45~74점 표준, 75점 이상 우수. 축별 56점 이상이면 유형 앞 글자. 점수가 낮은 축부터 정렬.</div>';

  // 02 합격 사례
  const caseTable = cases.length ? '<table style="' + T + '"><tr>' + th('합격 결과', '33%') + th('출발점', '20%') + th('제로베이스에서 한 일') + '</tr>' +
    cases.map(function (c) {
      return '<tr>' + td('<b>' + esc(c[0]) + '</b>') + td(esc(c[1]), 'color:' + INK2 + ';') + td(esc(c[2])) + '</tr>';
    }).join('') + '</table>' +
    '<div style="font-size:9.5px;color:' + INK3 + ';margin-top:5px">출처: 제로베이스 회원 합격 공유 (2026). 상담에서는 고객과 조건이 가까운 1~2건을 골라 소개.</div>'
    : '<div style="font-size:11px;color:' + INK2 + '">해당 연차 사례 없음</div>';

  // 03 전략
  const strat = '<table style="' + T + '"><tr>' +
    '<td style="width:50%;vertical-align:top;padding-right:10px"><div style="font-size:11px;font-weight:700;color:' + NAVY + ';border-bottom:1px solid ' + RULE + ';padding-bottom:4px;margin-bottom:6px">' + esc(r.stage) + ' 공통 전략</div>' + bullets((STAGE_STRATEGY[r.stageKey] || []).map(esc)) + '</td>' +
    '<td style="width:50%;vertical-align:top;padding-left:10px"><div style="font-size:11px;font-weight:700;color:' + NAVY + ';border-bottom:1px solid ' + RULE + ';padding-bottom:4px;margin-bottom:6px">' + esc(w.name) + ' 처방</div>' + bullets([esc(AXIS_STRATEGY[w.name] || '')]) +
    '<div style="font-size:11px;font-weight:700;color:' + NAVY + ';border-bottom:1px solid ' + RULE + ';padding-bottom:4px;margin:10px 0 6px">' + esc(r.plan.name) + ' 제공 내용</div>' + bullets((r.plan.items || []).map(esc)) + '</td>' +
    '</tr></table>';

  // 04 상담 가이드
  const lowQs = [];
  r.axes.forEach(function (a) { a.qa.forEach(function (x) { if (x.s !== null && x.s <= 1) lowQs.push([a.name, x.q, x.a]); }); });
  const steps = [
    ['오프닝', '유형 ' + esc(r.code) + ' ' + esc(r.typeName) + ', 종합 ' + r.total + '점을 함께 보며 시작'],
    ['고민 확인', r.concern ? '"' + esc(r.concern.title) + '"의 구체적 상황과 지금까지 시도한 방법 질문' : '현재 가장 막히는 지점 질문'],
    ['핵심 진단', esc(w.name) + ' ' + w.score + '점이 최저. ' + esc(w.flag)],
    ['해결 방향', esc((w.locked || []).join(', '))],
    ['사례 제시', cases.length ? cases.slice(0, 2).map(function (c) { return esc(c[1]) + ' → ' + esc(c[0]); }).join(' / ') : '-'],
    ['프로그램 제안', esc(r.plan.name) + ' 월 ' + esc(r.plan.monthly) + '. ' + esc(r.plan.why)]
  ];
  const guide = '<table style="' + T + '">' + steps.map(function (s, i) {
    return '<tr><td style="width:24px;padding:7px 0;vertical-align:top;font-size:11px;font-weight:700;color:' + NAVY2 + ';border-bottom:1px solid ' + RULE + '">' + (i + 1) + '</td>' +
      '<td style="width:86px;padding:7px 0;vertical-align:top;font-size:11px;font-weight:700;border-bottom:1px solid ' + RULE + '">' + s[0] + '</td>' +
      '<td style="padding:7px 0;font-size:11px;border-bottom:1px solid ' + RULE + '">' + s[1] + '</td></tr>';
  }).join('') + '</table>' +
    (lowQs.length ? '<div style="font-size:11px;font-weight:700;margin:12px 0 6px">확인이 필요한 응답 (1점 이하, ' + lowQs.length + '개)</div>' +
      '<table style="' + T + '"><tr>' + th('평가 축', '90px') + th('문항') + th('응답', '150px') + '</tr>' +
      lowQs.map(function (x) { return '<tr>' + td(esc(x[0]), 'color:' + INK2 + ';') + td(esc(x[1])) + td('<b style="color:' + RED + '">' + esc(x[2]) + '</b>') + '</tr>'; }).join('') + '</table>' : '');

  // 05 축별 상세
  const detail = r.axes.map(function (a, idx) {
    return '<div style="page-break-inside:avoid;margin-bottom:16px">' +
      '<table style="' + T + 'border-bottom:1px solid ' + INK + '"><tr>' +
      '<td style="padding:0 0 5px;font-size:12.5px;font-weight:700">' + (idx + 1) + '. ' + esc(a.name) + '  <span style="font-size:10.5px;font-weight:400;color:' + INK2 + '">' + esc(a.side) + '</span></td>' +
      '<td style="padding:0 0 5px;text-align:right;font-size:12.5px;font-weight:700;color:' + (a.score < 45 ? RED : NAVY) + '">' + a.score + '점  <span style="font-size:10.5px;font-weight:400;color:' + INK2 + '">' + esc(a.verdict) + '</span></td></tr></table>' +
      '<table style="' + T + 'margin-top:6px">' +
      '<tr><td style="width:96px;padding:4px 0;vertical-align:top;font-size:10.5px;font-weight:700;color:' + NAVY2 + '">진단</td><td style="padding:4px 0;font-size:11px">' + esc(a.flag) + '</td></tr>' +
      '<tr><td style="padding:4px 0;vertical-align:top;font-size:10.5px;font-weight:700;color:' + NAVY2 + '">함께 볼 부분</td><td style="padding:4px 0;font-size:11px">' + esc(a.point) + '</td></tr>' +
      '<tr><td style="padding:4px 0;vertical-align:top;font-size:10.5px;font-weight:700;color:' + NAVY2 + '">상담 의제</td><td style="padding:4px 0;font-size:11px">' + esc((a.locked || []).join(', ')) + '</td></tr></table>' +
      '<table style="' + T + 'margin-top:6px"><tr>' + th('No.', '34px') + th('문항') + th('응답', '160px') + th('점수', '40px') + '</tr>' +
      a.qa.map(function (x, i) {
        return '<tr>' + td('Q' + (i + 1), 'color:' + INK2 + ';') + td(esc(x.q), 'color:' + INK2 + ';') + td('<b>' + esc(x.a) + '</b>') +
          td(x.s === null ? '-' : x.s + '/3', 'text-align:right;' + (x.s !== null && x.s <= 1 ? 'color:' + RED + ';font-weight:700;' : 'color:' + INK2 + ';')) + '</tr>';
      }).join('') + '</table></div>';
  }).join('');

  return '<!doctype html><html><head><meta charset="utf-8"><style>' +
    'body{' + F + 'color:' + INK + ';font-size:11px;margin:30px 36px;line-height:1.6}' +
    'td,div,span,b{' + F + '}' +
    '</style></head><body>' +
    header + exec +
    sec('01', '취업/이직 공식 4개 축 진단', esc(w.name) + ' ' + w.score + '점으로 가장 낮고, ' + esc((r.strengths || []).length ? r.strengths.length + '개 축은 기준 이상' : '모든 축이 기준 미만')) + scoreTable +
    sec('02', '같은 연차 합격 사례 (' + esc(r.stage) + ')', '같은 연차에서 제로베이스와 함께 목표를 달성한 사례') + caseTable +
    sec('03', '제로베이스 전략', '연차 공통 전략과 약점 축 처방을 함께 제시') + strat +
    sec('04', '상담 진행 가이드', '아래 순서로 진단 결과를 해설하고 프로그램 제안으로 연결') + guide +
    sec('05', '축별 상세 진단', '') + detail +
    '<div style="border-top:1px solid ' + RULE + ';margin-top:20px;padding-top:8px;font-size:9.5px;color:' + INK3 + '">' +
    '본 결과지는 12개 자가 응답 문항으로 산출한 참고용 진단임. 고객 화면에는 유형까지만 공개되며, 상세 내용은 상담에서 안내. 응답 ID ' + esc(d['응답 ID'] || '-') + '</div>' +
    '</body></html>';
}

function reportHtml(d) {
  const NAVY = '#1d2483', LAV = '#eeeefb', INK = '#1a1d2e', INK2 = '#5a5f73', LINE = '#e3e4ee', ROSE = '#d23f6a';
  const axes = [
    ['업무 구조화', '구조화'], ['목표 구체화', '목표'], ['취업/이직 전략', '전략'], ['실전 준비도', '실전']
  ];
  const total = Number(d['종합 점수']) || 0;

  const bar = function (score) {
    const s = Math.max(0, Math.min(100, Number(score) || 0));
    const color = s < 45 ? ROSE : NAVY;
    return '<table style="width:100%;border-collapse:collapse"><tr>' +
      '<td style="width:' + s + '%;height:10px;background:' + color + ';border-radius:5px 0 0 5px"></td>' +
      '<td style="height:10px;background:' + LAV + '"></td></tr></table>';
  };

  const axisRows = axes.map(function (a) {
    const sc = d[a[0]];
    return '<tr>' +
      '<td style="padding:10px 0;width:110px;font-weight:700;color:' + INK + '">' + a[0] + '</td>' +
      '<td style="padding:10px 12px">' + bar(sc) + '</td>' +
      '<td style="padding:10px 0;width:60px;text-align:right;font-weight:800;color:' + INK + '">' + esc(sc) + '점</td>' +
      '<td style="padding:10px 0 10px 10px;width:70px;text-align:right;color:' + INK2 + ';font-size:11px">' + verdict(sc) + '</td>' +
      '</tr>';
  }).join('');

  const qaRows = axes.map(function (a) {
    return [1, 2, 3].map(function (i) {
      const k = a[1] + ' Q' + i;
      return '<tr><td style="padding:7px 10px;border-bottom:1px solid ' + LINE + ';color:' + INK2 + ';width:110px">' + (i === 1 ? a[0] : '') + '</td>' +
        '<td style="padding:7px 10px;border-bottom:1px solid ' + LINE + ';color:' + INK2 + ';width:40px">Q' + i + '</td>' +
        '<td style="padding:7px 10px;border-bottom:1px solid ' + LINE + ';color:' + INK + '">' + esc(d[k]) + '</td></tr>';
    }).join('');
  }).join('');

  const info = [
    ['이름', d['이름']], ['연락처', d['연락처']], ['구분', d['구분']], ['연차', d['연차']],
    ['신청 DB', d['신청 DB']], ['제출일시', d['제출일시']]
  ].filter(function (x) { return x[1]; }).map(function (x) {
    return '<tr><td style="padding:6px 0;color:' + INK2 + ';width:90px">' + x[0] + '</td><td style="padding:6px 0;color:' + INK + ';font-weight:600">' + esc(x[1]) + '</td></tr>';
  }).join('');

  const code = String(d['유형 코드'] || '').split('').map(function (ch) {
    return '<span style="display:inline-block;width:38px;height:44px;line-height:44px;margin-right:6px;text-align:center;border:1px solid rgba(255,255,255,.35);border-radius:8px;font-size:24px;font-weight:800">' + esc(ch) + '</span>';
  }).join('');

  return '<!doctype html><html><head><meta charset="utf-8"><style>' +
    'body{font-family:"Noto Sans KR","Malgun Gothic",sans-serif;color:' + INK + ';font-size:12px;margin:28px 34px;line-height:1.55}' +
    'h2{font-size:14px;margin:22px 0 8px;color:' + NAVY + '}' +
    '.box{border:1px solid ' + LINE + ';border-radius:10px;padding:14px 16px}' +
    '</style></head><body>' +

    '<div style="font-size:11px;color:' + INK2 + '">제로베이스 취업정보회사</div>' +
    '<div style="font-size:20px;font-weight:800;margin:2px 0 14px">취업/이직 미니 진단 결과지</div>' +

    '<div style="background:' + NAVY + ';color:#fff;border-radius:12px;padding:18px 20px">' +
    '<div style="font-size:11px;opacity:.8">' + esc(d['이름']) + '님의 유형</div>' +
    '<div style="margin:8px 0">' + code + '</div>' +
    '<div style="font-size:18px;font-weight:800">' + esc(d['유형 이름']) + '</div>' +
    '<div style="font-size:12px;opacity:.85;margin-top:6px">종합 점수 ' + total + '점 (' + verdict(total) + ')  |  가장 먼저 볼 축: ' + esc(d['가장 약한 축']) + '</div>' +
    '</div>' +

    '<h2>기본 정보</h2><div class="box"><table style="width:100%;border-collapse:collapse">' + info + '</table></div>' +

    '<h2>취업/이직 공식 4개 축 점수</h2><div class="box"><table style="width:100%;border-collapse:collapse">' + axisRows + '</table>' +
    '<div style="font-size:10.5px;color:' + INK2 + ';margin-top:6px">판정 기준: 45점 미만 보완 필요, 45~74점 표준, 75점 이상 우수. 축별 56점 이상이면 유형 앞 글자.</div></div>' +

    '<h2>상담 포인트</h2><div class="box"><table style="width:100%;border-collapse:collapse">' +
    '<tr><td style="padding:6px 0;color:' + INK2 + ';width:110px">가장 큰 고민</td><td style="padding:6px 0;font-weight:600">' + esc(d['가장 큰 고민']) + '</td></tr>' +
    '<tr><td style="padding:6px 0;color:' + INK2 + '">추천 서비스</td><td style="padding:6px 0;font-weight:600">' + esc(d['추천 서비스']) + '</td></tr>' +
    '<tr><td style="padding:6px 0;color:' + INK2 + '">추천 프로그램</td><td style="padding:6px 0;font-weight:800;color:' + NAVY + '">' + esc(d['추천 프로그램']) + '</td></tr>' +
    '</table></div>' +

    '<h2>문항별 응답</h2><table style="width:100%;border-collapse:collapse;font-size:11.5px">' + qaRows + '</table>' +

    '<div style="margin-top:18px;font-size:10px;color:' + INK2 + '">12개 자가 응답 문항으로 산출한 참고용 진단입니다. 세부 해설은 예약하신 상담에서 안내합니다.</div>' +
    '</body></html>';
}

/* ---------- 확인용 ---------- */

// 진단 페이지가 링크의 연락처로 앞서 남긴 DB 응답을 불러올 때 사용
// ?phone=01012345678&name=홍길동 → DB RAW 탭에서 연락처와 이름이 모두 맞는 가장 최근 응답을 돌려준다
// 연락처만으로는 조회되지 않게 이름까지 맞아야 응답한다 (개인정보 보호)
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.phone) return json({ ok: true, message: '미니 진단 응답 수집 중' });
  try {
    return json(lookupDb(p.phone, p.name || ''));
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

const DB_SHEET_ID = '1e7dIDsSq_Ja4bpWOosAgnv_12SY2papfe1Z-IyqNS_4'; // [ZB] SD 2.0
const DB_TAB = 'DB RAW';
// DB RAW 열 이름 → 화면에 보여줄 이름 (이 순서대로 보여줌)
const DB_FIELDS = [
  ['current_status', '현재 상황'], ['career_duration', '경력 기간'], ['target_company_size', '목표하는 기업 규모'],
  ['target_industry', '목표하는 산업 분야'], ['education', '최종 학력'], ['age_range', '나이'], ['employment_timing', '희망 취업/이직 시기']
];

function normPhone(v) {
  let d = String(v || '').replace(/\D/g, '');
  if (d.indexOf('82') === 0) d = '0' + d.slice(2);
  if (/^1[016789]/.test(d)) d = '0' + d;
  return d;
}

function lookupDb(phone, name) {
  const want = normPhone(phone), wantName = String(name).trim();
  if (want.length < 10 || !wantName) return { ok: true, found: false };
  const values = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(DB_TAB).getDataRange().getValues();
  const h = values[0].map(String), col = function (k) { return h.indexOf(k); };
  const cPhone = col('phone'), cName = col('name'), cDb = col('survey_name_snapshot'), cAt = col('submitted_at'), cExtra = col('additional_info');
  for (let r = values.length - 1; r >= 1; r--) { // 아래쪽이 최신
    const row = values[r];
    if (normPhone(row[cPhone]) !== want || String(row[cName]).trim() !== wantName) continue;
    const fields = [];
    DB_FIELDS.forEach(function (f) { const i = col(f[0]); if (i > -1 && String(row[i]).trim()) fields.push([f[1], String(row[i]).trim()]); });
    // additional_info 에 JSON으로 들어온 DB별 추가 문항도 함께 보여줌
    if (cExtra > -1 && String(row[cExtra]).trim()) {
      try {
        const extra = JSON.parse(row[cExtra]);
        Object.keys(extra).forEach(function (k) { const v = extra[k]; if (v !== '' && v !== null) fields.push([k, Array.isArray(v) ? v.join(', ') : String(v)]); });
      } catch (e) {}
    }
    return { ok: true, found: true, db: String(row[cDb]).replace(/[^\p{L}\p{N}\s\/()·,.~-]/gu, '').trim(), name: String(row[cName]).trim(), submittedAt: String(row[cAt]), fields: fields };
  }
  return { ok: true, found: false };
}

// 편집기에서 한 번 실행해 시트와 드라이브 권한을 모두 승인받는 용도
function authorize() {
  SpreadsheetApp.openById(SHEET_ID).getName();
  SpreadsheetApp.openById(DB_SHEET_ID).getName();
  DriveApp.getRootFolder().getName();
  DriveApp.getFileById(SHEET_ID).getName();
  Logger.log('권한 승인 완료');
}

// 편집기에서 직접 실행해 시트에 테스트 한 줄과 PDF가 생기는지 확인용
function testWrite() {
  const sample = {
    '제출일시': new Date().toLocaleString('ko-KR'), '이름': '테스트', '연락처': '010-0000-0000',
    '구분': '이직', '연차': '4~7년차', '가장 큰 고민': '대기업으로 옮기고 싶어요', '추천 서비스': '대기업 이직 전략 컨설팅',
    '추천 프로그램': '프레스티지', '유형 코드': 'SOPR', '유형 이름': '기회를 넓게 보는 전략가', '종합 점수': 58,
    '업무 구조화': 67, '목표 구체화': 33, '취업/이직 전략': 56, '실전 준비도': 78, '가장 약한 축': '목표 구체화',
    '구조화 Q1': '예시 응답', '목표 Q1': '예시 응답', '전략 Q1': '예시 응답', '실전 Q1': '예시 응답'
  };
  const res = doPost({ postData: { contents: JSON.stringify(sample) } });
  Logger.log(res.getContent());
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
