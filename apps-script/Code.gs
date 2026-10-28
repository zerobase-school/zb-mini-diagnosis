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
  const html = detail ? richHtml(d, detail) : reportHtml(d);
  const name = '미니진단_' + (d['이름'] || '이름없음') + '_' + String(d['연락처'] || '').replace(/\D/g, '').slice(-4) + '_' +
    Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmm') + '.pdf';
  const blob = Utilities.newBlob(html, 'text/html', 'report.html').getAs('application/pdf').setName(name);
  const file = getFolder().createFile(blob);
  try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return file.getUrl();
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

/* 상세 결과지 (세일즈팀 상담 준비용) */
function richHtml(d, r) {
  const NAVY = '#1d2483', NAVY2 = '#3a41c4', LAV = '#eeeefb', INK = '#1a1d2e', INK2 = '#5a5f73', INK3 = '#8a8fa3', LINE = '#e3e4ee', ROSE = '#d23f6a';
  const cell = 'padding:6px 0;vertical-align:top;';
  const bar = function (score) {
    const s = Math.max(0, Math.min(100, Number(score) || 0)), color = s < 45 ? ROSE : NAVY;
    return '<table style="width:100%;border-collapse:collapse"><tr>' +
      (s > 0 ? '<td style="width:' + s + '%;height:9px;background:' + color + '"></td>' : '') +
      (s < 100 ? '<td style="height:9px;background:' + LAV + '"></td>' : '') + '</tr></table>';
  };
  const h2 = function (num, t) { return '<div style="margin:22px 0 8px;padding-bottom:5px;border-bottom:2px solid ' + NAVY + ';font-size:14px;font-weight:800;color:' + NAVY + '">' + num + '. ' + t + '</div>'; };
  const kv = function (k, v, strong) { return v ? '<tr><td style="' + cell + 'color:' + INK2 + ';width:120px">' + k + '</td><td style="' + cell + (strong ? 'font-weight:800;color:' + NAVY : 'font-weight:600') + '">' + v + '</td></tr>' : ''; };
  const list = function (arr) { return '<ul style="margin:4px 0 0 16px;padding:0">' + arr.map(function (x) { return '<li style="margin:2px 0">' + esc(x) + '</li>'; }).join('') + '</ul>'; };
  const code = String(r.code || '').split('').map(function (ch) {
    return '<span style="display:inline-block;width:34px;height:40px;line-height:40px;margin-right:5px;text-align:center;border:1px solid rgba(255,255,255,.4);font-size:22px;font-weight:800">' + esc(ch) + '</span>';
  }).join('');

  // 1. 요약
  const summary = '<table style="width:100%;border-collapse:collapse">' +
    kv('가장 큰 고민', r.concern ? esc(r.concern.title) : '') +
    kv('추천 서비스', r.concern ? esc(r.concern.service) + '<div style="font-weight:400;color:' + INK2 + ';font-size:11px">' + esc(r.concern.desc) + '</div>' : '') +
    kv('추천 프로그램', esc(r.plan.name) + ' (월 ' + esc(r.plan.monthly) + ', 총 ' + esc(r.plan.total) + ', VAT 별도)', true) +
    kv('추천 이유', esc(r.plan.why)) +
    kv('가장 먼저 볼 축', esc(r.weakest)) +
    kv('강점', esc((r.strengths || []).join(', ') || '없음')) +
    kv('보완점', esc((r.gaps || []).join(', ') || '없음')) +
    '</table>';

  // 2. 축별 점수
  const axisTable = '<table style="width:100%;border-collapse:collapse">' + r.axes.map(function (a) {
    return '<tr><td style="padding:7px 0;width:95px;font-weight:700">' + esc(a.name) + '</td>' +
      '<td style="padding:7px 10px">' + bar(a.score) + '</td>' +
      '<td style="padding:7px 0;width:46px;text-align:right;font-weight:800">' + a.score + '점</td>' +
      '<td style="padding:7px 0 7px 10px;width:150px;color:' + INK2 + ';font-size:11px">' + esc(a.side) + ', ' + esc(a.strength) + '</td></tr>';
  }).join('') + '</table><div style="font-size:10px;color:' + INK3 + ';margin-top:4px">점수가 낮은 축부터 정렬. 45점 미만 보완 필요, 45~74점 표준, 75점 이상 우수. 축별 56점 이상이면 유형 앞 글자.</div>';

  // 3. 축별 상세
  const axisDetail = r.axes.map(function (a, idx) {
    const qa = '<table style="width:100%;border-collapse:collapse;font-size:11px;margin-top:6px">' + a.qa.map(function (x, i) {
      return '<tr><td style="padding:5px 6px;border-top:1px solid ' + LINE + ';color:' + INK2 + ';width:22px">Q' + (i + 1) + '</td>' +
        '<td style="padding:5px 6px;border-top:1px solid ' + LINE + ';color:' + INK2 + '">' + esc(x.q) + '</td>' +
        '<td style="padding:5px 6px;border-top:1px solid ' + LINE + ';font-weight:700;width:170px">' + esc(x.a) + '</td>' +
        '<td style="padding:5px 6px;border-top:1px solid ' + LINE + ';text-align:right;width:30px;color:' + (x.s !== null && x.s <= 1 ? ROSE : INK2) + '">' + (x.s === null ? '-' : x.s + '/3') + '</td></tr>';
    }).join('') + '</table>';
    return '<div style="border:1px solid ' + LINE + ';padding:12px 14px;margin-bottom:10px;page-break-inside:avoid">' +
      '<table style="width:100%;border-collapse:collapse"><tr><td style="font-size:13px;font-weight:800">' + (idx + 1) + ') ' + esc(a.name) +
      ' <span style="font-size:11px;font-weight:600;color:' + INK2 + '">' + esc(a.side) + '</span></td>' +
      '<td style="text-align:right;font-size:15px;font-weight:800;color:' + (a.score < 45 ? ROSE : NAVY) + '">' + a.score + '점 <span style="font-size:11px;color:' + INK2 + '">' + esc(a.verdict) + '</span></td></tr></table>' +
      '<div style="margin-top:6px"><b style="color:' + NAVY2 + '">눈에 띄는 점</b> ' + esc(a.flag) + '</div>' +
      '<div style="margin-top:4px"><b style="color:' + NAVY2 + '">함께 볼 부분</b> ' + esc(a.point) + '</div>' +
      '<div style="margin-top:4px"><b style="color:' + NAVY2 + '">상담에서 다룰 내용</b>' + list(a.locked || []) + '</div>' + qa + '</div>';
  }).join('');

  // 4. 상담 진행 가이드
  const w = r.axes[0];
  const lowQs = [];
  r.axes.forEach(function (a) { a.qa.forEach(function (x) { if (x.s !== null && x.s <= 1) lowQs.push(a.name + ': "' + x.q + '" → ' + x.a); }); });
  const guide = list([
    '오프닝: 결과 유형 "' + r.code + ' ' + r.typeName + '"과 종합 ' + r.total + '점(' + r.verdict + ')을 먼저 함께 보며 시작',
    r.concern ? '고민 확인: "' + r.concern.title + '" 고민의 구체적인 상황과 지금까지 시도한 방법 질문' : '고민 확인: 지금 가장 막히는 지점 질문',
    '핵심 진단: ' + w.name + '(' + w.score + '점)이 가장 낮음. ' + w.flag,
    '함께 확인할 문항 (1점 이하 응답): ' + (lowQs.length ? lowQs.length + '개' : '없음'),
    '해결 방향: ' + (w.locked || []).join(', '),
    '프로그램 제안: ' + r.plan.name + ' (월 ' + r.plan.monthly + ', 무이자 12개월). ' + r.plan.why,
    '참고 합격 사례: ' + (r.plan.cases || []).join(', ')
  ]) + (lowQs.length ? '<div style="margin-top:8px;font-size:11px;color:' + INK2 + '"><b>1점 이하 응답 문항</b>' + list(lowQs) + '</div>' : '');

  // 5. 앞서 남긴 DB 답변
  const dbTable = (r.dbAnswers && r.dbAnswers.length) ? h2(5, '앞서 남긴 DB 답변' + (r.db ? ' (' + esc(r.db) + ')' : '')) +
    '<table style="width:100%;border-collapse:collapse;font-size:11.5px">' + r.dbAnswers.map(function (x) {
      return '<tr><td style="padding:5px 0;border-bottom:1px solid ' + LINE + ';color:' + INK2 + ';width:160px">' + esc(x.k) + '</td><td style="padding:5px 0;border-bottom:1px solid ' + LINE + ';font-weight:600">' + esc(x.v) + '</td></tr>';
    }).join('') + '</table>' : '';

  return '<!doctype html><html><head><meta charset="utf-8"><style>' +
    'body{font-family:"Noto Sans KR","Malgun Gothic",sans-serif;color:' + INK + ';font-size:12px;margin:26px 32px;line-height:1.55}' +
    '</style></head><body>' +
    '<table style="width:100%;border-collapse:collapse"><tr><td><div style="font-size:11px;color:' + INK2 + '">제로베이스 취업정보회사 | 세일즈팀 상담 준비용</div>' +
    '<div style="font-size:20px;font-weight:800;margin-top:2px">취업/이직 미니 진단 결과지</div></td>' +
    '<td style="text-align:right;font-size:11px;color:' + INK2 + '">' + esc(d['제출일시']) + '</td></tr></table>' +

    '<div style="background:' + NAVY + ';color:#fff;padding:16px 18px;margin-top:12px">' +
    '<table style="width:100%;border-collapse:collapse;color:#fff"><tr><td>' +
    '<div style="font-size:11px;opacity:.8">' + esc(d['이름']) + '님의 유형' + (d['연락처'] ? ' | ' + esc(d['연락처']) : '') + '</div>' +
    '<div style="margin:7px 0">' + code + '</div><div style="font-size:17px;font-weight:800">' + esc(r.typeName) + '</div></td>' +
    '<td style="text-align:right;vertical-align:bottom"><div style="font-size:11px;opacity:.8">종합 점수</div><div style="font-size:30px;font-weight:800">' + r.total + '<span style="font-size:13px">점</span></div><div style="font-size:11px">' + esc(r.verdict) + '</div></td></tr></table>' +
    '<div style="font-size:11px;opacity:.85;margin-top:6px">' + esc([r.db, r.track, r.stage].filter(Boolean).join(' | ')) + '</div></div>' +

    h2(1, '한눈에 보기') + summary +
    h2(2, '취업/이직 공식 4개 축 점수') + axisTable +
    h2(3, '상담 진행 가이드') + guide +
    h2(4, '축별 상세 진단') + axisDetail +
    dbTable +
    '<div style="margin-top:16px;font-size:10px;color:' + INK3 + '">12개 자가 응답 문항으로 산출한 참고용 진단입니다. 고객 화면에서는 유형까지만 공개되고, 이 문서의 상세 내용은 상담에서 안내합니다.</div>' +
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
