/**
 * PH Korea 2026 — VOD 동의 응답 수집 (구글시트 자동 저장)
 *
 * 설치:
 *  1) 구글 드라이브에서 새 구글 스프레드시트 1개 생성 (예: "PH2026 VOD 동의 응답")
 *  2) 상단 메뉴 [확장 프로그램] > [Apps Script] 클릭
 *  3) 기본 코드를 지우고 이 파일 내용 전체를 붙여넣기 → 저장(디스크 아이콘)
 *  4) 우측 상단 [배포] > [새 배포] > 유형 톱니 > "웹 앱" 선택
 *       - 설명: VOD consent
 *       - 다음 사용자로 실행: 나
 *       - 액세스 권한이 있는 사용자: "모든 사용자" (Anyone)
 *     [배포] → 권한 승인 → 표시되는 "웹 앱 URL"(끝이 /exec)을 복사
 *  5) VOD.html 파일을 열어 상단의  var SUBMIT_URL = "https://script.google.com/macros/s/AKfycbyiynHBzLDSK9cy2p5FfltblPiQzH5enPWt16PVzi6mT7IaBlDEVt4ogZrbNYQO-2zG/exec";  에 그 URL을 붙여넣고 저장 후 다시 업로드
 *
 * 응답은 이 스프레드시트의 "응답" 시트에 강의별로 한 줄씩 쌓입니다.
 */

var SHEET_NAME = '응답';
var HEADERS = ['제출시각','연자ID','성함','소속','국가','회신이메일','강의코드','세션','강의제목','VOD동의'];

function doGet(e){
  var cb = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : 'callback';
  var out;
  try {
    var data = JSON.parse(e.parameter.data);
    saveRows_(data);
    out = { ok: true };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(out) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doPost(e){
  try {
    var data = JSON.parse(e.postData.contents);
    saveRows_(data);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveRows_(data){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  var now = new Date();
  var lecs = data.lectures || [];
  for (var i = 0; i < lecs.length; i++) {
    var l = lecs[i];
    sh.appendRow([now, data.id, data.name, data.aff, data.country, data.email,
                  l.code, l.session, l.title, l.consent]);
  }
}
