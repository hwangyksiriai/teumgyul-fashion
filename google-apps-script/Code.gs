/**
 * 틈결 오프라인 방문(마이크로) 신청서 → 구글시트 저장
 *
 * 설치 방법
 *  1. 신청 내역을 받을 구글시트를 새로 만들기
 *  2. 시트에서 [확장 프로그램] → [Apps Script] 열기
 *  3. 기본 Code.gs 내용을 지우고 이 파일 내용 전체 붙여넣기 → 저장
 *  4. [배포] → [새 배포] → 유형: 웹 앱
 *       - 실행 계정: 나
 *       - 액세스 권한: 모든 사용자
 *  5. 권한 승인 후 발급된 웹 앱 URL(…/exec)을 가이드 페이지의 GAS_WEB_APP_URL에 넣기
 *
 * 코드를 수정한 뒤에는 [배포] → [배포 관리] → 수정(연필) → 버전: 새 버전 으로 다시 배포해야 반영됩니다.
 */

const SHEET_NAME = '시트1';
const TIMEZONE = 'Asia/Seoul';

const COLUMNS = [
  { key: 'submittedAt',  label: '제출일시',          width: 150 },
  { key: 'name',         label: '이름',              width: 100 },
  { key: 'instagram',    label: '인스타그램',        width: 230 },
  { key: 'phone',        label: '휴대폰',            width: 130, text: true },
  { key: 'email',        label: '이메일',            width: 190 },
  { key: 'agreeVisit',   label: '12/4–5 방문 가능',  width: 120 },
  { key: 'agreeUpload',  label: '12/6–7 업로드',   width: 120 },
  { key: 'agreeSeongsu', label: '성수 방문 가능',    width: 110 },
  { key: 'agreeNoShow',  label: '노쇼 불가 동의',    width: 110 },
  { key: 'price',        label: '고료',              width: 80 }   // 5만원 / 10만원 / 15만원 페이지 구분
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    data.submittedAt = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    appendApplication_(data);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 브라우저로 웹 앱 URL을 열었을 때 연결 확인용
function doGet() {
  return json_({ ok: true, sheet: SHEET_NAME });
}

function appendApplication_(data) {
  const sheet = getSheet_();
  ensureHeader_(sheet);
  const row = sheet.getLastRow() + 1;
  const values = COLUMNS.map(c => String(data[c.key] == null ? '' : data[c.key]).trim());
  COLUMNS.forEach((c, i) => { if (c.text) sheet.getRange(row, i + 1).setNumberFormat('@'); });
  sheet.getRange(row, 1, 1, COLUMNS.length).setValues([values]).setVerticalAlignment('middle').setWrap(true);
}

function ensureHeader_(sheet) {
  if (sheet.getLastRow() > 0) {
    // 기존 시트: 새로 추가된 열 제목(예: 고료)만 채워 넣기
    const head = sheet.getRange(1, 1, 1, COLUMNS.length);
    const cur = head.getValues()[0];
    if (cur.some((v, i) => v !== COLUMNS[i].label)) {
      head.setValues([COLUMNS.map(c => c.label)])
        .setBackground('#B8489A').setFontColor('#FFFFFF').setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
      COLUMNS.forEach((c, i) => sheet.setColumnWidth(i + 1, c.width));
    }
    return;
  }
  sheet.getRange(1, 1, 1, COLUMNS.length)
    .setValues([COLUMNS.map(c => c.label)])
    .setBackground('#B8489A').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 34);
  COLUMNS.forEach((c, i) => sheet.setColumnWidth(i + 1, c.width));
}

function getSheet_() {
  const ss = SpreadsheetApp.openById('1POVacWQKlPuGoz3NnpE9bWfUBpltbJyniMeAJojParI');
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
