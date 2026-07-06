/**
 * 인터넷 가입 랜딩페이지 → 구글시트 연동용 Apps Script
 *
 * [설치 방법]
 * 1. 리드를 저장할 구글시트를 새로 만듭니다.
 * 2. 상단 메뉴 [확장 프로그램] → [Apps Script] 클릭
 * 3. 기본 코드(Code.gs)를 지우고 이 파일 내용을 전체 붙여넣기
 * 4. 상단 [배포] → [새 배포] 클릭
 *    - 유형: 웹 앱(Web app)
 *    - 실행 계정: 나(Me)
 *    - 액세스 권한: 모든 사용자(Anyone)
 * 5. 배포 후 나오는 웹 앱 URL을 복사
 * 6. index.html 안의 ENDPOINT_URL 값을 이 URL로 교체
 *
 * [주의]
 * - 액세스 권한을 "모든 사용자"로 열어야 외부 랜딩페이지에서 접수가 가능합니다.
 * - 이 URL은 누구나 데이터를 넣을 수 있는 공개 엔드포인트이므로,
 *   운영 단계에서는 스팸 방지를 위해 아래 SECRET_KEY 값을 랜덤 문자열로 바꾸고
 *   프론트엔드 payload에도 동일한 값을 추가하는 것을 권장합니다.
 */

const SHEET_NAME = "carrot";      // "carrot table" 스프레드시트의 실제 탭 이름에 맞춤
const SECRET_KEY = "wpdlTl316719";   // index.html의 SECRET_KEY와 동일해야 함
const SPREADSHEET_ID = "17a7UXxIta9opdeTYQU-2Fi8f6vzDgNEGalqFKEGwhAg"; // "carrot table" 시트 ID (URL의 /d/와 /edit 사이 문자열)
const RETENTION_DAYS = 7;         // 접수 후 이 일수가 지나면 자동 삭제

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 선택적 보안키 검증 (SECRET_KEY를 채운 경우에만 동작)
    if (SECRET_KEY && data.secret !== SECRET_KEY) {
      return ContentService
        .createTextOutput(JSON.stringify({ result: "error", message: "invalid secret" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }
    // 탭이 이미 있어도 데이터(헤더 포함)가 하나도 없으면 헤더 행을 추가합니다.
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["이름", "휴대전화", "거주 지역", "희망 상품", "접수일시"]);
    }

    sheet.appendRow([
      data.name || "",
      data.phone || "",
      data.region || "",
      data.plan || "",
      new Date()
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 배포 URL이 브라우저에서 정상 접속되는지 확인용 (GET 요청)
function doGet(e) {
  return ContentService
    .createTextOutput("OK - 이 URL은 POST 전용 웹훅입니다.")
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * [자동 삭제] RETENTION_DAYS(현재 7일)보다 오래된 데이터를 매일 삭제합니다.
 * 이 함수 자체는 자동으로 실행되지 않으며, 아래 installDailyCleanupTrigger()를
 * 딱 한 번 수동으로 실행해서 "매일 실행되는 트리거"를 등록해야 작동합니다.
 */
function cleanupOldLeads() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // 헤더만 있고 데이터 없음

  const dateColIndex = 5; // E열 = 접수일시
  const now = new Date();
  const cutoffMs = RETENTION_DAYS * 24 * 60 * 60 * 1000;

  // 아래에서 위로 훑어야 행을 지워도 인덱스가 안 꼬입니다.
  for (let row = lastRow; row >= 2; row--) {
    const cellValue = sheet.getRange(row, dateColIndex).getValue();
    if (!cellValue) continue; // 접수일시가 비어있으면 건너뜀 (수동 입력 데이터 등 보호)

    const recordedDate = new Date(cellValue);
    if (isNaN(recordedDate.getTime())) continue;

    if (now.getTime() - recordedDate.getTime() > cutoffMs) {
      sheet.deleteRow(row);
    }
  }
}

/**
 * [최초 1회 실행용] 위 cleanupOldLeads()를 매일 자동으로 실행하는 트리거를 등록합니다.
 * Apps Script 편집기에서 이 함수를 딱 한 번 "실행" 버튼으로 실행하면 됩니다.
 * (실행 시 권한 승인 창이 뜨면 이전처럼 승인해주세요)
 * 이미 등록된 동일한 트리거가 있으면 중복 등록되지 않도록 먼저 정리합니다.
 */
function installDailyCleanupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "cleanupOldLeads") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("cleanupOldLeads")
    .timeBased()
    .everyDays(1)
    .atHour(3) // 매일 새벽 3시경 실행
    .create();
}
