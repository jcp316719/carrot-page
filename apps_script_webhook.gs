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

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 선택적 보안키 검증 (SECRET_KEY를 채운 경우에만 동작)
    if (SECRET_KEY && data.secret !== SECRET_KEY) {
      return ContentService
        .createTextOutput(JSON.stringify({ result: "error", message: "invalid secret" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }
    // 탭이 이미 있어도 데이터(헤더 포함)가 하나도 없으면 헤더 행을 추가합니다.
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["이름", "휴대전화", "거주 지역", "희망 상품"]);
    }

    sheet.appendRow([
      data.name || "",
      data.phone || "",
      data.region || "",
      data.plan || ""
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
