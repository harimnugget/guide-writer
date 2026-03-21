/**
 * 영상 편집 가이드 작성기 — Google Apps Script
 *
 * [배포 방법]
 * 1. 구글 시트를 열고, 확장 프로그램 → Apps Script 클릭
 * 2. 이 코드를 전체 복사해서 Code.gs에 붙여넣기
 * 3. 상단 메뉴에서 배포 → 새 배포 클릭
 * 4. 유형 선택: 웹 앱
 *    - 설명: 가이드 작성기
 *    - 실행 계정: 나
 *    - 액세스 권한: 모든 사용자
 * 5. 배포 클릭 → 웹 앱 URL 복사
 * 6. guide-writer.html의 Apps Script URL 칸에 붙여넣기
 *
 * [시트 구조]
 * B열: #S (섹션명)
 * C열: 내용 (타임코드 포함)
 * 1행은 헤더로 사용 (#S | 내용)
 */

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const data = body.data; // [{ section, content }]

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    data.forEach(function(item) {
      // 기존에 같은 섹션이 있는지 확인 (B열 검색)
      const sectionCol = sheet.getRange('B:B').getValues();
      let existingRow = -1;

      for (var i = 0; i < sectionCol.length; i++) {
        if (sectionCol[i][0] === item.section) {
          existingRow = i + 1; // 1-indexed
          break;
        }
      }

      if (existingRow > 0) {
        // 기존 섹션에 내용 추가 (불릿 누적)
        var currentContent = sheet.getRange(existingRow, 3).getValue();
        if (currentContent) {
          currentContent = currentContent + '\n\n' + item.content;
        } else {
          currentContent = item.content;
        }
        sheet.getRange(existingRow, 3).setValue(currentContent);
      } else {
        // 새 섹션 추가 — 마지막 행 다음에 삽입
        var lastRow = sheet.getLastRow();
        sheet.getRange(lastRow + 1, 2).setValue(item.section);
        sheet.getRange(lastRow + 1, 3).setValue(item.content);
      }
    });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: '가이드 작성기 API가 정상 작동 중입니다.' }))
    .setMimeType(ContentService.MimeType.JSON);
}
