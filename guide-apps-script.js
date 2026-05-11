/**
 * 영상 편집 가이드 작성기 v2 — Google Apps Script
 *
 * 마스터 시트 1개 + 새 탭 자동 생성 방식.
 * 전송 시 "YYYY-MM-DD_프로젝트명" 탭에 데이터가 쌓이고,
 * 같은 탭이 이미 있으면 거기에 누적됩니다.
 *
 * 추가로 매 전송마다 지정된 백업 폴더에 JSON 파일도 함께 떨궈둡니다.
 * 백업 파일은 index.html의 "백업 가져오기" 기능으로 다시 불러올 수 있어요.
 *
 * ─────────────────────────────────────────────
 * [최초 설치 순서 — 한 번만]
 *
 * 1) 데이터를 받을 구글 스프레드시트를 새로 만든다 (이름 자유).
 * 2) URL의 /d/...../edit 사이 문자열이 시트 ID. 메모해둔다.
 * 3) 백업 파일이 쌓일 구글 드라이브 폴더를 새로 만든다 (이름 자유, 예: "영상가이드_백업").
 *    그 폴더에 들어가면 URL이 .../folders/<폴더_ID>. <폴더_ID>를 메모해둔다.
 * 4) 위 1)의 시트에서 확장 프로그램 → Apps Script 클릭.
 * 5) 기본 Code.gs 내용을 모두 지우고, 이 파일을 통째로 붙여넣는다.
 * 6) 아래 MASTER_SHEET_ID, BACKUP_FOLDER_ID 두 상수를 채우고 Cmd+S로 저장.
 *    백업 폴더를 안 쓰려면 BACKUP_FOLDER_ID를 ''로 두면 됨 (백업 생략).
 * 7) (검증) 함수 드롭다운에서 setupTest 선택 → ▶ 실행.
 *    첫 실행 시 권한 동의(고급 → 신뢰할 수 없음으로 이동 → 허용).
 *    실행 로그에 "✓ 시트 연결 성공"과 "✓ 백업 폴더 연결 성공"이 떠야 한다.
 * 8) 시트로 돌아가 [2099-01-01_연결테스트] 탭이 생기고, 백업 폴더에 테스트
 *    JSON 파일이 떨어졌는지 확인. 확인 후 그 탭/파일은 삭제해도 무방.
 * 9) 배포 → 새 배포 → 유형: 웹 앱 / 실행: 나 / 권한: 모든 사용자.
 *    웹 앱 URL을 복사해 index.html의 "Apps Script URL" 칸에 붙여넣는다.
 *    (기존 배포가 있으면 배포 관리 → 연필 → 새 버전 → 배포 업데이트.
 *     이러면 URL이 유지된다.)
 *
 * ─────────────────────────────────────────────
 * [시트 구조]
 * 각 탭 1행은 헤더, 2행부터 데이터.
 *   B열: #S (섹션명)
 *   C열: 내용 (불릿 누적, 줄바꿈 두 번으로 구분)
 *
 * [백업 JSON 형식]
 *   {
 *     version: 'v2',
 *     savedAt: ISO 문자열,
 *     projectName, tabName,
 *     data:    시트 전송 페이로드 그대로 ({section, content}[])
 *     entries: 클라이언트 미리보기 복원용 ({section, tc, content}[])  ← index.html이 함께 보내줌
 *   }
 *
 * [다중 컴퓨터/계정 공유]
 * - 웹앱 권한이 "모든 사용자"라 누구나 URL만 알면 같은 시트에 누적 가능.
 * - 시트 자체를 다른 계정에서 보려면 시트 공유 권한을 별도로 설정해야 함.
 * - 백업 폴더도 마찬가지 — 협업자에게 보기 권한을 주면 같은 백업을 볼 수 있음.
 */

// ⚠️ 여기에 마스터 시트 ID를 붙여넣으세요.
const MASTER_SHEET_ID = '여기에_시트_ID를_붙여넣으세요';

// ⚠️ 여기에 백업 폴더 ID를 붙여넣으세요. (백업을 안 쓸 거면 '' 빈 문자열)
const BACKUP_FOLDER_ID = '';

const TAB_HEADERS = ['', '#S', '내용'];
const CONTENT_DELIMITER = '\n\n';
const TAB_NAME_PATTERN = /^\d{4}-\d{2}-\d{2}_.+$/;


function doPost(e) {
  try {
    Logger.log('[doPost] body: ' + (e && e.postData ? e.postData.contents : '(no body)'));

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('요청 본문이 비어있습니다.');
    }

    const body = JSON.parse(e.postData.contents);
    const projectName = (body.projectName || '').toString().trim();
    const tabName = (body.tabName || '').toString().trim();
    const data = Array.isArray(body.data) ? body.data : [];

    if (!tabName) {
      throw new Error('tabName이 비어있습니다.');
    }
    if (!TAB_NAME_PATTERN.test(tabName)) {
      throw new Error('tabName 형식이 올바르지 않습니다: ' + tabName + ' (기대: YYYY-MM-DD_프로젝트명)');
    }
    if (data.length === 0) {
      throw new Error('전송할 항목(data)이 없습니다.');
    }
    if (!MASTER_SHEET_ID || MASTER_SHEET_ID.indexOf('여기에') === 0) {
      throw new Error('MASTER_SHEET_ID 상수가 설정되지 않았습니다. guide-apps-script.js 상단을 확인하세요.');
    }

    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet = findOrCreateTab_(ss, tabName);
    const count = appendToTab_(sheet, data);

    const backupResult = writeBackupFile_(body, count);

    const message = count + '개 항목을 [' + tabName + '] 탭에 기록했습니다.';
    Logger.log('[doPost] ✓ ' + message + ' (projectName=' + projectName + ')');

    return jsonResponse_({
      success: true,
      sheetUrl: ss.getUrl(),
      tabName: tabName,
      itemCount: count,
      backupFileName: backupResult.fileName,
      backupSkipped: backupResult.skipped,
      message: message
    });

  } catch (err) {
    Logger.log('[doPost] ✗ error: ' + err.message);
    if (err.stack) Logger.log(err.stack);
    return jsonResponse_({
      success: false,
      error: err.message,
      stack: err.stack || ''
    });
  }
}


function doGet(e) {
  try {
    const qs = (e && e.parameter) || {};

    if (qs.ping === '1') {
      Logger.log('[ping] ' + JSON.stringify(qs));
    }

    if (!MASTER_SHEET_ID || MASTER_SHEET_ID.indexOf('여기에') === 0) {
      return jsonResponse_({
        status: 'error',
        message: 'MASTER_SHEET_ID 상수가 설정되지 않았습니다.'
      });
    }

    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    return jsonResponse_({
      status: 'ok',
      message: '가이드 작성기 API가 정상 작동 중입니다.',
      sheetName: ss.getName(),
      sheetUrl: ss.getUrl()
    });
  } catch (err) {
    return jsonResponse_({
      status: 'error',
      message: err.message
    });
  }
}


/**
 * 탭을 찾거나 새로 만든다. 새 탭이면 헤더·고정행·컬럼 폭·정렬을 세팅.
 * - 모든 셀에 "가운데 맞춤"(수직 정렬 middle) 적용
 * - 줄바꿈 자동 wrap 활성화
 */
function findOrCreateTab_(ss, tabName) {
  let sheet = ss.getSheetByName(tabName);
  if (sheet) return sheet;

  sheet = ss.insertSheet(tabName);
  sheet.getRange(1, 1, 1, TAB_HEADERS.length).setValues([TAB_HEADERS]);
  sheet.getRange(1, 2, 1, 2)
    .setFontWeight('bold')
    .setBackground('#f0f0f0');
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 600);
  sheet.setFrozenRows(1);

  // 컬럼 A~C 전체에 미리 수직 가운데 맞춤 + 자동 줄바꿈 적용.
  // 컬럼 단위로 지정해두면 이후 들어오는 데이터에도 자동 상속됨.
  const allCols = sheet.getRange(1, 1, sheet.getMaxRows(), 3);
  allCols.setVerticalAlignment('middle');
  allCols.setWrap(true);

  Logger.log('[findOrCreateTab_] 새 탭 생성: ' + tabName);
  return sheet;
}


/**
 * 항목 배열을 시트에 누적 기록한다.
 * 같은 섹션이 이미 있으면 C열에 줄바꿈으로 누적, 없으면 새 행에 추가.
 * B열은 1회만 일괄 읽어 캐시한 뒤 사용 (range 호출 최소화).
 *
 * @returns {number} 기록한 항목 수
 */
function appendToTab_(sheet, items) {
  const lastRow = sheet.getLastRow();
  let sections;
  if (lastRow >= 2) {
    sections = sheet.getRange(2, 2, lastRow - 1, 1)
      .getValues()
      .map(function (row) { return row[0]; });
  } else {
    sections = [];
  }

  let writtenCount = 0;

  items.forEach(function (item) {
    if (!item || typeof item !== 'object') return;
    const section = (item.section || '').toString();
    const content = (item.content || '').toString();
    if (!section || !content) return;

    const idx = sections.indexOf(section);
    if (idx >= 0) {
      const rowNum = idx + 2;
      const current = sheet.getRange(rowNum, 3).getValue();
      const next = current
        ? current + CONTENT_DELIMITER + content
        : content;
      sheet.getRange(rowNum, 3).setValue(next);
    } else {
      const nextRow = sheet.getLastRow() + 1;
      sheet.getRange(nextRow, 2, 1, 2).setValues([[section, content]]);
      sections.push(section);
    }
    writtenCount++;
  });

  return writtenCount;
}


/**
 * 백업 폴더에 JSON 파일을 떨군다.
 * BACKUP_FOLDER_ID가 비어있으면 조용히 스킵.
 * 파일명은 "탭이름__HHmmss.json" — 같은 날·같은 프로젝트의 여러 전송이 충돌하지 않도록 시각 포함.
 *
 * @returns {{ fileName: string|null, skipped: boolean }}
 */
function writeBackupFile_(body, itemCount) {
  if (!BACKUP_FOLDER_ID) {
    Logger.log('[backup] BACKUP_FOLDER_ID 비어있음 — 스킵');
    return { fileName: null, skipped: true };
  }
  try {
    const folder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
    const tabName = body.tabName;
    const stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'HHmmss');
    const fileName = tabName + '__' + stamp + '.json';

    const payload = {
      version: 'v2',
      savedAt: new Date().toISOString(),
      projectName: body.projectName || '',
      tabName: tabName,
      itemCount: itemCount,
      data: Array.isArray(body.data) ? body.data : [],
      entries: Array.isArray(body.entries) ? body.entries : []
    };

    folder.createFile(
      fileName,
      JSON.stringify(payload, null, 2),
      'application/json'
    );
    Logger.log('[backup] ✓ ' + fileName);
    return { fileName: fileName, skipped: false };
  } catch (err) {
    Logger.log('[backup] ✗ 실패: ' + err.message);
    return { fileName: null, skipped: false };
  }
}


/**
 * 사용자가 Apps Script 편집기에서 직접 실행해 연결을 검증하는 함수.
 * 권한 동의 + MASTER_SHEET_ID 유효성 + 더미 탭 작성 + 백업 폴더 검증까지 한 번에 점검.
 */
function setupTest() {
  try {
    if (!MASTER_SHEET_ID || MASTER_SHEET_ID.indexOf('여기에') === 0) {
      Logger.log('✗ MASTER_SHEET_ID 상수가 설정되지 않았습니다.');
      Logger.log('  guide-apps-script.js 상단의 상수에 시트 ID를 채워주세요.');
      return;
    }

    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    Logger.log('✓ 시트 연결 성공: ' + ss.getName());
    Logger.log('✓ 시트 URL: ' + ss.getUrl());

    const testTabName = '2099-01-01_연결테스트';
    const sheet = findOrCreateTab_(ss, testTabName);
    appendToTab_(sheet, [
      { section: '#테스트', content: '- 00:00 setupTest 더미 행 (확인 후 탭 삭제해도 됨)' }
    ]);
    Logger.log('✓ 더미 행 작성 완료. 시트에서 [' + testTabName + '] 탭을 확인하세요.');

    // 백업 폴더 검증
    if (!BACKUP_FOLDER_ID) {
      Logger.log('· 백업 폴더 설정 안 됨 (BACKUP_FOLDER_ID 빈 문자열) — 백업 비활성화 상태');
    } else {
      try {
        const folder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
        Logger.log('✓ 백업 폴더 연결 성공: ' + folder.getName());
        Logger.log('✓ 백업 폴더 URL: ' + folder.getUrl());
        const testFile = folder.createFile(
          'setupTest_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'HHmmss') + '.json',
          JSON.stringify({ test: true, savedAt: new Date().toISOString() }, null, 2),
          'application/json'
        );
        Logger.log('✓ 백업 폴더에 테스트 파일 작성 완료: ' + testFile.getName());
        Logger.log('  (확인 후 폴더에서 삭제해도 됨)');
      } catch (e) {
        Logger.log('✗ 백업 폴더 검증 실패: ' + e.message);
        Logger.log('  BACKUP_FOLDER_ID 값이 정확한지 확인하세요.');
      }
    }
  } catch (err) {
    Logger.log('✗ 실패: ' + err.message);
    if (err.stack) Logger.log(err.stack);
  }
}


function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * 마스터 시트의 첫 번째 시트에 사용 안내를 채워넣는 함수.
 * Apps Script 편집기에서 한 번만 실행하면 됨.
 * 기존 첫 시트의 내용을 덮어쓰니, 데이터가 있다면 다른 탭으로 옮긴 뒤 실행하세요.
 */
function setupReadme() {
  try {
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);

    // "📖_사용법" 탭이 이미 있으면 그걸 쓰고, 없으면 첫 번째 시트를 사용해서 이름 변경
    let sheet = ss.getSheetByName('📖_사용법');
    if (!sheet) {
      sheet = ss.getSheets()[0];
      sheet.setName('📖_사용법');
    }

    sheet.clear();

    // 안내 줄. 각 행은 [텍스트, 스타일태그] 형태.
    // 스타일태그: 'title'/'heading'/'warn'/'body'/'sub'
    const rows = [
      ['📖 영상 편집 가이드 작성기 — 사용 안내', 'title'],
      ['', 'body'],
      ['이 시트는 영상 편집 가이드 작성기의 데이터 저장소입니다.', 'body'],
      ['index.html에서 입력한 가이드가 자동으로 새 탭에 누적됩니다.', 'body'],
      ['', 'body'],
      ['── 시트는 어떻게 쌓이나요? ──', 'heading'],
      ['• 화면에서 "마스터 시트로 전송"을 누를 때마다 새 탭이 자동으로 만들어집니다.', 'body'],
      ['• 탭 이름 형식: YYYY-MM-DD_프로젝트명  (예: 2026-05-11_현대위아50주년)', 'body'],
      ['• 같은 이름의 탭이 이미 있으면 거기에 누적됩니다.', 'body'],
      ['• 각 탭의 B열 = 섹션명(#S), C열 = 내용(타임코드 포함).', 'body'],
      ['• 새 탭은 자동으로 "가운데 맞춤" + "자동 줄바꿈" 적용된 상태로 만들어집니다.', 'sub'],
      ['', 'body'],
      ['── 백업은 어떻게 되나요? ──', 'heading'],
      ['• 전송할 때마다 브라우저의 다운로드 폴더에 JSON 파일이 자동으로 떨어집니다.', 'body'],
      ['• 파일 이름 형식: 2026-05-11_프로젝트명__HHMMSS.json', 'body'],
      ['• 이 시트가 손상되거나 삭제돼도 JSON 파일로 복원 가능합니다.', 'body'],
      ['• 다운로드 폴더가 가득 차지 않도록 주기적으로 본인의 백업 폴더로 옮겨주세요.', 'sub'],
      ['', 'body'],
      ['── 백업으로 복원하려면? ──', 'heading'],
      ['1. index.html 열기.', 'body'],
      ['2. 액션 바의 "📥 백업 가져오기" 버튼 클릭.', 'body'],
      ['3. 복원할 JSON 파일 선택.', 'body'],
      ['4. 미리보기·프로젝트명·섹션 칩이 자동으로 복원됩니다.', 'body'],
      ['', 'body'],
      ['── 공동 작업 ──', 'heading'],
      ['• 이 시트의 모든 탭은 여러 컴퓨터·여러 계정에서 동시에 누적·확인 가능합니다.', 'body'],
      ['• 같은 Apps Script URL을 쓰면 모두 같은 시트의 같은 위치에 쌓입니다.', 'body'],
      ['• 시트 접근 권한이 있는 사람은 누구나 모든 탭을 열람할 수 있어요.', 'body'],
      ['', 'body'],
      ['⚠ 편집자에게 공유할 때 — 매우 중요', 'warn'],
      ['이 마스터 시트를 편집자에게 그대로 공유하지 마세요.', 'body'],
      ['', 'body'],
      ['이유: 마스터 시트에는 모든 프로젝트의 가이드가 누적되어 있고, 본인의', 'body'],
      ['작업 아카이브 역할을 합니다. 그대로 공유하면 다른 프로젝트 정보까지', 'body'],
      ['편집자가 보게 됩니다.', 'body'],
      ['', 'body'],
      ['올바른 방법:', 'body'],
      ['1. 새 구글 스프레드시트를 별도로 만든다 (예: "현대위아50주년_편집자용").', 'body'],
      ['2. 이 마스터 시트에서 해당 프로젝트 탭을 연다.', 'body'],
      ['3. 전체 선택(Cmd+A) → 복사(Cmd+C).', 'body'],
      ['4. 새로 만든 편집자용 시트에 붙여넣기(Cmd+V).', 'body'],
      ['5. 그 새 시트만 편집자에게 공유한다.', 'body'],
      ['', 'body'],
      ['이렇게 하면 마스터 시트는 본인 아카이브로 그대로 두고,', 'body'],
      ['편집자에게는 해당 프로젝트만 깔끔하게 전달됩니다.', 'body'],
      ['', 'body'],
      ['── 도구 위치 ──', 'heading'],
      ['• 화면(입력 도구): index.html — 브라우저로 열기', 'body'],
      ['• 백엔드(이 코드): 확장 프로그램 → Apps Script', 'body'],
      ['• 백업 가져오기: index.html의 액션 바', 'body'],
    ];

    // 값 한 번에 쓰기 (성능 최적화)
    const values = rows.map(function (r) { return [r[0]]; });
    sheet.getRange(1, 1, values.length, 1).setValues(values);

    // 스타일 적용
    rows.forEach(function (r, i) {
      const row = i + 1;
      const range = sheet.getRange(row, 1);
      const kind = r[1];

      if (kind === 'title') {
        range.setFontSize(16)
          .setFontWeight('bold')
          .setBackground('#4a7cff')
          .setFontColor('#ffffff')
          .setVerticalAlignment('middle');
        sheet.setRowHeight(row, 44);
      } else if (kind === 'heading') {
        range.setFontSize(12)
          .setFontWeight('bold')
          .setBackground('#f0f2f5')
          .setFontColor('#1a1a1a')
          .setVerticalAlignment('middle');
        sheet.setRowHeight(row, 32);
      } else if (kind === 'warn') {
        range.setFontSize(13)
          .setFontWeight('bold')
          .setBackground('#fef3e2')
          .setFontColor('#b45309')
          .setVerticalAlignment('middle');
        sheet.setRowHeight(row, 36);
      } else if (kind === 'sub') {
        range.setFontSize(10)
          .setFontColor('#666666')
          .setFontStyle('italic')
          .setVerticalAlignment('middle');
      } else {
        range.setFontSize(11)
          .setFontColor('#1a1a1a')
          .setVerticalAlignment('middle');
      }
    });

    // 컬럼 폭 + 격자 숨기기
    sheet.setColumnWidth(1, 820);
    sheet.hideGridlines();

    // 첫 번째 시트로 이동
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(1);

    Logger.log('✓ 사용법 시트 작성 완료: 📖_사용법');
    Logger.log('  마스터 시트를 새로고침해서 확인해보세요.');
  } catch (err) {
    Logger.log('✗ setupReadme 실패: ' + err.message);
    if (err.stack) Logger.log(err.stack);
  }
}


/**
 * 백업 폴더가 왜 안 잡히는지 진단하는 함수.
 * 함수 드롭다운에서 debugBackupFolder 선택 → ▶ 실행 → 로그 확인.
 */
function debugBackupFolder() {
  Logger.log('===== 백업 폴더 디버그 v2 =====');
  try {
    Logger.log('현재 실행 계정: ' + Session.getActiveUser().getEmail());
  } catch (e) {
    Logger.log('현재 실행 계정: (userinfo.email 스코프 없음 — 확인 생략)');
  }
  Logger.log('BACKUP_FOLDER_ID 값: "' + BACKUP_FOLDER_ID + '"');
  Logger.log('BACKUP_FOLDER_ID 길이: ' + BACKUP_FOLDER_ID.length);

  // 0. 시트 접근은 되는지 (Drive 권한과 별개라 매니페스트 진단의 기준선)
  Logger.log('');
  Logger.log('--- [시도 0] SpreadsheetApp.openById (시트 권한 기준선) ---');
  try {
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    Logger.log('✓ 시트 접근 성공: ' + ss.getName());
  } catch (e) {
    Logger.log('✗ 시트 접근 실패: ' + e.message);
    Logger.log('  → 시트조차 안 되면 매니페스트 자체가 미적용. 권한 캐시 삭제 필요.');
  }

  // 0-1. 가장 단순한 Drive 호출 — 권한만 있으면 무조건 통과
  Logger.log('');
  Logger.log('--- [시도 0-1] DriveApp.getStorageLimit (Drive 권한 기준선) ---');
  try {
    const limit = DriveApp.getStorageLimit();
    Logger.log('✓ Drive 접근 가능. 저장 한도(바이트): ' + limit);
  } catch (e) {
    Logger.log('✗ Drive 접근 실패: ' + e.message);
    Logger.log('  → Drive 스코프 자체가 동의 안 됨. 새 권한 동의 필요.');
  }

  if (!BACKUP_FOLDER_ID) {
    Logger.log('✗ BACKUP_FOLDER_ID가 비어있음. 코드 상단에 값을 채워주세요.');
    return;
  }

  // 1. 직접 ID로 가져오기 시도
  Logger.log('');
  Logger.log('--- [시도 1] DriveApp.getFolderById ---');
  try {
    const folder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
    Logger.log('✓ 성공: 폴더 이름 = ' + folder.getName());
    Logger.log('  URL: ' + folder.getUrl());
    try {
      Logger.log('  소유자: ' + folder.getOwner().getEmail());
    } catch (e) {
      Logger.log('  소유자 정보 조회 실패 (' + e.message + ')');
    }
    try {
      const testFile = folder.createFile(
        'debugTest_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'HHmmss') + '.txt',
        '디버그 테스트',
        'text/plain'
      );
      Logger.log('✓ 폴더에 쓰기 권한 확인됨. 테스트 파일: ' + testFile.getName());
    } catch (e) {
      Logger.log('✗ 폴더에 쓰기 권한 없음: ' + e.message);
    }
    return;
  } catch (e) {
    Logger.log('✗ 실패: ' + e.message);
  }

  // 2. URL로 시도 (혹시 ID 형식 문제)
  Logger.log('');
  Logger.log('--- [시도 2] DriveApp.getFolderByUrl ---');
  try {
    const folder = DriveApp.getFolderByUrl(
      'https://drive.google.com/drive/folders/' + BACKUP_FOLDER_ID
    );
    Logger.log('✓ URL로는 성공: ' + folder.getName());
  } catch (e) {
    Logger.log('✗ 실패: ' + e.message);
  }

  // 3. 내 드라이브 최상위 폴더 일부를 나열 (계정 확인용)
  Logger.log('');
  Logger.log('--- [시도 3] 현재 계정의 폴더 목록 (상위 15개) ---');
  try {
    const it = DriveApp.getFolders();
    let count = 0;
    while (it.hasNext() && count < 15) {
      const f = it.next();
      const match = (f.getId() === BACKUP_FOLDER_ID) ? '  ← 일치!' : '';
      Logger.log('  · ' + f.getName() + '  /  ID=' + f.getId() + match);
      count++;
    }
    if (count === 0) {
      Logger.log('  (이 계정 드라이브에 폴더가 하나도 없음)');
    }
  } catch (e) {
    Logger.log('✗ 폴더 목록 조회 실패: ' + e.message);
    Logger.log('  → Drive 권한이 동의되지 않은 상태일 가능성.');
  }

  Logger.log('');
  Logger.log('===== 디버그 끝 =====');
  Logger.log('위 로그를 그대로 복사해서 보내주시면 원인 파악 가능합니다.');
}
