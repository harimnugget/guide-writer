/**
 * SNS 디스크립션 생성기 - Google Apps Script
 *
 * === 배포 방법 ===
 * 1. Google Sheets에서 새 스프레드시트 생성
 * 2. 확장 프로그램 > Apps Script 클릭
 * 3. 이 코드를 Code.gs에 붙여넣기
 * 4. 배포 > 새 배포 > 웹 앱 선택
 *    - 실행 사용자: "나"
 *    - 액세스: "모든 사용자"
 * 5. 배포 후 나오는 URL을 description-generator.html 설정에 입력
 *
 * === 시트 구조 (자동 생성됨) ===
 * Sheet: profiles
 * | profileId | name | platform | channelSamples | referenceSamples | analysis | createdAt | updatedAt |
 */

const SHEET_NAME = 'profiles';

// ===== 웹 앱 엔드포인트 =====

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    let result;
    switch (action) {
      case 'listProfiles':
        result = listProfiles(body.platform);
        break;
      case 'getProfile':
        result = getProfile(body.profileId);
        break;
      case 'saveProfile':
        result = saveProfile(body);
        break;
      case 'deleteProfile':
        result = deleteProfile(body.profileId);
        break;
      case 'generate':
        result = generate(body.apiKey, body.prompt);
        break;
      default:
        throw new Error('Unknown action: ' + action);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// GET 요청도 지원 (CORS preflight 등)
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 시트 초기화 =====

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'profileId', 'name', 'platform',
      'channelSamples', 'referenceSamples', 'analysis',
      'createdAt', 'updatedAt'
    ]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }

  return sheet;
}

// ===== 프로필 CRUD =====

function listProfiles(platform) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const profiles = [];
  for (let i = 1; i < data.length; i++) {
    const row = rowToObject(headers, data[i]);
    if (platform && row.platform !== platform) continue;

    const channelSamples = safeJsonParse(row.channelSamples, []);
    const referenceSamples = safeJsonParse(row.referenceSamples, []);

    profiles.push({
      profileId: row.profileId,
      name: row.name,
      platform: row.platform,
      sampleCount: channelSamples.length,
      refCount: referenceSamples.length,
      updatedAt: row.updatedAt
    });
  }

  return { profiles };
}

function getProfile(profileId) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    const row = rowToObject(headers, data[i]);
    if (row.profileId === profileId) {
      return {
        profile: {
          profileId: row.profileId,
          name: row.name,
          platform: row.platform,
          channelSamples: safeJsonParse(row.channelSamples, []),
          referenceSamples: safeJsonParse(row.referenceSamples, []),
          analysis: safeJsonParse(row.analysis, null),
          createdAt: row.createdAt,
          updatedAt: row.updatedAt
        }
      };
    }
  }

  throw new Error('프로필을 찾을 수 없습니다: ' + profileId);
}

function saveProfile(data) {
  const sheet = getOrCreateSheet();
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const now = new Date().toISOString();

  // 기존 프로필 업데이트
  if (data.profileId) {
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === data.profileId) {
        const rowNum = i + 1;
        sheet.getRange(rowNum, 2).setValue(data.name);
        sheet.getRange(rowNum, 3).setValue(data.platform);
        sheet.getRange(rowNum, 4).setValue(JSON.stringify(data.channelSamples || []));
        sheet.getRange(rowNum, 5).setValue(JSON.stringify(data.referenceSamples || []));
        sheet.getRange(rowNum, 6).setValue(JSON.stringify(data.analysis || null));
        sheet.getRange(rowNum, 8).setValue(now);
        return { profileId: data.profileId, updated: true };
      }
    }
  }

  // 새 프로필 생성
  const profileId = 'p_' + Utilities.getUuid().substring(0, 8);
  sheet.appendRow([
    profileId,
    data.name,
    data.platform,
    JSON.stringify(data.channelSamples || []),
    JSON.stringify(data.referenceSamples || []),
    JSON.stringify(data.analysis || null),
    now,
    now
  ]);

  return { profileId, created: true };
}

function deleteProfile(profileId) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === profileId) {
      sheet.deleteRow(i + 1);
      return { deleted: true };
    }
  }

  throw new Error('프로필을 찾을 수 없습니다');
}

// ===== Claude API 프록시 =====

function generate(apiKey, prompt) {
  const url = 'https://api.anthropic.com/v1/messages';

  const payload = {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 2048,
    messages: [{ role: 'user', content: prompt }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    throw new Error('Claude API 오류 (' + responseCode + '): ' + response.getContentText());
  }

  const result = JSON.parse(response.getContentText());
  const text = result.content[0].text;

  // JSON 파싱 시도
  try {
    const jsonMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (jsonMatch) {
      return { data: JSON.parse(jsonMatch[1].trim()) };
    }
    return { data: JSON.parse(text) };
  } catch (e) {
    const braceStart = text.indexOf('{');
    const braceEnd = text.lastIndexOf('}');
    if (braceStart !== -1 && braceEnd !== -1) {
      return { data: JSON.parse(text.substring(braceStart, braceEnd + 1)) };
    }
    throw new Error('AI 응답 파싱 실패');
  }
}

// ===== 유틸리티 =====

function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((h, i) => { obj[h] = row[i]; });
  return obj;
}

function safeJsonParse(str, fallback) {
  if (!str) return fallback;
  try { return JSON.parse(str); } catch (e) { return fallback; }
}
