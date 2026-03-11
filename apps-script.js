/**
 * ============================================================
 * 공개교육 참가자 관리기 - Google Apps Script Backend
 * ============================================================
 *
 * [설정 방법]
 * 1. 새 구글 시트를 만들고, 탭 3개 생성: 교육목록, 참가자, 설정
 * 2. 각 탭의 헤더(1행)를 아래와 같이 입력:
 *    - 교육목록: 교육ID | 교육명 | 교육일자 | FormID | 응답시트ID | Zoom링크 | 상태
 *    - 참가자:   교육ID | 타임스탬프 | 이름 | 전화번호 | 이메일 | 소속 | 신청글 | 입금여부 | 입금일 | 비고
 *    - 설정:     키 | 값
 * 3. 설정 탭에 기본 열 매핑 입력:
 *    이름열   | 2
 *    전화열   | 3
 *    메일열   | 4
 *    소속열   | 5
 *    신청글열 | 6
 * 4. 확장 프로그램 > Apps Script 클릭
 * 5. 기존 코드를 모두 삭제하고, 이 파일의 내용을 붙여넣기
 * 6. 저장 (Ctrl+S)
 * 7. 배포 > 새 배포 > 유형: 웹 앱 > 실행: 본인 > 액세스: 모든 사용자 > 배포
 * 8. 생성된 URL을 웹앱 설정에 붙여넣기
 *
 * [주의] 코드 수정 후 반드시 새 배포를 만들어야 변경사항 반영됩니다.
 */

// ── 시트 접근 헬퍼 ──

function getSheetByName(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getHeaders(sheet) {
  if (!sheet || sheet.getLastColumn() === 0) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
      }
      obj[headers[j]] = val;
    }
    rows.push(obj);
  }
  return rows;
}

function getSettings() {
  var sheet = getSheetByName('설정');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var settings = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) settings[data[i][0]] = data[i][1];
  }
  return settings;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET 요청 처리 ──

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'getEvents';

    switch (action) {
      case 'getEvents':
        return jsonResponse(getEvents());
      default:
        return jsonResponse({ success: false, error: '알 수 없는 action: ' + action });
    }
  } catch (error) {
    return jsonResponse({ success: false, error: error.message });
  }
}

// ── POST 요청 처리 ──

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

    switch (action) {
      case 'getParticipants':
        return jsonResponse(getParticipants(body.eventId));
      case 'addEvent':
        return jsonResponse(addEvent(body.data));
      case 'syncParticipants':
        return jsonResponse(syncParticipants(body.eventId));
      case 'updatePayment':
        return jsonResponse(updatePayment(body.eventId, body.email, body.paid, body.paidDate));
      case 'sendBulkEmail':
        return jsonResponse(sendBulkEmail(body.recipients, body.subject, body.htmlBody));
      default:
        return jsonResponse({ success: false, error: '알 수 없는 action: ' + action });
    }
  } catch (error) {
    return jsonResponse({ success: false, error: error.message });
  }
}

// ── 교육 목록 ──

function getEvents() {
  var sheet = getSheetByName('교육목록');
  if (!sheet) return { success: false, error: '교육목록 탭이 없습니다.' };
  var rows = sheetToObjects(sheet);
  return { success: true, data: rows };
}

// ── 새 교육 등록 ──

function addEvent(data) {
  var sheet = getSheetByName('교육목록');
  if (!sheet) return { success: false, error: '교육목록 탭이 없습니다.' };
  var headers = getHeaders(sheet);

  // FormID로 응답시트ID 자동 탐지
  var responseSheetId = '';
  if (data['FormID']) {
    try {
      var form = FormApp.openById(data['FormID']);
      responseSheetId = form.getDestinationId();
    } catch (err) {
      // Form에 아직 응답시트가 연결 안 된 경우
      responseSheetId = '(연결안됨)';
    }
  }
  data['응답시트ID'] = responseSheetId;

  var newRow = headers.map(function(h) {
    return data[h] !== undefined ? data[h] : '';
  });
  sheet.appendRow(newRow);

  return { success: true, action: 'addEvent', data: data };
}

// ── 참가자 조회 ──

function getParticipants(eventId) {
  var sheet = getSheetByName('참가자');
  if (!sheet) return { success: false, error: '참가자 탭이 없습니다.' };
  var all = sheetToObjects(sheet);
  var filtered = all.filter(function(r) { return String(r['교육ID']) === String(eventId); });
  return { success: true, data: filtered, count: filtered.length };
}

// ── Form 응답 동기화 ──

function syncParticipants(eventId) {
  // 1) 교육목록에서 해당 이벤트 찾기
  var eventSheet = getSheetByName('교육목록');
  if (!eventSheet) return { success: false, error: '교육목록 탭이 없습니다.' };
  var events = sheetToObjects(eventSheet);
  var event = null;
  for (var i = 0; i < events.length; i++) {
    if (String(events[i]['교육ID']) === String(eventId)) { event = events[i]; break; }
  }
  if (!event) return { success: false, error: '교육을 찾을 수 없습니다: ' + eventId };

  // 2) 응답시트ID 확인 (없으면 FormID로 재탐지)
  var responseSheetId = event['응답시트ID'];
  if (!responseSheetId || responseSheetId === '(연결안됨)') {
    if (!event['FormID']) return { success: false, error: 'FormID가 없습니다.' };
    try {
      var form = FormApp.openById(event['FormID']);
      responseSheetId = form.getDestinationId();
      // 교육목록에 응답시트ID 업데이트
      var evHeaders = getHeaders(eventSheet);
      var ridCol = evHeaders.indexOf('응답시트ID');
      var allEvData = eventSheet.getDataRange().getValues();
      for (var ei = 1; ei < allEvData.length; ei++) {
        if (String(allEvData[ei][0]) === String(eventId)) {
          eventSheet.getRange(ei + 1, ridCol + 1).setValue(responseSheetId);
          break;
        }
      }
    } catch (err) {
      return { success: false, error: 'Form 응답시트를 찾을 수 없습니다: ' + err.message };
    }
  }

  // 3) 응답시트 읽기
  var responseSS;
  try {
    responseSS = SpreadsheetApp.openById(responseSheetId);
  } catch (err) {
    return { success: false, error: '응답시트를 열 수 없습니다: ' + err.message };
  }
  var responseSheet = responseSS.getSheets()[0];
  var responseData = responseSheet.getDataRange().getValues();
  if (responseData.length < 2) return { success: true, message: '응답 데이터가 없습니다.', added: 0 };

  // 4) 설정에서 열 매핑 읽기
  var settings = getSettings();
  var nameCol = Number(settings['이름열'] || 2) - 1;
  var phoneCol = Number(settings['전화열'] || 3) - 1;
  var emailCol = Number(settings['메일열'] || 4) - 1;
  var orgCol = Number(settings['소속열'] || 5) - 1;
  var textCol = Number(settings['신청글열'] || 6) - 1;

  // 5) 기존 참가자 이메일 목록 (중복 방지)
  var partSheet = getSheetByName('참가자');
  if (!partSheet) return { success: false, error: '참가자 탭이 없습니다.' };
  var existing = sheetToObjects(partSheet);
  var existingEmails = {};
  existing.forEach(function(r) {
    if (String(r['교육ID']) === String(eventId) && r['이메일']) {
      existingEmails[String(r['이메일']).trim().toLowerCase()] = true;
    }
  });

  // 6) 새 참가자 추가
  var partHeaders = getHeaders(partSheet);
  var added = 0;
  for (var ri = 1; ri < responseData.length; ri++) {
    var row = responseData[ri];
    var email = String(row[emailCol] || '').trim();
    if (!email) continue;
    if (existingEmails[email.toLowerCase()]) continue;

    var timestamp = row[0];
    if (timestamp instanceof Date) {
      timestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    }

    var newData = {
      '교육ID': eventId,
      '타임스탬프': timestamp,
      '이름': String(row[nameCol] || '').trim(),
      '전화번호': String(row[phoneCol] || '').trim(),
      '이메일': email,
      '소속': String(row[orgCol] || '').trim(),
      '신청글': String(row[textCol] || '').trim(),
      '입금여부': '',
      '입금일': '',
      '비고': ''
    };

    var newRow = partHeaders.map(function(h) { return newData[h] || ''; });
    partSheet.appendRow(newRow);
    added++;
  }

  return { success: true, action: 'syncParticipants', added: added, total: responseData.length - 1 };
}

// ── 입금 상태 업데이트 ──

function updatePayment(eventId, email, paid, paidDate) {
  var sheet = getSheetByName('참가자');
  if (!sheet) return { success: false, error: '참가자 탭이 없습니다.' };
  var headers = getHeaders(sheet);
  var data = sheet.getDataRange().getValues();

  var eidCol = headers.indexOf('교육ID');
  var emailCol = headers.indexOf('이메일');
  var paidCol = headers.indexOf('입금여부');
  var dateCol = headers.indexOf('입금일');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][eidCol]) === String(eventId) &&
        String(data[i][emailCol]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      sheet.getRange(i + 1, paidCol + 1).setValue(paid ? 'Y' : '');
      sheet.getRange(i + 1, dateCol + 1).setValue(paid ? (paidDate || '') : '');
      return { success: true, action: 'updatePayment', email: email, paid: paid };
    }
  }
  return { success: false, error: '참가자를 찾을 수 없습니다: ' + email };
}

// ── 일괄 메일 발송 ──

function sendBulkEmail(recipients, subject, htmlBody) {
  if (!recipients || recipients.length === 0) {
    return { success: false, error: '수신자가 없습니다.' };
  }
  if (!subject || !htmlBody) {
    return { success: false, error: '제목과 본문을 입력하세요.' };
  }

  var sent = 0;
  var failed = [];
  for (var i = 0; i < recipients.length; i++) {
    try {
      GmailApp.sendEmail(recipients[i], subject, '', { htmlBody: htmlBody });
      sent++;
    } catch (err) {
      failed.push({ email: recipients[i], error: err.message });
    }
  }

  return {
    success: true,
    action: 'sendBulkEmail',
    sent: sent,
    failed: failed,
    total: recipients.length
  };
}
