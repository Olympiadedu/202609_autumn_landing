var APPLICATION_SHEET_NAME = '신청내역';
var CAMPUS_SHEET_NAME = '캠퍼스정보';
var SUBMISSION_LOG_SHEET_NAME = '전송테스트로그';
var SETTINGS_SHEET_NAME = '설정';
var SCHEDULE_SHEET_NAME = '일정관리';
var LEGACY_BASIC_SCHEDULE_SHEET_NAME = '기본일정';
var LEGACY_SPECIAL_SCHEDULE_SHEET_NAME = '특이일정';
var SETTINGS_SPREADSHEET_NAME = '랜딩페이지_정보';
var TEST_MODE_LABEL = '테스트 운영 (실제 사이트 미전송)';
var LIVE_MODE_LABEL = '라이브 운영 (실제 사이트 전송)';
var TEST_EMAIL_DISABLED_LABEL = '발송 안 함';
var TEST_EMAIL_ENABLED_LABEL = '테스트 메일 발송';
var SETTINGS_SPREADSHEET_CACHE_KEY = 'LANDING_SETTINGS_SPREADSHEET_ID';
var DUPLICATE_CHECK_ROW_LIMIT = 500;

function setupSheets() {
  setupSettingsSheet_(false);

  var applicationSheet = getOrCreateApplicationSheet_();
  normalizeApplicationSheet_(applicationSheet);
  ensureHeader_(applicationSheet, getApplicationHeaders_());

  var submissionLogSheet = getOrCreateApplicationNamedSheet_(SUBMISSION_LOG_SHEET_NAME);
  normalizeSubmissionLogSheet_(submissionLogSheet);
  ensureHeader_(submissionLogSheet, getSubmissionLogHeaders_());

  var campusSheet = getOrCreateCampusSheet_();
  ensureHeader_(campusSheet, getCampusHeaders_());

  setupScheduleSheets_();
}

function setupScheduleSheets() {
  return setupScheduleSheets_(false);
}

function setupScheduleSheet() {
  return setupScheduleSheets_(false);
}

function rebuildScheduleSheets() {
  return setupScheduleSheets_(true);
}

function rebuildScheduleSheet() {
  return setupScheduleSheets_(true);
}

function setupScheduleSheets_(replaceExisting) {
  var ss = getSettingsSpreadsheet_();
  var sheet = getOrCreateSheet_(ss, SCHEDULE_SHEET_NAME);
  var basicHeaders = getBasicScheduleHeaders_();
  var specialHeaders = getSpecialScheduleHeaders_();
  var basicRows = getDefaultBasicScheduleRows_();
  var specialRows = getDefaultSpecialScheduleRows_();

  if (replaceExisting) {
    sheet.getRange(1, 1, sheet.getMaxRows(), Math.min(sheet.getMaxColumns(), 15)).breakApart();
    sheet.clearContents();
  }

  if (sheet.getLastRow() < 2 || replaceExisting) {
    sheet.getRange(1, 1, 1, basicHeaders.length).merge().setValue('기본일정');
    sheet.getRange(1, 9, 1, specialHeaders.length).merge().setValue('특이일정');
    sheet.getRange(2, 1, 1, basicHeaders.length).setValues([basicHeaders]);
    sheet.getRange(2, 9, 1, specialHeaders.length).setValues([specialHeaders]);
    sheet.getRange(3, 1, basicRows.length, basicHeaders.length).setValues(basicRows);
    sheet.getRange(3, 9, specialRows.length, specialHeaders.length).setValues(specialRows);
    sheet.getRange(3, 9, specialRows.length, 1).setNumberFormat('yyyy-mm-dd');
  }

  formatScheduleSheet_(sheet, basicRows.length, specialRows.length);
  applyScheduleValidations_(sheet, basicRows.length, specialRows.length);

  if (replaceExisting) {
    [LEGACY_BASIC_SCHEDULE_SHEET_NAME, LEGACY_SPECIAL_SCHEDULE_SHEET_NAME].forEach(function (name) {
      var legacySheet = ss.getSheetByName(name);
      if (legacySheet) ss.deleteSheet(legacySheet);
    });
  }

  return {
    scheduleSheet: SCHEDULE_SHEET_NAME,
    replaced: !!replaceExisting,
    message: replaceExisting
      ? '일정관리 시트에 기본일정과 특이일정을 좌우 표로 다시 작성했습니다.'
      : '일정관리 시트가 준비되었습니다.'
  };
}

function getDefaultBasicScheduleRows_() {
  return [
    ['전체', '월~금', '전체', '전체', '신청가능', '14:00, 15:00, 16:00, 17:00, 18:00, 19:00', '평일 기본 시간'],
    ['전체', '토', '전체', '전체', '신청가능', '11:00, 12:00, 13:00, 14:00, 15:00', '토요일 기본 시간'],
    ['전체', '일', '전체', '전체', '휴무', '', '일요일 선택 불가'],
    ['광진', '토', '수학', '전체', '휴무', '', '특이일정 날짜는 별도 규칙 우선']
  ];
}

function getDefaultSpecialScheduleRows_() {
  var rows = [
    ['2026-08-17', '전체', '전체', '전체', '휴무', '', '모든 캠퍼스 미운영']
  ];
  var mainCampuses = ['광진', '성동', '동대문', '중랑'];

  ['2026-08-22', '2026-08-29'].forEach(function (date, index) {
    var round = (index + 1) + '차';

    mainCampuses.forEach(function (campus) {
      rows.push(
        [date, campus, '수학', '초등, 중등', '신청가능', '12:00', round],
        [date, campus, '영어', '초등, 중등', '신청가능', '11:00', round],
        [date, campus, '수학', '고등', '신청가능', '14:00', round],
        [date, campus, '영어', '고등', '신청가능', '14:00', round]
      );
    });

    rows.push(
      [date, '미사', '수학', '초등, 중등', '신청가능', '11:00', round],
      [date, '미사', '수학', '고등', '신청가능', '14:00', round],
      [date, '중계', '수학', '전체', '신청가능', '14:00', round],
      [date, '송파방이', '수학', '전체', '신청가능', '11:00', round]
    );
  });

  return rows;
}

function applyScheduleValidations_(sheet, basicRowCount, specialRowCount) {
  var campusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['전체', '광진', '성동', '동대문', '중랑', '미사', '중계', '송파방이'], true)
    .setAllowInvalid(false)
    .build();
  var subjectValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['전체', '수학', '영어'], true)
    .setAllowInvalid(false)
    .build();
  var gradeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['전체', '초등', '중등', '고등', '초등, 중등'], true)
    .setAllowInvalid(false)
    .build();
  var statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['신청가능', '휴무'], true)
    .setAllowInvalid(false)
    .build();
  var weekdayValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['월~금', '토', '일'], true)
    .setAllowInvalid(false)
    .build();

  var editableRows = Math.max(sheet.getMaxRows() - 2, 1);
  var basicRows = Math.max(editableRows, basicRowCount, 1);
  var specialRows = Math.max(editableRows, specialRowCount, 1);
  sheet.getRange(3, 1, basicRows, 1).setDataValidation(campusValidation);
  sheet.getRange(3, 2, basicRows, 1).setDataValidation(weekdayValidation);
  sheet.getRange(3, 3, basicRows, 1).setDataValidation(subjectValidation);
  sheet.getRange(3, 4, basicRows, 1).setDataValidation(gradeValidation);
  sheet.getRange(3, 5, basicRows, 1).setDataValidation(statusValidation);
  sheet.getRange(3, 10, specialRows, 1).setDataValidation(campusValidation);
  sheet.getRange(3, 11, specialRows, 1).setDataValidation(subjectValidation);
  sheet.getRange(3, 12, specialRows, 1).setDataValidation(gradeValidation);
  sheet.getRange(3, 13, specialRows, 1).setDataValidation(statusValidation);
}

function formatScheduleSheet_(sheet, basicRowCount, specialRowCount) {
  sheet.setFrozenRows(2);
  sheet.getRange(1, 1, 1, 7)
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setBackground('#dce6f8');
  sheet.getRange(1, 9, 1, 7)
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setBackground('#fce8df');
  sheet.getRange(2, 1, 1, 7).setFontWeight('bold').setBackground('#eef3fb');
  sheet.getRange(2, 9, 1, 7).setFontWeight('bold').setBackground('#fff2ec');
  sheet.getRange(3, 1, Math.max(basicRowCount, 1), 6)
    .setHorizontalAlignment('center');
  sheet.getRange(3, 9, Math.max(specialRowCount, 1), 6)
    .setHorizontalAlignment('center');
  sheet.getRange(3, 7, Math.max(basicRowCount, 1), 1).setWrap(true);
  sheet.getRange(3, 15, Math.max(specialRowCount, 1), 1).setWrap(true);
  sheet.setColumnWidth(8, 28);
  sheet.autoResizeColumns(1, 7);
  sheet.autoResizeColumns(9, 7);
}

function migrateScriptPropertiesToSettings() {
  var result = setupSettingsSheet_(true);
  installPendingSubmissionTrigger();
  result.pendingSubmissionTriggerInstalled = true;
  return result;
}

function migrateSheetSchemas() {
  var applicationSheet = getOrCreateApplicationSheet_();
  normalizeApplicationSheet_(applicationSheet);
  ensureHeader_(applicationSheet, getApplicationHeaders_());

  var submissionLogSheet = getOrCreateApplicationNamedSheet_(SUBMISSION_LOG_SHEET_NAME);
  normalizeSubmissionLogSheet_(submissionLogSheet);
  ensureHeader_(submissionLogSheet, getSubmissionLogHeaders_());
}

function testSampleApplication() {
  var sample = {
    studentName: '테스트학생',
    region: '서울/경기',
    campusId: 'songpa',
    campusName: '송파방이캠퍼스',
    campusLocation: '서울특별시 송파구',
    school: '테스트초등학교',
    grade: '초4',
    phone: '010-0000-0000',
    preferredDate: '2026-07-15',
    preferredTime: '14:00',
    subjectId: 'math',
    subjectName: '수학',
    referralSource: '랜딩',
    referralDetail: 'test',
    utmSource: 'test',
    utmMedium: 'landing',
    utmCampaign: 'sample'
  };

  var route = findCampusRoute_(sample);
  var row = appendApplication_(sample, route);
  var submission = processSiteSubmission_(sample, route, row);
  updateApplicationSubmissionStatus_(row, submission);
  return submission;
}

function doGet(e) {
  var action = e && e.parameter && e.parameter.action ? e.parameter.action : '';
  if (action === 'campus-options') {
    return campusOptionsResponse_(e);
  }
  if (action === 'redirect-site') {
    return redirectSiteResponse_(e);
  }

  return json_({
    ok: false,
    message: '지원하지 않는 요청입니다.'
  });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;

  try {
    var payload = parsePayload_(e);
    var route = findCampusRoute_(payload);

    lock.waitLock(30000);
    locked = true;
    var row = appendApplication_(payload, route);

    return json_({
      ok: true,
      row: row,
      queued: true,
      message: '신청내역 저장 완료'
    });
  } catch (err) {
    return json_({
      ok: false,
      message: err && err.message ? err.message : String(err)
    });
  } finally {
    if (locked) {
      try {
        lock.releaseLock();
      } catch (err) {
        // Lock may not have been acquired if parsing failed early.
      }
    }
  }
}

function processPendingSubmissions() {
  var lock = LockService.getScriptLock();
  var locked = false;

  try {
    lock.waitLock(30000);
    locked = true;

    var sheet = getOrCreateApplicationSheet_();
    normalizeApplicationSheet_(sheet);
    ensureHeader_(sheet, getApplicationHeaders_());

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { ok: true, processed: 0, message: '전송대기 신청이 없습니다.' };
    }

    var headers = getApplicationHeaders_();
    var statusCol = headers.indexOf('사이트전송상태') + 1;
    if (statusCol <= 0) {
      throw new Error('사이트전송상태 컬럼이 없습니다.');
    }

    var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var maxCount = Number(getSetting_('PENDING_SUBMIT_LIMIT') || 10);
    var processed = 0;

    for (var i = 0; i < values.length; i++) {
      if (processed >= maxCount) break;

      var rowNumber = i + 2;
      var row = values[i];
      var status = cellText_(row[statusCol - 1]);

      if (status !== '전송대기' && status !== '') continue;

      sheet.getRange(rowNumber, statusCol).setValue('전송중');
      SpreadsheetApp.flush();

      lock.releaseLock();
      locked = false;

      var payload = buildPayloadFromApplicationRow_(row);
      var route = findCampusRoute_(payload);
      var submission;

      try {
        submission = processSiteSubmission_(payload, route, rowNumber);
      } catch (siteErr) {
        submission = {
          mode: getSiteSubmitMode_(),
          inputSite: route.inputSite || '',
          routeMatched: !!route.matched,
          status: 'SITE_SUBMIT_ERROR',
          message: siteErr && siteErr.message ? siteErr.message : String(siteErr)
        };
      }

      lock.waitLock(30000);
      locked = true;
      updateApplicationSubmissionStatus_(rowNumber, submission);
      sendCampusNotification_(rowNumber, payload, route, submission);
      processed += 1;
    }

    return { ok: true, processed: processed };
  } finally {
    if (locked) {
      try {
        lock.releaseLock();
      } catch (err) {
        // Lock may not have been acquired if processing failed early.
      }
    }
  }
}

function installPendingSubmissionTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'processPendingSubmissions') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('processPendingSubmissions')
    .timeBased()
    .everyMinutes(1)
    .create();
}

function parsePayload_(e) {
  var raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
  var payload = JSON.parse(raw);

  requireField_(payload, 'studentName', '학생명');
  requireField_(payload, 'region', '지역');
  requireField_(payload, 'campusName', '캠퍼스명');
  requireField_(payload, 'school', '학교명');
  requireField_(payload, 'grade', '학년');
  requireField_(payload, 'phone', '연락처');
  requireField_(payload, 'preferredDate', '희망날짜');
  requireField_(payload, 'preferredTime', '희망시간');
  requireField_(payload, 'subjectName', '전형과목');

  payload.referralSource = payload.referralSource || '랜딩';
  payload.referralDetail = payload.referralDetail || '기타';
  payload.utmSource = payload.utmSource || '';
  payload.utmMedium = payload.utmMedium || '';
  payload.utmCampaign = payload.utmCampaign || '';
  return payload;
}

function campusOptionsResponse_(e) {
  var result = {
    ok: true,
    data: readCampusOptions_()
  };

  return jsonpOrJson_(result, e);
}

function redirectSiteResponse_(e) {
  var campusId = e && e.parameter ? cellText_(e.parameter.campusId) : '';
  var subjectId = e && e.parameter ? cellText_(e.parameter.subjectId) : '';
  var url = findRedirectSite_(campusId, subjectId);

  return jsonpOrJson_({
    ok: !!url,
    url: url,
    message: url ? '' : '이동사이트가 등록되지 않았습니다.'
  }, e);
}

function jsonpOrJson_(result, e) {
  var callback = e && e.parameter && e.parameter.callback ? e.parameter.callback : '';
  if (callback) {
    callback = String(callback).replace(/[^\w.$]/g, '');
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return json_(result);
}

function findRedirectSite_(campusId, subjectId) {
  if (!campusId || !subjectId) return '';

  var sheet = getOrCreateCampusSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return '';

  var headers = values[0].map(function (header) {
    return String(header || '').trim();
  });
  var campusIdCol = findHeaderIndex_(headers, 'campusId');
  var redirectSiteCol = findHeaderIndex_(headers, '이동사이트');
  if (redirectSiteCol < 0) {
    redirectSiteCol = findHeaderIndex_(headers, '이동 사이트');
  }
  var divisionCol = findHeaderIndex_(headers, '구분');
  var subjectCols = findAllHeaderIndexes_(headers, '전형과목');
  if (campusIdCol < 0 || redirectSiteCol < 0) return '';

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (cellText_(row[campusIdCol]) !== campusId) continue;

    var rowSubjects = divisionCol >= 0 ? [cellText_(row[divisionCol])] : [];
    subjectCols.forEach(function (col) {
      rowSubjects.push(cellText_(row[col]));
    });

    for (var s = 0; s < rowSubjects.length; s++) {
      if (makeSubjectId_(rowSubjects[s]) === subjectId) {
        return cellText_(row[redirectSiteCol]);
      }
    }
  }

  return '';
}

function readCampusOptions_() {
  var sheet = getOrCreateCampusSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return {
      campuses: [],
      subjects: [],
      scheduleRules: readScheduleRules_()
    };
  }

  var headers = values[0].map(function (header) {
    return String(header || '').trim();
  });
  var campusIdCol = findHeaderIndex_(headers, 'campusId');
  var divisionCol = findHeaderIndex_(headers, '구분');
  var regionCol = findHeaderIndex_(headers, '지역');
  var campusNameCol = findHeaderIndex_(headers, '캠퍼스명');
  var inputSiteCol = findHeaderIndex_(headers, '입력사이트');
  var redirectSiteCol = findHeaderIndex_(headers, '이동사이트');
  if (redirectSiteCol < 0) {
    redirectSiteCol = findHeaderIndex_(headers, '이동 사이트');
  }
  var saturdayBlockedCol = findHeaderIndex_(headers, '토요일선택불가');
  var subjectCols = findAllHeaderIndexes_(headers, '전형과목');

  var campusesById = {};
  var subjectsById = {};

  values.slice(1).forEach(function (row) {
    var campusName = cellText_(row[campusNameCol]);
    if (!campusName) return;

    var campusId = cellText_(row[campusIdCol]) || makeCampusId_(campusName);
    var region = cellText_(row[regionCol]);

    if (!campusesById[campusId]) {
      campusesById[campusId] = {
        id: campusId,
        name: campusName,
        olympiadName: campusName,
        location: region,
        region: region,
        inputSites: {},
        redirectSites: {}
      };
    }

    var subjectNames = [cellText_(row[divisionCol])];
    subjectCols.forEach(function (col) {
      subjectNames.push(cellText_(row[col]));
    });

    subjectNames.forEach(function (subjectName) {
      splitSubjectNames_(subjectName).forEach(function (name) {
        var subjectId = makeSubjectId_(name);
        if (!subjectId) return;

        if (!subjectsById[subjectId]) {
          subjectsById[subjectId] = {
            id: subjectId,
            name: name,
            olympiadName: name,
            campusIds: [],
            unavailableSaturdayCampusIds: []
          };
        }

        if (subjectsById[subjectId].campusIds.indexOf(campusId) === -1) {
          subjectsById[subjectId].campusIds.push(campusId);
        }
        if (inputSiteCol >= 0) {
          campusesById[campusId].inputSites[subjectId] = cellText_(row[inputSiteCol]);
        }
        if (redirectSiteCol >= 0) {
          campusesById[campusId].redirectSites[subjectId] = cellText_(row[redirectSiteCol]);
        }

        if (
          saturdayBlockedCol >= 0 &&
          String(row[saturdayBlockedCol]).toUpperCase() === 'Y' &&
          subjectsById[subjectId].unavailableSaturdayCampusIds.indexOf(campusId) === -1
        ) {
          subjectsById[subjectId].unavailableSaturdayCampusIds.push(campusId);
        }
      });
    });
  });

  return {
    campuses: Object.keys(campusesById).map(function (id) { return campusesById[id]; }),
    subjects: Object.keys(subjectsById).map(function (id) { return subjectsById[id]; }),
    scheduleRules: readScheduleRules_()
  };
}

function readScheduleRules_() {
  var ss = getSettingsSpreadsheet_();
  var scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);

  if (scheduleSheet && scheduleSheet.getLastRow() >= 3) {
    return {
      basic: readScheduleTableRange_(scheduleSheet, 1, false),
      special: readScheduleTableRange_(scheduleSheet, 9, true)
    };
  }

  return {
    basic: readScheduleSheet_(ss.getSheetByName(LEGACY_BASIC_SCHEDULE_SHEET_NAME), false),
    special: readScheduleSheet_(ss.getSheetByName(LEGACY_SPECIAL_SCHEDULE_SHEET_NAME), true)
  };
}

function readScheduleTableRange_(sheet, startColumn, includeDate) {
  var rowCount = sheet.getLastRow() - 1;
  if (rowCount < 2) return [];
  var values = sheet.getRange(2, startColumn, rowCount, 7).getValues();
  return scheduleRulesFromValues_(values, includeDate);
}

function readScheduleSheet_(sheet, includeDate) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  return scheduleRulesFromValues_(sheet.getDataRange().getValues(), includeDate);
}

function scheduleRulesFromValues_(values, includeDate) {
  if (!values || values.length < 2) return [];
  var headers = values[0].map(function (header) {
    return cellText_(header);
  });
  var dateCol = findHeaderIndex_(headers, '날짜');
  var campusCol = findHeaderIndex_(headers, '캠퍼스');
  var weekdayCol = findHeaderIndex_(headers, '요일');
  var subjectCol = findHeaderIndex_(headers, '과목');
  var gradeGroupCol = findHeaderIndex_(headers, '학년군');
  var statusCol = findHeaderIndex_(headers, '상태');
  var timesCol = findHeaderIndex_(headers, '가능시간');
  var noteCol = findHeaderIndex_(headers, '비고');

  return values.slice(1).map(function (row) {
    var rule = {
      campuses: splitScheduleValues_(row[campusCol]),
      subjects: splitScheduleValues_(row[subjectCol]),
      gradeGroups: splitScheduleValues_(row[gradeGroupCol]),
      status: cellText_(row[statusCol]) || '신청가능',
      times: splitScheduleTimes_(row[timesCol]),
      note: noteCol >= 0 ? cellText_(row[noteCol]) : ''
    };

    if (includeDate) {
      rule.date = scheduleDateText_(row[dateCol]);
    } else {
      rule.weekdays = splitScheduleValues_(row[weekdayCol]);
    }
    return rule;
  }).filter(function (rule) {
    return includeDate ? !!rule.date : rule.weekdays.length > 0;
  });
}

function splitScheduleValues_(value) {
  return cellText_(value)
    .split(/[,，]/)
    .map(function (item) { return item.trim(); })
    .filter(function (item) { return !!item; });
}

function splitScheduleTimes_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return [Utilities.formatDate(value, 'Asia/Seoul', 'HH:mm')];
  }

  return splitScheduleValues_(value).map(function (item) {
    var match = item.match(/(?:^|\s)(\d{1,2}:\d{2})(?::\d{2})?/);
    return match ? match[1] : item;
  });
}

function scheduleDateText_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'Asia/Seoul', 'yyyy-MM-dd');
  }

  var text = cellText_(value);
  var match = text.match(/^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})/);
  if (!match) return text;
  return match[1] + '-' + ('0' + match[2]).slice(-2) + '-' + ('0' + match[3]).slice(-2);
}

function appendApplication_(payload, route) {
  var sheet = getOrCreateApplicationSheet_();

  if (isDuplicateApplication_(sheet, payload)) {
    throw new Error('이미 접수된 신청입니다.');
  }

  var rowNumber = sheet.getLastRow() + 1;
  var values = [
    Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
    payload.studentName,
    payload.region,
    payload.campusName,
    payload.campusLocation || '',
    payload.school,
    payload.grade,
    payload.phone,
    payload.preferredDate,
    payload.preferredTime,
    payload.subjectName,
    payload.referralSource,
    payload.referralDetail,
    payload.campusId || '',
    payload.subjectId || '',
    route.inputSite || '',
    '전송대기',
    '접수',
    payload.utmSource || '',
    payload.utmMedium || '',
    payload.utmCampaign || ''
  ];
  sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);

  return rowNumber;
}

function updateApplicationSubmissionStatus_(row, submission) {
  var sheet = getOrCreateApplicationSheet_();
  var headers = getApplicationHeaders_();
  var siteStatusCol = headers.indexOf('사이트전송상태') + 1;
  var processStatusCol = headers.indexOf('처리상태') + 1;

  if (siteStatusCol > 0) {
    sheet.getRange(row, siteStatusCol).setValue(buildApplicationStatusText_(submission));
  }

  if (processStatusCol > 0) {
    sheet
      .getRange(row, processStatusCol)
      .setValue(submission && submission.mode === 'LIVE' ? '접수/전송처리' : '접수/전송테스트');
  }
}

function buildPayloadFromApplicationRow_(row) {
  return {
    studentName: cellText_(row[1]),
    region: cellText_(row[2]),
    campusName: cellText_(row[3]),
    campusLocation: cellText_(row[4]),
    school: cellText_(row[5]),
    grade: cellText_(row[6]),
    phone: cellText_(row[7]),
    preferredDate: formatSheetDate_(row[8]),
    preferredTime: formatSheetTime_(row[9]),
    subjectName: cellText_(row[10]),
    referralSource: cellText_(row[11]) || '기타',
    referralDetail: cellText_(row[12]) || '랜딩',
    campusId: cellText_(row[13]),
    subjectId: cellText_(row[14])
  };
}

function formatSheetDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'Asia/Seoul', 'yyyy-MM-dd');
  }

  return cellText_(value);
}

function formatSheetTime_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'Asia/Seoul', 'HH:mm');
  }

  return cellText_(value);
}

function buildApplicationStatusText_(submission) {
  if (!submission) return '전송결과없음';
  if (submission.mode === 'LIVE') {
    return submission.status + (submission.httpStatus ? ' (' + submission.httpStatus + ')' : '');
  }
  return 'DRY_RUN';
}

function findCampusRoute_(payload) {
  var sheet = getOrCreateCampusSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return { inputSite: '', note: '', matched: false };
  }

  var headers = values[0].map(function (header) {
    return String(header || '').trim();
  });
  var campusIdCol = findHeaderIndex_(headers, 'campusId');
  var campusNameCol = findHeaderIndex_(headers, '캠퍼스명');
  var inputSiteCol = findHeaderIndex_(headers, '입력사이트');
  var noteCol = findHeaderIndex_(headers, '특이사항');
  var siteCol = findHeaderIndex_(headers, 'SITE');
  var consultingIdxCol = findHeaderIndex_(headers, 'CONSULTING_IDX');
  var acadIdCol = findHeaderIndex_(headers, 'ACAD_ID');
  var areaIdCol = findHeaderIndex_(headers, 'AREA_ID');
  var areaNameCol = findHeaderIndex_(headers, 'AREA_NAME');
  var acadNameCol = findHeaderIndex_(headers, 'ACAD_NAME');
  var managerEmailCol = findHeaderIndex_(headers, '\uB2F4\uB2F9\uC790\uC774\uBA54\uC77C');
  var subjectCols = findAllHeaderIndexes_(headers, '전형과목');

  var payloadCampusId = cellText_(payload.campusId);
  var payloadCampusName = normalizeText_(payload.campusName);
  var payloadSubjectName = normalizeText_(payload.subjectName);

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowCampusId = cellText_(row[campusIdCol]);
    var rowCampusName = normalizeText_(row[campusNameCol]);
    var campusMatched = payloadCampusId
      ? rowCampusId === payloadCampusId
      : rowCampusName === payloadCampusName;

    if (!campusMatched) continue;

    for (var s = 0; s < subjectCols.length; s++) {
      if (normalizeText_(row[subjectCols[s]]) === payloadSubjectName) {
        return {
          inputSite: cellText_(row[inputSiteCol]),
          note: cellText_(row[noteCol]),
          site: cellText_(row[siteCol]),
          consultingIdx: cellText_(row[consultingIdxCol]),
          acadId: cellText_(row[acadIdCol]),
          areaId: cellText_(row[areaIdCol]),
          areaName: cellText_(row[areaNameCol]),
          acadName: cellText_(row[acadNameCol]),
          managerEmail: managerEmailCol >= 0 ? cellText_(row[managerEmailCol]) : '',
          matched: true
        };
      }
    }
  }

  return { inputSite: '', note: '', matched: false };
}

function processSiteSubmission_(payload, route, applicationRow) {
  var mode = getSiteSubmitMode_();
  var targetPayload = buildTargetSitePayload_(payload, route);
  var result = {
    mode: mode,
    inputSite: route.inputSite || '',
    routeMatched: !!route.matched,
    status: 'DRY_RUN'
  };

  if (mode !== 'LIVE') {
    logSubmissionAttempt_(payload, route, applicationRow, result.status, targetPayload, '');
    return result;
  }

  if (!isLiveSupportedSite_(route.inputSite)) {
    result.status = 'LIVE_UNSUPPORTED_SITE';
    result.message = '현재 LIVE 전송은 math.olympiad.ac, www.glec.co.kr, www.u2math.co.kr만 준비되어 있습니다.';
    logSubmissionAttempt_(payload, route, applicationRow, result.status, targetPayload, result.message);
    return result;
  }

  result = submitProcApply_(route.inputSite, targetPayload);
  logSubmissionAttempt_(payload, route, applicationRow, result.status, targetPayload, result.message || '');
  return result;
}

function buildTargetSitePayload_(payload, route) {
  var inputSite = route.inputSite || '';
  var routeConfig = getRouteSiteConfig_(route);

  if (inputSite.indexOf('math.olympiad.ac') >= 0) {
    return buildProcApplyPayload_(payload, routeConfig || getSiteConfig_('math'));
  }

  if (inputSite.indexOf('glec.co.kr') >= 0) {
    return buildProcApplyPayload_(payload, routeConfig || getSiteConfig_('glec'));
  }

  if (inputSite.indexOf('u2math.co.kr') >= 0) {
    return buildProcApplyPayload_(payload, routeConfig || getSiteConfig_('u2m'));
  }

  return buildUnknownSitePayload_(payload, inputSite);
}

function buildProcApplyPayload_(payload, siteConfig) {
  return {
    CONSULTING_IDX: siteConfig.consultingIdx,
    SITE: siteConfig.site,
    ACAD_ID: siteConfig.acadId,
    AREA_ID: siteConfig.areaId,
    AREA_NAME: siteConfig.areaName || '',
    ACAD_NAME: siteConfig.acadName || payload.campusName || '',
    RESERVED_DATE: payload.preferredDate + ' ' + payload.preferredTime,
    CHILD_NAME: payload.studentName,
    CHILD_GRADE: payload.grade,
    CHILD_SCHOOL: payload.school,
    PARENT_PHONE: onlyDigits_(payload.phone),
    KNOWLEDGE_SELECT: (payload.referralSource || '랜딩') + '(' + (payload.referralDetail || '기타') + ')',
    CONTENTS: '',
    APPLY_YN: 'N',
    PASS_YN: 'N',
    ENROLLMENT_YN: 'N',
    STATUS: 'N',
    TYPE_SUBJECT: payload.subjectName
  };
}

function getRouteSiteConfig_(route) {
  if (!route) return null;

  var inputSite = route.inputSite || '';
  var fallbackSite = '';

  if (inputSite.indexOf('math.olympiad.ac') >= 0) fallbackSite = 'math';
  if (inputSite.indexOf('glec.co.kr') >= 0) fallbackSite = 'glec';
  if (inputSite.indexOf('u2math.co.kr') >= 0) fallbackSite = 'u2m';

  var site = route.site || fallbackSite;
  var fallbackConfig = site ? getSiteConfig_(site) : null;

  if (!route.acadId && !route.areaId && !route.consultingIdx && !route.site) {
    return null;
  }

  return {
    site: site || (fallbackConfig && fallbackConfig.site) || '',
    consultingIdx: route.consultingIdx || (fallbackConfig && fallbackConfig.consultingIdx) || '844',
    acadId: route.acadId || (fallbackConfig && fallbackConfig.acadId) || '',
    areaId: route.areaId || (fallbackConfig && fallbackConfig.areaId) || '',
    areaName: route.areaName || '',
    acadName: route.acadName || ''
  };
}

function getSiteConfig_(site) {
  var configs = {
    math: {
      site: 'math',
      consultingIdx: getSetting_('MATH_CONSULTING_IDX') || '844',
      acadId: getSetting_('MATH_ACAD_ID') || '4',
      areaId: getSetting_('MATH_AREA_ID') || '105A'
    },
    glec: {
      site: 'glec',
      consultingIdx: getSetting_('GLEC_CONSULTING_IDX') || '844',
      acadId: getSetting_('GLEC_ACAD_ID') || '4',
      areaId: getSetting_('GLEC_AREA_ID') || '105A'
    },
    u2m: {
      site: 'u2m',
      consultingIdx: getSetting_('U2M_CONSULTING_IDX') || '844',
      acadId: getSetting_('U2M_ACAD_ID') || '20',
      areaId: getSetting_('U2M_AREA_ID') || '105O'
    }
  };

  return configs[site];
}

function buildUnknownSitePayload_(payload, inputSite) {
  return {
    NOTICE: '이 사이트는 아직 실제 POST 필드명이 확인되지 않았습니다. Network Payload가 필요합니다.',
    INPUT_SITE: inputSite,
    학생명: payload.studentName,
    지역: payload.region,
    캠퍼스명: payload.campusName,
    학교명: payload.school,
    학년: payload.grade,
    연락처: onlyDigits_(payload.phone),
    희망날짜: payload.preferredDate,
    희망시: payload.preferredTime,
    전형과목: payload.subjectName,
    알게된경로: payload.referralSource || '랜딩',
    알게된경로내용: payload.referralDetail || '기타'
  };
}

function submitProcApply_(inputSite, targetPayload) {
  var origin = getOrigin_(inputSite);
  var procApplyUrl = origin + '/Exam/ProcApply';
  var writeResponse = UrlFetchApp.fetch(inputSite, {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });
  var cookies = collectCookies_(writeResponse);
  var response = UrlFetchApp.fetch(procApplyUrl, {
    method: 'post',
    payload: targetPayload,
    muteHttpExceptions: true,
    followRedirects: false,
    headers: {
      Cookie: cookies,
      Origin: origin,
      Referer: inputSite,
      'X-Requested-With': 'XMLHttpRequest'
    }
  });

  return {
    mode: 'LIVE',
    inputSite: inputSite,
    procApplyUrl: procApplyUrl,
    status: response.getResponseCode() >= 200 && response.getResponseCode() < 300 ? 'LIVE_SENT' : 'LIVE_FAILED',
    httpStatus: response.getResponseCode(),
    message: response.getContentText('UTF-8').substring(0, 500)
  };
}

function isLiveSupportedSite_(inputSite) {
  inputSite = String(inputSite || '');
  return inputSite.indexOf('math.olympiad.ac') >= 0
    || inputSite.indexOf('glec.co.kr') >= 0
    || inputSite.indexOf('u2math.co.kr') >= 0;
}

function collectCookies_(response) {
  var headers = response.getAllHeaders();
  var setCookie = headers['Set-Cookie'] || headers['set-cookie'] || [];
  if (!Array.isArray(setCookie)) setCookie = [setCookie];
  return setCookie
    .map(function (cookie) { return String(cookie).split(';')[0]; })
    .filter(function (cookie) { return cookie; })
    .join('; ');
}

function logSubmissionAttempt_(payload, route, applicationRow, status, targetPayload, message) {
  var sheet = getOrCreateApplicationNamedSheet_(SUBMISSION_LOG_SHEET_NAME);
  normalizeSubmissionLogSheet_(sheet);
  ensureHeader_(sheet, getSubmissionLogHeaders_());

  sheet.appendRow([
    Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
    applicationRow,
    status,
    payload.campusName,
    payload.subjectName,
    route.inputSite || '',
    JSON.stringify(targetPayload),
    message || ''
  ]);
}

function sendCampusNotification_(applicationRow, payload, route, submission) {
  var sheet = getOrCreateApplicationSheet_();
  var statusCol = ensureNamedColumn_(sheet, 'EMAIL_STATUS');
  var sentAtCol = ensureNamedColumn_(sheet, 'EMAIL_SENT_AT');
  var currentStatus = cellText_(sheet.getRange(applicationRow, statusCol).getValue());

  if (
    currentStatus === 'SENT' ||
    currentStatus === 'DRY_RUN_SENT' ||
    currentStatus === 'TEST_EMAIL_SKIPPED'
  ) return;

  var mode = submission && submission.mode ? submission.mode : getSiteSubmitMode_();
  var isDryRun = mode !== 'LIVE';
  if (isDryRun && !shouldSendTestEmail_()) {
    sheet.getRange(applicationRow, statusCol).setValue('TEST_EMAIL_SKIPPED');
    sheet.getRange(applicationRow, sentAtCol).clearContent();
    return;
  }

  var recipients = normalizeEmailList_(route && route.managerEmail);
  if (!recipients) {
    sheet.getRange(applicationRow, statusCol).setValue('NO_RECIPIENT');
    return;
  }

  var subject = (isDryRun ? '[TEST] ' : '')
    + '[\uC218\uC2DC \uD559\uB825\uC9C4\uB2E8\uD3C9\uAC00 \uC2E0\uCCAD] '
    + payload.campusName + ' / ' + payload.subjectName;
  var body = [
    isDryRun ? '[DRY_RUN \uD14C\uC2A4\uD2B8 \uBA54\uC77C]' : '',
    '\uC218\uC2DC \uD559\uB825\uC9C4\uB2E8\uD3C9\uAC00 \uC2E0\uCCAD\uC774',
    '\uC811\uC218\uB418\uC5C8\uC2B5\uB2C8\uB2E4.',
    '',
    '\uD76C\uB9DD\uC77C\uC2DC : ' + payload.preferredDate + ' ' + payload.preferredTime,
    '\uC804\uD615\uACFC\uBAA9 : ' + payload.subjectName,
    '\uD559\uB144 : ' + payload.grade,
    '\uD559\uC0DD\uBA85 : ' + payload.studentName,
    '',
    '\uC790\uC138\uD55C \uB0B4\uC6A9\uC740 \uB4DC\uB9BC\uD50C\uB7EC\uC2A4 \uB610\uB294 \uD648\uD398\uC774\uC9C0\uC5D0\uC11C',
    '\uD655\uC778\uD558\uC2DC\uAE30 \uBC14\uB78D\uB2C8\uB2E4.'
  ].join('\n');

  var options = {
    name: getSetting_('MAIL_SENDER_NAME') || '\uC62C\uB9BC\uD53C\uC544\uB4DC\uAD50\uC721',
    htmlBody: buildCampusNotificationHtml_(payload, isDryRun)
  };
  var fromAlias = getSetting_('MAIL_FROM_ALIAS');
  if (fromAlias && GmailApp.getAliases().indexOf(fromAlias) >= 0) {
    options.from = fromAlias;
  }

  try {
    GmailApp.sendEmail(recipients, subject, body, options);
    sheet.getRange(applicationRow, statusCol).setValue(isDryRun ? 'DRY_RUN_SENT' : 'SENT');
    sheet.getRange(applicationRow, sentAtCol).setValue(
      Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')
    );
  } catch (err) {
    sheet.getRange(applicationRow, statusCol).setValue(
      'FAILED: ' + (err && err.message ? err.message : String(err)).substring(0, 200)
    );
  }
}

function buildCampusNotificationHtml_(payload, isDryRun) {
  var rows = [
    ['\uD76C\uB9DD\uC77C\uC2DC', payload.preferredDate + ' ' + payload.preferredTime],
    ['\uC804\uD615\uACFC\uBAA9', payload.subjectName],
    ['\uD559\uB144', payload.grade],
    ['\uD559\uC0DD\uBA85', payload.studentName]
  ];
  var tableRows = rows.map(function (row) {
    return '<tr>'
      + '<th style="padding:10px 14px;border:1px solid #d9d9d9;background:#f5f5f5;text-align:left;white-space:nowrap;font-weight:600;">'
      + escapeHtml_(row[0]) + '</th>'
      + '<td style="padding:10px 14px;border:1px solid #d9d9d9;">' + escapeHtml_(row[1]) + '</td>'
      + '</tr>';
  }).join('');

  return '<div style="font-family:Arial,\'Apple SD Gothic Neo\',\'Malgun Gothic\',sans-serif;color:#222;line-height:1.6;">'
    + (isDryRun ? '<p style="color:#d64500;font-weight:700;">[DRY_RUN \uD14C\uC2A4\uD2B8 \uBA54\uC77C]</p>' : '')
    + '<p style="margin:0 0 18px;">\uC218\uC2DC \uD559\uB825\uC9C4\uB2E8\uD3C9\uAC00 \uC2E0\uCCAD\uC774<br>\uC811\uC218\uB418\uC5C8\uC2B5\uB2C8\uB2E4.</p>'
    + '<table style="border-collapse:collapse;width:100%;max-width:520px;font-size:14px;">' + tableRows + '</table>'
    + '<p style="margin:18px 0 0;">\uC790\uC138\uD55C \uB0B4\uC6A9\uC740 \uB4DC\uB9BC\uD50C\uB7EC\uC2A4 \uB610\uB294 \uD648\uD398\uC774\uC9C0\uC5D0\uC11C<br>\uD655\uC778\uD558\uC2DC\uAE30 \uBC14\uB78D\uB2C8\uB2E4.</p>'
    + '</div>';
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function normalizeEmailList_(value) {
  return cellText_(value)
    .split(/[;,]/)
    .map(function (email) { return email.trim(); })
    .filter(function (email) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email); })
    .filter(function (email, index, list) { return list.indexOf(email) === index; })
    .join(',');
}

function ensureNamedColumn_(sheet, headerName) {
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(cellText_);
  var index = headers.indexOf(headerName);
  if (index >= 0) return index + 1;

  var newColumn = lastColumn + 1;
  sheet.getRange(1, newColumn).setValue(headerName);
  return newColumn;
}

function isDuplicateApplication_(sheet, payload) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  var firstRow = Math.max(2, lastRow - DUPLICATE_CHECK_ROW_LIMIT + 1);
  var rows = sheet.getRange(firstRow, 1, lastRow - firstRow + 1, 16).getValues();
  var phone = normalizeText_(payload.phone);
  var campusName = normalizeText_(payload.campusName);
  var preferredDate = normalizeText_(payload.preferredDate);
  var preferredTime = normalizeText_(payload.preferredTime);

  for (var i = 0; i < rows.length; i++) {
    var rowPhone = normalizeText_(rows[i][7]);
    var rowCampusName = normalizeText_(rows[i][3]);
    var rowDate = normalizeText_(rows[i][8]);
    var rowTime = normalizeText_(rows[i][9]);

    if (
      rowPhone === phone &&
      rowCampusName === campusName &&
      rowDate === preferredDate &&
      rowTime === preferredTime
    ) {
      return true;
    }
  }

  return false;
}

function getApplicationHeaders_() {
  return [
    '접수일시',
    '학생명',
    '지역',
    '캠퍼스명',
    '캠퍼스 위치',
    '학교명',
    '학년',
    '연락처',
    '희망날짜',
    '희망시간',
    '전형과목',
    '알게된 경로',
    '알게된 경로 내용',
    'campusId',
    'subjectId',
    '입력사이트',
    '사이트전송상태',
    '처리상태',
    'utm_source',
    'utm_medium',
    'utm_campaign'
  ];
}

function getSubmissionLogHeaders_() {
  return [
    '기록일시',
    '신청내역 행',
    '전송상태',
    '캠퍼스명',
    '전형과목',
    '입력사이트',
    '전송예정데이터',
    '메시지'
  ];
}

function getCampusHeaders_() {
  return [
    'campusId',
    '노출순서',
    '구분',
    '지역',
    '캠퍼스명',
    '전형과목',
    '입력사이트',
    '이동사이트',
    'SITE',
    'CONSULTING_IDX',
    'ACAD_ID',
    'AREA_ID',
    'AREA_NAME',
    'ACAD_NAME',
    '특이사항',
    '토요일선택불가'
  ];
}

function getBasicScheduleHeaders_() {
  return ['캠퍼스', '요일', '과목', '학년군', '상태', '가능시간', '비고'];
}

function getSpecialScheduleHeaders_() {
  return ['날짜', '캠퍼스', '과목', '학년군', '상태', '가능시간', '비고'];
}

function getSettingsDefinitions_() {
  return [
    {
      key: 'APPLICATION_SPREADSHEET_ID',
      defaultValue: '',
      description: '신청내역을 저장하는 랜딩페이지_신청 스프레드시트 ID'
    },
    {
      key: 'CAMPUS_SPREADSHEET_ID',
      defaultValue: '',
      description: '현재 랜딩페이지_정보 스프레드시트 ID (참고용)'
    },
    {
      key: 'MAIL_SENDER_NAME',
      defaultValue: '올림피아드교육',
      description: '캠퍼스 안내 메일에 표시할 발신자명'
    },
    {
      key: 'PENDING_SUBMIT_LIMIT',
      defaultValue: '10',
      description: '1회 처리할 전송대기 신청 건수'
    },
    {
      key: 'SITE_SUBMIT_MODE',
      defaultValue: TEST_MODE_LABEL,
      description: '테스트는 실제 사이트에 전송하지 않으며, 라이브만 실제 전송'
    },
    {
      key: 'TEST_EMAIL_SEND',
      defaultValue: TEST_EMAIL_DISABLED_LABEL,
      description: '테스트 운영일 때 캠퍼스 담당자에게 테스트 메일을 보낼지 선택'
    }
  ];
}

function setupSettingsSheet_(removeLegacyProperties) {
  var ss = getSettingsSpreadsheet_();
  var sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
  }

  var legacyProperties = PropertiesService.getScriptProperties().getProperties();
  var existingSettings = readSettingsFromSheet_(sheet);
  var definitions = getSettingsDefinitions_();
  var definitionByKey = {};
  var keys = [];

  definitions.forEach(function (definition) {
    definitionByKey[definition.key] = definition;
    keys.push(definition.key);
  });

  Object.keys(legacyProperties).forEach(function (key) {
    if (keys.indexOf(key) === -1) keys.push(key);
  });

  var rows = keys.map(function (key) {
    var definition = definitionByKey[key] || {
      defaultValue: '',
      description: '기존 스크립트 속성에서 이전된 설정'
    };
    var value = existingSettings[key];

    if (value === undefined || value === '') {
      value = legacyProperties[key] !== undefined
        ? legacyProperties[key]
        : definition.defaultValue;
    }

    if (key === 'CAMPUS_SPREADSHEET_ID') {
      value = ss.getId();
    }
    if (key === 'SITE_SUBMIT_MODE') {
      value = toSiteSubmitModeLabel_(value);
    }
    if (key === 'TEST_EMAIL_SEND') {
      value = toTestEmailSendLabel_(value);
    }

    return [key, value, definition.description];
  });

  sheet.getRange(1, 1, 1, 3).setValues([['설정항목', '설정값', '설명']]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  if (sheet.getLastRow() > rows.length + 1) {
    sheet
      .getRange(rows.length + 2, 1, sheet.getLastRow() - rows.length - 1, 3)
      .clearContent();
  }

  formatSettingsSheet_(sheet, rows);

  var legacyPropertiesRemoved = false;
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (
    removeLegacyProperties &&
    activeSpreadsheet &&
    activeSpreadsheet.getId() === ss.getId()
  ) {
    PropertiesService.getScriptProperties().deleteAllProperties();
    legacyPropertiesRemoved = true;
  }

  return {
    ok: true,
    spreadsheetName: ss.getName(),
    sheetName: SETTINGS_SHEET_NAME,
    settingsCount: rows.length,
    legacyPropertiesRemoved: legacyPropertiesRemoved,
    message: legacyPropertiesRemoved
      ? '설정 시트로 이전하고 기존 스크립트 속성을 삭제했습니다.'
      : '설정 시트를 만들었습니다. Apps Script가 랜딩페이지_정보에 연결된 경우에만 기존 스크립트 속성을 자동 삭제합니다.'
  };
}

function formatSettingsSheet_(sheet, rows) {
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 230);
  sheet.setColumnWidth(2, 330);
  sheet.setColumnWidth(3, 500);
  sheet
    .getRange(1, 1, 1, 3)
    .setBackground('#17213f')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  if (rows.length > 0) {
    sheet
      .getRange(2, 1, rows.length, 3)
      .setVerticalAlignment('middle')
      .setWrap(true);
  }

  var modeRow = -1;
  var testEmailRow = -1;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] === 'SITE_SUBMIT_MODE') {
      modeRow = i + 2;
    }
    if (rows[i][0] === 'TEST_EMAIL_SEND') {
      testEmailRow = i + 2;
    }
  }

  var conditionalRules = [];
  if (modeRow >= 0) {
    var modeCell = sheet.getRange(modeRow, 2);
    var modeValidation = SpreadsheetApp
      .newDataValidation()
      .requireValueInList([TEST_MODE_LABEL, LIVE_MODE_LABEL], true)
      .setAllowInvalid(false)
      .build();
    modeCell.setDataValidation(modeValidation).setFontWeight('bold');

    conditionalRules.push(
      SpreadsheetApp
        .newConditionalFormatRule()
        .whenTextEqualTo(TEST_MODE_LABEL)
        .setBackground('#e8f1ff')
        .setFontColor('#174ea6')
        .setRanges([modeCell])
        .build()
    );
    conditionalRules.push(
      SpreadsheetApp
        .newConditionalFormatRule()
        .whenTextEqualTo(LIVE_MODE_LABEL)
        .setBackground('#ffe2e2')
        .setFontColor('#b3261e')
        .setBold(true)
        .setRanges([modeCell])
        .build()
    );
  }

  if (testEmailRow >= 0) {
    var testEmailCell = sheet.getRange(testEmailRow, 2);
    var testEmailValidation = SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        [TEST_EMAIL_DISABLED_LABEL, TEST_EMAIL_ENABLED_LABEL],
        true
      )
      .setAllowInvalid(false)
      .build();
    testEmailCell.setDataValidation(testEmailValidation).setFontWeight('bold');

    conditionalRules.push(
      SpreadsheetApp
        .newConditionalFormatRule()
        .whenTextEqualTo(TEST_EMAIL_DISABLED_LABEL)
        .setBackground('#f1f3f4')
        .setFontColor('#5f6368')
        .setRanges([testEmailCell])
        .build()
    );
    conditionalRules.push(
      SpreadsheetApp
        .newConditionalFormatRule()
        .whenTextEqualTo(TEST_EMAIL_ENABLED_LABEL)
        .setBackground('#fff1d6')
        .setFontColor('#a14200')
        .setBold(true)
        .setRanges([testEmailCell])
        .build()
    );
  }

  sheet.setConditionalFormatRules(conditionalRules);
}

function readSettingsFromSheet_(sheet) {
  var settings = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return settings;

  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  values.forEach(function (row) {
    var key = cellText_(row[0]);
    if (key) settings[key] = cellText_(row[1]);
  });
  return settings;
}

function getSetting_(key) {
  var ss = getSettingsSpreadsheet_();
  var sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (sheet) {
    var settings = readSettingsFromSheet_(sheet);
    if (settings[key] !== undefined) {
      return settings[key];
    }
  }
  return getLegacyScriptProperty_(key);
}

function getSettingsSpreadsheet_() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (
    activeSpreadsheet &&
    (
      activeSpreadsheet.getName() === SETTINGS_SPREADSHEET_NAME ||
      activeSpreadsheet.getSheetByName(SETTINGS_SHEET_NAME) ||
      activeSpreadsheet.getSheetByName(CAMPUS_SHEET_NAME)
    )
  ) {
    CacheService
      .getScriptCache()
      .put(SETTINGS_SPREADSHEET_CACHE_KEY, activeSpreadsheet.getId(), 21600);
    return activeSpreadsheet;
  }

  var scriptCache = CacheService.getScriptCache();
  var cachedSpreadsheetId = scriptCache.get(SETTINGS_SPREADSHEET_CACHE_KEY);
  if (cachedSpreadsheetId) {
    try {
      return SpreadsheetApp.openById(cachedSpreadsheetId);
    } catch (cacheErr) {
      scriptCache.remove(SETTINGS_SPREADSHEET_CACHE_KEY);
    }
  }

  var legacySpreadsheetId = getLegacyScriptProperty_('CAMPUS_SPREADSHEET_ID');
  if (legacySpreadsheetId) {
    scriptCache.put(SETTINGS_SPREADSHEET_CACHE_KEY, legacySpreadsheetId, 21600);
    return SpreadsheetApp.openById(legacySpreadsheetId);
  }

  var namedSpreadsheet = findSettingsSpreadsheetByName_();
  if (namedSpreadsheet) {
    scriptCache.put(SETTINGS_SPREADSHEET_CACHE_KEY, namedSpreadsheet.getId(), 21600);
    return namedSpreadsheet;
  }

  throw new Error(
    'Apps Script를 랜딩페이지_정보 스프레드시트에 연결한 뒤 migrateScriptPropertiesToSettings 함수를 실행해 주세요.'
  );
}

function findSettingsSpreadsheetByName_() {
  var files = DriveApp.getFilesByName(SETTINGS_SPREADSHEET_NAME);

  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    var spreadsheet = SpreadsheetApp.openById(file.getId());
    if (
      spreadsheet.getSheetByName(SETTINGS_SHEET_NAME) &&
      spreadsheet.getSheetByName(CAMPUS_SHEET_NAME)
    ) {
      return spreadsheet;
    }
  }

  return null;
}

function getLegacyScriptProperty_(key) {
  return String(PropertiesService.getScriptProperties().getProperty(key) || '').trim();
}

function toSiteSubmitModeLabel_(value) {
  var normalized = normalizeText_(value).toUpperCase();
  if (
    normalized === 'LIVE' ||
    normalized === normalizeText_(LIVE_MODE_LABEL).toUpperCase() ||
    normalized === '라이브' ||
    normalized === '라이브운영'
  ) {
    return LIVE_MODE_LABEL;
  }
  return TEST_MODE_LABEL;
}

function toTestEmailSendLabel_(value) {
  var normalized = normalizeText_(value).toUpperCase();
  if (
    normalized === 'Y' ||
    normalized === 'YES' ||
    normalized === 'TRUE' ||
    normalized === 'SEND' ||
    normalized === normalizeText_(TEST_EMAIL_ENABLED_LABEL).toUpperCase() ||
    normalized === '발송' ||
    normalized === '테스트메일발송'
  ) {
    return TEST_EMAIL_ENABLED_LABEL;
  }
  return TEST_EMAIL_DISABLED_LABEL;
}

function getCampusSpreadsheet_() {
  return getSettingsSpreadsheet_();
}

function getApplicationSpreadsheet_() {
  var spreadsheetId = getSetting_('APPLICATION_SPREADSHEET_ID');
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }
  throw new Error('설정 시트의 APPLICATION_SPREADSHEET_ID 값이 비어 있습니다.');
}

function getSiteSubmitMode_() {
  var value = normalizeText_(getSetting_('SITE_SUBMIT_MODE')).toUpperCase();
  if (
    value === 'LIVE' ||
    value === normalizeText_(LIVE_MODE_LABEL).toUpperCase() ||
    value === '라이브' ||
    value === '라이브운영'
  ) {
    return 'LIVE';
  }
  return 'DRY_RUN';
}

function shouldSendTestEmail_() {
  return toTestEmailSendLabel_(getSetting_('TEST_EMAIL_SEND'))
    === TEST_EMAIL_ENABLED_LABEL;
}

function getOrCreateCampusSheet_() {
  return getOrCreateSheet_(getCampusSpreadsheet_(), CAMPUS_SHEET_NAME);
}

function getOrCreateApplicationSheet_() {
  return getOrCreateSheet_(getApplicationSpreadsheet_(), APPLICATION_SHEET_NAME);
}

function getOrCreateApplicationNamedSheet_(name) {
  return getOrCreateSheet_(getApplicationSpreadsheet_(), name);
}

function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function ensureHeader_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return;
  }

  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var current = sheet.getRange(1, 1, 1, Math.max(lastColumn, headers.length)).getValues()[0];
  if (current.join('') === '') {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  headers.forEach(function (header) {
    if (current.map(cellText_).indexOf(header) === -1) {
      var newColumn = sheet.getLastColumn() + 1;
      sheet.getRange(1, newColumn).setValue(header);
      current.push(header);
    }
  });
}

function normalizeApplicationSheet_(sheet) {
  renameHeaderIfMissing_(sheet, '희망시', '희망시간');
  deleteColumnsByHeader_(sheet, ['특이사항', '희망시']);
}

function normalizeSubmissionLogSheet_(sheet) {
  deleteColumnsByHeader_(sheet, ['특이사항']);
}

function renameHeaderIfMissing_(sheet, oldHeader, newHeader) {
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) return;

  var headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(cellText_);
  var oldIndex = headers.indexOf(oldHeader);
  var newIndex = headers.indexOf(newHeader);

  if (oldIndex >= 0 && newIndex < 0) {
    sheet.getRange(1, oldIndex + 1).setValue(newHeader);
  }
}

function deleteColumnsByHeader_(sheet, headersToDelete) {
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) return;

  var headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(cellText_);
  var columnIndexes = [];

  headers.forEach(function (header, index) {
    if (headersToDelete.indexOf(header) >= 0) {
      columnIndexes.push(index + 1);
    }
  });

  columnIndexes
    .sort(function (a, b) { return b - a; })
    .forEach(function (columnIndex) {
      sheet.deleteColumn(columnIndex);
    });
}

function requireField_(payload, key, label) {
  if (!payload[key]) {
    throw new Error(label + ' 값이 없습니다.');
  }
}

function normalizeText_(value) {
  return String(value == null ? '' : value).replace(/\s+/g, '').trim();
}

function onlyDigits_(value) {
  return String(value == null ? '' : value).replace(/\D/g, '');
}

function getOrigin_(url) {
  var match = String(url || '').match(/^(https?:\/\/[^\/]+)/i);
  return match ? match[1] : '';
}

function cellText_(value) {
  return String(value == null ? '' : value).trim();
}

function findHeaderIndex_(headers, name) {
  var index = headers.indexOf(name);
  return index >= 0 ? index : -1;
}

function findAllHeaderIndexes_(headers, name) {
  var indexes = [];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === name) indexes.push(i);
  }
  return indexes;
}

function splitSubjectNames_(value) {
  var text = cellText_(value).replace(/\s+/g, '');
  if (!text) return [];

  if (text === '수학+영어' || text === '영어+수학') {
    return ['수학+영어'];
  }

  return text
    .split(/[,/·ㆍ]+/g)
    .map(function (name) { return name.trim(); })
    .filter(function (name) { return name === '수학' || name === '영어' || name === '수학+영어'; });
}

function makeSubjectId_(name) {
  if (name === '수학') return 'math';
  if (name === '영어') return 'english';
  if (name === '수학+영어') return 'math_english';
  return '';
}

function makeCampusId_(name) {
  var clean = cellText_(name).replace(/\s+/g, '');
  if (clean.indexOf('송파') >= 0 || clean.indexOf('방이') >= 0) return 'songpa';
  if (clean.indexOf('미사') >= 0 || clean.indexOf('하남') >= 0) return 'misa';
  return encodeURIComponent(clean).replace(/%/g, '').toLowerCase();
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
