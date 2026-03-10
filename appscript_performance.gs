// ========== 설정 ==========
var PIPEDRIVE_API_KEY = 'cbc419beec83e32e7c9c50ab815eb0ab0508ea80';
var CUSTOM_FIELD_FIRST_PAYMENT_NOTICE = 'b24a25502fdeb48ac55986536dd8d449fe1ec494';
var CUSTOM_FIELD_COLLECTION_ORDER = '69b948cd5331d1c78a7d4045dae1be38be7a7177';
var CUSTOM_FIELD_BALANCE = '48c83ab2832b3e470154899dc3db5510221d321b';
var CUSTOM_FIELD_CUSTOMER_TYPE = '0ec37f587ba626b05d5db916d9e2f185e47f1abc';

var FIELD_REFUND_AMOUNT = '18f5d8f72f30db7d6abdc4aa862f64b9cb96409b';
var FIELD_APPLY_DATE = 'd63b4b92c9490208976c2fdd430cb55ee558b15e';
var FIELD_JUPJUP_PERSON = '3872c9f450beec097bbb5db2dff1ae118684d765';
var FIELD_DEFENSE_DATE = 'f362d92f24f82d2b27ecc093602489cf134170cd';
var FIELD_DEFENSE_PERSON = '32409a838adc8343885fc0c82fa4631d6a5087b9';

var FILTER_APPLY_SUCCESS = 1430754;
var FILTER_DEFENSE_SUCCESS = 1430989;
var FILTER_APPLY_ACTIVITY = 1431275;
var FILTER_DEFENSE_ACTIVITY = 1431330;

var CACHE_TTL = 21600; // 6시간

// ========== 시트 관리 헬퍼 ==========

function getOrCreateSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, 15).setValues([[
      'deal_id', 'title', 'value', 'first_payment_notice', 'won_time',
      'collection_order_date', 'balance', 'stage_name', 'customer_type',
      'add_time', 'apply_date', 'jupjup_person', 'defense_date', 'defense_person', 'refund_amount'
    ]]);
  }
  return sheet;
}

function getExistingDealIds(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var map = {};
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0]) map[ids[i][0]] = true;
  }
  return map;
}

function dealToRow(deal, stageMap) {
  return [
    deal.id,
    deal.title,
    deal.value || 0,
    deal[CUSTOM_FIELD_FIRST_PAYMENT_NOTICE] || '',
    deal.won_time || '',
    deal[CUSTOM_FIELD_COLLECTION_ORDER] || '',
    deal[CUSTOM_FIELD_BALANCE] || 0,
    (stageMap && stageMap[deal.stage_id]) ? stageMap[deal.stage_id] : '',
    deal[CUSTOM_FIELD_CUSTOMER_TYPE] || '',
    deal.add_time || '',
    deal[FIELD_APPLY_DATE] || '',
    deal[FIELD_JUPJUP_PERSON] ? String(deal[FIELD_JUPJUP_PERSON]) : '',
    deal[FIELD_DEFENSE_DATE] || '',
    deal[FIELD_DEFENSE_PERSON] ? String(deal[FIELD_DEFENSE_PERSON]) : '',
    parseFloat(deal[FIELD_REFUND_AMOUNT]) || 0
  ];
}

function writeDealsToSheet(sheet, deals, stageMap) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 15).clearContent();
  }
  if (deals.length === 0) return;
  var data = [];
  for (var i = 0; i < deals.length; i++) {
    data.push(dealToRow(deals[i], stageMap));
  }
  sheet.getRange(2, 1, data.length, 15).setValues(data);
}

function appendDealsToSheet(sheet, deals, stageMap) {
  if (deals.length === 0) return;
  var lastRow = sheet.getLastRow();
  var data = [];
  for (var i = 0; i < deals.length; i++) {
    data.push(dealToRow(deals[i], stageMap));
  }
  sheet.getRange(lastRow + 1, 1, data.length, 15).setValues(data);
}

function isRelevantDeal(deal) {
  return (deal[CUSTOM_FIELD_FIRST_PAYMENT_NOTICE] && deal[CUSTOM_FIELD_FIRST_PAYMENT_NOTICE] !== '') ||
         (deal[FIELD_APPLY_DATE] && deal[FIELD_APPLY_DATE] !== '') ||
         (deal[FIELD_DEFENSE_DATE] && deal[FIELD_DEFENSE_DATE] !== '');
}

// ========== 동기화 ==========

// 전체 동기화 (최초 1회 또는 수동 실행)
function syncPipedriveData() {
  var stageMap = fetchStageMap();

  // Won deals → 결제완료 시트
  var wonDeals = fetchDealsByStatus('won');
  var wonFiltered = [];
  for (var i = 0; i < wonDeals.length; i++) {
    if (isRelevantDeal(wonDeals[i])) wonFiltered.push(wonDeals[i]);
  }
  var wonSheet = getOrCreateSheet('결제완료');
  writeDealsToSheet(wonSheet, wonFiltered, stageMap);
  Logger.log('결제완료: ' + wonFiltered.length + '건');

  // Open deals → 미결제 시트
  var openDeals = fetchDealsByStatus('open');
  var openFiltered = [];
  for (var j = 0; j < openDeals.length; j++) {
    if (isRelevantDeal(openDeals[j])) openFiltered.push(openDeals[j]);
  }
  var openSheet = getOrCreateSheet('미결제');
  writeDealsToSheet(openSheet, openFiltered, stageMap);
  Logger.log('미결제: ' + openFiltered.length + '건');

  // 캐시 워밍업
  try {
    var json = buildPaymentJson();
    cachePaymentJson(json);
    Logger.log('캐시 워밍업 완료 (' + json.length + ' bytes)');
  } catch(e) {
    Logger.log('캐시 워밍업 실패: ' + e.message);
  }

  Logger.log('전체 동기화 완료');
}

// 배치 업데이트 (빠른 증분 동기화 - 트리거 전용)
// Open 거래만 다시 가져오고, 새로 성사된 거래는 자동으로 결제완료로 이동
function syncBatchUpdate() {
  var stageMap = fetchStageMap();

  // 1. 이전 미결제 ID 목록
  var openSheet = getOrCreateSheet('미결제');
  var prevOpenIds = getExistingDealIds(openSheet);

  // 2. 현재 Open 거래 조회 (이것만 Pipedrive에서 가져옴 → 빠름!)
  var openDeals = fetchDealsByStatus('open');
  var openFiltered = [];
  var currentOpenIds = {};
  for (var i = 0; i < openDeals.length; i++) {
    if (isRelevantDeal(openDeals[i])) {
      openFiltered.push(openDeals[i]);
      currentOpenIds[openDeals[i].id] = true;
    }
  }

  // 3. 미결제에서 사라진 거래 → 성사 여부 확인
  var wonSheet = getOrCreateSheet('결제완료');
  var wonIds = getExistingDealIds(wonSheet);
  var newWonDeals = [];

  var prevKeys = Object.keys(prevOpenIds);
  for (var k = 0; k < prevKeys.length; k++) {
    var id = prevKeys[k];
    if (currentOpenIds[id] || wonIds[id]) continue;
    // 이 거래가 성사되었는지 개별 확인
    var deal = fetchDealById(id);
    if (deal && deal.won_time && isRelevantDeal(deal)) {
      newWonDeals.push(deal);
      wonIds[deal.id] = true;
    }
    Utilities.sleep(100);
  }

  // 4. 새로 성사된 거래 → 결제완료 시트에 추가
  if (newWonDeals.length > 0) {
    appendDealsToSheet(wonSheet, newWonDeals, stageMap);
  }

  // 5. 미결제 시트 갱신
  writeDealsToSheet(openSheet, openFiltered, stageMap);

  // 6. 캐시 워밍업
  try {
    var json = buildPaymentJson();
    cachePaymentJson(json);
  } catch(e) {
    Logger.log('캐시 워밍업 실패: ' + e.message);
  }

  Logger.log('배치 완료: Open ' + openFiltered.length + '건, 신규Won ' + newWonDeals.length + '건');
}

// 배치 트리거 설정 (최초 1회 실행)
function setupBatchTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncBatchUpdate') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('syncBatchUpdate')
    .timeBased()
    .everyHours(2)
    .create();
  Logger.log('배치 트리거 설정 완료: 2시간마다 syncBatchUpdate 실행');
}

// ========== Pipedrive API ==========

function fetchStageMap() {
  var map = {};
  try {
    var url = 'https://api.pipedrive.com/v1/stages?api_token=' + PIPEDRIVE_API_KEY;
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(response.getContentText());
    if (json.success && json.data) {
      for (var i = 0; i < json.data.length; i++) {
        map[json.data[i].id] = json.data[i].name;
      }
    }
  } catch (e) {
    Logger.log('단계 조회 실패: ' + e.message);
  }
  return map;
}

function fetchDealsByStatus(status) {
  var deals = [];
  var start = 0;
  var limit = 500;
  var hasMore = true;

  while (hasMore) {
    var url = 'https://api.pipedrive.com/v1/deals?status=' + status + '&start=' + start + '&limit=' + limit + '&api_token=' + PIPEDRIVE_API_KEY;
    try {
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      var json = JSON.parse(response.getContentText());
      if (json.success && json.data) {
        deals = deals.concat(json.data);
        hasMore = (json.additional_data && json.additional_data.pagination && json.additional_data.pagination.more_items_in_collection) || false;
        start += limit;
        Utilities.sleep(200);
      } else {
        hasMore = false;
      }
    } catch (e) {
      Logger.log('API 오류 (' + status + '): ' + e.message);
      hasMore = false;
    }
  }
  return deals;
}

function fetchDealById(dealId) {
  try {
    var url = 'https://api.pipedrive.com/v1/deals/' + dealId + '?api_token=' + PIPEDRIVE_API_KEY;
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(resp.getContentText());
    return (json.success && json.data) ? json.data : null;
  } catch(e) {
    return null;
  }
}

function fetchDealsByFilter(filterId) {
  var deals = [];
  var start = 0;
  var limit = 500;
  var hasMore = true;

  while (hasMore) {
    var url = 'https://api.pipedrive.com/v1/deals?filter_id=' + filterId + '&start=' + start + '&limit=' + limit + '&api_token=' + PIPEDRIVE_API_KEY;
    try {
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      var json = JSON.parse(response.getContentText());
      if (json.success && json.data) {
        deals = deals.concat(json.data);
        hasMore = (json.additional_data && json.additional_data.pagination && json.additional_data.pagination.more_items_in_collection) || false;
        start += limit;
        Utilities.sleep(200);
      } else {
        hasMore = false;
      }
    } catch (e) {
      Logger.log('API 오류 (filter ' + filterId + '): ' + e.message);
      hasMore = false;
    }
  }
  return deals;
}

// ========== 결제 데이터 API ==========

function processSheetToArrays(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var rows = sheet.getDataRange().getValues().slice(1);
  var result = [];

  var toStr = function(v) {
    if (!v || v === '') return '';
    if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
    return String(v);
  };
  var getYear = function(v) {
    if (!v || v === '') return 0;
    if (v instanceof Date) return v.getFullYear();
    var s = String(v);
    var y = parseInt(s.substring(0, 4));
    return isNaN(y) ? 0 : y;
  };

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var noticeYear = getYear(row[3]);
    var wonYear = getYear(row[4]);
    if (noticeYear > 0 && noticeYear < 2025 && wonYear > 0 && wonYear < 2025) continue;
    // [deal_id, title, value, notice, won, collection, balance, stage, customer_type]
    result.push([
      row[0], row[1], Number(row[2]) || 0,
      toStr(row[3]), toStr(row[4]), toStr(row[5]),
      Number(row[6]) || 0, row[7] || '', row[8] || ''
    ]);
  }
  return result;
}

function buildPaymentJson() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wonSheet = ss.getSheetByName('결제완료');
  var openSheet = ss.getSheetByName('미결제');

  var allRows;
  if (wonSheet && openSheet) {
    allRows = processSheetToArrays(wonSheet).concat(processSheetToArrays(openSheet));
  } else {
    var oldSheet = ss.getSheetByName('결제데이터');
    allRows = oldSheet ? processSheetToArrays(oldSheet) : [];
  }

  // cols+rows 형식: 필드명 반복 없어서 응답 크기 40-50% 감소
  return JSON.stringify({
    success: true,
    cols: ['deal_id','title','value','first_payment_notice','won_time','collection_order_date','balance','stage_name','customer_type'],
    rows: allRows
  });
}

function cachePaymentJson(json) {
  var cache = CacheService.getScriptCache();
  try {
    cache.remove('paymentData');
    cache.remove('paymentData_chunks');
    for (var r = 0; r < 20; r++) cache.remove('paymentData_' + r);
  } catch(e) {}

  try {
    cache.put('paymentData', json, CACHE_TTL);
  } catch (e) {
    var chunkSize = 90000;
    var chunks = Math.ceil(json.length / chunkSize);
    var cacheMap = { 'paymentData_chunks': String(chunks) };
    for (var i = 0; i < chunks; i++) {
      cacheMap['paymentData_' + i] = json.substring(i * chunkSize, (i + 1) * chunkSize);
    }
    cache.putAll(cacheMap, CACHE_TTL);
  }
}

function getPaymentData() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('paymentData');
  if (!cached) {
    var chunksStr = cache.get('paymentData_chunks');
    if (chunksStr) {
      var numChunks = parseInt(chunksStr);
      var keys = [];
      for (var ci = 0; ci < numChunks; ci++) keys.push('paymentData_' + ci);
      var parts = cache.getAll(keys);
      cached = '';
      for (var ci2 = 0; ci2 < numChunks; ci2++) {
        cached += (parts['paymentData_' + ci2] || '');
      }
    }
  }

  if (cached) {
    return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
  }

  var json = buildPaymentJson();
  cachePaymentJson(json);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ========== 필터 기반 성과 요약 ==========

function calculatePerformanceByFilter() {
  Logger.log('=== 필터 기반 성과 계산 ===');

  var applyDeals = fetchDealsByFilter(FILTER_APPLY_SUCCESS);
  var applyCount = 0, applyAmount = 0;
  for (var i = 0; i < applyDeals.length; i++) {
    var ra = parseFloat(applyDeals[i][FIELD_REFUND_AMOUNT]) || 0;
    if (ra > 0) { applyCount++; applyAmount += ra; }
  }
  Logger.log('신청전환: ' + applyDeals.length + '건 조회, ' + applyCount + '건 환급액');

  var defenseDeals = fetchDealsByFilter(FILTER_DEFENSE_SUCCESS);
  var defenseCount = 0, defenseAmount = 0;
  for (var j = 0; j < defenseDeals.length; j++) {
    var rd = parseFloat(defenseDeals[j][FIELD_REFUND_AMOUNT]) || 0;
    if (rd > 0) { defenseCount++; defenseAmount += rd; }
  }
  Logger.log('취소방어: ' + defenseDeals.length + '건 조회, ' + defenseCount + '건 환급액');

  return {
    apply: { count: applyCount, amount: applyAmount },
    defense: { count: defenseCount, amount: defenseAmount },
    total: applyAmount + defenseAmount
  };
}

function getPerformanceData(year, month) {
  if (year && month) {
    updateFilterToMonth(FILTER_APPLY_SUCCESS, year, month);
    updateFilterToMonth(FILTER_DEFENSE_SUCCESS, year, month);
  }
  var result = calculatePerformanceByFilter();
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, data: result }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========== 필터 날짜 변경 (조건 보존) ==========

function updateFilterToMonth(filterId, year, month) {
  try {
    var getUrl = 'https://api.pipedrive.com/v1/filters/' + filterId + '?api_token=' + PIPEDRIVE_API_KEY;
    var getResp = UrlFetchApp.fetch(getUrl, { muteHttpExceptions: true });
    var filterData = JSON.parse(getResp.getContentText());
    if (!filterData.success) return;

    var monthStr = (month < 10 ? '0' + month : '' + month);
    var nextMonth = month + 1, nextYear = year;
    if (nextMonth > 12) { nextMonth = 1; nextYear = year + 1; }
    var nextMonthStr = (nextMonth < 10 ? '0' + nextMonth : '' + nextMonth);
    var startDate = year + '-' + monthStr + '-01';
    var endDate = nextYear + '-' + nextMonthStr + '-01';

    var conditions = filterData.data.conditions;
    if (!conditions.conditions || !conditions.conditions[0]) return;

    var firstGroup = conditions.conditions[0].conditions;
    var newGroup = [];
    var dateReplaced = false;

    for (var i = 0; i < firstGroup.length; i++) {
      var c = firstGroup[i];
      var isDateCond = (c.value === 'this_month' || c.value === 'last_month' || c.operator === '>' || c.operator === '<');
      if (isDateCond && !dateReplaced) {
        var obj = c.object || 'deal';
        var fid = c.field_id;
        newGroup.push({ object: obj, field_id: fid, operator: '>', value: startDate, extra_value: null });
        newGroup.push({ object: obj, field_id: fid, operator: '<', value: endDate, extra_value: null });
        dateReplaced = true;
      } else if (isDateCond && dateReplaced) {
        // skip
      } else {
        newGroup.push(c);
      }
    }
    conditions.conditions[0].conditions = newGroup;

    var putUrl = 'https://api.pipedrive.com/v1/filters/' + filterId + '?api_token=' + PIPEDRIVE_API_KEY;
    UrlFetchApp.fetch(putUrl, {
      method: 'put', contentType: 'application/json',
      payload: JSON.stringify({ conditions: conditions }), muteHttpExceptions: true
    });
    Logger.log('필터 ' + filterId + ' → ' + year + '-' + monthStr);
  } catch (e) { Logger.log('필터 업데이트 실패: ' + e.message); }
}

// ========== 활동수 현황 ==========

function fetchActivitiesByFilter(filterId) {
  var activities = [];
  var cursor = null;
  var hasMore = true;

  while (hasMore) {
    var url = 'https://api.pipedrive.com/api/v2/activities?filter_id=' + filterId + '&limit=500&api_token=' + PIPEDRIVE_API_KEY;
    if (cursor) url += '&cursor=' + cursor;

    try {
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      var json = JSON.parse(response.getContentText());
      if (json.success && json.data) {
        activities = activities.concat(json.data);
        cursor = (json.additional_data && json.additional_data.next_cursor) ? json.additional_data.next_cursor : null;
        hasMore = cursor !== null;
        Utilities.sleep(200);
      } else {
        hasMore = false;
      }
    } catch (e) {
      Logger.log('Activities API 오류: ' + e.message);
      hasMore = false;
    }
  }
  return activities;
}

function classifyActivitySubject(subject) {
  if (!subject) return null;
  var s = subject.toLowerCase();
  if (s.indexOf('줍줍콜') >= 0 && s.indexOf('부재') >= 0) return 'absent';
  if (s.indexOf('줍줍콜') >= 0 || s.indexOf('vip') >= 0) return 'activity';
  if (s.indexOf('사후관리') >= 0) return 'followup';
  if (s.indexOf('문자발송') >= 0 || s.indexOf('문자안내') >= 0) return 'sms';
  return 'other';
}

function classifyDefenseActivitySubject(subject) {
  if (!subject) return null;
  var s = subject.toLowerCase();
  if (s.indexOf('취소방어') >= 0 && s.indexOf('부재') >= 0) return 'absent';
  if (s.indexOf('취소방어') >= 0) return 'activity';
  if (s.indexOf('사후관리') >= 0) return 'followup';
  if (s.indexOf('문자') >= 0) return 'sms';
  return 'other';
}

function calculateApplyActivityCount() {
  var activities = fetchActivitiesByFilter(FILTER_APPLY_ACTIVITY);
  var counts = { activity: 0, absent: 0, followup: 0, sms: 0, other: 0 };
  for (var i = 0; i < activities.length; i++) {
    var cat = classifyActivitySubject(activities[i].subject || '');
    if (cat && counts.hasOwnProperty(cat)) counts[cat]++;
  }
  Logger.log('신청전환 활동: ' + JSON.stringify(counts));
  return counts;
}

function calculateDefenseActivityCount() {
  var activities = fetchActivitiesByFilter(FILTER_DEFENSE_ACTIVITY);
  var counts = { activity: 0, absent: 0, followup: 0, sms: 0, other: 0 };
  for (var i = 0; i < activities.length; i++) {
    var cat = classifyDefenseActivitySubject(activities[i].subject || '');
    if (cat && counts.hasOwnProperty(cat)) counts[cat]++;
  }
  Logger.log('취소방어 활동: ' + JSON.stringify(counts));
  return counts;
}

function getActivityData(year, month) {
  if (year && month) {
    updateFilterToMonth(FILTER_APPLY_ACTIVITY, year, month);
    updateFilterToMonth(FILTER_DEFENSE_ACTIVITY, year, month);
  }
  var applyActivity = calculateApplyActivityCount();
  var defenseActivity = calculateDefenseActivityCount();

  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      data: { apply: applyActivity, defense: defenseActivity }
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========== 수동 입력 데이터 ==========

function saveManualData(type, data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('수동입력');
  if (!sheet) return false;

  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  var values = sheet.getDataRange().getValues();
  var rowIndex = -1;

  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === today && values[i][1] === type) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 3).setValue(JSON.stringify(data));
  } else {
    sheet.appendRow([today, type, JSON.stringify(data)]);
  }
  return true;
}

function loadAllManualData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('수동입력');
  if (!sheet) return { coverage: null, application: null, defense: null, target: null };

  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  var values = sheet.getDataRange().getValues();
  var result = { coverage: null, application: null, defense: null, target: null };

  for (var i = 1; i < values.length; i++) {
    var rowDate = values[i][0];
    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, 'Asia/Seoul', 'yyyy-MM-dd');
    } else if (typeof rowDate === 'string') {
      rowDate = rowDate.trim();
    }
    if (rowDate === today && result.hasOwnProperty(values[i][1])) {
      try { result[values[i][1]] = JSON.parse(values[i][2]); } catch(e) {}
    }
  }
  return result;
}

// ========== 웹 앱 API ==========

function doGet(e) {
  var action = e.parameter.action || 'payment';

  if (action === 'payment') {
    return getPaymentData();
  } else if (action === 'performance') {
    var year = e.parameter.year ? parseInt(e.parameter.year) : null;
    var month = e.parameter.month ? parseInt(e.parameter.month) : null;
    return getPerformanceData(year, month);
  } else if (action === 'activity') {
    var aYear = e.parameter.year ? parseInt(e.parameter.year) : null;
    var aMonth = e.parameter.month ? parseInt(e.parameter.month) : null;
    return getActivityData(aYear, aMonth);
  } else if (action === 'manual') {
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, data: loadAllManualData() }))
      .setMimeType(ContentService.MimeType.JSON);
  } else if (action === 'saveManual') {
    return saveManualDataViaGet(e);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function saveManualDataViaGet(e) {
  try {
    var type = e.parameter.type;
    var dataStr = e.parameter.data;
    if (!type || !dataStr) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Missing params' })).setMimeType(ContentService.MimeType.JSON);
    }
    var data = JSON.parse(decodeURIComponent(dataStr));
    var result = saveManualData(type, data);
    return ContentService.createTextOutput(JSON.stringify({ success: result })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action === 'saveManual' && data.type && data.data) {
      var result = saveManualData(data.type, data.data);
      return ContentService.createTextOutput(JSON.stringify({ success: result })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Unknown action' })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}
