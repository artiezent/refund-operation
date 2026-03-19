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

// 서버 캐시 제거: 시트에서 직접 읽기 (항상 최신 데이터)

var SHEET_PAYMENT = '결제현황';
var SHEET_COLLECTION = '추심현황';
var SHEET_COVERAGE = '결제콜커버리지';

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

function classifyAndWrite(deals, stageMap, isWon) {
  var payRows = [], collRows = [];
  for (var i = 0; i < deals.length; i++) {
    var d = deals[i];
    var wy = parseInt(String(d.won_time || '').substring(0, 4)) || 0;
    var isPayment = (wy >= 2025 || isRelevantDeal(d));
    var isCollection = !!(d[CUSTOM_FIELD_COLLECTION_ORDER] && d[CUSTOM_FIELD_COLLECTION_ORDER] !== '');
    if (!isWon && !isRelevantDeal(d)) continue;
    var row = dealToRow(d, stageMap);
    if (isPayment) payRows.push(row);
    if (isCollection) collRows.push(row);
  }
  return { pay: payRows, coll: collRows };
}

function writeRowsToSheet(sheet, rows) {
  var cols = (rows.length > 0) ? rows[0].length : 15;
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, cols).clearContent();
  SpreadsheetApp.flush();
  if (rows.length === 0) return;
  var CHUNK = 5000;
  for (var s = 0; s < rows.length; s += CHUNK) {
    var chunk = rows.slice(s, s + CHUNK);
    sheet.getRange(s + 2, 1, chunk.length, cols).setValues(chunk);
    SpreadsheetApp.flush();
  }
}

function readSheetForSync(sheet) {
  var result = { wonRows: [], wonIds: {}, openIds: {} };
  if (!sheet || sheet.getLastRow() <= 1) return result;
  var rows = sheet.getDataRange().getValues().slice(1);
  for (var i = 0; i < rows.length; i++) {
    var id = rows[i][0];
    if (rows[i][4] && rows[i][4] !== '') {
      result.wonRows.push(rows[i]);
      result.wonIds[id] = true;
    } else {
      result.openIds[id] = true;
    }
  }
  return result;
}

// 전체 동기화 (최초 1회 또는 수동 실행)
function syncPipedriveData() {
  var stageMap = fetchStageMap();

  var wonDeals = fetchDealsByStatus('won');
  Logger.log('Won 조회: ' + wonDeals.length + '건');
  var wonResult = classifyAndWrite(wonDeals, stageMap, true);

  var openDeals = fetchDealsByStatus('open');
  Logger.log('Open 조회: ' + openDeals.length + '건');
  var openResult = classifyAndWrite(openDeals, stageMap, false);

  var paySheet = getOrCreateSheet(SHEET_PAYMENT);
  writeRowsToSheet(paySheet, wonResult.pay.concat(openResult.pay));
  Logger.log(SHEET_PAYMENT + ': ' + (wonResult.pay.length + openResult.pay.length) + '건');

  var collSheet = getOrCreateSheet(SHEET_COLLECTION);
  writeRowsToSheet(collSheet, wonResult.coll.concat(openResult.coll));
  Logger.log(SHEET_COLLECTION + ': ' + (wonResult.coll.length + openResult.coll.length) + '건');

  Logger.log('전체 동기화 완료');
}

// ========== 결제콜 커버리지 시트 ==========

var COVERAGE_COLS = [
  'deal_id', 'title', 'value', 'first_payment_notice', 'won_time',
  'days_to_payment', 'is_auto_paid', 'is_high_value',
  'has_payment_call', 'payment_call_date', 'payment_call_person',
  'has_hard_collection', 'hard_collection_date', 'hard_collection_person'
];

function getOrCreateCoverageSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_COVERAGE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_COVERAGE);
    sheet.getRange(1, 1, 1, COVERAGE_COLS.length).setValues([COVERAGE_COLS]);
    sheet.getRange(1, 1, 1, COVERAGE_COLS.length).setFontWeight('bold');
  }
  return sheet;
}

function syncCoverageData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var paySheet = ss.getSheetByName(SHEET_PAYMENT);
  if (!paySheet || paySheet.getLastRow() <= 1) {
    Logger.log('결제현황 시트가 비어있습니다. syncPipedriveData를 먼저 실행하세요.');
    return;
  }

  // 1. 결제현황에서 2026+ 알림톡 딜 추출
  var rows = paySheet.getDataRange().getValues().slice(1);
  var DAY_MS = 86400000;
  var deals = [];

  for (var i = 0; i < rows.length; i++) {
    var noticeStr = rows[i][3];
    if (!noticeStr || noticeStr === '') continue;
    var noticeDate = noticeStr instanceof Date ? noticeStr : new Date(noticeStr);
    if (isNaN(noticeDate.getTime()) || noticeDate.getFullYear() < 2026) continue;

    var wonStr = rows[i][4];
    var wonDate = (wonStr && wonStr !== '') ? (wonStr instanceof Date ? wonStr : new Date(wonStr)) : null;
    var value = Number(rows[i][2]) || 0;

    var daysToPayment = null;
    var isAutoPaid = false;
    if (wonDate && !isNaN(wonDate.getTime())) {
      var nd = new Date(noticeDate.getFullYear(), noticeDate.getMonth(), noticeDate.getDate());
      var wd = new Date(wonDate.getFullYear(), wonDate.getMonth(), wonDate.getDate());
      daysToPayment = Math.floor((wd - nd) / DAY_MS);
      isAutoPaid = (daysToPayment >= 0 && daysToPayment <= 2);
    }

    deals.push({
      dealId: rows[i][0],
      title: rows[i][1],
      value: value,
      noticeDate: noticeDate,
      noticeStr: Utilities.formatDate(noticeDate, 'Asia/Seoul', 'yyyy-MM-dd'),
      wonStr: wonDate ? Utilities.formatDate(wonDate, 'Asia/Seoul', 'yyyy-MM-dd') : '',
      daysToPayment: daysToPayment,
      isAutoPaid: isAutoPaid,
      isHighValue: (value >= 1000000)
    });
  }

  Logger.log('2026+ 알림톡 딜: ' + deals.length + '건, 자동결제: ' + deals.filter(function(d) { return d.isAutoPaid; }).length + '건');

  // 2. 사용자 목록 조회 (user_id → 이름 매핑)
  var userMap = {};
  try {
    var uResp = UrlFetchApp.fetch('https://api.pipedrive.com/v1/users?api_token=' + PIPEDRIVE_API_KEY, { muteHttpExceptions: true });
    var uJson = JSON.parse(uResp.getContentText());
    if (uJson.success && uJson.data) {
      for (var u = 0; u < uJson.data.length; u++) {
        userMap[uJson.data[u].id] = uJson.data[u].name;
      }
    }
    Logger.log('사용자 매핑: ' + Object.keys(userMap).length + '명');
  } catch(e) { Logger.log('사용자 조회 오류: ' + e.message); }

  // 3. 활동 조회: ☏ 결제 (_2) + 결제지연(강성) (code815199537)
  //    start_date=2025-07-01로 2026+ 알림톡 딜에 매칭될 활동만 조회
  var paymentCallMap = {};
  var hardCollectionMap = {};

  var types = [
    { type: '_2', map: paymentCallMap, label: '☏ 결제' },
    { type: 'code815199537', map: hardCollectionMap, label: '결제지연(강성)' }
  ];

  for (var t = 0; t < types.length; t++) {
    var start = 0, hasMore = true, pages = 0, total = 0;
    while (hasMore) {
      var url = 'https://api.pipedrive.com/v1/activities?type=' + types[t].type + '&done=1&user_id=0&start_date=2025-07-01&start=' + start + '&limit=500&api_token=' + PIPEDRIVE_API_KEY;
      try {
        var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        var json = JSON.parse(resp.getContentText());
        if (json.success && json.data) {
          for (var a = 0; a < json.data.length; a++) {
            var act = json.data[a];
            var dealId = act.deal_id;
            if (!dealId) continue;
            var dueDate = act.due_date || '';
            var userName = userMap[act.user_id] || String(act.user_id);
            var existing = types[t].map[dealId];
            if (!existing || dueDate < existing.date) {
              types[t].map[dealId] = { date: dueDate, person: userName };
            }
          }
          total += json.data.length;
          hasMore = (json.additional_data && json.additional_data.pagination && json.additional_data.pagination.more_items_in_collection) || false;
          start += 500;
          pages++;
          Utilities.sleep(200);
        } else { hasMore = false; }
      } catch(e) {
        Logger.log(types[t].label + ' API 오류: ' + e.message);
        hasMore = false;
      }
    }
    Logger.log(types[t].label + ': ' + total + '건, ' + Object.keys(types[t].map).length + '개 거래 (' + pages + '페이지)');
  }

  // 3. 시트에 쓰기
  var covRows = [];
  for (var d = 0; d < deals.length; d++) {
    var dl = deals[d];
    var pc = paymentCallMap[dl.dealId];
    var hc = hardCollectionMap[dl.dealId];
    covRows.push([
      dl.dealId, dl.title, dl.value, dl.noticeStr, dl.wonStr,
      dl.daysToPayment !== null ? dl.daysToPayment : '',
      dl.isAutoPaid ? 'Y' : 'N',
      dl.isHighValue ? 'Y' : 'N',
      pc ? 'Y' : 'N', pc ? pc.date : '', pc ? pc.person : '',
      hc ? 'Y' : 'N', hc ? hc.date : '', hc ? hc.person : ''
    ]);
  }

  var covSheet = getOrCreateCoverageSheet();
  writeRowsToSheet(covSheet, covRows);

  var payCallCount = covRows.filter(function(r) { return r[8] === 'Y'; }).length;
  var hardCount = covRows.filter(function(r) { return r[11] === 'Y'; }).length;
  Logger.log(SHEET_COVERAGE + ': ' + covRows.length + '건 (☏ 결제: ' + payCallCount + '건, 결제지연(강성): ' + hardCount + '건)');
}

// 배치 업데이트 (증분 동기화 - 성사 데이터는 절대 건드리지 않음)
function syncBatchUpdate() {
  var stageMap = fetchStageMap();

  function readSheetInfo(sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { wonCount: 0, wonIds: {}, openIds: {}, lastRow: lastRow };
    var col1 = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var col5 = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
    var wonCount = 0, wonIds = {}, openIds = {};
    for (var i = 0; i < col1.length; i++) {
      if (col5[i][0] && col5[i][0] !== '') { wonCount++; wonIds[col1[i][0]] = true; }
      else { openIds[col1[i][0]] = true; }
    }
    return { wonCount: wonCount, wonIds: wonIds, openIds: openIds, lastRow: lastRow };
  }

  var paySheet = getOrCreateSheet(SHEET_PAYMENT);
  var collSheet = getOrCreateSheet(SHEET_COLLECTION);
  var payInfo = readSheetInfo(paySheet);
  var collInfo = readSheetInfo(collSheet);

  Logger.log('기존: 결제현황 won=' + payInfo.wonCount + ' open=' + Object.keys(payInfo.openIds).length + ', 추심현황 won=' + collInfo.wonCount + ' open=' + Object.keys(collInfo.openIds).length);

  // Open 거래 조회
  var openDeals = fetchDealsByStatus('open');
  if (openDeals.length === 0) {
    Logger.log('⚠️ Open 거래 0건 - API 오류 가능성. 배치 중단.');
    return;
  }
  var currentOpenIds = {};
  for (var i = 0; i < openDeals.length; i++) currentOpenIds[openDeals[i].id] = true;

  // 이전 open에서 사라진 거래 → 성사 확인
  var allPrevOpen = {};
  var pid; for (pid in payInfo.openIds) allPrevOpen[pid] = true;
  var cid; for (cid in collInfo.openIds) allPrevOpen[cid] = true;
  var allWon = {};
  var wid; for (wid in payInfo.wonIds) allWon[wid] = true;
  var wid2; for (wid2 in collInfo.wonIds) allWon[wid2] = true;

  var newWonDeals = [];
  var prevKeys = Object.keys(allPrevOpen);
  for (var k = 0; k < prevKeys.length; k++) {
    var id = prevKeys[k];
    if (currentOpenIds[id] || allWon[id]) continue;
    var deal = fetchDealById(id);
    if (deal && deal.won_time) newWonDeals.push(deal);
    Utilities.sleep(100);
  }
  Logger.log('신규 Won: ' + newWonDeals.length + '건');

  // 시트 업데이트: won 행은 절대 건드리지 않음, open 영역만 교체
  function safeUpdate(sheet, info, sheetLabel) {
    var openStartRow = 2 + info.wonCount;

    // 1) 기존 open 행 삭제
    if (info.lastRow >= openStartRow) {
      sheet.getRange(openStartRow, 1, info.lastRow - openStartRow + 1, 15).clearContent();
      SpreadsheetApp.flush();
    }

    var writeRow = openStartRow;

    // 2) 신규 won 행 추가 (won 영역 뒤에 이어붙임)
    var newWonRows = [];
    for (var nw = 0; nw < newWonDeals.length; nw++) {
      var nd = newWonDeals[nw];
      var nwy = parseInt(String(nd.won_time || '').substring(0, 4)) || 0;
      var row = dealToRow(nd, stageMap);
      if (sheetLabel === 'pay' && (nwy >= 2025 || isRelevantDeal(nd))) newWonRows.push(row);
      if (sheetLabel === 'coll' && nd[CUSTOM_FIELD_COLLECTION_ORDER] && nd[CUSTOM_FIELD_COLLECTION_ORDER] !== '') newWonRows.push(row);
    }
    if (newWonRows.length > 0) {
      sheet.getRange(writeRow, 1, newWonRows.length, 15).setValues(newWonRows);
      writeRow += newWonRows.length;
      SpreadsheetApp.flush();
    }

    // 3) 현재 open 행 추가
    var openRows = [];
    for (var oi = 0; oi < openDeals.length; oi++) {
      var od = openDeals[oi];
      if (!isRelevantDeal(od)) continue;
      var orow = dealToRow(od, stageMap);
      if (sheetLabel === 'pay') openRows.push(orow);
      if (sheetLabel === 'coll' && od[CUSTOM_FIELD_COLLECTION_ORDER] && od[CUSTOM_FIELD_COLLECTION_ORDER] !== '') openRows.push(orow);
    }
    if (openRows.length > 0) {
      var CHUNK = 5000;
      for (var s = 0; s < openRows.length; s += CHUNK) {
        var chunk = openRows.slice(s, s + CHUNK);
        sheet.getRange(writeRow + s, 1, chunk.length, 15).setValues(chunk);
        SpreadsheetApp.flush();
      }
    }

    Logger.log(sheetLabel + ': won보존=' + info.wonCount + ', +won=' + newWonRows.length + ', open=' + openRows.length);
  }

  safeUpdate(paySheet, payInfo, 'pay');
  safeUpdate(collSheet, collInfo, 'coll');
  Logger.log('배치 완료');
}

// 배치 트리거 설정 (최초 1회 실행)
function setupBatchTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'syncBatchUpdate' || fn === 'syncCoverageData') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('syncBatchUpdate')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone('Asia/Seoul')
    .create();
  ScriptApp.newTrigger('syncCoverageData')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone('Asia/Seoul')
    .create();
  Logger.log('배치 트리거 설정 완료: 매일 오전 8시(KST) syncBatchUpdate, syncCoverageData 실행');
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

function processSheetToArrays(sheet, scope, filterYear, filterMonth) {
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
  var getMonth = function(v) {
    if (!v || v === '') return 0;
    if (v instanceof Date) return v.getMonth() + 1;
    var s = String(v);
    var m = parseInt(s.substring(5, 7));
    return isNaN(m) ? 0 : m;
  };

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var noticeYear = getYear(row[3]);
    var wonYear = getYear(row[4]);

    var include = true;
    if (scope === 'monthly' && filterYear && filterMonth) {
      var nm = getMonth(row[3]), wm = getMonth(row[4]);
      include = (noticeYear === filterYear && nm === filterMonth) || (wonYear === filterYear && wm === filterMonth);
    } else if (scope === 'yearly' && filterYear) {
      include = noticeYear === filterYear || wonYear === filterYear;
    }

    if (include) {
      result.push([
        row[0], row[1], Number(row[2]) || 0,
        toStr(row[3]), toStr(row[4]), toStr(row[5]),
        Number(row[6]) || 0, row[7] || '', row[8] || ''
      ]);
    }
  }
  return result;
}

var PAYMENT_COLS = ['deal_id','title','value','first_payment_notice','won_time','collection_order_date','balance','stage_name','customer_type'];

function buildPaymentJson(scope, year, month) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = (scope === 'collection') ? SHEET_COLLECTION : SHEET_PAYMENT;
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    var wonSheet = ss.getSheetByName('결제완료');
    var openSheet = ss.getSheetByName('미결제');
    var fallback = [];
    if (wonSheet) fallback = fallback.concat(processSheetToArrays(wonSheet, scope, year, month));
    if (openSheet) fallback = fallback.concat(processSheetToArrays(openSheet, scope, year, month));
    return JSON.stringify({ success: true, cols: PAYMENT_COLS, rows: fallback });
  }

  var filterScope = (scope === 'collection') ? null : scope;
  var allRows = processSheetToArrays(sheet, filterScope, year, month);
  return JSON.stringify({ success: true, cols: PAYMENT_COLS, rows: allRows });
}

function getPaymentData(scope, year, month) {
  var json = buildPaymentJson(scope, year, month);
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
  var result = { coverage: null, application: null, defense: null, target: null, paymentCoverage: null, pyPaymentCoverage: null };

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

function getCoverageData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_COVERAGE);
  if (!sheet || sheet.getLastRow() <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ success: true, cols: COVERAGE_COLS, rows: [] })).setMimeType(ContentService.MimeType.JSON);
  }
  var rows = sheet.getDataRange().getValues().slice(1);
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    result.push([
      r[0], r[1], Number(r[2]) || 0, String(r[3] instanceof Date ? Utilities.formatDate(r[3], 'Asia/Seoul', 'yyyy-MM-dd') : r[3] || ''),
      String(r[4] instanceof Date ? Utilities.formatDate(r[4], 'Asia/Seoul', 'yyyy-MM-dd') : r[4] || ''),
      r[5] !== '' ? Number(r[5]) : '', r[6] || '', r[7] || '',
      r[8] || '', String(r[9] instanceof Date ? Utilities.formatDate(r[9], 'Asia/Seoul', 'yyyy-MM-dd') : r[9] || ''), r[10] || '',
      r[11] || '', String(r[12] instanceof Date ? Utilities.formatDate(r[12], 'Asia/Seoul', 'yyyy-MM-dd') : r[12] || ''), r[13] || ''
    ]);
  }
  return ContentService.createTextOutput(JSON.stringify({ success: true, cols: COVERAGE_COLS, rows: result })).setMimeType(ContentService.MimeType.JSON);
}

// ========== 웹 앱 API ==========

function doGet(e) {
  var action = e.parameter.action || 'payment';

  if (action === 'payment') {
    var scope = e.parameter.scope || '';
    var pYear = e.parameter.year ? parseInt(e.parameter.year) : null;
    var pMonth = e.parameter.month ? parseInt(e.parameter.month) : null;
    return getPaymentData(scope, pYear, pMonth);
  } else if (action === 'performance') {
    var year = e.parameter.year ? parseInt(e.parameter.year) : null;
    var month = e.parameter.month ? parseInt(e.parameter.month) : null;
    return getPerformanceData(year, month);
  } else if (action === 'activity') {
    var aYear = e.parameter.year ? parseInt(e.parameter.year) : null;
    var aMonth = e.parameter.month ? parseInt(e.parameter.month) : null;
    return getActivityData(aYear, aMonth);
  } else if (action === 'coverage') {
    return getCoverageData();
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
