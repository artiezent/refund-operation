// ========== 설정 ==========
const PIPEDRIVE_API_KEY = 'cbc419beec83e32e7c9c50ab815eb0ab0508ea80';
const CUSTOM_FIELD_FIRST_PAYMENT_NOTICE = 'b24a25502fdeb48ac55986536dd8d449fe1ec494';
const CUSTOM_FIELD_COLLECTION_ORDER = '69b948cd5331d1c78a7d4045dae1be38be7a7177';  // 지급명령 단계

// 성과 요약용 커스텀 필드
const FIELD_REFUND_AMOUNT = '18f5d8f72f30db7d6abdc4aa862f64b9cb96409b';  // 조회 환급액

// 성과 요약용 필터 ID (Pipedrive에서 생성한 필터)
const FILTER_APPLY_SUCCESS = 1430754;    // 신청전환 성공 필터
const FILTER_DEFENSE_SUCCESS = 1430989;  // 취소방어 성공 필터

// 활동수 현황용 필터 ID
const FILTER_APPLY_ACTIVITY = 1431275;   // 신청전환 활동 필터
const FILTER_DEFENSE_ACTIVITY = 1431330; // 취소방어 활동 필터

// ========== 메인: 결제 데이터 동기화 (기존 유지) ==========
function syncPipedriveData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('결제데이터');
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  }
  
  const deals = fetchDealsWithPaymentNotice();
  
  if (deals.length === 0) {
    Logger.log('가져온 거래가 없습니다.');
    return;
  }
  
  const data = deals.map(deal => [
    deal.id,
    deal.title,
    deal.value || 0,
    deal[CUSTOM_FIELD_FIRST_PAYMENT_NOTICE] || '',
    deal.won_time || '',
    deal[CUSTOM_FIELD_COLLECTION_ORDER] || ''
  ]);
  
  sheet.getRange(2, 1, data.length, 6).setValues(data);
  Logger.log(`${data.length}개의 거래를 동기화했습니다.`);
}

function fetchDealsWithPaymentNotice() {
  let allDeals = [];
  
  const wonDeals = fetchDealsByStatus('won');
  allDeals = allDeals.concat(wonDeals);
  Logger.log(`Won 거래: ${wonDeals.length}개`);
  
  const openDeals = fetchDealsByStatus('open');
  allDeals = allDeals.concat(openDeals);
  Logger.log(`Open 거래: ${openDeals.length}개`);
  
  const filteredDeals = allDeals.filter(deal => {
    return deal[CUSTOM_FIELD_FIRST_PAYMENT_NOTICE] && 
           deal[CUSTOM_FIELD_FIRST_PAYMENT_NOTICE] !== '';
  });
  
  Logger.log(`결제안내가 있는 거래: ${filteredDeals.length}개`);
  return filteredDeals;
}

function fetchDealsByStatus(status) {
  let deals = [];
  let start = 0;
  const limit = 500;
  let hasMore = true;
  
  while (hasMore) {
    const url = `https://api.pipedrive.com/v1/deals?status=${status}&start=${start}&limit=${limit}&api_token=${PIPEDRIVE_API_KEY}`;
    
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const json = JSON.parse(response.getContentText());
      
      if (json.success && json.data) {
        deals = deals.concat(json.data);
        hasMore = json.additional_data?.pagination?.more_items_in_collection || false;
        start += limit;
        Utilities.sleep(200);
      } else {
        hasMore = false;
      }
    } catch (e) {
      Logger.log(`API 오류 (${status}): ` + e.message);
      hasMore = false;
    }
  }
  
  return deals;
}

// ========== 필터 기반 성과 요약 (신규) ==========

// 필터 ID로 거래 가져오기
function fetchDealsByFilter(filterId) {
  let deals = [];
  let start = 0;
  const limit = 500;
  let hasMore = true;
  
  while (hasMore) {
    const url = `https://api.pipedrive.com/v1/deals?filter_id=${filterId}&start=${start}&limit=${limit}&api_token=${PIPEDRIVE_API_KEY}`;
    
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const json = JSON.parse(response.getContentText());
      
      if (json.success && json.data) {
        deals = deals.concat(json.data);
        hasMore = json.additional_data?.pagination?.more_items_in_collection || false;
        start += limit;
        Utilities.sleep(200);
      } else {
        hasMore = false;
      }
    } catch (e) {
      Logger.log(`API 오류 (filter ${filterId}): ` + e.message);
      hasMore = false;
    }
  }
  
  return deals;
}

// 필터 기반 성과 계산
function calculatePerformanceByFilter() {
  Logger.log('=== 필터 기반 성과 계산 ===');
  
  // 1. 신청전환 성공
  const applyDeals = fetchDealsByFilter(FILTER_APPLY_SUCCESS);
  let applyCount = 0;
  let applyAmount = 0;
  
  applyDeals.forEach(deal => {
    const refundAmount = parseFloat(deal[FIELD_REFUND_AMOUNT]) || 0;
    if (refundAmount > 0) {
      applyCount++;
      applyAmount += refundAmount;
    }
  });
  
  Logger.log(`신청전환: ${applyDeals.length}건 조회, ${applyCount}건 환급액 있음, ${applyAmount.toLocaleString()}원`);
  
  // 2. 취소방어 성공
  const defenseDeals = fetchDealsByFilter(FILTER_DEFENSE_SUCCESS);
  let defenseCount = 0;
  let defenseAmount = 0;
  
  defenseDeals.forEach(deal => {
    const refundAmount = parseFloat(deal[FIELD_REFUND_AMOUNT]) || 0;
    if (refundAmount > 0) {
      defenseCount++;
      defenseAmount += refundAmount;
    }
  });
  
  Logger.log(`취소방어: ${defenseDeals.length}건 조회, ${defenseCount}건 환급액 있음, ${defenseAmount.toLocaleString()}원`);
  
  // 결과
  const total = applyAmount + defenseAmount;
  Logger.log(`\n=== 결과 ===`);
  Logger.log(`신청전환 성공: ${applyCount}건, ${applyAmount.toLocaleString()}원`);
  Logger.log(`취소방어 성공: ${defenseCount}건, ${defenseAmount.toLocaleString()}원`);
  Logger.log(`합계: ${total.toLocaleString()}원`);
  
  return {
    apply: { count: applyCount, amount: applyAmount },
    defense: { count: defenseCount, amount: defenseAmount },
    total: total
  };
}

// ========== 활동수 현황 (Activities API) ==========

// 필터 ID로 활동 가져오기 (API v2 사용)
function fetchActivitiesByFilter(filterId) {
  let activities = [];
  let cursor = null;
  let hasMore = true;
  
  while (hasMore) {
    let url = `https://api.pipedrive.com/api/v2/activities?filter_id=${filterId}&limit=500&api_token=${PIPEDRIVE_API_KEY}`;
    if (cursor) {
      url += `&cursor=${cursor}`;
    }
    
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const json = JSON.parse(response.getContentText());
      
      if (json.success && json.data) {
        activities = activities.concat(json.data);
        cursor = json.additional_data?.next_cursor || null;
        hasMore = cursor !== null;
        Utilities.sleep(200);
      } else {
        Logger.log('Activities API 응답 실패: ' + JSON.stringify(json));
        hasMore = false;
      }
    } catch (e) {
      Logger.log(`Activities API 오류 (filter ${filterId}): ` + e.message);
      hasMore = false;
    }
  }
  
  return activities;
}

// 활동 제목 분류 (신청 전환용)
function classifyActivitySubject(subject) {
  if (!subject) return null;
  
  const subjectLower = subject.toLowerCase();
  
  // 부재: "줍줍콜"과 "부재"가 모두 포함
  if (subjectLower.includes('줍줍콜') && subjectLower.includes('부재')) {
    return 'absent';
  }
  
  // 활동: "줍줍콜" 포함 (부재 제외) 또는 "vip" 포함
  if (subjectLower.includes('줍줍콜') || subjectLower.includes('vip')) {
    return 'activity';
  }
  
  // 사후관리
  if (subjectLower.includes('사후관리')) {
    return 'followup';
  }
  
  // 문자: "문자발송" 또는 "문자안내"
  if (subjectLower.includes('문자발송') || subjectLower.includes('문자안내')) {
    return 'sms';
  }
  
  return 'other';
}

// 활동 제목 분류 (취소방어용)
function classifyDefenseActivitySubject(subject) {
  if (!subject) return null;
  
  const subjectLower = subject.toLowerCase();
  
  // 부재: "취소방어"와 "부재"가 모두 포함
  if (subjectLower.includes('취소방어') && subjectLower.includes('부재')) {
    return 'absent';
  }
  
  // 활동: "취소방어" 포함 (부재 제외)
  if (subjectLower.includes('취소방어')) {
    return 'activity';
  }
  
  // 사후관리
  if (subjectLower.includes('사후관리')) {
    return 'followup';
  }
  
  // 문자: "문자" 포함
  if (subjectLower.includes('문자')) {
    return 'sms';
  }
  
  return 'other';
}

// 활동수 계산 (신청 전환)
function calculateApplyActivityCount() {
  Logger.log('=== 신청 전환 활동수 계산 ===');
  
  const activities = fetchActivitiesByFilter(FILTER_APPLY_ACTIVITY);
  Logger.log(`총 활동 조회: ${activities.length}건`);
  
  let counts = {
    activity: 0,  // 줍줍콜 + vip
    absent: 0,    // 줍줍콜 부재
    followup: 0,  // 사후관리
    sms: 0,       // 문자
    other: 0      // 기타
  };
  
  activities.forEach(act => {
    const subject = act.subject || '';
    const category = classifyActivitySubject(subject);
    if (category && counts.hasOwnProperty(category)) {
      counts[category]++;
    }
  });
  
  Logger.log(`활동(줍줍콜+VIP): ${counts.activity}건`);
  Logger.log(`부재: ${counts.absent}건`);
  Logger.log(`사후관리: ${counts.followup}건`);
  Logger.log(`문자: ${counts.sms}건`);
  Logger.log(`기타: ${counts.other}건`);
  
  return counts;
}

// 활동수 계산 (취소방어)
function calculateDefenseActivityCount() {
  Logger.log('=== 취소방어 활동수 계산 ===');
  
  const activities = fetchActivitiesByFilter(FILTER_DEFENSE_ACTIVITY);
  Logger.log(`총 활동 조회: ${activities.length}건`);
  
  let counts = {
    activity: 0,  // 취소방어
    absent: 0,    // 취소방어 부재
    followup: 0,  // 사후관리
    sms: 0,       // 문자
    other: 0      // 기타
  };
  
  activities.forEach(act => {
    const subject = act.subject || '';
    const category = classifyDefenseActivitySubject(subject);
    if (category && counts.hasOwnProperty(category)) {
      counts[category]++;
    }
  });
  
  Logger.log(`활동(취소방어): ${counts.activity}건`);
  Logger.log(`부재: ${counts.absent}건`);
  Logger.log(`사후관리: ${counts.followup}건`);
  Logger.log(`문자: ${counts.sms}건`);
  Logger.log(`기타: ${counts.other}건`);
  
  return counts;
}

// ========== 수동 입력 데이터 (Google Sheets 저장) ==========

// 수동 데이터 저장
function saveManualData(type, data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('수동입력');
  if (!sheet) {
    Logger.log('수동입력 시트를 찾을 수 없습니다.');
    return false;
  }
  
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // 기존 데이터 찾기 (같은 날짜, 같은 타입)
  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === today && values[i][1] === type) {
      rowIndex = i + 1; // 1-based index
      break;
    }
  }
  
  const jsonData = JSON.stringify(data);
  
  if (rowIndex > 0) {
    // 기존 데이터 업데이트
    sheet.getRange(rowIndex, 3).setValue(jsonData);
    Logger.log(`${type} 데이터 업데이트: ${today}`);
  } else {
    // 새 데이터 추가
    sheet.appendRow([today, type, jsonData]);
    Logger.log(`${type} 데이터 추가: ${today}`);
  }
  
  return true;
}

// 수동 데이터 읽기 (오늘 날짜)
function loadManualData(type) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('수동입력');
  if (!sheet) {
    Logger.log('수동입력 시트를 찾을 수 없습니다.');
    return null;
  }
  
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  for (let i = 1; i < values.length; i++) {
    // 날짜 비교 (Date 객체 또는 문자열 모두 처리)
    let rowDate = values[i][0];
    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, 'Asia/Seoul', 'yyyy-MM-dd');
    } else if (typeof rowDate === 'string') {
      rowDate = rowDate.trim();
    }
    
    if (rowDate === today && values[i][1] === type) {
      try {
        return JSON.parse(values[i][2]);
      } catch (e) {
        Logger.log(`JSON 파싱 오류 (${type}): ${e.message}`);
        return null;
      }
    }
  }
  
  return null;
}

// 모든 수동 데이터 읽기 (오늘 날짜)
function loadAllManualData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('수동입력');
  if (!sheet) {
    Logger.log('수동입력 시트를 찾을 수 없습니다.');
    return { coverage: null, application: null, defense: null };
  }
  
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  Logger.log('오늘 날짜: ' + today);
  
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  Logger.log('시트 데이터 행 수: ' + values.length);
  
  const result = {
    coverage: null,
    application: null,
    defense: null,
    target: null
  };
  
  for (let i = 1; i < values.length; i++) {
    // 날짜 비교 (Date 객체 또는 문자열 모두 처리)
    let rowDate = values[i][0];
    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, 'Asia/Seoul', 'yyyy-MM-dd');
    } else if (typeof rowDate === 'string') {
      rowDate = rowDate.trim();
    }
    
    Logger.log(`행 ${i}: 날짜="${rowDate}", 타입="${values[i][1]}"`);
    
    if (rowDate === today) {
      const type = values[i][1];
      if (result.hasOwnProperty(type)) {
        try {
          result[type] = JSON.parse(values[i][2]);
          Logger.log(`${type} 데이터 로드 성공`);
        } catch (e) {
          Logger.log(`JSON 파싱 오류 (${type}): ${e.message}`);
        }
      }
    }
  }
  
  Logger.log('최종 결과: ' + JSON.stringify(result));
  return result;
}

// ========== 웹 앱 API ==========
function doGet(e) {
  const action = e.parameter.action || 'payment';
  
  if (action === 'payment') {
    return getPaymentData();
  } else if (action === 'performance') {
    return getPerformanceData();
  } else if (action === 'activity') {
    return getActivityData();
  } else if (action === 'manual') {
    return getManualData();
  } else if (action === 'saveManual') {
    // GET으로 수동 데이터 저장 (CORS 우회)
    return saveManualDataViaGet(e);
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET 요청으로 수동 데이터 저장
function saveManualDataViaGet(e) {
  try {
    const type = e.parameter.type;
    const dataStr = e.parameter.data;
    
    if (!type || !dataStr) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: 'Missing type or data' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const data = JSON.parse(decodeURIComponent(dataStr));
    const result = saveManualData(type, data);
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: result }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log('saveManualDataViaGet 오류: ' + error.message);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// POST 요청 처리 (수동 데이터 저장)
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    if (action === 'saveManual') {
      const type = data.type;
      const payload = data.data;
      
      if (!type || !payload) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: 'Missing type or data' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      
      const result = saveManualData(type, payload);
      
      return ContentService
        .createTextOutput(JSON.stringify({ success: result }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'Unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log('doPost 오류: ' + error.message);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 결제 데이터 API (기존 유지)
function getPaymentData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('결제데이터');
  const data = sheet.getDataRange().getValues();
  
  const rows = data.slice(1);
  
  const toStr = (v) => {
    if (!v || v === '') return '';
    if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
    return String(v);
  };

  const result = rows.map(row => ({
    deal_id: row[0],
    title: row[1],
    value: Number(row[2]) || 0,
    first_payment_notice: toStr(row[3]),
    won_time: toStr(row[4]),
    collection_order_date: toStr(row[5])
  }));
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, data: result }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 성과 요약 API (필터 기반)
function getPerformanceData() {
  const result = calculatePerformanceByFilter();
  
  return ContentService
    .createTextOutput(JSON.stringify({ 
      success: true, 
      data: result 
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 활동수 API
function getActivityData() {
  const applyActivity = calculateApplyActivityCount();
  const defenseActivity = calculateDefenseActivityCount();
  
  return ContentService
    .createTextOutput(JSON.stringify({ 
      success: true, 
      data: {
        apply: applyActivity,
        defense: defenseActivity
      }
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 수동 입력 데이터 API (읽기)
function getManualData() {
  const data = loadAllManualData();
  
  return ContentService
    .createTextOutput(JSON.stringify({ 
      success: true, 
      data: data
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========== 테스트 함수 ==========

// 성과 테스트 (필터 기반)
function testPerformance() {
  calculatePerformanceByFilter();
}

// 신청전환만 테스트
function testApplyByFilter() {
  const deals = fetchDealsByFilter(FILTER_APPLY_SUCCESS);
  Logger.log('신청전환 필터 조회: ' + deals.length + '건');
  
  let totalAmount = 0;
  let count = 0;
  
  deals.forEach(deal => {
    const refundAmount = parseFloat(deal[FIELD_REFUND_AMOUNT]) || 0;
    if (refundAmount > 0) {
      totalAmount += refundAmount;
      count++;
    }
  });
  
  Logger.log('조회환급액 있는 거래: ' + count + '건');
  Logger.log('신청전환 성공 금액: ' + totalAmount.toLocaleString() + '원');
}

// 취소방어만 테스트
function testDefenseByFilter() {
  const deals = fetchDealsByFilter(FILTER_DEFENSE_SUCCESS);
  Logger.log('취소방어 필터 조회: ' + deals.length + '건');
  
  let totalAmount = 0;
  let count = 0;
  
  deals.forEach(deal => {
    const refundAmount = parseFloat(deal[FIELD_REFUND_AMOUNT]) || 0;
    if (refundAmount > 0) {
      totalAmount += refundAmount;
      count++;
    }
  });
  
  Logger.log('조회환급액 있는 거래: ' + count + '건');
  Logger.log('취소방어 성공 금액: ' + totalAmount.toLocaleString() + '원');
}

// API 연결 테스트
function testConnection() {
  const url = `https://api.pipedrive.com/v1/deals?status=won&limit=1&api_token=${PIPEDRIVE_API_KEY}`;
  const response = UrlFetchApp.fetch(url);
  Logger.log(response.getContentText());
}

// 활동수 테스트 (신청 전환)
function testApplyActivity() {
  calculateApplyActivityCount();
}

// 활동수 테스트 (취소방어)
function testDefenseActivity() {
  calculateDefenseActivityCount();
}

// 활동 제목 샘플 확인 (디버깅용)
function checkActivitySubjects() {
  const activities = fetchActivitiesByFilter(FILTER_APPLY_ACTIVITY);
  
  Logger.log('=== 활동 제목 샘플 (처음 20개) ===');
  for (let i = 0; i < Math.min(20, activities.length); i++) {
    const act = activities[i];
    const subject = act.subject || '(제목없음)';
    const category = classifyActivitySubject(subject);
    Logger.log(`${i+1}. "${subject}" → ${category}`);
  }
}

// 지급명령 단계 필드 디버깅
function testCollectionOrderField() {
  Logger.log('=== 지급명령 단계 필드 확인 ===');
  Logger.log('필드 키: ' + CUSTOM_FIELD_COLLECTION_ORDER);
  
  // Won 거래 중 처음 5개에서 해당 필드 확인
  const url = `https://api.pipedrive.com/v1/deals?status=won&start=0&limit=5&api_token=${PIPEDRIVE_API_KEY}`;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const json = JSON.parse(response.getContentText());
  
  if (json.success && json.data) {
    json.data.forEach((deal, i) => {
      const fieldValue = deal[CUSTOM_FIELD_COLLECTION_ORDER];
      Logger.log(`거래 ${i+1} (ID: ${deal.id}, ${deal.title})`);
      Logger.log(`  → 지급명령 단계 값: "${fieldValue}" (타입: ${typeof fieldValue})`);
      
      // 해당 거래의 모든 커스텀 필드 키 중 '69b9'로 시작하는 것 찾기
      const keys = Object.keys(deal).filter(k => k.includes('69b9'));
      if (keys.length > 0) {
        Logger.log(`  → '69b9' 포함 키: ${keys.join(', ')}`);
        keys.forEach(k => Logger.log(`     ${k} = "${deal[k]}"`));
      }
    });
  }
  
  // 시트에서 F열 데이터 확인
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('결제데이터');
  const lastRow = Math.min(sheet.getLastRow(), 20);
  if (lastRow > 1) {
    const fCol = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
    let hasData = 0;
    fCol.forEach((row, i) => {
      if (row[0] && row[0] !== '') {
        hasData++;
        if (hasData <= 5) {
          Logger.log(`시트 F열 (행 ${i+2}): "${row[0]}"`);
        }
      }
    });
    Logger.log(`F열에 데이터가 있는 행: ${hasData}개 (처음 ${lastRow-1}행 중)`);
  }
}
