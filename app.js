// Daily KPI Dashboard - Main Application

// DOM Elements
const tabButtons = document.querySelectorAll('.tab-btn');
const tabContents = document.querySelectorAll('.tab-content');
const toggleButtons = document.querySelectorAll('.toggle-input-btn');
const currentDateEl = document.getElementById('currentDate');

// Storage Keys
const STORAGE_KEYS = {
    coverage: 'dailyKpi_coverage',
    activity: 'dailyKpi_activity',
    application: 'dailyKpi_application',
    defense: 'dailyKpi_defense',
    performance: 'dailyKpi_performance'
};

// 커버리지 파싱 규칙 (Pipedrive)
const COVERAGE_PARSE_RULES = {
    successCount: {
        keyword: '전환 성공 건수',
        offsetLines: 5,
        type: 'count'
    },
    contactCount: {
        keyword: '컨택 진행 건수 (구간 전체)',
        offsetLines: 5,
        type: 'count'
    },
    unconvertedCount: {
        keyword: '미전환 건수 (구간 전체)',
        offsetLines: 5,
        type: 'count'
    },
    successAmount: {
        keyword: '전환 성공 금액',
        offsetLines: 5,
        type: 'amount'
    },
    contactAmount: {
        keyword: '컨택 진행 금액 (구간 전체)',
        offsetLines: 5,
        type: 'amount'
    },
    unconvertedAmount: {
        keyword: '미전환 금액 (구간 전체)',
        offsetLines: 5,
        type: 'amount'
    }
};

// Initialize App
function init() {
    setupTabNavigation();
    setupToggleButtons();
    displayCurrentDate();
    
    // 성과 요약 자동 로드
    initPerformanceSummary();
    
    // 활동수 자동 로드
    loadActivityStatus();
    
    // 수동 입력 데이터 로드 (Google Sheets에서)
    loadManualDataFromSheets();
}

// Tab Navigation
function setupTabNavigation() {
    tabButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const tabId = btn.dataset.tab;
            
            tabButtons.forEach(b => b.classList.remove('active'));
            tabContents.forEach(c => c.classList.remove('active'));
            
            btn.classList.add('active');
            document.getElementById(tabId).classList.add('active');
        });
    });
}

// Toggle Input Panels
function setupToggleButtons() {
    toggleButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const targetId = btn.dataset.target;
            const panel = document.getElementById(targetId);
            
            // Toggle active state
            btn.classList.toggle('active');
            panel.classList.toggle('active');
            
            // Update icon
            const icon = btn.querySelector('.toggle-icon');
            icon.textContent = btn.classList.contains('active') ? '×' : '+';
        });
    });
}

// Display Current Date
function displayCurrentDate() {
    const now = new Date();
    const options = { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric', 
        weekday: 'long' 
    };
    currentDateEl.textContent = now.toLocaleDateString('ko-KR', options);
}

// ==========================================
// 커버리지 현황 (Pipedrive)
// ==========================================

async function applyCoverageData() {
    const input = document.getElementById('coverageDataInput');
    const inputText = input.value.trim();
    
    if (!inputText) {
        showToast('데이터를 입력해주세요.', 'error');
        return;
    }
    
    const data = parsePipedriveData(inputText);
    
    // 건수 커버리지 계산
    const countNumerator = data.successCount + data.contactCount;
    const countDenominator = data.successCount + data.unconvertedCount;
    const countRate = countDenominator > 0 ? (countNumerator / countDenominator) * 100 : 0;
    
    // 금액 커버리지 계산
    const amountNumerator = data.successAmount + data.contactAmount;
    const amountDenominator = data.successAmount + data.unconvertedAmount;
    const amountRate = amountDenominator > 0 ? (amountNumerator / amountDenominator) * 100 : 0;
    
    // UI 업데이트
    updateKPICard('coverage-count', countRate, {
        'success-count': formatNumber(data.successCount),
        'contact-count': formatNumber(data.contactCount),
        'unconverted-count': formatNumber(data.unconvertedCount)
    });
    
    updateKPICard('coverage-amount', amountRate, {
        'success-amount': '₩' + formatNumber(data.successAmount),
        'contact-amount': '₩' + formatNumber(data.contactAmount),
        'unconverted-amount': '₩' + formatNumber(data.unconvertedAmount)
    });
    
    // Google Sheets에 저장
    const saveData = {
        ...data,
        countRate,
        amountRate
    };
    
    showToast('저장 중...');
    const saved = await saveManualDataToSheets('coverage', saveData);
    
    // 입력 패널 닫기
    closeInputPanel('coverage-input');
    
    if (saved) {
        showToast('커버리지 현황이 업데이트되었습니다!');
    } else {
        showToast('저장에 실패했습니다. 다시 시도해주세요.', 'error');
    }
}

function clearCoverageData() {
    document.getElementById('coverageDataInput').value = '';
    
    updateKPICard('coverage-count', 0, {
        'success-count': '0',
        'contact-count': '0',
        'unconverted-count': '0'
    });
    
    updateKPICard('coverage-amount', 0, {
        'success-amount': '₩0',
        'contact-amount': '₩0',
        'unconverted-amount': '₩0'
    });
    
    localStorage.removeItem(STORAGE_KEYS.coverage);
    showToast('커버리지 데이터가 초기화되었습니다.');
}

function parsePipedriveData(text) {
    const lines = text.split('\n').map(line => line.trim());
    const result = {};
    
    for (const [key, rule] of Object.entries(COVERAGE_PARSE_RULES)) {
        const value = extractValue(lines, rule);
        result[key] = value;
    }
    
    return result;
}

function extractValue(lines, rule) {
    const keywordIndex = lines.findIndex(line => line.includes(rule.keyword));
    
    if (keywordIndex === -1) {
        console.log(`키워드 "${rule.keyword}"를 찾을 수 없음`);
        return 0;
    }
    
    for (let i = 1; i <= rule.offsetLines + 2; i++) {
        const targetIndex = keywordIndex + i;
        if (targetIndex >= lines.length) break;
        
        const line = lines[targetIndex];
        
        if (rule.type === 'count') {
            const countMatch = line.match(/^[\d,]+$/);
            if (countMatch) {
                return parseNumber(countMatch[0]);
            }
            const numberOnlyMatch = line.match(/^([\d,]+)\s*$/);
            if (numberOnlyMatch) {
                return parseNumber(numberOnlyMatch[1]);
            }
        } else if (rule.type === 'amount') {
            const amountMatch = line.match(/^₩([\d,]+)/);
            if (amountMatch) {
                return parseNumber(amountMatch[1]);
            }
        }
    }
    
    console.log(`"${rule.keyword}"의 값을 찾을 수 없음`);
    return 0;
}

// ==========================================
// 활동수 현황 (Pipedrive)
// ==========================================

// 활동수 파싱 규칙
const ACTIVITY_PARSE_RULES = {
    // 신청 전환 (줍줍콜)
    applyActivity: { keyword: '줍줍 활동', offsetLines: 5 },
    applyAbsent: { keyword: '줍줍 부재', offsetLines: 5 },
    applyFollowup: { keyword: '줍줍콜 사후관리', offsetLines: 5 },
    applySms: { keyword: '신청 문자', offsetLines: 5 },
    // 취소 방어
    defenseActivity: { keyword: '취소 활동', offsetLines: 5 },
    defenseAbsent: { keyword: '취소 부재', offsetLines: 5 },
    defenseFollowup: { keyword: '취소방어 사후관리', offsetLines: 5 },
    defenseSms: { keyword: '취소 문자', offsetLines: 5 }
};

function applyActivityData() {
    const input = document.getElementById('activityDataInput');
    const inputText = input.value.trim();
    
    if (!inputText) {
        showToast('데이터를 입력해주세요.', 'error');
        return;
    }
    
    const data = parseActivityData(inputText);
    
    // 신청 전환 계산
    const applyTotal = data.applyActivity + data.applyAbsent;
    const applyExtra = data.applyFollowup + data.applySms;
    
    // 취소 방어 계산
    const defenseTotal = data.defenseActivity + data.defenseAbsent;
    const defenseExtra = data.defenseFollowup + data.defenseSms;
    
    // 최대값 계산 (바 높이 비율용)
    const maxValue = Math.max(applyTotal, applyExtra, defenseTotal, defenseExtra, 1);
    
    // UI 업데이트 - 신청 전환 (스택 바)
    updateTotalActivityBar('apply', data.applyActivity, data.applyAbsent, maxValue);
    updateExtraActivityBar('apply', data.applyFollowup, data.applySms, maxValue);
    document.getElementById('apply-total-value').textContent = formatNumber(applyTotal);
    document.getElementById('apply-extra-value').textContent = formatNumber(applyExtra);
    // 큰 값이 왼쪽에 오도록 막대 순서 변경
    updateBarGroupOrder('apply', applyTotal, applyExtra);
    
    // 툴팁 업데이트 - 신청 전환
    document.getElementById('tt-apply-activity').textContent = formatNumber(data.applyActivity);
    document.getElementById('tt-apply-absent').textContent = formatNumber(data.applyAbsent);
    document.getElementById('tt-apply-total').textContent = formatNumber(applyTotal);
    document.getElementById('tt-apply-followup').textContent = formatNumber(data.applyFollowup);
    document.getElementById('tt-apply-sms').textContent = formatNumber(data.applySms);
    document.getElementById('tt-apply-extra').textContent = formatNumber(applyExtra);
    
    // UI 업데이트 - 취소 방어 (스택 바)
    updateTotalActivityBar('defense', data.defenseActivity, data.defenseAbsent, maxValue);
    updateExtraActivityBar('defense', data.defenseFollowup, data.defenseSms, maxValue);
    document.getElementById('defense-total-value').textContent = formatNumber(defenseTotal);
    document.getElementById('defense-extra-value').textContent = formatNumber(defenseExtra);
    // 큰 값이 왼쪽에 오도록 막대 순서 변경
    updateBarGroupOrder('defense', defenseTotal, defenseExtra);
    
    // 툴팁 업데이트 - 취소 방어
    document.getElementById('tt-defense-activity').textContent = formatNumber(data.defenseActivity);
    document.getElementById('tt-defense-absent').textContent = formatNumber(data.defenseAbsent);
    document.getElementById('tt-defense-total').textContent = formatNumber(defenseTotal);
    document.getElementById('tt-defense-followup').textContent = formatNumber(data.defenseFollowup);
    document.getElementById('tt-defense-sms').textContent = formatNumber(data.defenseSms);
    document.getElementById('tt-defense-extra').textContent = formatNumber(defenseExtra);
    
    // 저장
    const saveData = {
        ...data,
        applyTotal,
        applyExtra,
        defenseTotal,
        defenseExtra,
        date: new Date().toDateString(),
        rawText: inputText
    };
    localStorage.setItem(STORAGE_KEYS.activity, JSON.stringify(saveData));
    
    // 입력 패널 닫기
    closeInputPanel('activity-input');
    
    // 스냅샷 자동 저장
    saveSnapshot();
    
    showToast('활동수 현황이 업데이트되었습니다!');
}

function parseActivityData(text) {
    const lines = text.split('\n').map(line => line.trim());
    const result = {};
    
    for (const [key, rule] of Object.entries(ACTIVITY_PARSE_RULES)) {
        result[key] = extractActivityValue(lines, rule);
    }
    
    return result;
}

function extractActivityValue(lines, rule) {
    const keywordIndex = lines.findIndex(line => line === rule.keyword || line.startsWith(rule.keyword));
    
    if (keywordIndex === -1) {
        console.log(`활동수 키워드 "${rule.keyword}"를 찾을 수 없음`);
        return 0;
    }
    
    // 키워드 아래로 검색하여 숫자 찾기
    for (let i = 1; i <= rule.offsetLines + 2; i++) {
        const targetIndex = keywordIndex + i;
        if (targetIndex >= lines.length) break;
        
        const line = lines[targetIndex];
        
        // 순수 숫자 또는 콤마가 포함된 숫자 찾기
        const countMatch = line.match(/^[\d,]+$/);
        if (countMatch) {
            return parseNumber(countMatch[0]);
        }
    }
    
    console.log(`"${rule.keyword}"의 값을 찾을 수 없음`);
    return 0;
}

function updateActivityBar(id, value, maxValue) {
    const bar = document.getElementById(`${id}-bar`);
    if (bar) {
        const heightPercent = maxValue > 0 ? (value / maxValue) * 100 : 0;
        bar.style.height = Math.max(heightPercent, 5) + '%';
        bar.dataset.value = value;
    }
}

// 스택 바 업데이트 (세그먼트 비율로 나누기)
function updateStackedBar(prefix, segment1Value, segment2Value, maxValue) {
    const bar = document.getElementById(`${prefix}-bar`);
    const segment1 = document.getElementById(`${prefix}-activity-segment`) || document.getElementById(`${prefix}-followup-segment`);
    const segment2 = document.getElementById(`${prefix}-absent-segment`) || document.getElementById(`${prefix}-sms-segment`);
    
    if (!bar) return;
    
    const total = segment1Value + segment2Value;
    const heightPercent = maxValue > 0 ? (total / maxValue) * 100 : 0;
    bar.style.height = Math.max(heightPercent, 5) + '%';
    bar.dataset.value = total;
    
    if (segment1 && segment2 && total > 0) {
        const seg1Ratio = segment1Value / total;
        const seg2Ratio = segment2Value / total;
        segment1.style.flex = seg1Ratio;
        segment2.style.flex = seg2Ratio;
    }
}

// 총활동 스택바 업데이트
function updateTotalActivityBar(prefix, activityValue, absentValue, maxValue) {
    const bar = document.getElementById(`${prefix}-total-bar`);
    const activitySeg = document.getElementById(`${prefix}-activity-segment`);
    const absentSeg = document.getElementById(`${prefix}-absent-segment`);
    
    if (!bar) return;
    
    const total = activityValue + absentValue;
    const heightPercent = maxValue > 0 ? (total / maxValue) * 100 : 0;
    bar.style.height = Math.max(heightPercent, 5) + '%';
    bar.dataset.value = total;
    
    if (activitySeg && absentSeg && total > 0) {
        activitySeg.style.flex = activityValue / total;
        absentSeg.style.flex = absentValue / total;
        // 큰 값이 위에 오도록 order 설정 (order 작은 값이 위)
        if (activityValue >= absentValue) {
            activitySeg.style.order = 1;
            absentSeg.style.order = 2;
        } else {
            activitySeg.style.order = 2;
            absentSeg.style.order = 1;
        }
    }
}

// 부가활동 스택바 업데이트
function updateExtraActivityBar(prefix, followupValue, smsValue, maxValue) {
    const bar = document.getElementById(`${prefix}-extra-bar`);
    const followupSeg = document.getElementById(`${prefix}-followup-segment`);
    const smsSeg = document.getElementById(`${prefix}-sms-segment`);
    
    if (!bar) return;
    
    const total = followupValue + smsValue;
    const heightPercent = maxValue > 0 ? (total / maxValue) * 100 : 0;
    bar.style.height = Math.max(heightPercent, 5) + '%';
    bar.dataset.value = total;
    
    if (followupSeg && smsSeg && total > 0) {
        followupSeg.style.flex = followupValue / total;
        smsSeg.style.flex = smsValue / total;
        // 큰 값이 위에 오도록 order 설정 (order 작은 값이 위)
        if (followupValue >= smsValue) {
            followupSeg.style.order = 1;
            smsSeg.style.order = 2;
        } else {
            followupSeg.style.order = 2;
            smsSeg.style.order = 1;
        }
    }
}

// 막대 그룹 순서 변경 (큰 값이 왼쪽에)
function updateBarGroupOrder(prefix, totalValue, extraValue) {
    const totalBarGroup = document.getElementById(`${prefix}-total-bar`)?.closest('.bar-group');
    const extraBarGroup = document.getElementById(`${prefix}-extra-bar`)?.closest('.bar-group');
    
    if (totalBarGroup && extraBarGroup) {
        if (totalValue >= extraValue) {
            totalBarGroup.style.order = 1;
            extraBarGroup.style.order = 2;
        } else {
            totalBarGroup.style.order = 2;
            extraBarGroup.style.order = 1;
        }
    }
}

function clearActivityData() {
    document.getElementById('activityDataInput').value = '';
    
    // 바 초기화
    ['apply-total', 'apply-extra', 'defense-total', 'defense-extra'].forEach(id => {
        const bar = document.getElementById(`${id}-bar`);
        const value = document.getElementById(`${id}-value`);
        if (bar) bar.style.height = '5%';
        if (value) value.textContent = '0';
    });
    
    // 세그먼트 초기화
    const segmentIds = [
        'apply-activity-segment', 'apply-absent-segment',
        'apply-followup-segment', 'apply-sms-segment',
        'defense-activity-segment', 'defense-absent-segment',
        'defense-followup-segment', 'defense-sms-segment'
    ];
    segmentIds.forEach(id => {
        const seg = document.getElementById(id);
        if (seg) seg.style.flex = '0';
    });
    
    // 툴팁 초기화
    const tooltipIds = [
        'tt-apply-activity', 'tt-apply-absent', 'tt-apply-total',
        'tt-apply-followup', 'tt-apply-sms', 'tt-apply-extra',
        'tt-defense-activity', 'tt-defense-absent', 'tt-defense-total',
        'tt-defense-followup', 'tt-defense-sms', 'tt-defense-extra'
    ];
    tooltipIds.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.textContent = '0';
    });
    
    localStorage.removeItem(STORAGE_KEYS.activity);
    showToast('활동수 데이터가 초기화되었습니다.');
}

// ==========================================
// 총조회대비 신청전환 (Pipedrive)
// ==========================================

async function applyApplicationData() {
    const input = document.getElementById('applicationDataInput');
    const inputText = input.value.trim();
    
    if (!inputText) {
        showToast('데이터를 입력해주세요.', 'error');
        return;
    }
    
    const lines = inputText.split('\n').map(line => line.trim());
    
    // 새 키워드로 금액 추출
    const totalView = findAmountByKeyword(lines, 'KPI_총조회');
    const totalApply = findAmountByKeyword(lines, 'KPI_신청');
    const applyConvert = findAmountByKeyword(lines, 'KPI_신청전환');
    
    console.log('신청전환 파싱 결과:', { totalView, totalApply, applyConvert });
    
    // 총조회 대비 전체 신청률
    const totalApplyRate = totalView > 0 ? (totalApply / totalView) * 100 : 0;
    
    // 총조회 대비 전환 성공률
    const applySuccessRate = totalView > 0 ? (applyConvert / totalView) * 100 : 0;
    
    // UI 업데이트 - 전체 신청률
    updateKPICard('apply-total-rate', totalApplyRate, {
        'total-view-amount': '₩' + formatNumber(totalView),
        'total-apply-amount': '₩' + formatNumber(totalApply)
    });
    
    // UI 업데이트 - 전환 성공률
    updateKPICard('apply-success-rate', applySuccessRate, {
        'total-view-amount2': '₩' + formatNumber(totalView),
        'apply-convert-amount': '₩' + formatNumber(applyConvert)
    });
    
    // Google Sheets에 저장
    const saveData = {
        totalView,
        totalApply,
        applyConvert,
        totalApplyRate,
        applySuccessRate
    };
    
    showToast('저장 중...');
    const saved = await saveManualDataToSheets('application', saveData);
    
    closeInputPanel('application-input');
    
    if (saved) {
        showToast('신청전환 현황이 업데이트되었습니다!');
    } else {
        showToast('저장에 실패했습니다. 다시 시도해주세요.', 'error');
    }
}

// 키워드로 금액 찾기 (정확히 해당 줄이 키워드와 일치)
function findAmountByKeyword(lines, keyword) {
    // 줄이 정확히 키워드와 같은지 확인
    const keywordIndex = lines.findIndex(line => line === keyword);
    
    if (keywordIndex === -1) {
        console.log(`키워드 "${keyword}"를 찾을 수 없음`);
        return 0;
    }
    
    console.log(`키워드 "${keyword}" 찾음 (라인 ${keywordIndex})`);
    
    // 키워드 아래 7줄 범위에서 ₩로 시작하는 금액 찾기 (- 제외)
    for (let i = 1; i <= 7; i++) {
        const targetIndex = keywordIndex + i;
        if (targetIndex >= lines.length) break;
        
        const line = lines[targetIndex];
        console.log(`  라인 ${targetIndex}: "${line}"`);
        
        // ₩로 시작하고 -가 아닌 금액 찾기
        const amountMatch = line.match(/^₩([\d,]+)/);
        if (amountMatch) {
            console.log(`  → 금액 발견: ${amountMatch[1]}`);
            return parseNumber(amountMatch[1]);
        }
    }
    
    console.log(`"${keyword}"의 값을 찾을 수 없음`);
    return 0;
}

// 정확한 키워드로 금액 찾기 (excludeWords 포함 시 제외)
function findAmountByExactKeyword(lines, keyword, excludeWords = []) {
    const keywordIndex = lines.findIndex(line => {
        if (!line.includes(keyword)) return false;
        // 제외 단어가 포함되어 있으면 제외
        for (const word of excludeWords) {
            if (line.includes(word)) return false;
        }
        return true;
    });
    
    if (keywordIndex === -1) {
        console.log(`정확한 키워드 "${keyword}"를 찾을 수 없음 (제외: ${excludeWords})`);
        return 0;
    }
    
    // 키워드 아래 7줄 범위에서 ₩로 시작하는 금액 찾기
    for (let i = 1; i <= 7; i++) {
        const targetIndex = keywordIndex + i;
        if (targetIndex >= lines.length) break;
        
        const line = lines[targetIndex];
        const amountMatch = line.match(/^₩([\d,]+)/);
        if (amountMatch) {
            return parseNumber(amountMatch[1]);
        }
    }
    
    console.log(`"${keyword}"의 값을 찾을 수 없음`);
    return 0;
}

function clearApplicationData() {
    document.getElementById('applicationDataInput').value = '';
    
    updateKPICard('apply-total-rate', 0, {
        'total-view-amount': '₩0',
        'total-apply-amount': '₩0'
    });
    
    updateKPICard('apply-success-rate', 0, {
        'total-view-amount2': '₩0',
        'apply-convert-amount': '₩0'
    });
    
    localStorage.removeItem(STORAGE_KEYS.application);
    showToast('신청전환 데이터가 초기화되었습니다.');
}

// ==========================================
// 총검토대비 취소방어 (Pipedrive)
// ==========================================

async function applyDefenseData() {
    const input = document.getElementById('defenseDataInput');
    const inputText = input.value.trim();
    
    if (!inputText) {
        showToast('데이터를 입력해주세요.', 'error');
        return;
    }
    
    const lines = inputText.split('\n').map(line => line.trim());
    
    // 새 키워드로 금액 추출
    const cancelRequest = findAmountByKeyword(lines, 'KPI_취소전체');
    const cancelAvailable = findAmountByKeyword(lines, 'KPI_취소검토');
    const cancelSuccess = findAmountByKeyword(lines, 'KPI_취소성공');
    
    console.log('취소방어 파싱 결과:', { cancelRequest, cancelAvailable, cancelSuccess });
    
    // 전체취소 대비 검토완료율
    const reviewRate = cancelRequest > 0 ? (cancelAvailable / cancelRequest) * 100 : 0;
    
    // 검토완료대비 취소방어 성공률
    const defenseRate = cancelAvailable > 0 ? (cancelSuccess / cancelAvailable) * 100 : 0;
    
    // UI 업데이트 - 검토완료율
    updateKPICard('cancel-review-rate', reviewRate, {
        'cancel-request-amount': '₩' + formatNumber(cancelRequest),
        'cancel-available-amount': '₩' + formatNumber(cancelAvailable)
    });
    
    // UI 업데이트 - 방어 성공률
    updateKPICard('cancel-defense-rate', defenseRate, {
        'cancel-available-amount2': '₩' + formatNumber(cancelAvailable),
        'cancel-success-amount': '₩' + formatNumber(cancelSuccess)
    });
    
    // Google Sheets에 저장
    const saveData = {
        cancelRequest,
        cancelAvailable,
        cancelSuccess,
        reviewRate,
        defenseRate
    };
    
    showToast('저장 중...');
    const saved = await saveManualDataToSheets('defense', saveData);
    
    closeInputPanel('defense-input');
    
    if (saved) {
        showToast('취소방어 현황이 업데이트되었습니다!');
    } else {
        showToast('저장에 실패했습니다. 다시 시도해주세요.', 'error');
    }
}

function clearDefenseData() {
    document.getElementById('defenseDataInput').value = '';
    
    updateKPICard('cancel-review-rate', 0, {
        'cancel-request-amount': '₩0',
        'cancel-available-amount': '₩0'
    });
    
    updateKPICard('cancel-defense-rate', 0, {
        'cancel-available-amount2': '₩0',
        'cancel-success-amount': '₩0'
    });
    
    localStorage.removeItem(STORAGE_KEYS.defense);
    showToast('취소방어 데이터가 초기화되었습니다.');
}

// ==========================================
// 공통 함수들
// ==========================================

function updateKPICard(prefix, rate, values) {
    const rateEl = document.getElementById(`${prefix}-rate`);
    const barEl = document.getElementById(`${prefix}-bar`);
    
    if (!rateEl || !barEl) return;
    
    rateEl.textContent = rate.toFixed(1) + '%';
    
    const barWidth = Math.min(rate, 100);
    barEl.style.width = barWidth + '%';
    
    rateEl.classList.remove('success', 'warning', 'danger');
    barEl.classList.remove('success', 'warning', 'danger');
    
    let colorClass = '';
    if (rate >= 100) {
        colorClass = 'success';
    } else if (rate >= 80) {
        colorClass = 'warning';
    } else if (rate > 0) {
        colorClass = 'danger';
    }
    
    if (colorClass) {
        rateEl.classList.add(colorClass);
        barEl.classList.add(colorClass);
    }
    
    for (const [id, value] of Object.entries(values)) {
        const el = document.getElementById(id);
        if (el) {
            el.textContent = value;
        }
    }
}

function parseNumber(str) {
    if (!str) return 0;
    const cleaned = str.replace(/,/g, '').replace(/[^0-9.-]/g, '');
    const num = parseFloat(cleaned);
    return isNaN(num) ? 0 : num;
}

function formatNumber(num) {
    if (num === 0 || num === null || num === undefined) return '0';
    return Math.round(num).toLocaleString('ko-KR');
}

function closeInputPanel(panelId) {
    const panel = document.getElementById(panelId);
    const btn = document.querySelector(`[data-target="${panelId}"]`);
    
    if (panel && btn) {
        panel.classList.remove('active');
        btn.classList.remove('active');
        const icon = btn.querySelector('.toggle-icon');
        if (icon) icon.textContent = '+';
    }
}

function loadAllSavedData() {
    // 성과 요약 데이터 로드
    loadPerformanceData();
    
    // 커버리지 데이터 로드
    loadCoverageData();
    
    // 활동수 데이터 로드
    loadActivityData();
    
    // 신청전환 데이터 로드
    loadApplicationData();
    
    // 취소방어 데이터 로드
    loadDefenseData();
}

// ==========================================
// 성과 요약 (직접 입력)
// ==========================================

function applyPerformanceData() {
    const applyInput = document.getElementById('applySuccessInput');
    const defenseInput = document.getElementById('defenseSuccessInput');
    const targetInput = document.getElementById('targetAmountInput');
    
    // 입력값 파싱 (콤마 제거)
    const applyAmount = parseNumber(applyInput.value);
    const defenseAmount = parseNumber(defenseInput.value);
    const targetAmount = parseNumber(targetInput.value);
    
    if (applyAmount === 0 && defenseAmount === 0 && targetAmount === 0) {
        showToast('금액을 입력해주세요.', 'error');
        return;
    }
    
    // 합계 및 진행률 계산
    const totalAmount = applyAmount + defenseAmount;
    const progressRate = targetAmount > 0 ? (totalAmount / targetAmount) * 100 : 0;
    
    // UI 업데이트
    updatePerformanceUI(applyAmount, defenseAmount, totalAmount, targetAmount, progressRate);
    
    // 저장
    const saveData = {
        applyAmount,
        defenseAmount,
        targetAmount,
        totalAmount,
        progressRate,
        date: new Date().toDateString()
    };
    localStorage.setItem(STORAGE_KEYS.performance, JSON.stringify(saveData));
    
    // 입력 패널 닫기
    closeInputPanel('performance-input');
    
    // 스냅샷 자동 저장
    saveSnapshot();
    
    showToast('성과 요약이 업데이트되었습니다!');
}

function clearPerformanceData() {
    document.getElementById('applySuccessInput').value = '';
    document.getElementById('defenseSuccessInput').value = '';
    document.getElementById('targetAmountInput').value = '';
    
    updatePerformanceUI(0, 0, 0, 0, 0);
    
    localStorage.removeItem(STORAGE_KEYS.performance);
    showToast('성과 요약이 초기화되었습니다.');
}

function updatePerformanceUI(applyAmount, defenseAmount, totalAmount, targetAmount, progressRate) {
    // 금액 표시
    document.getElementById('perf-apply-amount').textContent = '₩' + formatNumber(applyAmount);
    document.getElementById('perf-defense-amount').textContent = '₩' + formatNumber(defenseAmount);
    document.getElementById('perf-total-amount').textContent = '₩' + formatNumber(totalAmount);
    document.getElementById('perf-target-amount').textContent = '₩' + formatNumber(targetAmount);
    
    // 진행률 표시
    document.getElementById('perf-progress-rate').textContent = progressRate.toFixed(1) + '%';
    
    // 진행률 바 (최대 100%)
    const progressFill = document.getElementById('perf-progress-fill');
    progressFill.style.width = Math.min(progressRate, 100) + '%';
    
    // 진행률 색상
    const rateEl = document.getElementById('perf-progress-rate');
    rateEl.style.color = '';
    if (progressRate >= 100) {
        rateEl.style.color = 'var(--accent-success)';
    } else if (progressRate >= 80) {
        rateEl.style.color = 'var(--accent-primary)';
    } else if (progressRate >= 50) {
        rateEl.style.color = 'var(--accent-warning)';
    } else if (progressRate > 0) {
        rateEl.style.color = 'var(--accent-danger)';
    }
}

function loadPerformanceData() {
    try {
        const savedData = localStorage.getItem(STORAGE_KEYS.performance);
        if (savedData) {
            const data = JSON.parse(savedData);
            
            const today = new Date().toDateString();
            if (data.date === today) {
                // 입력 필드 복원
                document.getElementById('applySuccessInput').value = data.applyAmount || '';
                document.getElementById('defenseSuccessInput').value = data.defenseAmount || '';
                document.getElementById('targetAmountInput').value = data.targetAmount || '';
                
                // UI 업데이트
                updatePerformanceUI(
                    data.applyAmount || 0,
                    data.defenseAmount || 0,
                    data.totalAmount || 0,
                    data.targetAmount || 0,
                    data.progressRate || 0
                );
            }
        }
    } catch (e) {
        console.error('성과 요약 데이터 로드 실패:', e);
    }
}

function loadApplicationData() {
    try {
        const savedData = localStorage.getItem(STORAGE_KEYS.application);
        if (savedData) {
            const data = JSON.parse(savedData);
            
            const today = new Date().toDateString();
            if (data.date === today) {
                updateKPICard('apply-total-rate', data.totalApplyRate, {
                    'total-view-amount': '₩' + formatNumber(data.totalView),
                    'total-apply-amount': '₩' + formatNumber(data.totalApply)
                });
                
                updateKPICard('apply-success-rate', data.applySuccessRate, {
                    'total-view-amount2': '₩' + formatNumber(data.totalView),
                    'apply-convert-amount': '₩' + formatNumber(data.applyConvert)
                });
                
                if (data.rawText) {
                    document.getElementById('applicationDataInput').value = data.rawText;
                }
            }
        }
    } catch (e) {
        console.error('신청전환 데이터 로드 실패:', e);
    }
}

function loadDefenseData() {
    try {
        const savedData = localStorage.getItem(STORAGE_KEYS.defense);
        if (savedData) {
            const data = JSON.parse(savedData);
            
            const today = new Date().toDateString();
            if (data.date === today) {
                updateKPICard('cancel-review-rate', data.reviewRate, {
                    'cancel-request-amount': '₩' + formatNumber(data.cancelRequest),
                    'cancel-available-amount': '₩' + formatNumber(data.cancelAvailable)
                });
                
                updateKPICard('cancel-defense-rate', data.defenseRate, {
                    'cancel-available-amount2': '₩' + formatNumber(data.cancelAvailable),
                    'cancel-success-amount': '₩' + formatNumber(data.cancelSuccess)
                });
                
                if (data.rawText) {
                    document.getElementById('defenseDataInput').value = data.rawText;
                }
            }
        }
    } catch (e) {
        console.error('취소방어 데이터 로드 실패:', e);
    }
}

function loadActivityData() {
    try {
        const savedData = localStorage.getItem(STORAGE_KEYS.activity);
        if (savedData) {
            const data = JSON.parse(savedData);
            
            const today = new Date().toDateString();
            if (data.date === today) {
                // 최대값 계산
                const maxValue = Math.max(data.applyTotal, data.applyExtra, data.defenseTotal, data.defenseExtra, 1);
                
                // UI 업데이트 - 신청 전환 (스택 바)
                updateTotalActivityBar('apply', data.applyActivity, data.applyAbsent, maxValue);
                updateExtraActivityBar('apply', data.applyFollowup, data.applySms, maxValue);
                document.getElementById('apply-total-value').textContent = formatNumber(data.applyTotal);
                document.getElementById('apply-extra-value').textContent = formatNumber(data.applyExtra);
                updateBarGroupOrder('apply', data.applyTotal, data.applyExtra);
                
                // 툴팁 업데이트 - 신청 전환
                document.getElementById('tt-apply-activity').textContent = formatNumber(data.applyActivity);
                document.getElementById('tt-apply-absent').textContent = formatNumber(data.applyAbsent);
                document.getElementById('tt-apply-total').textContent = formatNumber(data.applyTotal);
                document.getElementById('tt-apply-followup').textContent = formatNumber(data.applyFollowup);
                document.getElementById('tt-apply-sms').textContent = formatNumber(data.applySms);
                document.getElementById('tt-apply-extra').textContent = formatNumber(data.applyExtra);
                
                // UI 업데이트 - 취소 방어 (스택 바)
                updateTotalActivityBar('defense', data.defenseActivity, data.defenseAbsent, maxValue);
                updateExtraActivityBar('defense', data.defenseFollowup, data.defenseSms, maxValue);
                document.getElementById('defense-total-value').textContent = formatNumber(data.defenseTotal);
                document.getElementById('defense-extra-value').textContent = formatNumber(data.defenseExtra);
                updateBarGroupOrder('defense', data.defenseTotal, data.defenseExtra);
                
                // 툴팁 업데이트 - 취소 방어
                document.getElementById('tt-defense-activity').textContent = formatNumber(data.defenseActivity);
                document.getElementById('tt-defense-absent').textContent = formatNumber(data.defenseAbsent);
                document.getElementById('tt-defense-total').textContent = formatNumber(data.defenseTotal);
                document.getElementById('tt-defense-followup').textContent = formatNumber(data.defenseFollowup);
                document.getElementById('tt-defense-sms').textContent = formatNumber(data.defenseSms);
                document.getElementById('tt-defense-extra').textContent = formatNumber(data.defenseExtra);
                
                if (data.rawText) {
                    document.getElementById('activityDataInput').value = data.rawText;
                }
            }
        }
    } catch (e) {
        console.error('활동수 데이터 로드 실패:', e);
    }
}

function loadCoverageData() {
    try {
        const savedData = localStorage.getItem(STORAGE_KEYS.coverage);
        if (savedData) {
            const data = JSON.parse(savedData);
            
            const today = new Date().toDateString();
            if (data.date === today) {
                updateKPICard('coverage-count', data.countRate, {
                    'success-count': formatNumber(data.successCount),
                    'contact-count': formatNumber(data.contactCount),
                    'unconverted-count': formatNumber(data.unconvertedCount)
                });
                
                updateKPICard('coverage-amount', data.amountRate, {
                    'success-amount': '₩' + formatNumber(data.successAmount),
                    'contact-amount': '₩' + formatNumber(data.contactAmount),
                    'unconverted-amount': '₩' + formatNumber(data.unconvertedAmount)
                });
                
                if (data.rawText) {
                    document.getElementById('coverageDataInput').value = data.rawText;
                }
            }
        }
    } catch (e) {
        console.error('커버리지 데이터 로드 실패:', e);
    }
}

function showToast(message, type = 'success') {
    const existingToast = document.querySelector('.toast');
    if (existingToast) {
        existingToast.remove();
    }
    
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    document.body.appendChild(toast);
    
    setTimeout(() => {
        toast.classList.add('show');
    }, 10);
    
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => {
            toast.remove();
        }, 300);
    }, 3000);
}

// ==========================================
// 이미지 캡처 및 요약 기능
// ==========================================

// 화면 캡처 후 클립보드에 복사
async function captureAndCopy() {
    try {
        showToast('캡처 중...');
        
        // 입력 패널과 모달 숨기기
        const inputPanels = document.querySelectorAll('.input-panel');
        const modal = document.getElementById('summaryModal');
        const hideOnCapture = document.querySelectorAll('.hide-on-capture');
        inputPanels.forEach(p => p.classList.remove('active'));
        modal.classList.remove('active');
        hideOnCapture.forEach(el => el.style.display = 'none');
        
        // 현재 활성 탭 컨텐츠만 캡처
        const target = document.querySelector('.tab-content.active');
        
        // html-to-image로 PNG blob 생성
        const blob = await htmlToImage.toBlob(target, {
            backgroundColor: '#f5f7fa',
            pixelRatio: 2
        });
        
        try {
            await navigator.clipboard.write([
                new ClipboardItem({ 'image/png': blob })
            ]);
            showToast('이미지가 클립보드에 복사되었습니다!');
        } catch (err) {
            console.error('클립보드 복사 실패:', err);
            showToast('클립보드 복사에 실패했습니다.', 'error');
        }
        
        // 숨긴 요소 복원
        hideOnCapture.forEach(el => el.style.display = '');
    } catch (err) {
        console.error('캡처 실패:', err);
        showToast('캡처에 실패했습니다.', 'error');
        // 오류 시에도 복원
        document.querySelectorAll('.hide-on-capture').forEach(el => el.style.display = '');
    }
}

// 화면 캡처 후 이미지로 저장
async function captureAndSave() {
    try {
        showToast('캡처 중...');
        
        // 입력 패널과 모달 숨기기
        const inputPanels = document.querySelectorAll('.input-panel');
        const modal = document.getElementById('summaryModal');
        const hideOnCapture = document.querySelectorAll('.hide-on-capture');
        inputPanels.forEach(p => p.classList.remove('active'));
        modal.classList.remove('active');
        hideOnCapture.forEach(el => el.style.display = 'none');
        
        // 현재 활성 탭 컨텐츠만 캡처
        const target = document.querySelector('.tab-content.active');
        
        // html-to-image로 PNG data URL 생성
        const dataUrl = await htmlToImage.toPng(target, {
            backgroundColor: '#f5f7fa',
            pixelRatio: 2
        });
        
        // 파일명 생성 (날짜 포함)
        const today = new Date();
        const dateStr = today.toISOString().split('T')[0];
        const filename = `KPI_Dashboard_${dateStr}.png`;
        
        // 다운로드 링크 생성
        const link = document.createElement('a');
        link.download = filename;
        link.href = dataUrl;
        link.click();
        
        showToast('이미지가 저장되었습니다!');
        
        // 숨긴 요소 복원
        hideOnCapture.forEach(el => el.style.display = '');
    } catch (err) {
        console.error('저장 실패:', err);
        showToast('저장에 실패했습니다.', 'error');
        // 오류 시에도 복원
        document.querySelectorAll('.hide-on-capture').forEach(el => el.style.display = '');
    }
}

// 요약 모달 표시
function showSummary() {
    const modal = document.getElementById('summaryModal');
    const textarea = document.getElementById('summaryText');
    
    const activeTab = document.querySelector('.tab-btn.active')?.dataset.tab || 'conversion';
    
    if (activeTab === 'payment') {
        textarea.value = generatePaymentSummary();
    } else if (activeTab === 'collection') {
        textarea.value = generateCollectionSummary();
    } else {
        textarea.value = generateSummary();
    }
    modal.classList.add('active');
}

// 요약 모달 닫기
function closeSummary() {
    const modal = document.getElementById('summaryModal');
    modal.classList.remove('active');
}

// 요약 텍스트 복사
async function copySummary() {
    const textarea = document.getElementById('summaryText');
    try {
        await navigator.clipboard.writeText(textarea.value);
        showToast('요약이 복사되었습니다!');
    } catch (err) {
        // 폴백: 구식 방법
        textarea.select();
        document.execCommand('copy');
        showToast('요약이 복사되었습니다!');
    }
}

// 금액을 억단위로 변환 (예: ₩8,945,453,599 → 89.45억, ₩25,434,487 → 0.25억)
function formatToEok(amountStr) {
    // ₩, 콤마 제거하고 숫자만 추출
    const numStr = amountStr.replace(/[₩,\s]/g, '');
    const num = parseInt(numStr, 10);
    
    if (isNaN(num) || num === 0) return '0억';
    
    const eok = num / 100000000; // 억 단위
    return `${eok.toFixed(2)}억`;
}

// KPI 요약 텍스트 생성
function generateSummary() {
    const today = new Date();
    const dateStr = `${today.getFullYear()}년 ${today.getMonth() + 1}월 ${today.getDate()}일`;
    
    let summary = `📊 Daily KPI 요약 (${dateStr})\n`;
    summary += `${'='.repeat(40)}\n\n`;
    
    // 성과 요약
    summary += `🏆 성과 요약\n`;
    summary += `-`.repeat(30) + `\n`;
    
    const perfApplyAmount = document.getElementById('perf-apply-amount')?.textContent || '₩0';
    const perfApplyCount = document.getElementById('perf-apply-count')?.textContent || '0건';
    const perfDefenseAmount = document.getElementById('perf-defense-amount')?.textContent || '₩0';
    const perfDefenseCount = document.getElementById('perf-defense-count')?.textContent || '0건';
    const perfTotalAmount = document.getElementById('perf-total-amount')?.textContent || '₩0';
    const perfTargetDisplay = document.getElementById('perf-target-display')?.textContent || '미설정';
    const perfProgressRate = document.getElementById('perf-progress-rate')?.textContent || '0%';
    
    summary += `• 신청전환 성공: ${formatToEok(perfApplyAmount)} (${perfApplyCount})\n`;
    summary += `• 취소방어 성공: ${formatToEok(perfDefenseAmount)} (${perfDefenseCount})\n`;
    summary += `• 합계: ${formatToEok(perfTotalAmount)}\n`;
    summary += `• 목표: ${perfTargetDisplay} / 진행률: ${perfProgressRate}\n\n`;
    
    // 커버리지 현황
    summary += `📌 커버리지 현황\n`;
    summary += `-`.repeat(30) + `\n`;
    
    const coverageCountRate = document.getElementById('coverage-count-rate')?.textContent || '0%';
    const coverageAmountRate = document.getElementById('coverage-amount-rate')?.textContent || '0%';
    const successCount = document.getElementById('success-count')?.textContent || '0';
    const contactCount = document.getElementById('contact-count')?.textContent || '0';
    const successAmount = document.getElementById('success-amount')?.textContent || '₩0';
    const contactAmount = document.getElementById('contact-amount')?.textContent || '₩0';
    
    summary += `• 건수 커버리지: ${coverageCountRate}\n`;
    summary += `  - 전환 성공: ${successCount}건 / 컨택 진행: ${contactCount}건\n`;
    summary += `• 금액 커버리지: ${coverageAmountRate}\n`;
    summary += `  - 전환 성공: ${formatToEok(successAmount)} / 컨택 진행: ${formatToEok(contactAmount)}\n\n`;
    
    // 활동수 현황
    summary += `📌 활동수 현황\n`;
    summary += `-`.repeat(30) + `\n`;
    
    const applyTotalValue = document.getElementById('apply-total-value')?.textContent || '0';
    const applyExtraValue = document.getElementById('apply-extra-value')?.textContent || '0';
    const defenseTotalValue = document.getElementById('defense-total-value')?.textContent || '0';
    const defenseExtraValue = document.getElementById('defense-extra-value')?.textContent || '0';
    
    summary += `• 신청 전환 활동수\n`;
    summary += `  - 총 활동: ${applyTotalValue} / 부가활동: ${applyExtraValue}\n`;
    summary += `• 취소 방어 활동수\n`;
    summary += `  - 총 활동: ${defenseTotalValue} / 부가활동: ${defenseExtraValue}\n\n`;
    
    // 총조회대비 신청전환
    summary += `📌 총조회대비 신청전환\n`;
    summary += `-`.repeat(30) + `\n`;
    
    const applyTotalRate = document.getElementById('apply-total-rate-rate')?.textContent || '0%';
    const applySuccessRate = document.getElementById('apply-success-rate-rate')?.textContent || '0%';
    const totalViewAmount = document.getElementById('total-view-amount')?.textContent || '₩0';
    const totalApplyAmount = document.getElementById('total-apply-amount')?.textContent || '₩0';
    const applyConvertAmount = document.getElementById('apply-convert-amount')?.textContent || '₩0';
    
    summary += `• 총조회 대비 전체 신청률: ${applyTotalRate}\n`;
    summary += `  - 총조회: ${formatToEok(totalViewAmount)} / 조회 신청: ${formatToEok(totalApplyAmount)}\n`;
    summary += `• 총조회 대비 전환 성공률: ${applySuccessRate}\n`;
    summary += `  - 총조회: ${formatToEok(totalViewAmount)} / 전환 성공: ${formatToEok(applyConvertAmount)}\n\n`;
    
    // 총검토대비 취소방어
    summary += `📌 총검토대비 취소방어\n`;
    summary += `-`.repeat(30) + `\n`;
    
    const cancelReviewRate = document.getElementById('cancel-review-rate-rate')?.textContent || '0%';
    const cancelDefenseRate = document.getElementById('cancel-defense-rate-rate')?.textContent || '0%';
    const cancelRequestAmount = document.getElementById('cancel-request-amount')?.textContent || '₩0';
    const cancelAvailableAmount = document.getElementById('cancel-available-amount')?.textContent || '₩0';
    const cancelSuccessAmount = document.getElementById('cancel-success-amount')?.textContent || '₩0';
    
    summary += `• 전체취소 대비 검토완료율: ${cancelReviewRate}\n`;
    summary += `  - 전체취소: ${formatToEok(cancelRequestAmount)} / 검토완료: ${formatToEok(cancelAvailableAmount)}\n`;
    summary += `• 검토완료 대비 방어 성공률: ${cancelDefenseRate}\n`;
    summary += `  - 검토완료: ${formatToEok(cancelAvailableAmount)} / 취소방어 성공: ${formatToEok(cancelSuccessAmount)}\n`;
    
    return summary;
}

// 결제파트 요약 텍스트 생성
function generatePaymentSummary() {
    const today = new Date();
    const dateStr = `${today.getFullYear()}년 ${today.getMonth() + 1}월 ${today.getDate()}일`;

    let summary = `💳 결제파트 요약 (${dateStr})\n`;
    summary += `${'='.repeat(40)}\n\n`;

    const monthLabel = document.getElementById('selected-payment-month')?.textContent || '';
    const monthAmount = document.getElementById('monthly-payment-amount')?.textContent || '₩0';
    const monthCount = document.getElementById('monthly-payment-count')?.textContent || '0건';

    summary += `📌 ${monthLabel || '월별'} 결제금액\n`;
    summary += `-`.repeat(30) + `\n`;
    summary += `• 결제금액: ${formatToEok(monthAmount)} (${monthCount})\n\n`;

    const weekLabel = document.getElementById('selected-week')?.textContent || '';
    const refundCount = document.getElementById('weekly-refund-count')?.textContent || '0건';
    const refundAmount = document.getElementById('weekly-refund-amount')?.textContent || '₩0';
    const samedayCount = document.getElementById('weekly-sameday-count')?.textContent || '0건';
    const samedayRate = document.getElementById('weekly-sameday-rate')?.textContent || '0%';
    const within30dCount = document.getElementById('weekly-30d-count')?.textContent || '0건';
    const within30dRate = document.getElementById('weekly-30d-rate')?.textContent || '0%';

    summary += `📌 주차별 결제 KPI (${weekLabel})\n`;
    summary += `-`.repeat(30) + `\n`;
    summary += `• 환급 완료: ${refundCount} / ${formatToEok(refundAmount)}\n`;
    summary += `• 당일 결제: ${samedayCount} (${samedayRate})\n`;
    summary += `• 30일이내 결제: ${within30dCount} (${within30dRate})\n\n`;

    summary += `📌 기간별 결제율\n`;
    summary += `-`.repeat(30) + `\n`;
    const periods = [0, 3, 7, 21, 30, 60];
    const periodLabels = { 0: '당일', 3: '3일', 7: '7일', 21: '21일', 30: '30일', 60: '60일' };
    periods.forEach(d => {
        const cntRate = document.getElementById(`weekly-rate-cnt-${d}d`)?.textContent || '-';
        const amtRate = document.getElementById(`weekly-rate-amt-${d}d`)?.textContent || '-';
        summary += `• ${periodLabels[d]}이내: 건수 ${cntRate} / 금액 ${amtRate}\n`;
    });

    return summary;
}

// 추심 요약 텍스트 생성
function generateCollectionSummary() {
    const today = new Date();
    const dateStr = `${today.getFullYear()}년 ${today.getMonth() + 1}월 ${today.getDate()}일`;

    let summary = `⚖️ 추심 요약 (${dateStr})\n`;
    summary += `${'='.repeat(40)}\n\n`;

    const monthLabel = document.getElementById('selected-collection-month')?.textContent || '';

    const refundAmount = document.getElementById('col-refund-amount')?.textContent || '₩0';
    const refundCount = document.getElementById('col-refund-count')?.textContent || '0건';
    const transferAmount = document.getElementById('col-transfer-amount')?.textContent || '₩0';
    const transferCount = document.getElementById('col-transfer-count')?.textContent || '0건';
    const ratioValue = document.getElementById('col-ratio-value')?.textContent || '0%';

    summary += `📌 추심 KPI (${monthLabel})\n`;
    summary += `-`.repeat(30) + `\n`;
    summary += `• 환급완료: ${formatToEok(refundAmount)} (${refundCount})\n`;
    summary += `• 이관총액: ${formatToEok(transferAmount)} (${transferCount})\n`;
    summary += `• 이관비율: ${ratioValue}\n\n`;

    const yearLabel = document.getElementById('selected-collection-year')?.textContent || '';
    summary += `📌 연간 이관 결제 추적 (${yearLabel})\n`;
    summary += `-`.repeat(30) + `\n`;

    const rows = document.querySelectorAll('#yearly-collection-body .yt-row');
    rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        if (cells.length < 3) return;
        const monthTd = cells[0]?.textContent?.trim() || '';
        const transferTd = cells[1]?.textContent?.trim().replace(/\s+/g, ' ') || '';
        const paidTd = cells[2]?.textContent?.trim().replace(/\s+/g, ' ') || '';
        if (transferTd === '- 0건') return;
        summary += `• ${monthTd}: 이관 ${transferTd} / 성사 ${paidTd}\n`;
    });

    return summary;
}

// ==========================================
// 히스토리 (스냅샷) 기능
// ==========================================

// 오늘 날짜 키 생성
function getTodayKey() {
    const today = new Date();
    return today.toISOString().split('T')[0]; // YYYY-MM-DD
}

// 현재 데이터를 스냅샷으로 저장
function saveSnapshot() {
    const dateKey = getTodayKey();
    
    // 현재 모든 데이터 수집 (STORAGE_KEYS 사용)
    const snapshot = {
        date: dateKey,
        timestamp: new Date().toISOString(),
        performance: JSON.parse(localStorage.getItem(STORAGE_KEYS.performance) || '{}'),
        coverage: JSON.parse(localStorage.getItem(STORAGE_KEYS.coverage) || '{}'),
        activity: JSON.parse(localStorage.getItem(STORAGE_KEYS.activity) || '{}'),
        application: JSON.parse(localStorage.getItem(STORAGE_KEYS.application) || '{}'),
        defense: JSON.parse(localStorage.getItem(STORAGE_KEYS.defense) || '{}')
    };
    
    // 기존 스냅샷 목록 가져오기
    const snapshots = JSON.parse(localStorage.getItem('kpi_snapshots') || '{}');
    
    // 오늘 날짜로 저장 (덮어쓰기)
    snapshots[dateKey] = snapshot;
    
    // 저장
    localStorage.setItem('kpi_snapshots', JSON.stringify(snapshots));
    
    console.log('스냅샷 저장됨:', dateKey);
}

// 스냅샷 불러오기
function loadSnapshot(dateKey) {
    const snapshots = JSON.parse(localStorage.getItem('kpi_snapshots') || '{}');
    const snapshot = snapshots[dateKey];
    
    if (!snapshot) {
        showToast('해당 날짜의 데이터가 없습니다.', 'error');
        return;
    }
    
    // 각 데이터 복원 (STORAGE_KEYS 사용)
    if (snapshot.performance && Object.keys(snapshot.performance).length > 0) {
        localStorage.setItem(STORAGE_KEYS.performance, JSON.stringify(snapshot.performance));
    }
    if (snapshot.coverage && Object.keys(snapshot.coverage).length > 0) {
        localStorage.setItem(STORAGE_KEYS.coverage, JSON.stringify(snapshot.coverage));
    }
    if (snapshot.activity && Object.keys(snapshot.activity).length > 0) {
        localStorage.setItem(STORAGE_KEYS.activity, JSON.stringify(snapshot.activity));
    }
    if (snapshot.application && Object.keys(snapshot.application).length > 0) {
        localStorage.setItem(STORAGE_KEYS.application, JSON.stringify(snapshot.application));
    }
    if (snapshot.defense && Object.keys(snapshot.defense).length > 0) {
        localStorage.setItem(STORAGE_KEYS.defense, JSON.stringify(snapshot.defense));
    }
    
    // 화면 갱신
    loadPerformanceData();
    loadCoverageData();
    loadActivityData();
    loadApplyConversionData();
    loadCancelDefenseData();
    
    closeHistory();
    showToast(`${formatDateKorean(dateKey)} 데이터를 불러왔습니다!`);
}

// 스냅샷 삭제
function deleteSnapshot(dateKey) {
    if (!confirm(`${formatDateKorean(dateKey)} 데이터를 삭제하시겠습니까?`)) {
        return;
    }
    
    const snapshots = JSON.parse(localStorage.getItem('kpi_snapshots') || '{}');
    delete snapshots[dateKey];
    localStorage.setItem('kpi_snapshots', JSON.stringify(snapshots));
    
    showHistory(); // 목록 갱신
    showToast('삭제되었습니다.');
}

// 날짜를 한국어로 포맷
function formatDateKorean(dateKey) {
    const date = new Date(dateKey);
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const dayNames = ['일', '월', '화', '수', '목', '금', '토'];
    const dayName = dayNames[date.getDay()];
    return `${year}년 ${month}월 ${day}일 (${dayName})`;
}

// 히스토리 모달 표시
function showHistory() {
    const modal = document.getElementById('historyModal');
    const listContainer = document.getElementById('historyList');
    
    const snapshots = JSON.parse(localStorage.getItem('kpi_snapshots') || '{}');
    const dates = Object.keys(snapshots).sort().reverse(); // 최신 날짜 먼저
    
    if (dates.length === 0) {
        listContainer.innerHTML = `
            <div class="history-empty">
                <div class="history-empty-icon">📭</div>
                <p>저장된 히스토리가 없습니다.</p>
                <p style="font-size: 12px; margin-top: 8px;">데이터를 입력하면 자동으로 저장됩니다.</p>
            </div>
        `;
    } else {
        const todayKey = getTodayKey();
        listContainer.innerHTML = dates.map(dateKey => {
            const snapshot = snapshots[dateKey];
            const isToday = dateKey === todayKey;
            const time = new Date(snapshot.timestamp).toLocaleTimeString('ko-KR', { 
                hour: '2-digit', 
                minute: '2-digit' 
            });
            
            return `
                <div class="history-item ${isToday ? 'active' : ''}">
                    <div>
                        <div class="history-date">${formatDateKorean(dateKey)}</div>
                        <div class="history-info">마지막 저장: ${time} ${isToday ? '(오늘)' : ''}</div>
                    </div>
                    <div class="history-actions">
                        <button class="history-btn history-btn-load" onclick="loadSnapshot('${dateKey}')">불러오기</button>
                        <button class="history-btn history-btn-delete" onclick="deleteSnapshot('${dateKey}')">삭제</button>
                    </div>
                </div>
            `;
        }).join('');
    }
    
    modal.classList.add('active');
}

// 히스토리 모달 닫기
function closeHistory() {
    const modal = document.getElementById('historyModal');
    modal.classList.remove('active');
}

// ==========================================
// 성과 요약 (자동 로드 - 필터 기반)
// ==========================================

// Google Apps Script 웹 앱 URL
const PERFORMANCE_API_URL = 'https://script.google.com/macros/s/AKfycbyquFPhmxs4O7kyAAkF6MsvXYIluufzkqzcGHI_yZtp80NosTVVD1IuvTB3n9n81frbAg/exec';

// 성과 요약 초기화
function initPerformanceSummary() {
    updatePerfMonthDisplay();
    loadTargetAmount();  // 저장된 목표금액 로드
    loadPerformanceSummary();
}

// 월 표시 업데이트 (현재 월 표시)
function updatePerfMonthDisplay() {
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1;
    
    const monthEl = document.getElementById('selected-perf-month');
    if (monthEl) {
        monthEl.textContent = `${year}년 ${month}월`;
    }
}

// 성과 데이터 로드 (필터 기반 - Pipedrive 필터가 "이번 달"로 설정됨)
async function loadPerformanceSummary() {
    const loadingEl = document.getElementById('perf-loading');
    const errorEl = document.getElementById('perf-error');
    const contentEl = document.getElementById('perf-content');
    
    // 로딩 상태 표시
    if (loadingEl) loadingEl.style.display = 'flex';
    if (errorEl) errorEl.style.display = 'none';
    if (contentEl) contentEl.style.display = 'none';
    
    try {
        const url = `${PERFORMANCE_API_URL}?action=performance`;
        console.log('성과 API 호출:', url);
        
        const response = await fetch(url);
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const result = await response.json();
        console.log('성과 API 응답:', result);
        
        if (!result.success) {
            throw new Error('성과 데이터를 가져오는데 실패했습니다.');
        }
        
        // UI 업데이트
        updatePerformanceSummaryUI(result.data);
        
        // 마지막 동기화 시간 표시
        const syncTimeEl = document.getElementById('perf-sync-time');
        if (syncTimeEl) {
            syncTimeEl.textContent = new Date().toLocaleString('ko-KR');
        }
        
        // 컨텐츠 표시
        if (loadingEl) loadingEl.style.display = 'none';
        if (contentEl) contentEl.style.display = 'block';
        
    } catch (error) {
        console.error('성과 데이터 로드 실패:', error);
        
        if (loadingEl) loadingEl.style.display = 'none';
        if (errorEl) {
            errorEl.style.display = 'flex';
            const errorMsgEl = document.getElementById('perf-error-message');
            if (errorMsgEl) {
                errorMsgEl.textContent = error.message || '네트워크 오류가 발생했습니다.';
            }
        }
    }
}

// 성과 요약 UI 업데이트
function updatePerformanceSummaryUI(data) {
    // 신청전환 성공
    const applyAmount = data.apply?.amount || 0;
    const applyCount = data.apply?.count || 0;
    document.getElementById('perf-apply-amount').textContent = '₩' + formatNumber(applyAmount);
    document.getElementById('perf-apply-count').textContent = applyCount + '건';
    
    // 취소방어 성공
    const defenseAmount = data.defense?.amount || 0;
    const defenseCount = data.defense?.count || 0;
    document.getElementById('perf-defense-amount').textContent = '₩' + formatNumber(defenseAmount);
    document.getElementById('perf-defense-count').textContent = defenseCount + '건';
    
    // 합계
    const totalAmount = data.total || (applyAmount + defenseAmount);
    document.getElementById('perf-total-amount').textContent = '₩' + formatNumber(totalAmount);
    document.getElementById('perf-apply-summary').textContent = '₩' + formatNumber(applyAmount);
    document.getElementById('perf-defense-summary').textContent = '₩' + formatNumber(defenseAmount);
    
    // 목표 대비 진행률 업데이트
    updateProgressRate(totalAmount);
}

// 목표금액 저장
async function saveTargetAmount() {
    const input = document.getElementById('perf-target-input');
    const inputValue = input.value.trim();
    
    // 억 단위 파싱 (예: "5억" → 500000000)
    let targetAmount = parseTargetInput(inputValue);
    
    // localStorage에 저장 (백업용)
    localStorage.setItem('perf_target_amount', targetAmount.toString());
    
    // Google Sheets에 저장
    showToast('저장 중...');
    const saveData = { amount: targetAmount };
    const saved = await saveManualDataToSheets('target', saveData);
    
    // 현재 합계로 진행률 업데이트
    const totalText = document.getElementById('perf-total-amount').textContent;
    const totalAmount = parseNumber(totalText.replace('₩', '').replace(/,/g, ''));
    updateProgressRate(totalAmount);
    
    if (saved) {
        showToast('목표금액이 저장되었습니다!');
    } else {
        showToast('저장에 실패했습니다. 로컬에만 저장됩니다.', 'error');
    }
}

// 목표금액 입력 파싱 (억 단위 지원)
function parseTargetInput(value) {
    if (!value) return 0;
    
    // "5억", "5.5억", "55000만" 등 파싱
    value = value.replace(/,/g, '').replace(/\s/g, '');
    
    if (value.includes('억')) {
        const num = parseFloat(value.replace('억', ''));
        return Math.round(num * 100000000);
    } else if (value.includes('만')) {
        const num = parseFloat(value.replace('만', ''));
        return Math.round(num * 10000);
    } else {
        return parseFloat(value) || 0;
    }
}

// 목표금액 로드
function loadTargetAmount() {
    const saved = localStorage.getItem('perf_target_amount');
    if (saved) {
        const amount = parseInt(saved);
        const input = document.getElementById('perf-target-input');
        if (input && amount > 0) {
            // 억 단위로 표시
            const eok = amount / 100000000;
            if (eok >= 1) {
                input.value = eok % 1 === 0 ? `${eok}억` : `${eok.toFixed(1)}억`;
            } else {
                input.value = formatNumber(amount);
            }
        }
        return amount;
    }
    return 0;
}

// 진행률 업데이트
function updateProgressRate(totalAmount) {
    const targetAmount = loadTargetAmount();
    
    // 목표 금액 표시
    const targetDisplay = document.getElementById('perf-target-display');
    if (targetDisplay) {
        if (targetAmount > 0) {
            const eok = targetAmount / 100000000;
            targetDisplay.textContent = eok >= 1 ? `${eok.toFixed(1)}억` : '₩' + formatNumber(targetAmount);
        } else {
            targetDisplay.textContent = '미설정';
        }
    }
    
    // 진행률 계산
    const progressRate = targetAmount > 0 ? (totalAmount / targetAmount) * 100 : 0;
    
    // 진행률 표시
    const rateEl = document.getElementById('perf-progress-rate');
    if (rateEl) {
        rateEl.textContent = progressRate.toFixed(1) + '%';
        rateEl.style.color = '';
        if (progressRate >= 100) {
            rateEl.style.color = 'var(--accent-success)';
        } else if (progressRate >= 80) {
            rateEl.style.color = 'var(--accent-primary)';
        } else if (progressRate >= 50) {
            rateEl.style.color = 'var(--accent-warning)';
        } else if (progressRate > 0) {
            rateEl.style.color = 'var(--accent-danger)';
        }
    }
    
    // 진행률 바
    const progressFill = document.getElementById('perf-progress-fill');
    if (progressFill) {
        progressFill.style.width = Math.min(progressRate, 100) + '%';
    }
}

// ==========================================
// 수동 입력 데이터 (Google Sheets 저장/로드)
// ==========================================

// Google Sheets에서 수동 데이터 로드
async function loadManualDataFromSheets() {
    try {
        const url = `${PERFORMANCE_API_URL}?action=manual`;
        console.log('수동 데이터 API 호출:', url);
        
        const response = await fetch(url);
        const result = await response.json();
        
        console.log('수동 데이터 API 응답:', result);
        
        if (!result.success) {
            console.error('수동 데이터 로드 실패');
            return;
        }
        
        const data = result.data;
        
        // 커버리지 데이터 적용
        if (data.coverage) {
            applyCoverageFromData(data.coverage);
        }
        
        // 신청전환 데이터 적용
        if (data.application) {
            applyApplicationFromData(data.application);
        }
        
        // 취소방어 데이터 적용
        if (data.defense) {
            applyDefenseFromData(data.defense);
        }
        
        // 목표금액 데이터 적용
        if (data.target && data.target.amount) {
            applyTargetFromData(data.target);
        }
        
        console.log('수동 데이터 로드 완료');
        
    } catch (error) {
        console.error('수동 데이터 로드 오류:', error);
    }
}

// Google Sheets에 수동 데이터 저장 (GET 요청으로 변경 - CORS 우회)
async function saveManualDataToSheets(type, data) {
    try {
        // 데이터를 URL 파라미터로 인코딩
        const encodedData = encodeURIComponent(JSON.stringify(data));
        const url = `${PERFORMANCE_API_URL}?action=saveManual&type=${type}&data=${encodedData}`;
        
        console.log(`${type} 저장 요청:`, url.substring(0, 200) + '...');
        
        const response = await fetch(url);
        const result = await response.json();
        
        console.log(`${type} 저장 결과:`, result);
        
        return result.success;
        
    } catch (error) {
        console.error(`${type} 저장 오류:`, error);
        return false;
    }
}

// 커버리지 데이터 적용 (API에서 로드된 데이터)
function applyCoverageFromData(data) {
    updateKPICard('coverage-count', data.countRate || 0, {
        'success-count': formatNumber(data.successCount || 0),
        'contact-count': formatNumber(data.contactCount || 0),
        'unconverted-count': formatNumber(data.unconvertedCount || 0)
    });
    
    updateKPICard('coverage-amount', data.amountRate || 0, {
        'success-amount': '₩' + formatNumber(data.successAmount || 0),
        'contact-amount': '₩' + formatNumber(data.contactAmount || 0),
        'unconverted-amount': '₩' + formatNumber(data.unconvertedAmount || 0)
    });
}

// 신청전환 데이터 적용 (API에서 로드된 데이터)
function applyApplicationFromData(data) {
    updateKPICard('apply-total-rate', data.totalApplyRate || 0, {
        'total-view-amount': '₩' + formatNumber(data.totalView || 0),
        'total-apply-amount': '₩' + formatNumber(data.totalApply || 0)
    });
    
    updateKPICard('apply-success-rate', data.applySuccessRate || 0, {
        'total-view-amount2': '₩' + formatNumber(data.totalView || 0),
        'apply-convert-amount': '₩' + formatNumber(data.applyConvert || 0)
    });
}

// 취소방어 데이터 적용 (API에서 로드된 데이터)
function applyDefenseFromData(data) {
    updateKPICard('cancel-review-rate', data.reviewRate || 0, {
        'cancel-request-amount': '₩' + formatNumber(data.cancelRequest || 0),
        'cancel-available-amount': '₩' + formatNumber(data.cancelAvailable || 0)
    });
    
    updateKPICard('cancel-defense-rate', data.defenseRate || 0, {
        'cancel-available-amount2': '₩' + formatNumber(data.cancelAvailable || 0),
        'cancel-success-amount': '₩' + formatNumber(data.cancelSuccess || 0)
    });
}

// 목표금액 데이터 적용 (API에서 로드된 데이터)
function applyTargetFromData(data) {
    const amount = data.amount || 0;
    
    // localStorage에도 저장 (백업)
    localStorage.setItem('perf_target_amount', amount.toString());
    
    // 입력 필드 업데이트
    const input = document.getElementById('perf-target-input');
    if (input && amount > 0) {
        const eok = amount / 100000000;
        if (eok >= 1) {
            input.value = eok % 1 === 0 ? `${eok}억` : `${eok.toFixed(1)}억`;
        } else {
            input.value = formatNumber(amount);
        }
    }
    
    // 진행률 업데이트 (현재 합계 기준)
    const totalText = document.getElementById('perf-total-amount')?.textContent || '₩0';
    const totalAmount = parseNumber(totalText.replace('₩', '').replace(/,/g, ''));
    updateProgressRate(totalAmount);
    
    console.log('목표금액 적용:', amount);
}

// ==========================================
// 활동수 현황 (자동 로드 - 필터 기반)
// ==========================================

// 활동수 데이터 로드
async function loadActivityStatus() {
    const loadingEl = document.getElementById('activity-loading');
    const errorEl = document.getElementById('activity-error');
    const contentEl = document.getElementById('activity-content');
    
    // 로딩 상태 표시
    if (loadingEl) loadingEl.style.display = 'flex';
    if (errorEl) errorEl.style.display = 'none';
    if (contentEl) contentEl.style.display = 'none';
    
    try {
        const url = `${PERFORMANCE_API_URL}?action=activity`;
        console.log('활동수 API 호출:', url);
        
        const response = await fetch(url);
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const result = await response.json();
        console.log('활동수 API 응답:', result);
        
        if (!result.success) {
            throw new Error('활동수 데이터를 가져오는데 실패했습니다.');
        }
        
        // UI 업데이트
        updateActivityStatusUI(result.data);
        
        // 마지막 동기화 시간 표시
        const syncTimeEl = document.getElementById('activity-sync-time');
        if (syncTimeEl) {
            syncTimeEl.textContent = new Date().toLocaleString('ko-KR');
        }
        
        // 컨텐츠 표시
        if (loadingEl) loadingEl.style.display = 'none';
        if (contentEl) contentEl.style.display = 'block';
        
    } catch (error) {
        console.error('활동수 데이터 로드 실패:', error);
        
        if (loadingEl) loadingEl.style.display = 'none';
        if (errorEl) {
            errorEl.style.display = 'flex';
            const errorMsgEl = document.getElementById('activity-error-message');
            if (errorMsgEl) {
                errorMsgEl.textContent = error.message || '네트워크 오류가 발생했습니다.';
            }
        }
    }
}

// 활동수 UI 업데이트
function updateActivityStatusUI(data) {
    const apply = data.apply || {};
    const defense = data.defense || {};
    
    // 신청 전환
    const applyActivity = apply.activity || 0;  // 줍줍콜 + VIP
    const applyAbsent = apply.absent || 0;      // 부재
    const applyFollowup = apply.followup || 0;  // 사후관리
    const applySms = apply.sms || 0;            // 문자
    
    const applyTotal = applyActivity + applyAbsent;
    const applyExtra = applyFollowup + applySms;
    
    // 취소 방어 (필터 미설정 시 0)
    const defenseActivity = defense.activity || 0;
    const defenseAbsent = defense.absent || 0;
    const defenseFollowup = defense.followup || 0;
    const defenseSms = defense.sms || 0;
    
    const defenseTotal = defenseActivity + defenseAbsent;
    const defenseExtra = defenseFollowup + defenseSms;
    
    // 최대값 계산 (바 높이 비율용)
    const maxValue = Math.max(applyTotal, applyExtra, defenseTotal, defenseExtra, 1);
    
    // UI 업데이트 - 신청 전환 (스택 바)
    updateTotalActivityBar('apply', applyActivity, applyAbsent, maxValue);
    updateExtraActivityBar('apply', applyFollowup, applySms, maxValue);
    document.getElementById('apply-total-value').textContent = formatNumber(applyTotal);
    document.getElementById('apply-extra-value').textContent = formatNumber(applyExtra);
    updateBarGroupOrder('apply', applyTotal, applyExtra);
    
    // 툴팁 업데이트 - 신청 전환
    document.getElementById('tt-apply-activity').textContent = formatNumber(applyActivity);
    document.getElementById('tt-apply-absent').textContent = formatNumber(applyAbsent);
    document.getElementById('tt-apply-total').textContent = formatNumber(applyTotal);
    document.getElementById('tt-apply-followup').textContent = formatNumber(applyFollowup);
    document.getElementById('tt-apply-sms').textContent = formatNumber(applySms);
    document.getElementById('tt-apply-extra').textContent = formatNumber(applyExtra);
    
    // UI 업데이트 - 취소 방어 (스택 바)
    updateTotalActivityBar('defense', defenseActivity, defenseAbsent, maxValue);
    updateExtraActivityBar('defense', defenseFollowup, defenseSms, maxValue);
    document.getElementById('defense-total-value').textContent = formatNumber(defenseTotal);
    document.getElementById('defense-extra-value').textContent = formatNumber(defenseExtra);
    updateBarGroupOrder('defense', defenseTotal, defenseExtra);
    
    // 툴팁 업데이트 - 취소 방어
    document.getElementById('tt-defense-activity').textContent = formatNumber(defenseActivity);
    document.getElementById('tt-defense-absent').textContent = formatNumber(defenseAbsent);
    document.getElementById('tt-defense-total').textContent = formatNumber(defenseTotal);
    document.getElementById('tt-defense-followup').textContent = formatNumber(defenseFollowup);
    document.getElementById('tt-defense-sms').textContent = formatNumber(defenseSms);
    document.getElementById('tt-defense-extra').textContent = formatNumber(defenseExtra);
}

// ==========================================
// 결제파트 대시보드 (주차별 계획)
// ==========================================

// Google Apps Script 웹 앱 URL (결제용 - 같은 URL 사용)
const PAYMENT_API_URL = PERFORMANCE_API_URL;

// 결제+추심 공용 데이터 캐시 (같은 API에서 가져오므로 공유)
let sharedDataPromise = null;
async function fetchSharedData() {
    if (!sharedDataPromise) {
        sharedDataPromise = (async () => {
            const maxRetries = 3;
            for (let attempt = 1; attempt <= maxRetries; attempt++) {
                try {
                    const response = await fetch(PAYMENT_API_URL);
                    if (!response.ok) throw new Error(`HTTP ${response.status}`);
                    const result = await response.json();
                    if (!result.success) throw new Error('API 실패');
                    return result.data;
                } catch (e) {
                    console.warn(`데이터 로드 시도 ${attempt}/${maxRetries} 실패:`, e.message);
                    if (attempt === maxRetries) {
                        sharedDataPromise = null;
                        throw e;
                    }
                    await new Promise(r => setTimeout(r, 1000 * attempt));
                }
            }
        })();
    }
    return sharedDataPromise;
}

// 결제 데이터 캐시
let paymentDataCache = null;
let selectedWeekStart = getWeekStart(new Date()); // 현재 주의 월요일
let selectedPaymentMonth = new Date(); // 현재 월

// 결제 목표
const PAYMENT_TARGETS = {
    3: 80,   // 3일 이내 80%
    30: 90,  // 30일 이내 90%
    60: 95   // 60일 이내 95%
};

// API 날짜를 KST로 변환 (날짜 문자열 → KST Date 객체)
function parseToKST(dateString) {
    if (!dateString) return null;
    
    // 이미 Date 객체인 경우
    if (dateString instanceof Date) return dateString;
    
    // ISO 형식 또는 기타 형식 파싱
    const date = new Date(dateString);
    
    // 유효하지 않은 날짜
    if (isNaN(date.getTime())) return null;
    
    // UTC로 해석된 시간을 KST로 보정
    // API가 UTC로 내려준다고 가정하고, 9시간 추가
    // 단, 날짜만 있는 경우(T가 없는 경우) 그대로 사용
    if (typeof dateString === 'string' && !dateString.includes('T')) {
        // 날짜만 있는 경우 (예: "2026-01-26") - 로컬 타임존으로 해석
        const parts = dateString.split('-');
        if (parts.length === 3) {
            return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
        }
    }
    
    // ISO 형식인 경우 UTC+9 적용
    if (typeof dateString === 'string' && dateString.includes('T')) {
        const utcTime = date.getTime();
        const kstOffset = 9 * 60 * 60 * 1000; // 9시간
        return new Date(utcTime + kstOffset);
    }
    
    return date;
}

// KST 기준 날짜만 비교 (시간 무시)
function isSameDateKST(date1, date2) {
    if (!date1 || !date2) return false;
    const d1 = parseToKST(date1);
    const d2 = parseToKST(date2);
    if (!d1 || !d2) return false;
    
    return d1.getFullYear() === d2.getFullYear() &&
           d1.getMonth() === d2.getMonth() &&
           d1.getDate() === d2.getDate();
}

// KST 기준 일수 차이 계산
function getDaysDiffKST(startDate, endDate) {
    const start = parseToKST(startDate);
    const end = parseToKST(endDate);
    if (!start || !end) return null;
    
    // 날짜만 비교 (시간 무시)
    const startDay = new Date(start.getFullYear(), start.getMonth(), start.getDate());
    const endDay = new Date(end.getFullYear(), end.getMonth(), end.getDate());
    
    return Math.floor((endDay - startDay) / (1000 * 60 * 60 * 24));
}

// 주의 월요일 구하기
function getWeekStart(date) {
    const d = new Date(date);
    const day = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1); // 월요일로 조정
    d.setDate(diff);
    d.setHours(0, 0, 0, 0);
    return d;
}

// 주의 일요일 구하기
function getWeekEnd(weekStart) {
    const d = new Date(weekStart);
    d.setDate(d.getDate() + 6);
    d.setHours(23, 59, 59, 999);
    return d;
}

// 결제파트 탭 초기화
function initPaymentTab() {
    // 탭 클릭 시 데이터 로드
    const paymentTabBtn = document.querySelector('[data-tab="payment"]');
    if (paymentTabBtn) {
        paymentTabBtn.addEventListener('click', () => {
            if (!paymentDataCache) {
                loadPaymentData();
            }
        });
    }
    
    // 주차 표시 업데이트
    updateWeekDisplay();
}

// 주차 표시 업데이트
function updateWeekDisplay() {
    const weekEl = document.getElementById('selected-week');
    if (weekEl) {
        const weekEnd = getWeekEnd(selectedWeekStart);
        const formatDate = (d) => {
            const yy = String(d.getFullYear()).slice(2);
            const mm = String(d.getMonth() + 1).padStart(2, '0');
            const dd = String(d.getDate()).padStart(2, '0');
            return `${yy}/${mm}/${dd}`;
        };
        weekEl.textContent = `${formatDate(selectedWeekStart)} ~ ${formatDate(weekEnd)}`;
    }
}

// 주차 변경
function changeWeek(delta) {
    selectedWeekStart.setDate(selectedWeekStart.getDate() + (delta * 7));
    updateWeekDisplay();
    
    if (paymentDataCache) {
        calculateAndDisplayWeeklyKPIs(paymentDataCache);
    }
}

// 결제 데이터 로드
async function loadPaymentData() {
    const loadingEl = document.getElementById('payment-loading');
    const errorEl = document.getElementById('payment-error');
    const contentEl = document.getElementById('payment-content');
    
    // 로딩 상태 표시
    if (loadingEl) loadingEl.style.display = 'flex';
    if (errorEl) errorEl.style.display = 'none';
    if (contentEl) contentEl.style.display = 'none';
    
    try {
        paymentDataCache = await fetchSharedData();
        
        // 월별 표시 초기화 및 계산
        updatePaymentMonthDisplay();
        calculateAndDisplayMonthlyPayment(paymentDataCache);
        
        // 주차별 KPI 계산 및 표시
        calculateAndDisplayWeeklyKPIs(paymentDataCache);
        
        // 마지막 동기화 시간 표시
        const syncTimeEl = document.getElementById('last-sync-time');
        if (syncTimeEl) {
            syncTimeEl.textContent = new Date().toLocaleString('ko-KR');
        }
        
        // 컨텐츠 표시
        if (loadingEl) loadingEl.style.display = 'none';
        if (contentEl) contentEl.style.display = 'block';
        
    } catch (error) {
        console.error('결제 데이터 로드 실패:', error);
        
        if (loadingEl) loadingEl.style.display = 'none';
        if (errorEl) {
            errorEl.style.display = 'flex';
            const errorMsgEl = document.getElementById('payment-error-message');
            if (errorMsgEl) {
                errorMsgEl.textContent = error.message || '네트워크 오류가 발생했습니다.';
            }
        }
    }
}

// ==========================================
// 월별 결제금액 섹션
// ==========================================

// 월 선택 변경
function changePaymentMonth(delta) {
    selectedPaymentMonth.setMonth(selectedPaymentMonth.getMonth() + delta);
    updatePaymentMonthDisplay();
    
    if (paymentDataCache) {
        calculateAndDisplayMonthlyPayment(paymentDataCache);
    }
}

// 월 표시 업데이트
function updatePaymentMonthDisplay() {
    const year = selectedPaymentMonth.getFullYear();
    const month = selectedPaymentMonth.getMonth() + 1;
    
    document.getElementById('selected-payment-month').textContent = `${year}년 ${month}월`;
    document.getElementById('monthly-payment-title').textContent = `${month}월 결제금액`;
}

// 월별 결제금액 계산 및 표시
function calculateAndDisplayMonthlyPayment(data) {
    const year = selectedPaymentMonth.getFullYear();
    const month = selectedPaymentMonth.getMonth(); // 0-based
    
    // 성사일자(won_time)가 해당 월에 속하는 거래 필터링
    const monthlyDeals = data.filter(deal => {
        if (!deal.won_time || deal.won_time === '') return false;
        
        const wonDate = parseToKST(deal.won_time);
        if (!wonDate) return false;
        
        return wonDate.getFullYear() === year && wonDate.getMonth() === month;
    });
    
    // 총 결제금액 및 건수 계산
    const totalAmount = monthlyDeals.reduce((sum, deal) => {
        const amount = parseInt(deal.value) || 0;
        return sum + amount;
    }, 0);
    
    const totalCount = monthlyDeals.length;
    
    // UI 업데이트
    document.getElementById('monthly-payment-amount').textContent = '₩' + formatNumber(totalAmount);
    document.getElementById('monthly-payment-count').textContent = totalCount + '건';
    
    console.log(`=== ${year}년 ${month + 1}월 결제금액 ===`);
    console.log(`결제 건수: ${totalCount}건`);
    console.log(`결제 금액: ₩${formatNumber(totalAmount)}`);
}

// 디버그: 특정 고객 검색 (콘솔에서 searchCustomer('이름') 으로 사용)
window.searchCustomer = function(keyword) {
    if (!paymentDataCache) {
        console.log('데이터가 로드되지 않았습니다.');
        return;
    }
    
    const results = paymentDataCache.filter(deal => {
        const name = deal.person_name || deal.title || '';
        return name.includes(keyword);
    });
    
    if (results.length === 0) {
        console.log(`'${keyword}' 검색 결과 없음`);
        return;
    }
    
    console.log(`=== '${keyword}' 검색 결과 (${results.length}건) ===`);
    results.forEach(deal => {
        const wonDateRaw = deal.won_time || '(없음)';
        const wonDateKST = deal.won_time ? parseToKST(deal.won_time) : null;
        const wonDateKSTStr = wonDateKST ? wonDateKST.toLocaleString('ko-KR') : '(없음)';
        
        console.log('---');
        console.log(`ID: ${deal.id}`);
        console.log(`고객명: ${deal.person_name || deal.title}`);
        console.log(`금액: ${deal.value}`);
        console.log(`성사일(원본): ${wonDateRaw}`);
        console.log(`성사일(KST): ${wonDateKSTStr}`);
        console.log(`최초결제안내일: ${deal.first_payment_notice || '(없음)'}`);
        
        if (wonDateKST) {
            console.log(`→ KST 기준 월: ${wonDateKST.getFullYear()}년 ${wonDateKST.getMonth() + 1}월`);
        }
    });
    
    return results;
};

// 월별 로우데이터 다운로드
function downloadMonthlyRawData() {
    if (!paymentDataCache) {
        showToast('데이터가 없습니다. 먼저 데이터를 로드해주세요.', 'error');
        return;
    }
    
    const year = selectedPaymentMonth.getFullYear();
    const month = selectedPaymentMonth.getMonth(); // 0-based
    
    // 성사일자(won_time)가 해당 월에 속하는 거래 필터링
    const monthlyDeals = paymentDataCache.filter(deal => {
        if (!deal.won_time || deal.won_time === '') return false;
        
        const wonDate = parseToKST(deal.won_time);
        if (!wonDate) return false;
        
        return wonDate.getFullYear() === year && wonDate.getMonth() === month;
    });
    
    if (monthlyDeals.length === 0) {
        showToast('해당 월에 데이터가 없습니다.', 'error');
        return;
    }
    
    // CSV 헤더
    const headers = [
        '거래ID',
        '고객명',
        '금액',
        '성사일(원본)',
        '성사일(KST)',
        '최초결제안내일'
    ];
    
    // CSV 데이터 생성
    const rows = monthlyDeals.map(deal => {
        const wonDateKST = parseToKST(deal.won_time);
        const noticeDate = deal.first_payment_notice ? parseToKST(deal.first_payment_notice) : null;
        
        return [
            deal.id || '',
            deal.person_name || deal.title || '',
            deal.value || 0,
            deal.won_time || '',  // 원본 값
            wonDateKST ? formatDateForCSV(wonDateKST) : '',  // KST 변환값
            noticeDate ? formatDateForCSV(noticeDate) : ''
        ];
    });
    
    // CSV 문자열 생성
    const BOM = '\uFEFF';
    const csvContent = BOM + [
        headers.join(','),
        ...rows.map(row => row.map(cell => {
            const cellStr = String(cell);
            if (cellStr.includes(',') || cellStr.includes('\n') || cellStr.includes('"')) {
                return '"' + cellStr.replace(/"/g, '""') + '"';
            }
            return cellStr;
        }).join(','))
    ].join('\n');
    
    // 파일 다운로드
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    const fileName = `월별결제_${year}년${month + 1}월_${formatDateForCSV(new Date())}.csv`;
    
    link.setAttribute('href', url);
    link.setAttribute('download', fileName);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    showToast(`${year}년 ${month + 1}월 로우데이터 (${monthlyDeals.length}건) 다운로드 완료!`);
}

// 주차별 KPI 계산 및 표시
function calculateAndDisplayWeeklyKPIs(data) {
    const weekStart = selectedWeekStart;
    const weekEnd = getWeekEnd(weekStart);
    
    console.log(`=== 주차별 KPI 계산 ===`);
    console.log(`기간: ${weekStart.toLocaleDateString()} ~ ${weekEnd.toLocaleDateString()}`);
    
    // 1. 환급 완료 (해당 기간 내 '최초결제안내일' 기준) - KST 적용
    const refundDeals = data.filter(deal => {
        if (!deal.first_payment_notice) return false;
        const noticeDate = parseToKST(deal.first_payment_notice);
        if (!noticeDate) return false;
        return noticeDate >= weekStart && noticeDate <= weekEnd;
    });
    
    const refundCount = refundDeals.length;
    const refundAmount = refundDeals.reduce((sum, deal) => sum + (deal.value || 0), 0);
    
    // 환급 완료 UI 업데이트
    document.getElementById('weekly-refund-count').textContent = formatNumber(refundCount) + '건';
    document.getElementById('weekly-refund-amount').textContent = '₩' + formatNumber(refundAmount);
    
    console.log(`환급 완료: ${refundCount}건, ₩${formatNumber(refundAmount)}`);
    
    // 2. 당일 결제 (최초결제안내일 = 성사일) - KST 적용
    const sameDayDeals = refundDeals.filter(deal => {
        if (!deal.won_time || deal.won_time === '') return false;
        // KST 기준 같은 날 비교
        return isSameDateKST(deal.first_payment_notice, deal.won_time);
    });
    
    const sameDayCount = sameDayDeals.length;
    const sameDayAmount = sameDayDeals.reduce((sum, deal) => sum + (deal.value || 0), 0);
    const sameDayCountRate = refundCount > 0 ? (sameDayCount / refundCount) * 100 : 0;
    
    // 당일 결제 UI 업데이트
    document.getElementById('weekly-sameday-count').textContent = formatNumber(sameDayCount) + '건';
    document.getElementById('weekly-sameday-amount').textContent = '₩' + formatNumber(sameDayAmount);
    updateRateDisplay('weekly-sameday-rate', sameDayCountRate);
    
    console.log(`당일 결제: ${sameDayCount}건, ₩${formatNumber(sameDayAmount)}, ${sameDayCountRate.toFixed(1)}%`);
    
    // 3. 30일이내 결제 (안내일부터 30일째까지 = daysDiff <= 29) - KST 적용
    const within30dDeals = refundDeals.filter(deal => {
        if (!deal.won_time || deal.won_time === '') return false;
        const daysDiff = getDaysDiffKST(deal.first_payment_notice, deal.won_time);
        if (daysDiff === null) return false;
        return daysDiff >= 0 && daysDiff <= 29;
    });
    
    const within30dCount = within30dDeals.length;
    const within30dAmount = within30dDeals.reduce((sum, deal) => sum + (deal.value || 0), 0);
    const within30dCountRate = refundCount > 0 ? (within30dCount / refundCount) * 100 : 0;
    
    // 30일이내 결제 UI 업데이트
    document.getElementById('weekly-30d-count').textContent = formatNumber(within30dCount) + '건';
    document.getElementById('weekly-30d-amount').textContent = '₩' + formatNumber(within30dAmount);
    updateRateDisplay('weekly-30d-rate', within30dCountRate);
    
    console.log(`30일이내 결제: ${within30dCount}건, ₩${formatNumber(within30dAmount)}, ${within30dCountRate.toFixed(1)}%`);
    
    // 4. 기간별 결제율 상세 테이블 업데이트
    const periods = [0, 3, 7, 21, 30, 60];
    const periodLabelsMap = { 0: '당일', 3: '3일 이내', 7: '7일 이내', 21: '21일 이내', 30: '30일 이내', 60: '60일 이내' };
    const paymentPeriodDeals = {};

    periods.forEach(days => {
        const result = calculateWeeklyConversionRate(refundDeals, days);
        updateWeeklyConversionRow(days, result, refundCount, refundAmount);
        paymentPeriodDeals[days] = result.convertedDeals;
    });

    // 미결제 거래 (환급 완료 중 아직 결제 안된 것)
    const unpaidDeals = refundDeals.filter(d => !d.won_time || d.won_time === '');
    paymentPeriodDeals['unpaid'] = unpaidDeals;

    // 행 클릭 이벤트 바인딩
    const periodRows = document.querySelectorAll('tr[data-period]');
    periodRows.forEach(row => {
        row.style.cursor = 'pointer';
        const newRow = row.cloneNode(true);
        row.parentNode.replaceChild(newRow, row);
        newRow.style.cursor = 'pointer';
        newRow.addEventListener('click', () => {
            const days = parseInt(newRow.getAttribute('data-period'));
            const deals = paymentPeriodDeals[days];
            if (!deals || deals.length === 0) return;

            periodRows.forEach(r => r.classList.remove('yt-cell-active'));
            document.querySelectorAll('tr[data-period]').forEach(r => r.classList.remove('yt-cell-active'));
            newRow.classList.add('yt-cell-active');

            const weekLabel = `${weekStart.toLocaleDateString('ko-KR')} ~ ${weekEnd.toLocaleDateString('ko-KR')}`;
            showPaymentDetail(deals, `${weekLabel} 환급 → ${periodLabelsMap[days]} 결제`);
        });
    });
}

// 주차별 전환율 계산 (KST 적용)
// N일이내 = 안내일(1일째)부터 N일째까지
// 예: 01-01 안내 → 01-03 성사 = 3일째 = 3일이내에 포함
function calculateWeeklyConversionRate(refundDeals, targetDays) {
    const convertedDeals = refundDeals.filter(deal => {
        if (!deal.won_time || deal.won_time === '') return false;
        const daysDiff = getDaysDiffKST(deal.first_payment_notice, deal.won_time);
        if (daysDiff === null) return false;
        
        if (targetDays === 0) {
            return daysDiff === 0;
        } else {
            return daysDiff >= 0 && daysDiff <= targetDays - 1;
        }
    });
    
    const convertedCount = convertedDeals.length;
    const convertedAmount = convertedDeals.reduce((sum, deal) => sum + (deal.value || 0), 0);
    
    return {
        convertedCount,
        convertedAmount,
        convertedDeals
    };
}

// 주차별 전환율 행 업데이트
function updateWeeklyConversionRow(days, result, totalCount, totalAmount) {
    const countRate = totalCount > 0 ? (result.convertedCount / totalCount) * 100 : 0;
    const amountRate = totalAmount > 0 ? (result.convertedAmount / totalAmount) * 100 : 0;
    
    // 건수 전환율
    const countRateEl = document.getElementById(`weekly-rate-cnt-${days}d`);
    if (countRateEl) {
        countRateEl.textContent = countRate.toFixed(1) + '%';
        countRateEl.className = 'rate-value';
        if (countRate >= 70) countRateEl.classList.add('rate-high');
        else if (countRate >= 40) countRateEl.classList.add('rate-medium');
        else if (countRate > 0) countRateEl.classList.add('rate-low');
    }
    
    // 금액 전환율
    const amountRateEl = document.getElementById(`weekly-rate-amt-${days}d`);
    if (amountRateEl) {
        amountRateEl.textContent = amountRate.toFixed(1) + '%';
        amountRateEl.className = 'rate-value rate-amount';
        if (amountRate >= 70) amountRateEl.classList.add('rate-high');
        else if (amountRate >= 40) amountRateEl.classList.add('rate-medium');
        else if (amountRate > 0) amountRateEl.classList.add('rate-low');
    }
    
    // 건수 상세
    const countDetailEl = document.getElementById(`weekly-detail-cnt-${days}d`);
    if (countDetailEl) {
        countDetailEl.textContent = `${result.convertedCount} / ${totalCount}`;
    }
    
    // 금액 상세 (억 단위)
    const amountDetailEl = document.getElementById(`weekly-detail-amt-${days}d`);
    if (amountDetailEl) {
        const convertedEok = (result.convertedAmount / 100000000).toFixed(2);
        const totalEok = (totalAmount / 100000000).toFixed(2);
        amountDetailEl.textContent = `${convertedEok}억 / ${totalEok}억`;
    }
    
    // 목표 달성 여부 (3일, 30일, 60일만)
    const statusEl = document.getElementById(`weekly-status-${days}d`);
    if (statusEl && PAYMENT_TARGETS[days]) {
        const target = PAYMENT_TARGETS[days];
        if (countRate >= target) {
            statusEl.innerHTML = '<span class="status-achieved">✅</span>';
        } else if (countRate >= target * 0.8) {
            statusEl.innerHTML = '<span class="status-warning">⚠️</span>';
        } else if (totalCount > 0) {
            statusEl.innerHTML = '<span class="status-failed">❌</span>';
        } else {
            statusEl.textContent = '-';
        }
    }
    
    // 행 색상 업데이트
    const rowEl = document.querySelector(`tr[data-period="${days}"]`);
    if (rowEl) {
        rowEl.removeAttribute('data-rate');
        if (countRate >= 70) {
            rowEl.setAttribute('data-rate', 'high');
        } else if (countRate >= 40) {
            rowEl.setAttribute('data-rate', 'medium');
        } else if (countRate > 0) {
            rowEl.setAttribute('data-rate', 'low');
        }
    }
}

function showPaymentDetail(deals, label) {
    const panel = document.getElementById('payment-detail-panel');
    const titleEl = document.getElementById('payment-detail-title');
    const tbodyEl = document.getElementById('payment-detail-tbody');
    if (!panel || !titleEl || !tbodyEl) return;

    const totalAmt = deals.reduce((s, d) => s + (parseInt(d.value) || 0), 0);
    titleEl.textContent = `${label} — ${deals.length}건, ₩${totalAmt.toLocaleString()}`;

    const sorted = [...deals].sort((a, b) => (parseInt(b.value) || 0) - (parseInt(a.value) || 0));

    tbodyEl.innerHTML = sorted.map((d, i) => {
        const val = parseInt(d.value) || 0;
        const noticeDate = d.first_payment_notice || '-';
        const wonDate = d.won_time || '미결제';
        const daysDiff = (d.won_time && d.first_payment_notice)
            ? getDaysDiffKST(d.first_payment_notice, d.won_time)
            : null;
        const daysLabel = daysDiff !== null ? `${daysDiff}일` : '-';
        return `<tr>
            <td>${i + 1}</td>
            <td>${d.title || '-'}</td>
            <td class="td-amount">₩${val.toLocaleString()}</td>
            <td>${noticeDate}</td>
            <td>${wonDate}</td>
            <td>${daysLabel}</td>
        </tr>`;
    }).join('');

    panel.style.display = 'block';
    panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function closePaymentDetail() {
    const panel = document.getElementById('payment-detail-panel');
    if (panel) panel.style.display = 'none';
    document.querySelectorAll('tr[data-period]').forEach(r => r.classList.remove('yt-cell-active'));
}

// 결제율 표시 업데이트 (색상 포함)
function updateRateDisplay(elementId, rate) {
    const el = document.getElementById(elementId);
    if (el) {
        el.textContent = rate.toFixed(1) + '%';
        el.classList.remove('warning', 'danger');
        if (rate >= 70) {
            // 기본 초록색 유지
        } else if (rate >= 40) {
            el.classList.add('warning');
        } else {
            el.classList.add('danger');
        }
    }
}

// 주차별 로우데이터 엑셀(CSV) 다운로드
function downloadWeeklyRawData() {
    if (!paymentDataCache) {
        showToast('데이터가 없습니다. 먼저 데이터를 로드해주세요.', 'error');
        return;
    }
    
    const weekStart = selectedWeekStart;
    const weekEnd = getWeekEnd(weekStart);
    
    // 해당 주차의 결제 안내 데이터 필터링 (KST 적용)
    const weeklyDeals = paymentDataCache.filter(deal => {
        if (!deal.first_payment_notice) return false;
        const noticeDate = parseToKST(deal.first_payment_notice);
        if (!noticeDate) return false;
        return noticeDate >= weekStart && noticeDate <= weekEnd;
    });
    
    if (weeklyDeals.length === 0) {
        showToast('해당 주차에 데이터가 없습니다.', 'error');
        return;
    }
    
    // CSV 헤더
    const headers = [
        '거래ID',
        '고객명',
        '금액',
        '최초결제안내일',
        '성사일',
        '일수차이',
        '결제여부',
        '당일결제',
        '3일이내',
        '7일이내',
        '21일이내',
        '30일이내',
        '60일이내'
    ];
    
    // CSV 데이터 생성 (KST 적용)
    const rows = weeklyDeals.map(deal => {
        const noticeDate = parseToKST(deal.first_payment_notice);
        const wonDate = deal.won_time ? parseToKST(deal.won_time) : null;
        const daysDiff = getDaysDiffKST(deal.first_payment_notice, deal.won_time);
        const isPaid = wonDate !== null && deal.won_time;
        
        // 기간별 포함 여부 계산
        const isSameDay = isPaid && daysDiff === 0;
        const isWithin3d = isPaid && daysDiff >= 0 && daysDiff <= 2;
        const isWithin7d = isPaid && daysDiff >= 0 && daysDiff <= 6;
        const isWithin21d = isPaid && daysDiff >= 0 && daysDiff <= 20;
        const isWithin30d = isPaid && daysDiff >= 0 && daysDiff <= 29;
        const isWithin60d = isPaid && daysDiff >= 0 && daysDiff <= 59;
        
        return [
            deal.id || '',
            deal.person_name || deal.title || '',
            deal.value || 0,
            formatDateForCSV(noticeDate),
            wonDate ? formatDateForCSV(wonDate) : '',
            daysDiff !== null ? daysDiff : '',
            isPaid ? 'Y' : 'N',
            isSameDay ? 'Y' : 'N',
            isWithin3d ? 'Y' : 'N',
            isWithin7d ? 'Y' : 'N',
            isWithin21d ? 'Y' : 'N',
            isWithin30d ? 'Y' : 'N',
            isWithin60d ? 'Y' : 'N'
        ];
    });
    
    // CSV 문자열 생성 (BOM 추가로 엑셀 한글 깨짐 방지)
    const BOM = '\uFEFF';
    const csvContent = BOM + [
        headers.join(','),
        ...rows.map(row => row.map(cell => {
            // 콤마나 줄바꿈이 포함된 셀은 따옴표로 감싸기
            const cellStr = String(cell);
            if (cellStr.includes(',') || cellStr.includes('\n') || cellStr.includes('"')) {
                return '"' + cellStr.replace(/"/g, '""') + '"';
            }
            return cellStr;
        }).join(','))
    ].join('\n');
    
    // 파일 다운로드
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const formatDate = (d) => {
        const yy = String(d.getFullYear()).slice(2);
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return `${yy}${mm}${dd}`;
    };
    const filename = `결제_로우데이터_${formatDate(weekStart)}-${formatDate(weekEnd)}.csv`;
    
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
    
    showToast(`${weeklyDeals.length}건 데이터가 다운로드되었습니다.`);
}

// CSV용 날짜 포맷
function formatDateForCSV(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

// ==========================================
// 추심 대시보드
// ==========================================

const COLLECTION_API_URL = PERFORMANCE_API_URL;

let collectionDataCache = null;
let selectedCollectionMonth = new Date();
let selectedCollectionYearVal = new Date().getFullYear();

function initCollectionTab() {
    const collectionTabBtn = document.querySelector('[data-tab="collection"]');
    if (collectionTabBtn) {
        collectionTabBtn.addEventListener('click', () => {
            if (!collectionDataCache) {
                loadCollectionData();
            }
        });
    }
    updateCollectionMonthDisplay();
    updateCollectionYearDisplay();
}

function updateCollectionMonthDisplay() {
    const year = selectedCollectionMonth.getFullYear();
    const month = selectedCollectionMonth.getMonth() + 1;
    const monthEl = document.getElementById('selected-collection-month');
    if (monthEl) {
        monthEl.textContent = `${year}년 ${month}월`;
    }
}

function changeCollectionMonth(delta) {
    selectedCollectionMonth.setMonth(selectedCollectionMonth.getMonth() + delta);
    updateCollectionMonthDisplay();
    if (collectionDataCache) {
        calculateAndDisplayCollectionKPIs(collectionDataCache);
    }
}

function changeCollectionYear(delta) {
    selectedCollectionYearVal += delta;
    updateCollectionYearDisplay();
    if (collectionDataCache) {
        calculateYearlyCollectionTracking(collectionDataCache);
    }
}

function updateCollectionYearDisplay() {
    const el = document.getElementById('selected-collection-year');
    if (el) el.textContent = `${selectedCollectionYearVal}년`;
}

async function loadCollectionData() {
    const loadingEl = document.getElementById('collection-loading');
    const errorEl = document.getElementById('collection-error');
    const contentEl = document.getElementById('collection-content');

    if (loadingEl) loadingEl.style.display = 'flex';
    if (errorEl) errorEl.style.display = 'none';
    if (contentEl) contentEl.style.display = 'none';

    try {
        collectionDataCache = await fetchSharedData();

        calculateAndDisplayCollectionKPIs(collectionDataCache);
        calculateYearlyCollectionTracking(collectionDataCache);

        const syncTimeEl = document.getElementById('collection-last-sync-time');
        if (syncTimeEl) syncTimeEl.textContent = new Date().toLocaleString('ko-KR');

        if (loadingEl) loadingEl.style.display = 'none';
        if (contentEl) contentEl.style.display = 'block';
    } catch (error) {
        console.error('추심 데이터 로드 실패:', error);
        if (loadingEl) loadingEl.style.display = 'none';
        if (errorEl) {
            errorEl.style.display = 'flex';
            const errorMsgEl = document.getElementById('collection-error-message');
            if (errorMsgEl) errorMsgEl.textContent = error.message || '네트워크 오류가 발생했습니다.';
        }
    }
}

function calculateAndDisplayCollectionKPIs(data) {
    const year = selectedCollectionMonth.getFullYear();
    const month = selectedCollectionMonth.getMonth();

    let prevYear = year;
    let prevMonth = month - 1;
    if (prevMonth < 0) { prevMonth = 11; prevYear--; }

    const refundBadgeEl = document.getElementById('col-refund-badge');
    const transferBadgeEl = document.getElementById('col-transfer-badge');
    if (refundBadgeEl) refundBadgeEl.textContent = `${prevMonth + 1}월 기준`;
    if (transferBadgeEl) transferBadgeEl.textContent = `${month + 1}월 기준`;

    console.log(`=== 추심 KPI (${year}년 ${month + 1}월) ===`);
    console.log(`환급완료 기준: ${prevYear}년 ${prevMonth + 1}월 first_payment_notice`);
    console.log(`이관총액 기준: ${year}년 ${month + 1}월 collection_order_date`);

    const refundDeals = data.filter(deal => {
        if (!deal.first_payment_notice) return false;
        const d = parseToKST(deal.first_payment_notice);
        if (!d) return false;
        return d.getFullYear() === prevYear && d.getMonth() === prevMonth;
    });

    const refundCount = refundDeals.length;
    const refundAmount = refundDeals.reduce((sum, deal) => sum + (parseInt(deal.value) || 0), 0);

    const transferDeals = data.filter(deal => {
        if (!deal.collection_order_date || deal.collection_order_date === '') return false;
        const d = parseToKST(deal.collection_order_date);
        if (!d) return false;
        return d.getFullYear() === year && d.getMonth() === month;
    });

    const transferCount = transferDeals.length;
    const transferAmount = transferDeals.reduce((sum, deal) => sum + (parseInt(deal.value) || 0), 0);

    const ratio = refundAmount > 0 ? (transferAmount / refundAmount) * 100 : 0;

    console.log(`환급완료: ${refundCount}건, ₩${formatNumber(refundAmount)}`);
    console.log(`이관총액: ${transferCount}건, ₩${formatNumber(transferAmount)}`);
    console.log(`비율: ${ratio.toFixed(1)}%`);

    document.getElementById('col-refund-amount').textContent = '₩' + formatNumber(refundAmount);
    document.getElementById('col-refund-count').textContent = refundCount + '건';

    document.getElementById('col-transfer-amount').textContent = '₩' + formatNumber(transferAmount);
    document.getElementById('col-transfer-count').textContent = transferCount + '건';

    const ratioEl = document.getElementById('col-ratio-value');
    if (ratioEl) {
        ratioEl.textContent = ratio.toFixed(1) + '%';
        ratioEl.style.color = '';
        if (ratio >= 30) ratioEl.style.color = 'var(--accent-danger)';
        else if (ratio >= 20) ratioEl.style.color = 'var(--accent-warning)';
        else if (ratio > 0) ratioEl.style.color = 'var(--accent-success)';
    }

    document.getElementById('col-ratio-refund').textContent = '₩' + formatNumber(refundAmount);
    document.getElementById('col-ratio-transfer').textContent = '₩' + formatNumber(transferAmount);
}

function calculateYearlyCollectionTracking(data) {
    const year = selectedCollectionYearVal;
    const tbody = document.getElementById('yearly-collection-body');
    if (!tbody) return;

    const MONTH_BUCKETS = 12;
    const ytBucketStore = {};
    let rows = '';

    for (let m = 0; m < 12; m++) {
        let prevYear = year;
        let prevMonth = m - 1;
        if (prevMonth < 0) { prevMonth = 11; prevYear--; }
        const refundAmount = data.filter(deal => {
            if (!deal.first_payment_notice) return false;
            const d = parseToKST(deal.first_payment_notice);
            return d && d.getFullYear() === prevYear && d.getMonth() === prevMonth;
        }).reduce((s, d) => s + (parseInt(d.value) || 0), 0);

        const transferDeals = data.filter(deal => {
            if (!deal.collection_order_date || deal.collection_order_date === '') return false;
            const d = parseToKST(deal.collection_order_date);
            if (!d) return false;
            return d.getFullYear() === year && d.getMonth() === m;
        });

        const totalCount = transferDeals.length;
        const totalAmount = transferDeals.reduce((s, d) => s + (parseInt(d.value) || 0), 0);

        if (totalCount === 0) {
            rows += `<tr class="yt-row">
                <td class="yt-cell-month">${m + 1}월</td>
                <td class="yt-cell-num yt-sticky-transfer" style="left:48px;"><div class="yt-rate-main" style="font-size:12px;">-</div><div class="yt-rate-detail">0건</div></td>
                <td class="yt-cell-num yt-sticky-paid" style="left:148px;">-</td>
                <td class="yt-cell-rate" colspan="${MONTH_BUCKETS + 2}" style="text-align:center;color:var(--text-secondary);">이관 데이터 없음</td>
            </tr>`;
            continue;
        }

        const buckets = {};
        for (let i = 0; i < MONTH_BUCKETS; i++) buckets[i] = [];
        buckets.over = [];
        buckets.unpaid = [];

        transferDeals.forEach(deal => {
            if (!deal.won_time || deal.won_time === '') {
                buckets.unpaid.push(deal);
                return;
            }
            const orderDate = parseToKST(deal.collection_order_date);
            const wonDate = parseToKST(deal.won_time);
            if (!orderDate || !wonDate) { buckets.unpaid.push(deal); return; }
            const md = (wonDate.getFullYear() - orderDate.getFullYear()) * 12 + (wonDate.getMonth() - orderDate.getMonth());
            if (md >= MONTH_BUCKETS) {
                buckets.over.push(deal);
            } else {
                const key = Math.max(0, md);
                buckets[key].push(deal);
            }
        });

        ytBucketStore[m] = { transfer: transferDeals, buckets };

        const totalEok = (totalAmount / 100000000).toFixed(1);

        function cellHtml(bucket, isUnpaid, clickKey) {
            const cnt = bucket.length;
            const amt = bucket.reduce((s, d) => s + (parseInt(d.value) || 0), 0);
            const amtEok = amt / 100000000;
            const amtFull = '₩' + amt.toLocaleString();
            const clickAttr = cnt > 0 ? ` data-click="${clickKey}"` : '';

            if (cnt === 0) {
                return `<td class="yt-cell-rate"><div class="yt-rate-main" style="color:var(--text-secondary);font-size:11px;opacity:0.4;">-</div></td>`;
            }

            let amtDisplay;
            if (amtEok >= 1) {
                amtDisplay = amtEok.toFixed(1) + '억';
            } else {
                const amtMan = Math.round(amt / 10000);
                amtDisplay = amtMan.toLocaleString() + '만';
            }

            let colorClass = '';
            if (isUnpaid) {
                if (amtEok > 0) colorClass = 'yt-low';
            } else {
                if (amtEok >= 1) colorClass = 'yt-high';
                else if (amtEok >= 0.1) colorClass = 'yt-medium';
                else if (amt > 0) colorClass = 'yt-low';
            }

            return `<td class="yt-cell-rate ${colorClass}" data-tip="${amtFull}"${clickAttr}>
                <div class="yt-rate-main">${amtDisplay}</div>
                <div class="yt-rate-detail">${cnt}건</div>
            </td>`;
        }

        let monthCells = '';
        for (let i = 0; i < MONTH_BUCKETS; i++) {
            monthCells += cellHtml(buckets[i], false, `${m}-${i}`);
        }

        const paidDeals = transferDeals.filter(d => d.won_time && d.won_time !== '');
        const paidAmount = paidDeals.reduce((s, d) => s + (parseInt(d.value) || 0), 0);
        const paidEok = paidAmount / 100000000;
        const paidDisplay = paidEok >= 1 ? paidEok.toFixed(1) + '억' : Math.round(paidAmount / 10000).toLocaleString() + '만';
        const paidRate = totalAmount > 0 ? (paidAmount / totalAmount * 100).toFixed(0) : 0;

        const paidPctNum = parseInt(paidRate);
        let paidColorCls = '';
        if (paidPctNum >= 50) paidColorCls = 'yt-paid-high';
        else if (paidPctNum >= 20) paidColorCls = 'yt-paid-mid';
        else if (paidPctNum > 0) paidColorCls = 'yt-paid-low';

        const transferClickAttr = totalCount > 0 ? ` data-click="${m}-transfer"` : '';
        const paidClickAttr = paidDeals.length > 0 ? ` data-click="${m}-paid"` : '';

        rows += `<tr class="yt-row">
            <td class="yt-cell-month">${m + 1}월</td>
            <td class="yt-cell-num yt-sticky-transfer" style="left:48px;" data-tip="₩${totalAmount.toLocaleString()}"${transferClickAttr}><div class="yt-rate-main" style="font-size:12px;white-space:nowrap;">${totalEok}억 (${refundAmount > 0 ? (totalAmount / refundAmount * 100).toFixed(0) : 0}%)</div><div class="yt-rate-detail">${totalCount}건</div></td>
            <td class="yt-cell-num yt-sticky-paid ${paidColorCls}" style="left:148px;font-weight:700;" data-tip="₩${paidAmount.toLocaleString()}"${paidClickAttr}><div class="yt-rate-main" style="font-size:12px;white-space:nowrap;">${paidAmount > 0 ? paidDisplay : '-'}</div><div class="yt-rate-detail">${paidDeals.length}건 (${paidRate}%)</div></td>
            ${monthCells}
            ${cellHtml(buckets.over, false, `${m}-over`)}
            ${cellHtml(buckets.unpaid, true, `${m}-unpaid`)}
        </tr>`;
    }

    tbody.innerHTML = rows;

    // 셀 클릭 이벤트 바인딩
    tbody.querySelectorAll('[data-click]').forEach(cell => {
        cell.addEventListener('click', () => {
            const key = cell.getAttribute('data-click');
            const [monthStr, bucketKey] = key.split('-');
            const monthIdx = parseInt(monthStr);
            const store = ytBucketStore[monthIdx];
            if (!store) return;

            let deals, label;
            const monthLabel = `${monthIdx + 1}월`;

            if (bucketKey === 'transfer') {
                deals = store.transfer;
                label = `${year}년 ${monthLabel} 이관 전체`;
            } else if (bucketKey === 'paid') {
                deals = store.transfer.filter(d => d.won_time && d.won_time !== '');
                label = `${year}년 ${monthLabel} 성사 (결제 완료)`;
            } else if (bucketKey === 'over') {
                deals = store.buckets.over;
                label = `${year}년 ${monthLabel} 이관 → +12개월 이후 결제`;
            } else if (bucketKey === 'unpaid') {
                deals = store.buckets.unpaid;
                label = `${year}년 ${monthLabel} 이관 → 미결제`;
            } else {
                const idx = parseInt(bucketKey);
                deals = store.buckets[idx];
                label = `${year}년 ${monthLabel} 이관 → +${idx}개월 후 결제`;
            }

            if (!deals || deals.length === 0) return;

            // 선택 하이라이트
            tbody.querySelectorAll('.yt-cell-active').forEach(el => el.classList.remove('yt-cell-active'));
            cell.classList.add('yt-cell-active');

            showYtDetail(deals, label);
        });
    });
}

function showYtDetail(deals, label) {
    const panel = document.getElementById('yt-detail-panel');
    const titleEl = document.getElementById('yt-detail-title');
    const tbodyEl = document.getElementById('yt-detail-tbody');
    if (!panel || !titleEl || !tbodyEl) return;

    const totalAmt = deals.reduce((s, d) => s + (parseInt(d.value) || 0), 0);
    titleEl.textContent = `${label} — ${deals.length}건, ₩${totalAmt.toLocaleString()}`;

    const sorted = [...deals].sort((a, b) => (parseInt(b.value) || 0) - (parseInt(a.value) || 0));

    tbodyEl.innerHTML = sorted.map((d, i) => {
        const val = parseInt(d.value) || 0;
        const orderDate = d.collection_order_date || '-';
        const wonDate = d.won_time || '미결제';
        return `<tr>
            <td>${i + 1}</td>
            <td>${d.title || '-'}</td>
            <td class="td-amount">₩${val.toLocaleString()}</td>
            <td>${orderDate}</td>
            <td>${wonDate}</td>
        </tr>`;
    }).join('');

    panel.style.display = 'block';
    panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function closeYtDetail() {
    const panel = document.getElementById('yt-detail-panel');
    if (panel) panel.style.display = 'none';
    const tbody = document.getElementById('yearly-collection-body');
    if (tbody) tbody.querySelectorAll('.yt-cell-active').forEach(el => el.classList.remove('yt-cell-active'));
}

// data-tip 툴팁 (body에 붙여서 overflow 무시)
(function initTooltip() {
    const tip = document.createElement('div');
    tip.className = 'yt-tooltip';
    tip.style.opacity = '0';
    document.body.appendChild(tip);

    document.addEventListener('mouseover', (e) => {
        const el = e.target.closest('[data-tip]');
        if (!el) { tip.style.opacity = '0'; return; }
        tip.textContent = el.getAttribute('data-tip');
        tip.style.opacity = '1';
        const rect = el.getBoundingClientRect();
        tip.style.left = (rect.left + rect.width / 2 - tip.offsetWidth / 2) + 'px';
        tip.style.top = (rect.bottom + 6) + 'px';
    });

    document.addEventListener('mouseout', (e) => {
        const el = e.target.closest('[data-tip]');
        if (el) tip.style.opacity = '0';
    });
})();

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    init();
    initPaymentTab();
    initCollectionTab();
});
