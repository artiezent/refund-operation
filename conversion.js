// 전환파트 Dashboard - conversion.js

// ==========================================
// 상수 & 상태
// ==========================================

const STORAGE_KEYS = {
    coverage: 'dailyKpi_coverage',
    activity: 'dailyKpi_activity',
    application: 'dailyKpi_application',
    defense: 'dailyKpi_defense'
};

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

const ACTIVITY_PARSE_RULES = {
    applyActivity: { keyword: '줍줍 활동', offsetLines: 5 },
    applyAbsent: { keyword: '줍줍 부재', offsetLines: 5 },
    applyFollowup: { keyword: '줍줍콜 사후관리', offsetLines: 5 },
    applySms: { keyword: '신청 문자', offsetLines: 5 },
    defenseActivity: { keyword: '취소 활동', offsetLines: 5 },
    defenseAbsent: { keyword: '취소 부재', offsetLines: 5 },
    defenseFollowup: { keyword: '취소방어 사후관리', offsetLines: 5 },
    defenseSms: { keyword: '취소 문자', offsetLines: 5 }
};

let selectedPerfMonth = new Date();

// ==========================================
// 날짜 표시
// ==========================================

function displayCurrentDate() {
    const now = new Date();
    const options = { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric', 
        weekday: 'long' 
    };
    const el = document.getElementById('currentDate');
    if (el) el.textContent = now.toLocaleDateString('ko-KR', options);
}

// ==========================================
// 커버리지 현황
// ==========================================

async function applyCoverageData() {
    const input = document.getElementById('coverageDataInput');
    const inputText = input.value.trim();
    
    if (!inputText) {
        showToast('데이터를 입력해주세요.', 'error');
        return;
    }
    
    const data = parsePipedriveData(inputText);
    
    const countNumerator = data.successCount + data.contactCount;
    const countDenominator = data.successCount + data.unconvertedCount;
    const countRate = countDenominator > 0 ? (countNumerator / countDenominator) * 100 : 0;
    
    const amountNumerator = data.successAmount + data.contactAmount;
    const amountDenominator = data.successAmount + data.unconvertedAmount;
    const amountRate = amountDenominator > 0 ? (amountNumerator / amountDenominator) * 100 : 0;
    
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
    
    const saveData = {
        ...data,
        countRate,
        amountRate
    };
    
    showToast('저장 중...');
    const saved = await saveManualDataToSheets('coverage', saveData);
    
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
// 활동수 현황
// ==========================================

function applyActivityData() {
    const input = document.getElementById('activityDataInput');
    const inputText = input.value.trim();
    
    if (!inputText) {
        showToast('데이터를 입력해주세요.', 'error');
        return;
    }
    
    const data = parseActivityData(inputText);
    
    const applyTotal = data.applyActivity + data.applyAbsent;
    const applyExtra = data.applyFollowup + data.applySms;
    
    const defenseTotal = data.defenseActivity + data.defenseAbsent;
    const defenseExtra = data.defenseFollowup + data.defenseSms;
    
    const maxValue = Math.max(applyTotal, applyExtra, defenseTotal, defenseExtra, 1);
    
    updateTotalActivityBar('apply', data.applyActivity, data.applyAbsent, maxValue);
    updateExtraActivityBar('apply', data.applyFollowup, data.applySms, maxValue);
    document.getElementById('apply-total-value').textContent = formatNumber(applyTotal);
    document.getElementById('apply-extra-value').textContent = formatNumber(applyExtra);
    updateBarGroupOrder('apply', applyTotal, applyExtra);
    
    document.getElementById('tt-apply-activity').textContent = formatNumber(data.applyActivity);
    document.getElementById('tt-apply-absent').textContent = formatNumber(data.applyAbsent);
    document.getElementById('tt-apply-total').textContent = formatNumber(applyTotal);
    document.getElementById('tt-apply-followup').textContent = formatNumber(data.applyFollowup);
    document.getElementById('tt-apply-sms').textContent = formatNumber(data.applySms);
    document.getElementById('tt-apply-extra').textContent = formatNumber(applyExtra);
    
    updateTotalActivityBar('defense', data.defenseActivity, data.defenseAbsent, maxValue);
    updateExtraActivityBar('defense', data.defenseFollowup, data.defenseSms, maxValue);
    document.getElementById('defense-total-value').textContent = formatNumber(defenseTotal);
    document.getElementById('defense-extra-value').textContent = formatNumber(defenseExtra);
    updateBarGroupOrder('defense', defenseTotal, defenseExtra);
    
    document.getElementById('tt-defense-activity').textContent = formatNumber(data.defenseActivity);
    document.getElementById('tt-defense-absent').textContent = formatNumber(data.defenseAbsent);
    document.getElementById('tt-defense-total').textContent = formatNumber(defenseTotal);
    document.getElementById('tt-defense-followup').textContent = formatNumber(data.defenseFollowup);
    document.getElementById('tt-defense-sms').textContent = formatNumber(data.defenseSms);
    document.getElementById('tt-defense-extra').textContent = formatNumber(defenseExtra);
    
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
    
    closeInputPanel('activity-input');
    
    if (typeof saveSnapshot === 'function') saveSnapshot();
    
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
    
    for (let i = 1; i <= rule.offsetLines + 2; i++) {
        const targetIndex = keywordIndex + i;
        if (targetIndex >= lines.length) break;
        
        const line = lines[targetIndex];
        
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
        if (activityValue >= absentValue) {
            activitySeg.style.order = 1;
            absentSeg.style.order = 2;
        } else {
            activitySeg.style.order = 2;
            absentSeg.style.order = 1;
        }
    }
}

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
        if (followupValue >= smsValue) {
            followupSeg.style.order = 1;
            smsSeg.style.order = 2;
        } else {
            followupSeg.style.order = 2;
            smsSeg.style.order = 1;
        }
    }
}

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
    const input = document.getElementById('activityDataInput');
    if (input) input.value = '';
    
    ['apply-total', 'apply-extra', 'defense-total', 'defense-extra'].forEach(id => {
        const bar = document.getElementById(`${id}-bar`);
        const value = document.getElementById(`${id}-value`);
        if (bar) bar.style.height = '5%';
        if (value) value.textContent = '0';
    });
    
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
// 총조회대비 신청전환
// ==========================================

async function applyApplicationData() {
    const input = document.getElementById('applicationDataInput');
    const inputText = input.value.trim();
    
    if (!inputText) {
        showToast('데이터를 입력해주세요.', 'error');
        return;
    }
    
    const lines = inputText.split('\n').map(line => line.trim());
    
    const totalView = findAmountByKeyword(lines, 'KPI_총조회');
    const totalApply = findAmountByKeyword(lines, 'KPI_신청');
    const applyConvert = findAmountByKeyword(lines, 'KPI_신청전환');
    
    console.log('신청전환 파싱 결과:', { totalView, totalApply, applyConvert });
    
    const totalApplyRate = totalView > 0 ? (totalApply / totalView) * 100 : 0;
    const applySuccessRate = totalView > 0 ? (applyConvert / totalView) * 100 : 0;
    
    updateKPICard('apply-total-rate', totalApplyRate, {
        'total-view-amount': '₩' + formatNumber(totalView),
        'total-apply-amount': '₩' + formatNumber(totalApply)
    });
    
    updateKPICard('apply-success-rate', applySuccessRate, {
        'total-view-amount2': '₩' + formatNumber(totalView),
        'apply-convert-amount': '₩' + formatNumber(applyConvert)
    });
    
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
// 총검토대비 취소방어
// ==========================================

async function applyDefenseData() {
    const input = document.getElementById('defenseDataInput');
    const inputText = input.value.trim();
    
    if (!inputText) {
        showToast('데이터를 입력해주세요.', 'error');
        return;
    }
    
    const lines = inputText.split('\n').map(line => line.trim());
    
    const cancelRequest = findAmountByKeyword(lines, 'KPI_취소전체');
    const cancelAvailable = findAmountByKeyword(lines, 'KPI_취소검토');
    const cancelSuccess = findAmountByKeyword(lines, 'KPI_취소성공');
    
    console.log('취소방어 파싱 결과:', { cancelRequest, cancelAvailable, cancelSuccess });
    
    const reviewRate = cancelRequest > 0 ? (cancelAvailable / cancelRequest) * 100 : 0;
    const defenseRate = cancelAvailable > 0 ? (cancelSuccess / cancelAvailable) * 100 : 0;
    
    updateKPICard('cancel-review-rate', reviewRate, {
        'cancel-request-amount': '₩' + formatNumber(cancelRequest),
        'cancel-available-amount': '₩' + formatNumber(cancelAvailable)
    });
    
    updateKPICard('cancel-defense-rate', defenseRate, {
        'cancel-available-amount2': '₩' + formatNumber(cancelAvailable),
        'cancel-success-amount': '₩' + formatNumber(cancelSuccess)
    });
    
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
// 성과 요약 (Performance Summary)
// ==========================================

function initPerformanceSummary() {
    updatePerfMonthDisplay();
    loadTargetAmount();
    loadPerformanceSummary();
}

function updatePerfMonthDisplay() {
    const year = selectedPerfMonth.getFullYear();
    const month = selectedPerfMonth.getMonth() + 1;
    const monthEl = document.getElementById('selected-perf-month');
    if (monthEl) monthEl.textContent = `${year}년 ${month}월`;
}

function changePerfMonth(delta) {
    selectedPerfMonth.setMonth(selectedPerfMonth.getMonth() + delta);
    updatePerfMonthDisplay();
    loadPerformanceSummary();
    loadActivityStatus();
}

function savePerfToLocal(year, month, data) {
    try {
        const key = `perf_${year}_${month}`;
        localStorage.setItem(key, JSON.stringify({ data, ts: Date.now() }));
    } catch (e) {}
}

function loadPerfFromLocal(year, month) {
    try {
        const key = `perf_${year}_${month}`;
        const raw = localStorage.getItem(key);
        if (!raw) return null;
        return JSON.parse(raw);
    } catch (e) { return null; }
}

async function loadPerformanceSummary() {
    const loadingEl = document.getElementById('perf-loading');
    const errorEl = document.getElementById('perf-error');
    const contentEl = document.getElementById('perf-content');
    
    const year = selectedPerfMonth.getFullYear();
    const month = selectedPerfMonth.getMonth() + 1;
    
    if (loadingEl) loadingEl.style.display = 'flex';
    if (errorEl) errorEl.style.display = 'none';
    if (contentEl) contentEl.style.display = 'none';
    
    try {
        const url = `${PERFORMANCE_API_URL}?action=performance&year=${year}&month=${month}`;
        const response = await fetch(url);
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        const json = await response.json();
        if (!json.success) throw new Error('성과 데이터를 가져오는데 실패했습니다.');
        
        updatePerformanceSummaryUI(json.data);
        
        const syncTimeEl = document.getElementById('perf-sync-time');
        if (syncTimeEl) syncTimeEl.textContent = new Date().toLocaleString('ko-KR');
        
        if (loadingEl) loadingEl.style.display = 'none';
        if (contentEl) contentEl.style.display = 'block';
        
    } catch (error) {
        console.error('성과 데이터 로드 실패:', error);
        if (loadingEl) loadingEl.style.display = 'none';
        if (errorEl) {
            errorEl.style.display = 'flex';
            const errorMsgEl = document.getElementById('perf-error-message');
            if (errorMsgEl) errorMsgEl.textContent = error.message || '데이터 로딩에 실패했습니다.';
        }
    }
}

function calculatePerformanceFromData(data, year, month) {
    let applyCount = 0, applyAmount = 0;
    let defenseCount = 0, defenseAmount = 0;
    
    data.forEach(d => {
        const val = d._refundAmount || 0;
        
        if (d._applyDate && d._hasJupjupPerson && val > 0) {
            if (d._applyDate.getFullYear() === year && d._applyDate.getMonth() + 1 === month) {
                applyCount++;
                applyAmount += val;
            }
        }
        
        if (d._defenseDate && d._hasDefensePerson && val > 0) {
            if (d._defenseDate.getFullYear() === year && d._defenseDate.getMonth() + 1 === month) {
                defenseCount++;
                defenseAmount += val;
            }
        }
    });
    
    return {
        apply: { count: applyCount, amount: applyAmount },
        defense: { count: defenseCount, amount: defenseAmount },
        total: applyAmount + defenseAmount
    };
}

function updatePerformanceSummaryUI(data) {
    const applyAmount = data.apply?.amount || 0;
    const applyCount = data.apply?.count || 0;
    document.getElementById('perf-apply-amount').textContent = '₩' + formatNumber(applyAmount);
    document.getElementById('perf-apply-count').textContent = applyCount + '건';
    
    const defenseAmount = data.defense?.amount || 0;
    const defenseCount = data.defense?.count || 0;
    document.getElementById('perf-defense-amount').textContent = '₩' + formatNumber(defenseAmount);
    document.getElementById('perf-defense-count').textContent = defenseCount + '건';
    
    const totalAmount = data.total || (applyAmount + defenseAmount);
    document.getElementById('perf-total-amount').textContent = '₩' + formatNumber(totalAmount);
    document.getElementById('perf-apply-summary').textContent = '₩' + formatNumber(applyAmount);
    document.getElementById('perf-defense-summary').textContent = '₩' + formatNumber(defenseAmount);
    
    updateProgressRate(totalAmount);
}

// ==========================================
// 목표금액
// ==========================================

async function saveTargetAmount() {
    const input = document.getElementById('perf-target-input');
    const inputValue = input.value.trim();
    
    let targetAmount = parseTargetInput(inputValue);
    
    localStorage.setItem('perf_target_amount', targetAmount.toString());
    
    showToast('저장 중...');
    const saveData = { amount: targetAmount };
    const saved = await saveManualDataToSheets('target', saveData);
    
    const totalText = document.getElementById('perf-total-amount').textContent;
    const totalAmount = parseNumber(totalText.replace('₩', '').replace(/,/g, ''));
    updateProgressRate(totalAmount);
    
    if (saved) {
        showToast('목표금액이 저장되었습니다!');
    } else {
        showToast('저장에 실패했습니다. 로컬에만 저장됩니다.', 'error');
    }
}

function parseTargetInput(value) {
    if (!value) return 0;
    
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

function loadTargetAmount() {
    const saved = localStorage.getItem('perf_target_amount');
    if (saved) {
        const amount = parseInt(saved);
        const input = document.getElementById('perf-target-input');
        if (input && amount > 0) {
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

function updateProgressRate(totalAmount) {
    const targetAmount = loadTargetAmount();
    
    const targetDisplay = document.getElementById('perf-target-display');
    if (targetDisplay) {
        if (targetAmount > 0) {
            const eok = targetAmount / 100000000;
            targetDisplay.textContent = eok >= 1 ? `${eok.toFixed(1)}억` : '₩' + formatNumber(targetAmount);
        } else {
            targetDisplay.textContent = '미설정';
        }
    }
    
    const progressRate = targetAmount > 0 ? (totalAmount / targetAmount) * 100 : 0;
    
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
    
    const progressFill = document.getElementById('perf-progress-fill');
    if (progressFill) {
        progressFill.style.width = Math.min(progressRate, 100) + '%';
    }
}

// ==========================================
// 수동 데이터 적용 (API → UI)
// ==========================================

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

function applyTargetFromData(data) {
    const amount = data.amount || 0;
    
    localStorage.setItem('perf_target_amount', amount.toString());
    
    const input = document.getElementById('perf-target-input');
    if (input && amount > 0) {
        const eok = amount / 100000000;
        if (eok >= 1) {
            input.value = eok % 1 === 0 ? `${eok}억` : `${eok.toFixed(1)}억`;
        } else {
            input.value = formatNumber(amount);
        }
    }
    
    const totalText = document.getElementById('perf-total-amount')?.textContent || '₩0';
    const totalAmount = parseNumber(totalText.replace('₩', '').replace(/,/g, ''));
    updateProgressRate(totalAmount);
    
    console.log('목표금액 적용:', amount);
}

// ==========================================
// 활동수 현황 (자동 로드 - 필터 기반)
// ==========================================

async function loadActivityStatus() {
    const loadingEl = document.getElementById('activity-loading');
    const errorEl = document.getElementById('activity-error');
    const contentEl = document.getElementById('activity-content');
    
    if (loadingEl) loadingEl.style.display = 'flex';
    if (errorEl) errorEl.style.display = 'none';
    if (contentEl) contentEl.style.display = 'none';
    
    try {
        const year = selectedPerfMonth.getFullYear();
        const month = selectedPerfMonth.getMonth() + 1;
        const url = `${PERFORMANCE_API_URL}?action=activity&year=${year}&month=${month}`;
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
        
        updateActivityStatusUI(result.data);
        
        const syncTimeEl = document.getElementById('activity-sync-time');
        if (syncTimeEl) {
            syncTimeEl.textContent = new Date().toLocaleString('ko-KR');
        }
        
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

function updateActivityStatusUI(data) {
    const apply = data.apply || {};
    const defense = data.defense || {};
    
    const applyActivity = apply.activity || 0;
    const applyAbsent = apply.absent || 0;
    const applyFollowup = apply.followup || 0;
    const applySms = apply.sms || 0;
    
    const applyTotal = applyActivity + applyAbsent;
    const applyExtra = applyFollowup + applySms;
    
    const defenseActivity = defense.activity || 0;
    const defenseAbsent = defense.absent || 0;
    const defenseFollowup = defense.followup || 0;
    const defenseSms = defense.sms || 0;
    
    const defenseTotal = defenseActivity + defenseAbsent;
    const defenseExtra = defenseFollowup + defenseSms;
    
    const maxValue = Math.max(applyTotal, applyExtra, defenseTotal, defenseExtra, 1);
    
    updateTotalActivityBar('apply', applyActivity, applyAbsent, maxValue);
    updateExtraActivityBar('apply', applyFollowup, applySms, maxValue);
    document.getElementById('apply-total-value').textContent = formatNumber(applyTotal);
    document.getElementById('apply-extra-value').textContent = formatNumber(applyExtra);
    updateBarGroupOrder('apply', applyTotal, applyExtra);
    
    document.getElementById('tt-apply-activity').textContent = formatNumber(applyActivity);
    document.getElementById('tt-apply-absent').textContent = formatNumber(applyAbsent);
    document.getElementById('tt-apply-total').textContent = formatNumber(applyTotal);
    document.getElementById('tt-apply-followup').textContent = formatNumber(applyFollowup);
    document.getElementById('tt-apply-sms').textContent = formatNumber(applySms);
    document.getElementById('tt-apply-extra').textContent = formatNumber(applyExtra);
    
    updateTotalActivityBar('defense', defenseActivity, defenseAbsent, maxValue);
    updateExtraActivityBar('defense', defenseFollowup, defenseSms, maxValue);
    document.getElementById('defense-total-value').textContent = formatNumber(defenseTotal);
    document.getElementById('defense-extra-value').textContent = formatNumber(defenseExtra);
    updateBarGroupOrder('defense', defenseTotal, defenseExtra);
    
    document.getElementById('tt-defense-activity').textContent = formatNumber(defenseActivity);
    document.getElementById('tt-defense-absent').textContent = formatNumber(defenseAbsent);
    document.getElementById('tt-defense-total').textContent = formatNumber(defenseTotal);
    document.getElementById('tt-defense-followup').textContent = formatNumber(defenseFollowup);
    document.getElementById('tt-defense-sms').textContent = formatNumber(defenseSms);
    document.getElementById('tt-defense-extra').textContent = formatNumber(defenseExtra);
}

// ==========================================
// 초기화
// ==========================================

async function init() {
    setupToggleButtons();
    displayCurrentDate();
    initPerformanceSummary();
    loadActivityStatus();
    
    try {
        const data = await loadManualData();
        
        if (data.coverage) {
            applyCoverageFromData(data.coverage);
        }
        if (data.application) {
            applyApplicationFromData(data.application);
        }
        if (data.defense) {
            applyDefenseFromData(data.defense);
        }
        if (data.target && data.target.amount) {
            applyTargetFromData(data.target);
        }
        
        console.log('수동 데이터 로드 완료');
    } catch (error) {
        console.error('수동 데이터 로드 오류:', error);
    }
}

document.addEventListener('DOMContentLoaded', init);
