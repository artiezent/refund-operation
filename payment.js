// payment.js - 결제파트 (Payment Monthly) Dashboard

// ==========================================
// State
// ==========================================

let paymentDataCache = null;
let selectedWeekStart = getWeekStart(new Date());
let selectedPaymentMonth = new Date();

const PAYMENT_TARGETS = {
    3: 80,
    30: 90,
    60: 95
};

// ==========================================
// Init
// ==========================================

document.addEventListener('DOMContentLoaded', init);

async function init() {
    setupToggleButtons();
    displayCurrentDate();
    updateWeekDisplay();
    loadPaymentTarget();

    loadPaymentData();
    loadCoverageForPayment();
}

function displayCurrentDate() {
    const el = document.getElementById('currentDate');
    if (!el) return;
    const now = new Date();
    const options = { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long' };
    el.textContent = now.toLocaleDateString('ko-KR', options);
}

// ==========================================
// Data Loading
// ==========================================

async function loadPaymentData() {
    const loadingEl = document.getElementById('payment-loading');
    const errorEl = document.getElementById('payment-error');
    const contentEl = document.getElementById('payment-content');

    if (loadingEl) loadingEl.style.display = 'flex';
    if (errorEl) errorEl.style.display = 'none';
    if (contentEl) contentEl.style.display = 'none';

    try {
        const year = selectedPaymentMonth.getFullYear();
        const month = selectedPaymentMonth.getMonth() + 1;
        paymentDataCache = await fetchPaymentData('monthly', year, month);

        updatePaymentMonthDisplay();
        calculateAndDisplayMonthlyPayment(paymentDataCache);
        calculateAndDisplayWeeklyKPIs(paymentDataCache);

        const syncTimeEl = document.getElementById('last-sync-time');
        if (syncTimeEl) {
            syncTimeEl.textContent = new Date().toLocaleString('ko-KR');
        }

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

function refreshAllData() {
    paymentDataCache = null;
    loadPaymentData();
}

// ==========================================
// Week Navigation
// ==========================================

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

function changeWeek(delta) {
    selectedWeekStart.setDate(selectedWeekStart.getDate() + (delta * 7));
    updateWeekDisplay();

    if (paymentDataCache) {
        calculateAndDisplayWeeklyKPIs(paymentDataCache);
    }
}

// ==========================================
// Month Navigation
// ==========================================

function changePaymentMonth(delta) {
    selectedPaymentMonth.setMonth(selectedPaymentMonth.getMonth() + delta);
    updatePaymentMonthDisplay();
    loadPaymentData();
    loadCoverageForPayment();
}

function updatePaymentMonthDisplay() {
    const year = selectedPaymentMonth.getFullYear();
    const month = selectedPaymentMonth.getMonth() + 1;

    document.getElementById('selected-payment-month').textContent = `${year}년 ${month}월`;
    document.getElementById('monthly-payment-title').textContent = `${month}월 결제금액`;
}

// ==========================================
// Monthly Payment Calculation
// ==========================================

function calculateAndDisplayMonthlyPayment(data) {
    const year = selectedPaymentMonth.getFullYear();
    const month = selectedPaymentMonth.getMonth();

    const refundDeals = data.filter(d => d._noticeDate && d._noticeDate.getFullYear() === year && d._noticeDate.getMonth() === month);
    const refundAmount = refundDeals.reduce((s, d) => s + d._value, 0);
    const refundCount = refundDeals.length;

    document.getElementById('monthly-refund-amount').textContent = '₩' + formatNumber(refundAmount);
    document.getElementById('monthly-refund-count').textContent = refundCount + '건';
    const refundTitleEl = document.getElementById('monthly-refund-title');
    if (refundTitleEl) refundTitleEl.textContent = `${month + 1}월 환급완료`;

    const monthlyDeals = data.filter(d => d._wonDate && d._wonDate.getFullYear() === year && d._wonDate.getMonth() === month);
    const totalAmount = monthlyDeals.reduce((s, d) => s + d._value, 0);
    const totalCount = monthlyDeals.length;

    document.getElementById('monthly-payment-amount').textContent = '₩' + formatNumber(totalAmount);
    document.getElementById('monthly-payment-count').textContent = totalCount + '건';

    updatePaymentProgress(totalAmount);
    calculateCollectionPaid(data, 'mp', year, month);

    console.log(`=== ${year}년 ${month + 1}월 ===`);
    console.log(`환급완료: ${refundCount}건, ₩${formatNumber(refundAmount)}`);
    console.log(`결제금액: ${totalCount}건, ₩${formatNumber(totalAmount)}`);
}

// ==========================================
// 추심결제완료 (stage_name 기준)
// ==========================================

function calculateCollectionPaid(data, prefix, yearFilter, monthFilter) {
    const colPaidDeals = data.filter(d => {
        if (!d._hasWon || !d._wonDate) return false;
        const stage = (d.stage_name || '').toString();
        if (!stage.includes('추심') || !stage.includes('결제완료')) return false;
        if (d._wonDate.getFullYear() !== yearFilter) return false;
        if (monthFilter !== undefined && d._wonDate.getMonth() !== monthFilter) return false;
        return true;
    });
    const colPaidCount = colPaidDeals.length;
    const colPaidAmount = colPaidDeals.reduce((s, d) => s + d._value, 0);

    const periodPaidDeals = data.filter(d => {
        if (!d._hasWon || !d._wonDate) return false;
        if (d._wonDate.getFullYear() !== yearFilter) return false;
        if (monthFilter !== undefined && d._wonDate.getMonth() !== monthFilter) return false;
        return true;
    });
    const periodPaidTotal = periodPaidDeals.reduce((s, d) => s + d._value, 0);
    const ratio = periodPaidTotal > 0 ? (colPaidAmount / periodPaidTotal * 100) : 0;

    const amtEl = document.getElementById(`${prefix}-colpaid-amount`);
    if (amtEl) amtEl.textContent = '₩' + formatNumber(colPaidAmount);
    const cntEl = document.getElementById(`${prefix}-colpaid-count`);
    if (cntEl) cntEl.textContent = colPaidCount + '건';
    const ratioEl = document.getElementById(`${prefix}-colpaid-ratio`);
    if (ratioEl) ratioEl.textContent = ratio.toFixed(1) + '%';
}

// ==========================================
// Payment Target
// ==========================================

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

function savePaymentTarget() {
    const input = document.getElementById('payment-target-input');
    const value = input.value.trim();
    const target = parseTargetInput(value);
    localStorage.setItem('payment_target_amount', target.toString());

    const totalText = document.getElementById('monthly-payment-amount').textContent;
    const totalAmount = parseNumber(totalText.replace(/[₩,]/g, ''));
    updatePaymentProgress(totalAmount);

    if (target > 0) {
        showToast('결제 목표금액이 저장되었습니다!');
    }
}

function loadPaymentTarget() {
    const saved = localStorage.getItem('payment_target_amount');
    if (saved) {
        const amount = parseInt(saved);
        const input = document.getElementById('payment-target-input');
        if (input && amount > 0) {
            const eok = amount / 100000000;
            if (eok >= 1) {
                input.value = eok % 1 === 0 ? `${eok}억` : `${eok.toFixed(1)}억`;
            } else {
                const man = amount / 10000;
                input.value = man >= 1 ? `${Math.round(man)}만` : formatNumber(amount);
            }
        }
        return amount;
    }
    return 0;
}

function updatePaymentProgress(totalAmount) {
    const target = loadPaymentTarget();

    const displayEl = document.getElementById('payment-target-display');
    if (displayEl) {
        if (target > 0) {
            const eok = target / 100000000;
            displayEl.textContent = eok >= 1 ? `₩${eok % 1 === 0 ? eok : eok.toFixed(1)}억` : '₩' + formatNumber(target);
        } else {
            displayEl.textContent = '미설정';
        }
    }

    const rate = target > 0 ? (totalAmount / target) * 100 : 0;

    const rateEl = document.getElementById('payment-progress-rate');
    if (rateEl) {
        rateEl.textContent = rate.toFixed(1) + '%';
        rateEl.style.color = '';
        if (rate >= 100) rateEl.style.color = 'var(--accent-success)';
        else if (rate >= 80) rateEl.style.color = 'var(--accent-primary)';
        else if (rate > 0) rateEl.style.color = 'var(--accent-danger)';
    }

    const fillEl = document.getElementById('payment-progress-fill');
    if (fillEl) fillEl.style.width = Math.min(rate, 100) + '%';
}

// ==========================================
// Weekly KPI Calculation
// ==========================================

function calculateAndDisplayWeeklyKPIs(data) {
    const weekStart = selectedWeekStart;
    const weekEnd = getWeekEnd(weekStart);

    console.log(`=== 주차별 KPI 계산 ===`);
    console.log(`기간: ${weekStart.toLocaleDateString()} ~ ${weekEnd.toLocaleDateString()}`);

    const refundDeals = data.filter(d => d._noticeDate && d._noticeDate >= weekStart && d._noticeDate <= weekEnd);

    const refundCount = refundDeals.length;
    const refundAmount = refundDeals.reduce((s, d) => s + d._value, 0);

    document.getElementById('weekly-refund-count').textContent = formatNumber(refundCount) + '건';
    document.getElementById('weekly-refund-amount').textContent = '₩' + formatNumber(refundAmount);

    console.log(`환급 완료: ${refundCount}건, ₩${formatNumber(refundAmount)}`);

    const sameDayDeals = refundDeals.filter(d => d._hasWon && d._daysDiff === 0);

    const sameDayCount = sameDayDeals.length;
    const sameDayAmount = sameDayDeals.reduce((s, d) => s + d._value, 0);
    const sameDayCountRate = refundCount > 0 ? (sameDayCount / refundCount) * 100 : 0;

    document.getElementById('weekly-sameday-count').textContent = formatNumber(sameDayCount) + '건';
    document.getElementById('weekly-sameday-amount').textContent = '₩' + formatNumber(sameDayAmount);
    updateRateDisplay('weekly-sameday-rate', sameDayCountRate);

    console.log(`당일 결제: ${sameDayCount}건, ₩${formatNumber(sameDayAmount)}, ${sameDayCountRate.toFixed(1)}%`);

    const within30dDeals = refundDeals.filter(d => d._hasWon && d._daysDiff !== null && d._daysDiff >= 0 && d._daysDiff <= 29);

    const within30dCount = within30dDeals.length;
    const within30dAmount = within30dDeals.reduce((s, d) => s + d._value, 0);
    const within30dCountRate = refundCount > 0 ? (within30dCount / refundCount) * 100 : 0;

    document.getElementById('weekly-30d-count').textContent = formatNumber(within30dCount) + '건';
    document.getElementById('weekly-30d-amount').textContent = '₩' + formatNumber(within30dAmount);
    updateRateDisplay('weekly-30d-rate', within30dCountRate);

    console.log(`30일이내 결제: ${within30dCount}건, ₩${formatNumber(within30dAmount)}, ${within30dCountRate.toFixed(1)}%`);

    const periods = [0, 3, 7, 21, 30, 60];
    const periodLabelsMap = { 0: '당일', 3: '3일 이내', 7: '7일 이내', 21: '21일 이내', 30: '30일 이내', 60: '60일 이내' };
    const paymentPeriodDeals = {};

    periods.forEach(days => {
        const result = calculateWeeklyConversionRate(refundDeals, days);
        updateWeeklyConversionRow(days, result, refundCount, refundAmount);
        paymentPeriodDeals[days] = result.convertedDeals;
    });

    const unpaidDeals = refundDeals.filter(d => !d._hasWon);
    paymentPeriodDeals['unpaid'] = unpaidDeals;

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

// ==========================================
// Weekly Conversion Rate
// ==========================================

function calculateWeeklyConversionRate(refundDeals, targetDays) {
    const convertedDeals = refundDeals.filter(d => {
        if (!d._hasWon || d._daysDiff === null) return false;
        return targetDays === 0 ? d._daysDiff === 0 : (d._daysDiff >= 0 && d._daysDiff <= targetDays - 1);
    });

    const convertedCount = convertedDeals.length;
    const convertedAmount = convertedDeals.reduce((s, d) => s + d._value, 0);

    return {
        convertedCount,
        convertedAmount,
        convertedDeals
    };
}

function updateWeeklyConversionRow(days, result, totalCount, totalAmount) {
    const countRate = totalCount > 0 ? (result.convertedCount / totalCount) * 100 : 0;
    const amountRate = totalAmount > 0 ? (result.convertedAmount / totalAmount) * 100 : 0;

    const countRateEl = document.getElementById(`weekly-rate-cnt-${days}d`);
    if (countRateEl) {
        countRateEl.textContent = countRate.toFixed(1) + '%';
        countRateEl.className = 'rate-value';
        if (countRate >= 70) countRateEl.classList.add('rate-high');
        else if (countRate >= 40) countRateEl.classList.add('rate-medium');
        else if (countRate > 0) countRateEl.classList.add('rate-low');
    }

    const amountRateEl = document.getElementById(`weekly-rate-amt-${days}d`);
    if (amountRateEl) {
        amountRateEl.textContent = amountRate.toFixed(1) + '%';
        amountRateEl.className = 'rate-value rate-amount';
        if (amountRate >= 70) amountRateEl.classList.add('rate-high');
        else if (amountRate >= 40) amountRateEl.classList.add('rate-medium');
        else if (amountRate > 0) amountRateEl.classList.add('rate-low');
    }

    const countDetailEl = document.getElementById(`weekly-detail-cnt-${days}d`);
    if (countDetailEl) {
        countDetailEl.textContent = `${result.convertedCount} / ${totalCount}`;
    }

    const amountDetailEl = document.getElementById(`weekly-detail-amt-${days}d`);
    if (amountDetailEl) {
        const convertedEok = (result.convertedAmount / 100000000).toFixed(2);
        const totalEok = (totalAmount / 100000000).toFixed(2);
        amountDetailEl.textContent = `${convertedEok}억 / ${totalEok}억`;
    }

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

// ==========================================
// Rate Display
// ==========================================

function updateRateDisplay(elementId, rate) {
    const el = document.getElementById(elementId);
    if (el) {
        el.textContent = rate.toFixed(1) + '%';
        el.classList.remove('warning', 'danger');
        if (rate >= 70) {
            // default green
        } else if (rate >= 40) {
            el.classList.add('warning');
        } else {
            el.classList.add('danger');
        }
    }
}

// ==========================================
// Payment Detail Panel
// ==========================================

function showPaymentDetail(deals, label) {
    const panel = document.getElementById('payment-detail-panel');
    const titleEl = document.getElementById('payment-detail-title');
    const tbodyEl = document.getElementById('payment-detail-tbody');
    if (!panel || !titleEl || !tbodyEl) return;

    const totalAmt = deals.reduce((s, d) => s + d._value, 0);
    titleEl.textContent = `${label} — ${deals.length}건, ₩${totalAmt.toLocaleString()}`;

    const sorted = [...deals].sort((a, b) => b._value - a._value);

    tbodyEl.innerHTML = sorted.map((d, i) => {
        const daysLabel = d._daysDiff !== null ? `${d._daysDiff}일` : '-';
        return `<tr>
            <td>${i + 1}</td>
            <td>${d.title || '-'}</td>
            <td class="td-amount">₩${d._value.toLocaleString()}</td>
            <td>${d.first_payment_notice || '-'}</td>
            <td>${d.won_time || '미결제'}</td>
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

// ==========================================
// CSV Downloads
// ==========================================

function formatDateForCSV(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function downloadMonthlyRawData() {
    if (!paymentDataCache) {
        showToast('데이터가 없습니다. 먼저 데이터를 로드해주세요.', 'error');
        return;
    }

    const year = selectedPaymentMonth.getFullYear();
    const month = selectedPaymentMonth.getMonth();

    const monthlyDeals = paymentDataCache.filter(d => d._wonDate && d._wonDate.getFullYear() === year && d._wonDate.getMonth() === month);

    if (monthlyDeals.length === 0) {
        showToast('해당 월에 데이터가 없습니다.', 'error');
        return;
    }

    const headers = ['거래ID', '고객명', '금액', '성사일(원본)', '성사일(KST)', '최초결제안내일'];

    const rows = monthlyDeals.map(deal => [
        deal.id || '',
        deal.person_name || deal.title || '',
        deal._value,
        deal.won_time || '',
        deal._wonDate ? formatDateForCSV(deal._wonDate) : '',
        deal._noticeDate ? formatDateForCSV(deal._noticeDate) : ''
    ]);

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

function downloadWeeklyRawData() {
    if (!paymentDataCache) {
        showToast('데이터가 없습니다. 먼저 데이터를 로드해주세요.', 'error');
        return;
    }

    const weekStart = selectedWeekStart;
    const weekEnd = getWeekEnd(weekStart);

    const weeklyDeals = paymentDataCache.filter(d => d._noticeDate && d._noticeDate >= weekStart && d._noticeDate <= weekEnd);

    if (weeklyDeals.length === 0) {
        showToast('해당 주차에 데이터가 없습니다.', 'error');
        return;
    }

    const headers = [
        '거래ID', '고객명', '금액', '최초결제안내일', '성사일', '일수차이',
        '결제여부', '당일결제', '3일이내', '7일이내', '21일이내', '30일이내', '60일이내'
    ];

    const rows = weeklyDeals.map(deal => {
        const dd = deal._daysDiff;
        const isPaid = deal._hasWon;
        return [
            deal.id || '',
            deal.person_name || deal.title || '',
            deal._value,
            deal._noticeDate ? formatDateForCSV(deal._noticeDate) : '',
            deal._wonDate ? formatDateForCSV(deal._wonDate) : '',
            dd !== null ? dd : '',
            isPaid ? 'Y' : 'N',
            isPaid && dd === 0 ? 'Y' : 'N',
            isPaid && dd >= 0 && dd <= 2 ? 'Y' : 'N',
            isPaid && dd >= 0 && dd <= 6 ? 'Y' : 'N',
            isPaid && dd >= 0 && dd <= 20 ? 'Y' : 'N',
            isPaid && dd >= 0 && dd <= 29 ? 'Y' : 'N',
            isPaid && dd >= 0 && dd <= 59 ? 'Y' : 'N'
        ];
    });

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

// ==========================================
// Payment Coverage (자동 로드)
// ==========================================

async function loadCoverageForPayment() {
    const statusEl = document.getElementById('mp-cov-status');
    try {
        const allCov = await fetchCoverageData();
        const year = selectedPaymentMonth.getFullYear();
        const month = selectedPaymentMonth.getMonth() + 1;
        const monthly = filterCoverageByMonth(allCov, year, month);
        renderCoverageCards(monthly, 'mp');
        if (statusEl) { statusEl.textContent = `${monthly.length}건`; statusEl.className = 'kpi-badge badge-green'; }
    } catch (e) {
        console.error('커버리지 로드 실패:', e);
        if (statusEl) { statusEl.textContent = '로드 실패'; statusEl.className = 'kpi-badge badge-red'; }
    }
}

function renderCoverageCards(data, prefix) {
    const high = calcCoverageKPI(data, true);
    const total = calcCoverageKPI(data, false);

    const setColor = (el, rate) => {
        if (!el) return;
        el.style.color = '';
        if (rate >= 90) el.style.color = 'var(--accent-success)';
        else if (rate >= 70) el.style.color = 'var(--accent-primary)';
        else if (rate > 0) el.style.color = 'var(--accent-danger)';
    };

    const highRateEl = document.getElementById(`${prefix}-high-coverage-rate`);
    if (highRateEl) { highRateEl.textContent = high.contactRate.toFixed(1) + '%'; setColor(highRateEl, high.contactRate); }
    const highDetailEl = document.getElementById(`${prefix}-high-coverage-detail`);
    if (highDetailEl) highDetailEl.textContent = `${high.contactedCount}건 / ${high.totalCount}건`;
    const highDaysEl = document.getElementById(`${prefix}-high-avg-days`);
    if (highDaysEl) highDaysEl.textContent = high.avgDays > 0 ? high.avgDays.toFixed(1) + '일' : '-';

    const totalRateEl = document.getElementById(`${prefix}-total-coverage-rate`);
    if (totalRateEl) { totalRateEl.textContent = total.contactRate.toFixed(1) + '%'; setColor(totalRateEl, total.contactRate); }
    const totalDetailEl = document.getElementById(`${prefix}-total-coverage-detail`);
    if (totalDetailEl) totalDetailEl.textContent = `${total.contactedCount}건 / ${total.totalCount}건`;
    const totalDaysEl = document.getElementById(`${prefix}-total-avg-days`);
    if (totalDaysEl) totalDaysEl.textContent = total.avgDays > 0 ? total.avgDays.toFixed(1) + '일' : '-';
    const totalFill = document.getElementById(`${prefix}-total-coverage-fill`);
    if (totalFill) totalFill.style.width = Math.min(total.contactRate, 100) + '%';
}

// ==========================================
// 요약
// ==========================================

function generateSummary() {
    const today = new Date();
    const dateStr = `${today.getFullYear()}년 ${today.getMonth() + 1}월 ${today.getDate()}일`;
    let s = `💳 결제파트 요약 (${dateStr})\n${'='.repeat(40)}\n\n`;

    const ml = document.getElementById('selected-payment-month')?.textContent || '';
    s += `📌 ${ml} 결제금액\n${'-'.repeat(30)}\n`;
    s += `• 환급완료: ${document.getElementById('monthly-refund-amount')?.textContent || '₩0'} (${document.getElementById('monthly-refund-count')?.textContent || '0건'})\n`;
    s += `• 결제금액: ${document.getElementById('monthly-payment-amount')?.textContent || '₩0'} (${document.getElementById('monthly-payment-count')?.textContent || '0건'})\n`;
    s += `• 추심결제완료: ${document.getElementById('mp-colpaid-amount')?.textContent || '₩0'} (${document.getElementById('mp-colpaid-count')?.textContent || '0건'})\n\n`;

    const wl = document.getElementById('selected-week')?.textContent || '';
    s += `📌 주차별 KPI (${wl})\n${'-'.repeat(30)}\n`;
    s += `• 환급완료: ${document.getElementById('weekly-refund-count')?.textContent || '0건'} / ${document.getElementById('weekly-refund-amount')?.textContent || '₩0'}\n`;
    s += `• 당일결제: ${document.getElementById('weekly-sameday-count')?.textContent || '0건'} (${document.getElementById('weekly-sameday-rate')?.textContent || '0%'})\n`;
    s += `• 30일이내: ${document.getElementById('weekly-30d-count')?.textContent || '0건'} (${document.getElementById('weekly-30d-rate')?.textContent || '0%'})\n\n`;

    s += `📌 기간별 결제율\n${'-'.repeat(30)}\n`;
    [0, 3, 7, 21, 30, 60].forEach(d => {
        const label = d === 0 ? '당일' : `${d}일`;
        s += `• ${label}이내: 건수 ${document.getElementById(`weekly-rate-cnt-${d}d`)?.textContent || '-'} / 금액 ${document.getElementById(`weekly-rate-amt-${d}d`)?.textContent || '-'}\n`;
    });
    return s;
}

// ==========================================
// Debug: Customer Search
// ==========================================

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
        const wonDateKSTStr = deal._wonDate ? deal._wonDate.toLocaleString('ko-KR') : '(없음)';
        console.log('---');
        console.log(`ID: ${deal.id}`);
        console.log(`고객명: ${deal.person_name || deal.title}`);
        console.log(`금액: ${deal._value}`);
        console.log(`성사일(원본): ${deal.won_time || '(없음)'}`);
        console.log(`성사일(KST): ${wonDateKSTStr}`);
        console.log(`최초결제안내일: ${deal.first_payment_notice || '(없음)'}`);
        if (deal._wonDate) {
            console.log(`→ KST 기준 월: ${deal._wonDate.getFullYear()}년 ${deal._wonDate.getMonth() + 1}월`);
        }
    });

    return results;
};
