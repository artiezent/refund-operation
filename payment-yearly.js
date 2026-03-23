// payment-yearly.js - 결제파트 누적 대시보드

let pyYear = new Date().getFullYear();
let pyDataLoaded = false;
let paymentYearlyDataCache = null;

// ==========================================
// 데이터 로드
// ==========================================

async function loadPaymentYearlyData() {
    const loadingEl = document.getElementById('py-loading');
    const errorEl = document.getElementById('py-error');
    const contentEl = document.getElementById('py-content');

    if (loadingEl) loadingEl.style.display = 'flex';
    if (errorEl) errorEl.style.display = 'none';
    if (contentEl) contentEl.style.display = 'none';

    try {
        const data = await fetchPaymentData('yearly', pyYear);
        paymentYearlyDataCache = data;
        pyDataLoaded = true;
        calculatePaymentYearlyKPIs(data);

        const syncEl = document.getElementById('py-last-sync-time');
        if (syncEl) syncEl.textContent = new Date().toLocaleString('ko-KR');

        if (loadingEl) loadingEl.style.display = 'none';
        if (contentEl) contentEl.style.display = 'block';
    } catch (error) {
        if (loadingEl) loadingEl.style.display = 'none';
        if (errorEl) {
            errorEl.style.display = 'flex';
            const msgEl = document.getElementById('py-error-message');
            if (msgEl) msgEl.textContent = error.message || '네트워크 오류';
        }
    }
}

function refreshAllData() {
    paymentYearlyDataCache = null;
    pyDataLoaded = false;
    loadPaymentYearlyData();
}

// ==========================================
// 연도 변경
// ==========================================

function changePyYear(delta) {
    pyYear += delta;
    updatePyYearDisplay();
    if (paymentYearlyDataCache) {
        calculatePaymentYearlyKPIs(paymentYearlyDataCache);
    }
    loadCoverageForYearly();
}

function updatePyYearDisplay() {
    const el = document.getElementById('py-selected-year');
    if (el) el.textContent = `${pyYear}년`;
    const refundTitle = document.getElementById('py-refund-title');
    if (refundTitle) refundTitle.textContent = `${pyYear}년 환급완료`;
    const payTitle = document.getElementById('py-payment-title');
    if (payTitle) payTitle.textContent = `${pyYear}년 결제금액`;
}

// ==========================================
// KPI 계산 (app.js 동일 로직)
// ==========================================

function calculatePaymentYearlyKPIs(data) {
    const year = pyYear;

    // 1. 연간 누적 환급완료 (최초결제안내일 기준)
    const refundDeals = data.filter(d => d._noticeDate && d._noticeDate.getFullYear() === year);
    const refundCount = refundDeals.length;
    const refundAmount = refundDeals.reduce((s, d) => s + d._value, 0);

    document.getElementById('py-refund-amount').textContent = '₩' + formatNumber(refundAmount);
    document.getElementById('py-refund-count').textContent = refundCount + '건';

    // 2. 연간 누적 결제금액 (성사일 기준)
    const paidDeals = data.filter(d => d._wonDate && d._wonDate.getFullYear() === year);
    const paidCount = paidDeals.length;
    const paidAmount = paidDeals.reduce((s, d) => s + d._value, 0);

    document.getElementById('py-payment-amount').textContent = '₩' + formatNumber(paidAmount);
    document.getElementById('py-payment-count').textContent = paidCount + '건';

    updatePyProgress(paidAmount);

    // 추심-결제완료
    calculateCollectionPaid(data, 'py', year);

    // 3. 연간 누적 결제율
    const pyPeriods = [0, 7, 30, 90];
    pyPeriods.forEach(days => {
        const converted = refundDeals.filter(d => {
            if (!d._hasWon || d._daysDiff === null) return false;
            return days === 0 ? d._daysDiff === 0 : (d._daysDiff >= 0 && d._daysDiff <= days - 1);
        });

        const cntRate = refundCount > 0 ? (converted.length / refundCount) * 100 : 0;
        const convertedAmt = converted.reduce((s, d) => s + d._value, 0);
        const amtRate = refundAmount > 0 ? (convertedAmt / refundAmount) * 100 : 0;

        const cntRateEl = document.getElementById(`py-rate-cnt-${days}d`);
        if (cntRateEl) {
            cntRateEl.textContent = cntRate.toFixed(1) + '%';
            cntRateEl.className = 'rate-value';
            if (cntRate >= 70) cntRateEl.classList.add('rate-high');
            else if (cntRate >= 40) cntRateEl.classList.add('rate-medium');
            else if (cntRate > 0) cntRateEl.classList.add('rate-low');
        }

        const amtRateEl = document.getElementById(`py-rate-amt-${days}d`);
        if (amtRateEl) {
            amtRateEl.textContent = amtRate.toFixed(1) + '%';
            amtRateEl.className = 'rate-value rate-amount';
            if (amtRate >= 70) amtRateEl.classList.add('rate-high');
            else if (amtRate >= 40) amtRateEl.classList.add('rate-medium');
            else if (amtRate > 0) amtRateEl.classList.add('rate-low');
        }

        const cntDetailEl = document.getElementById(`py-detail-cnt-${days}d`);
        if (cntDetailEl) cntDetailEl.textContent = `${converted.length} / ${refundCount}`;

        const amtDetailEl = document.getElementById(`py-detail-amt-${days}d`);
        if (amtDetailEl) {
            amtDetailEl.textContent = `${(convertedAmt / 100000000).toFixed(2)}억 / ${(refundAmount / 100000000).toFixed(2)}억`;
        }
    });

    // 4. 추심 이관 비율
    const transferDeals = data.filter(d => d._collectionDate && d._collectionDate.getFullYear() === year);
    const transferCount = transferDeals.length;
    const transferAmount = transferDeals.reduce((s, d) => s + d._value, 0);
    const colRate = refundAmount > 0 ? (transferAmount / refundAmount) * 100 : 0;

    document.getElementById('py-collection-rate').textContent = colRate.toFixed(1) + '%';
    document.getElementById('py-col-refund').textContent = '₩' + formatNumber(refundAmount);
    document.getElementById('py-col-transfer').textContent = '₩' + formatNumber(transferAmount);
    document.getElementById('py-col-transfer-count').textContent = transferCount + '건';

    const colFill = document.getElementById('py-collection-fill');
    if (colFill) colFill.style.width = Math.min(colRate, 100) + '%';
}

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
// 연간 목표
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

function savePaymentYearlyTarget() {
    const input = document.getElementById('py-target-input');
    const target = parseTargetInput(input.value.trim());
    localStorage.setItem('py_target_amount', target.toString());

    const totalText = document.getElementById('py-payment-amount').textContent;
    const totalAmount = parseNumber(totalText.replace(/[₩,]/g, ''));
    updatePyProgress(totalAmount);

    if (target > 0) showToast('연간 목표금액이 저장되었습니다!');
}

function loadPaymentYearlyTarget() {
    const saved = localStorage.getItem('py_target_amount');
    if (saved) {
        const amount = parseInt(saved);
        const input = document.getElementById('py-target-input');
        if (input && amount > 0) {
            const eok = amount / 100000000;
            if (eok >= 1) input.value = eok % 1 === 0 ? `${eok}억` : `${eok.toFixed(1)}억`;
            else input.value = Math.round(amount / 10000) + '만';
        }
        return amount;
    }
    return 0;
}

function updatePyProgress(totalAmount) {
    const target = loadPaymentYearlyTarget();

    const displayEl = document.getElementById('py-target-display');
    if (displayEl) {
        if (target > 0) {
            const eok = target / 100000000;
            displayEl.textContent = eok >= 1 ? `₩${eok % 1 === 0 ? eok : eok.toFixed(1)}억` : '₩' + formatNumber(target);
        } else {
            displayEl.textContent = '미설정';
        }
    }

    const rate = target > 0 ? (totalAmount / target) * 100 : 0;
    const rateEl = document.getElementById('py-progress-rate');
    if (rateEl) {
        rateEl.textContent = rate.toFixed(1) + '%';
        rateEl.style.color = '';
        if (rate >= 100) rateEl.style.color = 'var(--accent-success)';
        else if (rate >= 80) rateEl.style.color = 'var(--accent-primary)';
        else if (rate > 0) rateEl.style.color = 'var(--accent-danger)';
    }

    const fillEl = document.getElementById('py-progress-fill');
    if (fillEl) fillEl.style.width = Math.min(rate, 100) + '%';
}

// ==========================================
// 결제콜 커버리지 (자동 로드)
// ==========================================

async function loadCoverageForYearly() {
    const statusEl = document.getElementById('py-cov-status');
    try {
        const allCov = await fetchCoverageData();
        const yearly = filterCoverageByYear(allCov, pyYear);
        renderCoverageCards(yearly, 'py');
        if (statusEl) { statusEl.textContent = `${yearly.length}건`; statusEl.className = 'kpi-badge badge-green'; }
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
// 날짜 표시
// ==========================================

function displayCurrentDate() {
    const now = new Date();
    const options = { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long' };
    const dateStr = now.toLocaleDateString('ko-KR', options);
    const el = document.getElementById('currentDate');
    if (el) el.textContent = dateStr;
}

// ==========================================
// 요약
// ==========================================

function generateSummary() {
    const today = new Date();
    const dateStr = `${today.getFullYear()}년 ${today.getMonth() + 1}월 ${today.getDate()}일`;
    let s = `💳 결제파트 누적 요약 (${dateStr})\n${'='.repeat(40)}\n\n`;

    const yl = document.getElementById('selected-py-year')?.textContent || '';
    s += `📌 ${yl} 연간 결제\n${'-'.repeat(30)}\n`;
    s += `• 환급완료: ${document.getElementById('yearly-refund-amount')?.textContent || '₩0'} (${document.getElementById('yearly-refund-count')?.textContent || '0건'})\n`;
    s += `• 결제금액: ${document.getElementById('yearly-payment-amount')?.textContent || '₩0'} (${document.getElementById('yearly-payment-count')?.textContent || '0건'})\n`;
    s += `• 추심결제완료: ${document.getElementById('py-colpaid-amount')?.textContent || '₩0'} (${document.getElementById('py-colpaid-count')?.textContent || '0건'})\n\n`;

    s += `📌 연간 커버리지\n${'-'.repeat(30)}\n`;
    s += `• 고액 컨택률: ${document.getElementById('py-high-coverage-rate')?.textContent || '-'} (${document.getElementById('py-high-coverage-detail')?.textContent || '-'})\n`;
    s += `• 전체 컨택률: ${document.getElementById('py-total-coverage-rate')?.textContent || '-'} (${document.getElementById('py-total-coverage-detail')?.textContent || '-'})\n\n`;

    s += `📌 월별 결제 추적\n${'-'.repeat(30)}\n`;
    document.querySelectorAll('#py-yearly-body tr').forEach(row => {
        const cells = row.querySelectorAll('td');
        if (cells.length < 3) return;
        const m = cells[0]?.textContent?.trim() || '';
        const refund = cells[1]?.textContent?.trim().replace(/\s+/g, ' ') || '';
        const paid = cells[2]?.textContent?.trim().replace(/\s+/g, ' ') || '';
        if (refund.includes('₩0') && paid.includes('₩0')) return;
        s += `• ${m}: 환급 ${refund} / 결제 ${paid}\n`;
    });
    return s;
}

// ==========================================
// 초기화
// ==========================================

async function init() {
    setupToggleButtons();
    displayCurrentDate();
    updatePyYearDisplay();
    loadPaymentYearlyTarget();

    loadPaymentYearlyData();
    loadCoverageForYearly();
}

document.addEventListener('DOMContentLoaded', () => {
    init();
});
