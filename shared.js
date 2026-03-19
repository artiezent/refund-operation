// shared.js - 공용 유틸리티

const PERFORMANCE_API_URL = 'https://script.google.com/macros/s/AKfycbwSt_945KvB4jGIj0DKjZFWl9Wr2_OZOs1LnyO4OKAs8YT0wtd-eH3BNo8b0ypXjb50GQ/exec';

// ==========================================
// 날짜 유틸
// ==========================================

function parseToKST(dateString) {
    if (!dateString) return null;
    if (dateString instanceof Date) return dateString;
    const date = new Date(dateString);
    if (isNaN(date.getTime())) return null;
    if (typeof dateString === 'string' && !dateString.includes('T')) {
        const parts = dateString.split('-');
        if (parts.length === 3) {
            return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
        }
    }
    if (typeof dateString === 'string' && dateString.includes('T')) {
        return new Date(date.getTime() + 9 * 60 * 60 * 1000);
    }
    return date;
}

function isSameDateKST(date1, date2) {
    if (!date1 || !date2) return false;
    const d1 = parseToKST(date1);
    const d2 = parseToKST(date2);
    if (!d1 || !d2) return false;
    return d1.getFullYear() === d2.getFullYear() &&
           d1.getMonth() === d2.getMonth() &&
           d1.getDate() === d2.getDate();
}

function getDaysDiffKST(startDate, endDate) {
    const start = parseToKST(startDate);
    const end = parseToKST(endDate);
    if (!start || !end) return null;
    const startDay = new Date(start.getFullYear(), start.getMonth(), start.getDate());
    const endDay = new Date(end.getFullYear(), end.getMonth(), end.getDate());
    return Math.floor((endDay - startDay) / (1000 * 60 * 60 * 24));
}

function getWeekStart(date) {
    const d = new Date(date);
    const day = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1);
    d.setDate(diff);
    d.setHours(0, 0, 0, 0);
    return d;
}

function getWeekEnd(weekStart) {
    const d = new Date(weekStart);
    d.setDate(d.getDate() + 6);
    d.setHours(23, 59, 59, 999);
    return d;
}

// ==========================================
// 숫자/포맷 유틸
// ==========================================

function parseNumber(str) {
    if (!str) return 0;
    const cleaned = String(str).replace(/,/g, '').replace(/[^0-9.-]/g, '');
    const num = parseFloat(cleaned);
    return isNaN(num) ? 0 : num;
}

function formatNumber(num) {
    if (num === 0 || num === null || num === undefined) return '0';
    return Math.round(num).toLocaleString('ko-KR');
}

// ==========================================
// UI 유틸
// ==========================================

function showToast(message, type = 'success') {
    const existingToast = document.querySelector('.toast');
    if (existingToast) existingToast.remove();
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    document.body.appendChild(toast);
    setTimeout(() => toast.classList.add('show'), 10);
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

function closeInputPanel(panelId) {
    const panel = document.getElementById(panelId);
    if (panel) {
        panel.classList.remove('active');
        const btn = document.querySelector(`[data-target="${panelId}"]`);
        if (btn) {
            btn.classList.remove('active');
            const icon = btn.querySelector('.toggle-icon');
            if (icon) icon.textContent = '+';
        }
    }
}

function setupToggleButtons() {
    document.querySelectorAll('.toggle-input-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const targetId = btn.dataset.target;
            const panel = document.getElementById(targetId);
            btn.classList.toggle('active');
            panel.classList.toggle('active');
            const icon = btn.querySelector('.toggle-icon');
            icon.textContent = btn.classList.contains('active') ? '×' : '+';
        });
    });
}

function updateKPICard(prefix, rate, values) {
    const rateEl = document.getElementById(`${prefix}-rate`);
    const barEl = document.getElementById(`${prefix}-bar`);
    if (!rateEl || !barEl) return;
    rateEl.textContent = rate.toFixed(1) + '%';
    barEl.style.width = Math.min(rate, 100) + '%';
    rateEl.classList.remove('success', 'warning', 'danger');
    barEl.classList.remove('success', 'warning', 'danger');
    let colorClass = '';
    if (rate >= 100) colorClass = 'success';
    else if (rate >= 80) colorClass = 'warning';
    else if (rate > 0) colorClass = 'danger';
    if (colorClass) { rateEl.classList.add(colorClass); barEl.classList.add(colorClass); }
    for (const [id, value] of Object.entries(values)) {
        const el = document.getElementById(id);
        if (el) el.textContent = value;
    }
}

// ==========================================
// 키워드 파싱
// ==========================================

function findAmountByKeyword(lines, keyword) {
    const keywordIndex = lines.findIndex(line => line === keyword);
    if (keywordIndex === -1) return 0;
    for (let i = 1; i <= 7; i++) {
        const targetIndex = keywordIndex + i;
        if (targetIndex >= lines.length) break;
        const amountMatch = lines[targetIndex].match(/^₩([\d,]+)/);
        if (amountMatch) return parseNumber(amountMatch[1]);
    }
    return 0;
}

function findAmountByExactKeyword(lines, keyword, excludeWords = []) {
    const keywordIndex = lines.findIndex(line => {
        if (!line.includes(keyword)) return false;
        for (const word of excludeWords) { if (line.includes(word)) return false; }
        return true;
    });
    if (keywordIndex === -1) return 0;
    for (let i = 1; i <= 7; i++) {
        const targetIndex = keywordIndex + i;
        if (targetIndex >= lines.length) break;
        const amountMatch = lines[targetIndex].match(/^₩([\d,]+)/);
        if (amountMatch) return parseNumber(amountMatch[1]);
    }
    return 0;
}

// ==========================================
// 데이터 처리
// ==========================================

function convertRowsToObjects(result) {
    if (result.data) return result.data;
    if (!result.cols || !result.rows) return [];
    const cols = result.cols;
    return result.rows.map(r => {
        const obj = {};
        for (let i = 0; i < cols.length; i++) obj[cols[i]] = r[i];
        return obj;
    });
}

function preprocessData(data) {
    const DAY_MS = 86400000;
    return data.map(deal => {
        const d = Object.assign({}, deal);
        d._value = Number(deal.value) || 0;
        d._noticeDate = deal.first_payment_notice ? parseToKST(deal.first_payment_notice) : null;
        d._wonDate = (deal.won_time && deal.won_time !== '') ? parseToKST(deal.won_time) : null;
        d._collectionDate = (deal.collection_order_date && deal.collection_order_date !== '') ? parseToKST(deal.collection_order_date) : null;
        d._hasWon = !!d._wonDate;
        d._hasCustomerType = !!(deal.customer_type && deal.customer_type.toString().trim() !== '');
        d._applyDate = (deal.apply_date && deal.apply_date !== '') ? parseToKST(deal.apply_date) : null;
        d._hasJupjupPerson = !!(deal.jupjup_person && deal.jupjup_person.toString().trim() !== '');
        d._defenseDate = (deal.defense_date && deal.defense_date !== '') ? parseToKST(deal.defense_date) : null;
        d._hasDefensePerson = !!(deal.defense_person && deal.defense_person.toString().trim() !== '');
        d._refundAmount = Number(deal.refund_amount) || 0;
        if (d._noticeDate && d._wonDate) {
            const s = new Date(d._noticeDate.getFullYear(), d._noticeDate.getMonth(), d._noticeDate.getDate());
            const e = new Date(d._wonDate.getFullYear(), d._wonDate.getMonth(), d._wonDate.getDate());
            d._daysDiff = Math.floor((e - s) / DAY_MS);
        } else {
            d._daysDiff = null;
        }
        return d;
    });
}

// ==========================================
// API 통신
// ==========================================

async function saveManualDataToSheets(type, data) {
    try {
        const encodedData = encodeURIComponent(JSON.stringify(data));
        const url = `${PERFORMANCE_API_URL}?action=saveManual&type=${type}&data=${encodedData}`;
        const response = await fetch(url);
        const result = await response.json();
        return result.success;
    } catch (error) {
        console.error(`${type} 저장 오류:`, error);
        return false;
    }
}

async function fetchPaymentData(scope, year, month) {
    let url = `${PERFORMANCE_API_URL}?action=payment&scope=${scope}`;
    if (year) url += `&year=${year}`;
    if (month) url += `&month=${month}`;
    const response = await fetch(url);
    const result = await response.json();
    if (!result.success) throw new Error('API 실패');
    return preprocessData(convertRowsToObjects(result));
}

async function loadManualData() {
    try {
        const url = `${PERFORMANCE_API_URL}?action=manual`;
        const response = await fetch(url);
        const result = await response.json();
        if (!result.success) return {};
        return result.data || {};
    } catch (e) {
        console.error('수동 데이터 로드 오류:', e);
        return {};
    }
}

// ==========================================
// 결제콜 커버리지 공용
// ==========================================

let _coverageCache = null;

async function fetchCoverageData() {
    if (_coverageCache) return _coverageCache;
    const url = `${PERFORMANCE_API_URL}?action=coverage`;
    const resp = await fetch(url);
    const result = await resp.json();
    if (!result.success) throw new Error(result.error || 'coverage API 실패');
    const cols = result.cols;
    _coverageCache = result.rows.map(r => {
        const obj = {};
        for (let i = 0; i < cols.length; i++) obj[cols[i]] = r[i];
        return obj;
    });
    return _coverageCache;
}

function calcCoverageKPI(data, highOnly) {
    const target = highOnly
        ? data.filter(d => d.is_high_value === 'Y' && d.is_auto_paid === 'N')
        : data.filter(d => d.is_auto_paid === 'N');
    const totalCount = target.length;

    const contacted = target.filter(d => d.has_payment_call === 'Y' || d.has_hard_collection === 'Y');
    const contactedCount = contacted.length;
    const contactRate = totalCount > 0 ? (contactedCount / totalCount * 100) : 0;

    function covDaysDiff(d1, d2) {
        if (!d1 || !d2) return Infinity;
        const a = new Date(d1), b = new Date(d2);
        if (isNaN(a) || isNaN(b)) return Infinity;
        return Math.floor((b - a) / 86400000);
    }

    function getFirstContact(deal) {
        const dates = [];
        if (deal.has_payment_call === 'Y' && deal.payment_call_date) dates.push(deal.payment_call_date);
        if (deal.has_hard_collection === 'Y' && deal.hard_collection_date) dates.push(deal.hard_collection_date);
        if (dates.length === 0) return null;
        dates.sort();
        return dates[0];
    }

    const contactDays = contacted.map(d => {
        const first = getFirstContact(d);
        return first ? covDaysDiff(d.first_payment_notice, first) : null;
    }).filter(d => d !== null && d !== Infinity);

    const avgDays = contactDays.length > 0 ? contactDays.reduce((a, b) => a + b, 0) / contactDays.length : 0;

    const pool = highOnly ? data.filter(d => d.is_high_value === 'Y') : data;
    const autoCount = pool.filter(d => d.is_auto_paid === 'Y').length;

    return { totalCount, contactedCount, contactRate, avgDays, autoCount, notContacted: totalCount - contactedCount };
}

function filterCoverageByMonth(data, year, month) {
    return data.filter(d => {
        if (!d.first_payment_notice) return false;
        const parts = d.first_payment_notice.split('-');
        return parseInt(parts[0]) === year && parseInt(parts[1]) === month;
    });
}

function filterCoverageByYear(data, year) {
    return data.filter(d => {
        if (!d.first_payment_notice) return false;
        return parseInt(d.first_payment_notice.split('-')[0]) === year;
    });
}
