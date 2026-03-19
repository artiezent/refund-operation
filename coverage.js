// coverage.js - 결제콜 커버리지 대시보드

let allData = [];
let selectedYear = new Date().getFullYear();
let selectedMonth = new Date().getMonth() + 1;
let highPeriod = 'month';
let totalPeriod = 'month';
let currentTab = 'high_noauto';

function displayCurrentDate() {
    const el = document.getElementById('currentDate');
    if (el) el.textContent = new Date().toLocaleDateString('ko-KR', { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long' });
}

function updateMonthDisplay() {
    document.getElementById('monthDisplay').textContent = `${selectedYear}년 ${selectedMonth}월`;
}

function changeMonth(delta) {
    selectedMonth += delta;
    if (selectedMonth > 12) { selectedMonth = 1; selectedYear++; }
    if (selectedMonth < 1) { selectedMonth = 12; selectedYear--; }
    updateMonthDisplay();
    render();
}

function setPeriod(section, period) {
    if (section === 'high') highPeriod = period;
    else totalPeriod = period;
    const toggleId = section === 'high' ? 'highToggle' : 'totalToggle';
    document.querySelectorAll(`#${toggleId} .period-btn`).forEach(b => b.classList.remove('active'));
    document.querySelector(`#${toggleId} .period-btn[onclick="setPeriod('${section}','${period}')"]`).classList.add('active');
    render();
}

async function loadData() {
    document.getElementById('loading').style.display = 'block';
    document.getElementById('content').style.display = 'none';
    try {
        const url = `${PERFORMANCE_API_URL}?action=coverage`;
        console.log('커버리지 API 호출:', url);
        const resp = await fetch(url);
        const result = await resp.json();
        console.log('커버리지 API 응답:', result.success, '건수:', result.rows?.length);
        if (!result.success) throw new Error(result.error || 'API 실패');
        const cols = result.cols;
        allData = result.rows.map(r => {
            const obj = {};
            for (let i = 0; i < cols.length; i++) obj[cols[i]] = r[i];
            return obj;
        });
        render();
        document.getElementById('loading').style.display = 'none';
        document.getElementById('content').style.display = 'block';
    } catch (e) {
        console.error('커버리지 로드 실패:', e);
        document.getElementById('loading').textContent = '데이터 로드 실패: ' + e.message;
    }
}

function daysDiff(dateStr1, dateStr2) {
    if (!dateStr1 || !dateStr2) return Infinity;
    const d1 = new Date(dateStr1), d2 = new Date(dateStr2);
    if (isNaN(d1) || isNaN(d2)) return Infinity;
    return Math.floor((d2 - d1) / 86400000);
}

function getFirstContactDate(deal) {
    const dates = [];
    if (deal.has_payment_call === 'Y' && deal.payment_call_date) dates.push(deal.payment_call_date);
    if (deal.has_hard_collection === 'Y' && deal.hard_collection_date) dates.push(deal.hard_collection_date);
    if (dates.length === 0) return null;
    dates.sort();
    return dates[0];
}

function getMonthlyData() {
    return allData.filter(d => {
        if (!d.first_payment_notice) return false;
        const parts = d.first_payment_notice.split('-');
        return parseInt(parts[0]) === selectedYear && parseInt(parts[1]) === selectedMonth;
    });
}

function getYearlyData() {
    return allData.filter(d => {
        if (!d.first_payment_notice) return false;
        return parseInt(d.first_payment_notice.split('-')[0]) === 2026;
    });
}

function calcKPI(data, highOnly) {
    const target = highOnly
        ? data.filter(d => d.is_high_value === 'Y' && d.is_auto_paid === 'N')
        : data.filter(d => d.is_auto_paid === 'N');
    const totalCount = target.length;

    const contacted = target.filter(d => d.has_payment_call === 'Y' || d.has_hard_collection === 'Y');
    const contactedCount = contacted.length;
    const contactRate = totalCount > 0 ? (contactedCount / totalCount * 100) : 0;

    const contactDays = contacted.map(d => {
        const firstDate = getFirstContactDate(d);
        return firstDate ? daysDiff(d.first_payment_notice, firstDate) : null;
    }).filter(d => d !== null && d !== Infinity);

    const avgDays = contactDays.length > 0 ? contactDays.reduce((a, b) => a + b, 0) / contactDays.length : 0;
    const sorted = contactDays.slice().sort((a, b) => a - b);
    const medianDays = sorted.length > 0 ? sorted[Math.floor(sorted.length / 2)] : 0;

    const pool = highOnly ? data.filter(d => d.is_high_value === 'Y') : data;
    const autoCount = pool.filter(d => d.is_auto_paid === 'Y').length;
    const notContacted = totalCount - contactedCount;

    return { totalCount, contactedCount, contactRate, avgDays, medianDays, contactDays, autoCount, notContacted, target };
}

function render() {
    renderSection('high');
    renderSection('total');
    renderTable();
}

function renderSection(section) {
    const period = section === 'high' ? highPeriod : totalPeriod;
    const data = period === 'month' ? getMonthlyData() : getYearlyData();
    const highOnly = section === 'high';
    const kpi = calcKPI(data, highOnly);
    const targetId = section === 'high' ? 'highKPI' : 'totalKPI';
    const subjectLabel = highOnly ? '고액+비자동' : '비자동';

    const rateColor = kpi.contactRate >= 80 ? 'good' : kpi.contactRate >= 50 ? 'warn' : 'neutral';
    const daysColor = kpi.avgDays <= 5 ? 'good' : kpi.avgDays <= 10 ? 'warn' : 'neutral';
    const barColor = kpi.contactRate >= 80 ? 'green' : 'purple';

    document.getElementById(targetId).innerHTML = `
        <div class="kpi-grid">
            <div class="kpi-card primary">
                <div class="kpi-title">컨택률 (자동결제 제외)</div>
                <div class="kpi-value ${rateColor}">${kpi.contactRate.toFixed(1)}%</div>
                <div class="kpi-sub">${kpi.contactedCount} / ${kpi.totalCount}건</div>
            </div>
            <div class="kpi-card primary">
                <div class="kpi-title">평균 컨택 소요일</div>
                <div class="kpi-value ${daysColor}">${kpi.avgDays.toFixed(1)}일</div>
                <div class="kpi-sub">중앙값: ${kpi.medianDays}일</div>
            </div>
        </div>
        <div class="bar-wrap">
            <div class="bar-fill ${barColor}" style="width:${Math.min(kpi.contactRate, 100)}%">
                <span>${kpi.contactRate.toFixed(1)}%</span>
            </div>
        </div>
        <div class="bar-label">
            <span>컨택 ${kpi.contactedCount}건</span>
            <span>미컨택 ${kpi.notContacted}건 · 자동결제 ${kpi.autoCount}건 제외</span>
        </div>
        <div class="detail-row">
            <div class="detail-item">
                <div class="dt-label">대상 (${subjectLabel})</div>
                <div class="dt-value">${kpi.totalCount}건</div>
            </div>
            <div class="detail-item">
                <div class="dt-label">컨택 완료</div>
                <div class="dt-value">${kpi.contactedCount}건</div>
            </div>
            <div class="detail-item">
                <div class="dt-label">미컨택</div>
                <div class="dt-value">${kpi.notContacted}건</div>
            </div>
            <div class="detail-item">
                <div class="dt-label">소요일 분포</div>
                <div class="dt-value" style="font-size:13px;">
                    ${renderDaysDistribution(kpi.contactDays)}
                </div>
            </div>
        </div>
    `;
}

function renderDaysDistribution(days) {
    if (days.length === 0) return '-';
    const d3 = days.filter(d => d <= 3).length;
    const d7 = days.filter(d => d > 3 && d <= 7).length;
    const d14 = days.filter(d => d > 7 && d <= 14).length;
    const d30 = days.filter(d => d > 14 && d <= 30).length;
    const over = days.filter(d => d > 30).length;
    const parts = [];
    if (d3) parts.push(`~3일: ${d3}`);
    if (d7) parts.push(`~7일: ${d7}`);
    if (d14) parts.push(`~14일: ${d14}`);
    if (d30) parts.push(`~30일: ${d30}`);
    if (over) parts.push(`30일+: ${over}`);
    return parts.join('<br>');
}

function showTab(tab) {
    currentTab = tab;
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelector(`.tab-btn[onclick="showTab('${tab}')"]`).classList.add('active');
    renderTable();
}

function renderTable() {
    const monthData = getMonthlyData();
    const yearData = getYearlyData();
    const useMonth = highPeriod === 'month';
    const baseData = useMonth ? monthData : yearData;

    let filtered;
    switch (currentTab) {
        case 'high_noauto':
            filtered = baseData.filter(d => d.is_high_value === 'Y' && d.is_auto_paid === 'N');
            break;
        case 'high_all':
            filtered = baseData.filter(d => d.is_high_value === 'Y');
            break;
        case 'all_noauto':
            filtered = baseData.filter(d => d.is_auto_paid === 'N');
            break;
        case 'nocall':
            filtered = baseData.filter(d => d.is_auto_paid === 'N' && d.has_payment_call === 'N' && d.has_hard_collection === 'N');
            break;
        default:
            filtered = baseData;
    }

    document.getElementById('tableHead').innerHTML = `<tr>
        <th>거래명</th><th>가치</th><th>알림톡</th><th>결제일</th><th>결제일수</th>
        <th>자동</th><th>고액</th><th>☏결제</th><th>결제일자</th><th>결제담당</th>
        <th>강성</th><th>강성일자</th><th>강성담당</th><th>첫컨택</th><th>소요일</th>
    </tr>`;

    const badge = (v) => `<span class="badge ${v === 'Y' ? 'y' : 'n'}">${v}</span>`;

    document.getElementById('tableBody').innerHTML = filtered.map(d => {
        const firstContact = getFirstContactDate(d);
        const contactDays = firstContact ? daysDiff(d.first_payment_notice, firstContact) : null;
        return `<tr>
            <td><a href="https://raw-competition.pipedrive.com/deal/${d.deal_id}" target="_blank" style="color:#6366f1;">${d.title || d.deal_id}</a></td>
            <td style="text-align:right;">₩${formatNumber(d.value)}</td>
            <td>${d.first_payment_notice || ''}</td>
            <td>${d.won_time || '-'}</td>
            <td style="text-align:center;">${d.days_to_payment !== '' ? d.days_to_payment : '-'}</td>
            <td style="text-align:center;">${d.is_auto_paid === 'Y' ? '<span class="badge auto">자동</span>' : badge('N')}</td>
            <td style="text-align:center;">${badge(d.is_high_value)}</td>
            <td style="text-align:center;">${badge(d.has_payment_call)}</td>
            <td>${d.payment_call_date || ''}</td>
            <td>${d.payment_call_person || ''}</td>
            <td style="text-align:center;">${badge(d.has_hard_collection)}</td>
            <td>${d.hard_collection_date || ''}</td>
            <td>${d.hard_collection_person || ''}</td>
            <td>${firstContact || '-'}</td>
            <td style="text-align:center;font-weight:600;${contactDays !== null && contactDays <= 3 ? 'color:#16a34a' : contactDays !== null ? 'color:#d97706' : ''}">${contactDays !== null ? contactDays + '일' : '-'}</td>
        </tr>`;
    }).join('');

    document.getElementById('tableCount').textContent = `${filtered.length}건 표시`;
}

document.addEventListener('DOMContentLoaded', () => {
    displayCurrentDate();
    updateMonthDisplay();
    loadData();
});
