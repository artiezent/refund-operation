// collection.js - 추심 대시보드

let collectionDataCache = null;
let selectedCollectionMonth = new Date();
let selectedCollectionYearVal = new Date().getFullYear();

const LOCAL_CACHE_KEY = 'kpi_collection_data';
const LOCAL_CACHE_TS_KEY = 'kpi_collection_data_ts';
const LOCAL_CACHE_TTL = 10 * 60 * 1000;

// ==========================================
// 로컬 캐시
// ==========================================

function saveToLocalCache(rawData) {
    try {
        localStorage.setItem(LOCAL_CACHE_KEY, JSON.stringify(rawData));
        localStorage.setItem(LOCAL_CACHE_TS_KEY, Date.now().toString());
    } catch (e) { /* quota exceeded */ }
}

function loadFromLocalCache() {
    try {
        const ts = localStorage.getItem(LOCAL_CACHE_TS_KEY);
        const raw = localStorage.getItem(LOCAL_CACHE_KEY);
        if (!ts || !raw) return null;
        return { data: JSON.parse(raw), age: Date.now() - parseInt(ts) };
    } catch (e) { return null; }
}

// ==========================================
// 데이터 로드
// ==========================================

async function fetchCollectionFromNetwork() {
    const url = `${PERFORMANCE_API_URL}?action=payment&scope=collection`;
    const response = await fetch(url);
    const result = await response.json();
    if (!result.success) throw new Error('API 실패');
    const data = convertRowsToObjects(result);
    saveToLocalCache(data);
    return preprocessData(data);
}

let paymentRefundCache = null;
let paymentRefundYear = null;

async function fetchPaymentRefundData() {
    const year = selectedCollectionYearVal;
    if (paymentRefundCache && paymentRefundYear === year) return paymentRefundCache;

    const [thisYearData, prevYearData] = await Promise.all([
        fetchPaymentData('yearly', year),
        fetchPaymentData('yearly', year - 1)
    ]);
    paymentRefundCache = [...prevYearData, ...thisYearData];
    paymentRefundYear = year;
    return paymentRefundCache;
}

async function loadCollectionData() {
    const loadingEl = document.getElementById('collection-loading');
    const errorEl = document.getElementById('collection-error');
    const contentEl = document.getElementById('collection-content');

    if (loadingEl) loadingEl.style.display = 'flex';
    if (errorEl) errorEl.style.display = 'none';
    if (contentEl) contentEl.style.display = 'none';

    try {
        const cached = loadFromLocalCache();
        if (cached && cached.age < LOCAL_CACHE_TTL) {
            collectionDataCache = preprocessData(cached.data);
            showCollectionDashboard();
            fetchCollectionFromNetwork().then(freshData => {
                if (freshData) {
                    collectionDataCache = freshData;
                    showCollectionDashboard();
                }
            }).catch(() => {});
            return;
        }
        if (cached) {
            collectionDataCache = preprocessData(cached.data);
            showCollectionDashboard();
            fetchCollectionFromNetwork().then(freshData => {
                if (freshData) {
                    collectionDataCache = freshData;
                    showCollectionDashboard();
                }
            }).catch(() => {});
            return;
        }
        collectionDataCache = await fetchCollectionFromNetwork();
        showCollectionDashboard();
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

async function showCollectionDashboard() {
    const paymentData = await fetchPaymentRefundData().catch(() => null);
    calculateAndDisplayCollectionKPIs(collectionDataCache, paymentData);
    calculateYearlyCollectionTracking(collectionDataCache, paymentData);

    const syncTimeEl = document.getElementById('collection-last-sync-time');
    if (syncTimeEl) syncTimeEl.textContent = new Date().toLocaleString('ko-KR');

    const loadingEl = document.getElementById('collection-loading');
    const contentEl = document.getElementById('collection-content');
    if (loadingEl) loadingEl.style.display = 'none';
    if (contentEl) contentEl.style.display = 'block';
}

function refreshAllData() {
    collectionDataCache = null;
    localStorage.removeItem(LOCAL_CACHE_KEY);
    localStorage.removeItem(LOCAL_CACHE_TS_KEY);
    loadCollectionData();
}

// ==========================================
// 월/연 네비게이션
// ==========================================

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
        calculateAndDisplayCollectionKPIs(collectionDataCache, paymentRefundCache);
    }
}

async function changeCollectionYear(delta) {
    selectedCollectionYearVal += delta;
    paymentRefundCache = null;
    updateCollectionYearDisplay();
    if (collectionDataCache) {
        const paymentData = await fetchPaymentRefundData().catch(() => null);
        calculateYearlyCollectionTracking(collectionDataCache, paymentData);
    }
}

function updateCollectionYearDisplay() {
    const el = document.getElementById('selected-collection-year');
    if (el) el.textContent = `${selectedCollectionYearVal}년`;
}

// ==========================================
// 추심 KPI 계산
// ==========================================

function calculateAndDisplayCollectionKPIs(data, paymentData) {
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
    console.log(`환급완료 기준: ${prevYear}년 ${prevMonth + 1}월 first_payment_notice (결제현황)`);
    console.log(`이관총액 기준: ${year}년 ${month + 1}월 collection_order_date`);

    const refundSource = paymentData || data;
    const refundDeals = refundSource.filter(d => d._noticeDate && d._noticeDate.getFullYear() === prevYear && d._noticeDate.getMonth() === prevMonth);
    const refundCount = refundDeals.length;
    const refundAmount = refundDeals.reduce((s, d) => s + d._value, 0);

    const transferDeals = data.filter(d => d._collectionDate && d._collectionDate.getFullYear() === year && d._collectionDate.getMonth() === month);
    const transferCount = transferDeals.length;
    const transferAmount = transferDeals.reduce((s, d) => s + d._value, 0);

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

// ==========================================
// 연간 이관 결제 추적
// ==========================================

let ytBucketStore = {};

function calculateYearlyCollectionTracking(data, paymentData) {
    const year = selectedCollectionYearVal;
    const tbody = document.getElementById('yearly-collection-body');
    if (!tbody) return;

    const MONTH_BUCKETS = 12;
    ytBucketStore = {};
    let rows = '';
    const refundSource = paymentData || data;

    for (let m = 0; m < 12; m++) {
        let prevYear = year;
        let prevMonth = m - 1;
        if (prevMonth < 0) { prevMonth = 11; prevYear--; }
        const refundAmount = refundSource.filter(d => d._noticeDate && d._noticeDate.getFullYear() === prevYear && d._noticeDate.getMonth() === prevMonth).reduce((s, d) => s + d._value, 0);

        const transferDeals = data.filter(d => d._collectionDate && d._collectionDate.getFullYear() === year && d._collectionDate.getMonth() === m);
        const totalCount = transferDeals.length;
        const totalAmount = transferDeals.reduce((s, d) => s + d._value, 0);

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
            if (!deal._hasWon) {
                buckets.unpaid.push(deal);
                return;
            }
            if (!deal._collectionDate || !deal._wonDate) { buckets.unpaid.push(deal); return; }
            const md = (deal._wonDate.getFullYear() - deal._collectionDate.getFullYear()) * 12 + (deal._wonDate.getMonth() - deal._collectionDate.getMonth());
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
            const amt = bucket.reduce((s, d) => s + d._value, 0);
            const amtEok = amt / 100000000;
            const amtFull = '₩' + amt.toLocaleString();
            const clickAttr = cnt > 0 ? ` data-click="${clickKey}"` : '';

            if (cnt === 0) {
                return `<td class="yt-cell-rate" data-amt="0" data-cnt="0"><div class="yt-rate-main" style="color:var(--text-secondary);font-size:11px;opacity:0.4;">-</div></td>`;
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

            return `<td class="yt-cell-rate ${colorClass}" data-tip="${amtFull}" data-amt="${amt}" data-cnt="${cnt}"${clickAttr}>
                <div class="yt-rate-main">${amtDisplay}</div>
                <div class="yt-rate-detail">${cnt}건</div>
            </td>`;
        }

        let monthCells = '';
        for (let i = 0; i < MONTH_BUCKETS; i++) {
            monthCells += cellHtml(buckets[i], false, `${m}-${i}`);
        }

        const paidDeals = transferDeals.filter(d => d._hasWon);
        const paidAmount = paidDeals.reduce((s, d) => s + d._value, 0);
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
            <td class="yt-cell-num yt-sticky-transfer" style="left:48px;" data-tip="₩${totalAmount.toLocaleString()}" data-amt="${totalAmount}" data-cnt="${totalCount}"${transferClickAttr}><div class="yt-rate-main" style="font-size:12px;white-space:nowrap;">${totalEok}억 (${refundAmount > 0 ? (totalAmount / refundAmount * 100).toFixed(0) : 0}%)</div><div class="yt-rate-detail">${totalCount}건</div></td>
            <td class="yt-cell-num yt-sticky-paid ${paidColorCls}" style="left:148px;font-weight:700;" data-tip="₩${paidAmount.toLocaleString()}" data-amt="${paidAmount}" data-cnt="${paidDeals.length}"${paidClickAttr}><div class="yt-rate-main" style="font-size:12px;white-space:nowrap;">${paidAmount > 0 ? paidDisplay : '-'}</div><div class="yt-rate-detail">${paidDeals.length}건 (${paidRate}%)</div></td>
            ${monthCells}
            ${cellHtml(buckets.over, false, `${m}-over`)}
            ${cellHtml(buckets.unpaid, true, `${m}-unpaid`)}
        </tr>`;
    }

    tbody.innerHTML = rows;

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
                deals = store.transfer.filter(d => d._hasWon);
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

            tbody.querySelectorAll('.yt-cell-active').forEach(el => el.classList.remove('yt-cell-active'));
            cell.classList.add('yt-cell-active');

            showYtDetail(deals, label);
        });
    });

    initYtDragSelect(tbody);
}

// ==========================================
// 드래그 선택 합계
// ==========================================

function initYtDragSelect(tbody) {
    let isDragging = false;
    let selectedCells = [];
    const table = tbody.closest('table');

    function getCellPos(td) {
        const row = td.parentElement;
        return { row: row.rowIndex, col: td.cellIndex };
    }

    function getCellsInRange(start, end) {
        const minR = Math.min(start.row, end.row);
        const maxR = Math.max(start.row, end.row);
        const minC = Math.min(start.col, end.col);
        const maxC = Math.max(start.col, end.col);
        const cells = [];
        const rows = table.rows;
        for (let r = minR; r <= maxR; r++) {
            for (let c = minC; c <= maxC; c++) {
                const cell = rows[r]?.cells[c];
                if (cell && cell.dataset.amt !== undefined) cells.push(cell);
            }
        }
        return cells;
    }

    function showDragTooltip(cells, e) {
        let totalAmt = 0, totalCnt = 0;
        cells.forEach(c => {
            totalAmt += parseInt(c.dataset.amt) || 0;
            totalCnt += parseInt(c.dataset.cnt) || 0;
        });
        if (cells.length <= 1) { hideDragTooltip(); return; }

        let tip = document.getElementById('yt-drag-tooltip');
        if (!tip) {
            tip = document.createElement('div');
            tip.id = 'yt-drag-tooltip';
            tip.style.cssText = 'position:fixed;z-index:9999;padding:8px 14px;background:#1e1e2e;border:1px solid var(--accent-primary);border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,0.5);pointer-events:none;font-size:13px;line-height:1.6;white-space:nowrap;opacity:1;';
            document.body.appendChild(tip);
        }
        const eok = totalAmt / 100000000;
        const amtStr = eok >= 1 ? `₩${eok.toFixed(2)}억` : `₩${totalAmt.toLocaleString()}`;
        tip.innerHTML = `<div style="font-weight:700;color:var(--accent-primary);">${amtStr}</div><div style="color:var(--text-secondary);">${totalCnt}건 · ${cells.length}셀</div>`;
        tip.style.display = 'block';
        tip.style.left = (e.clientX + 16) + 'px';
        tip.style.top = (e.clientY - 10) + 'px';
    }

    function hideDragTooltip() {
        const tip = document.getElementById('yt-drag-tooltip');
        if (tip) tip.style.display = 'none';
    }

    function clearSelection() {
        selectedCells.forEach(c => c.classList.remove('yt-drag-selected'));
        selectedCells = [];
    }

    let startPos = null;

    tbody.addEventListener('mousedown', (e) => {
        const td = e.target.closest('td[data-amt]');
        if (!td) return;
        e.preventDefault();
        isDragging = true;
        startPos = getCellPos(td);
        clearSelection();
        selectedCells = [td];
        td.classList.add('yt-drag-selected');
    });

    tbody.addEventListener('mousemove', (e) => {
        if (!isDragging || !startPos) return;
        const td = e.target.closest('td[data-amt]');
        if (!td) return;
        const endPos = getCellPos(td);
        clearSelection();
        selectedCells = getCellsInRange(startPos, endPos);
        selectedCells.forEach(c => c.classList.add('yt-drag-selected'));
        showDragTooltip(selectedCells, e);
    });

    document.addEventListener('mouseup', () => {
        if (isDragging) {
            isDragging = false;
            startPos = null;
            setTimeout(() => {
                clearSelection();
                hideDragTooltip();
            }, 2000);
        }
    });
}

// ==========================================
// 상세 패널
// ==========================================

let ytDetailCurrentDeals = [];
let ytDetailCurrentLabel = '';
let ytDetailSortedByBalance = false;

function showYtDetail(deals, label) {
    const panel = document.getElementById('yt-detail-panel');
    const titleEl = document.getElementById('yt-detail-title');
    if (!panel || !titleEl) return;

    ytDetailCurrentDeals = [...deals].sort((a, b) => b._value - a._value);
    ytDetailCurrentLabel = label;
    ytDetailSortedByBalance = false;

    const headerEl = document.getElementById('yt-balance-header');
    if (headerEl) headerEl.textContent = '잔액 ▼';

    renderYtDetailRows(ytDetailCurrentDeals, label);
    panel.style.display = 'block';
    panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function renderYtDetailRows(deals, label) {
    const titleEl = document.getElementById('yt-detail-title');
    const tbodyEl = document.getElementById('yt-detail-tbody');
    if (!titleEl || !tbodyEl) return;

    const totalAmt = deals.reduce((s, d) => s + d._value, 0);
    titleEl.textContent = `${label} — ${deals.length}건, ₩${totalAmt.toLocaleString()}`;

    tbodyEl.innerHTML = deals.map((d, i) => {
        const val = d._value;
        const bal = Number(d.balance) || 0;
        const orderDate = d.collection_order_date || '-';
        const wonDate = d.won_time || '미결제';
        const stage = d.stage_name || '-';
        const dealUrl = d.deal_id ? `https://raw-competition.pipedrive.com/deal/${d.deal_id}` : '';
        const titleHtml = dealUrl
            ? `<a href="${dealUrl}" target="_blank" rel="noopener" style="color:var(--accent-primary);text-decoration:none;">${d.title || '-'}</a>`
            : (d.title || '-');
        return `<tr>
            <td>${i + 1}</td>
            <td>${titleHtml}</td>
            <td>${stage}</td>
            <td class="td-amount">₩${val.toLocaleString()}</td>
            <td class="td-amount">₩${bal.toLocaleString()}</td>
            <td>${orderDate}</td>
            <td>${wonDate}</td>
        </tr>`;
    }).join('');
}

function sortYtDetailByBalance() {
    if (!ytDetailCurrentDeals.length) return;

    ytDetailSortedByBalance = !ytDetailSortedByBalance;

    const headerEl = document.getElementById('yt-balance-header');

    let sorted;
    if (ytDetailSortedByBalance) {
        sorted = [...ytDetailCurrentDeals].sort((a, b) => (Number(b.balance) || 0) - (Number(a.balance) || 0));
        if (headerEl) headerEl.textContent = '잔액 ▲';
    } else {
        sorted = [...ytDetailCurrentDeals].sort((a, b) => b._value - a._value);
        if (headerEl) headerEl.textContent = '잔액 ▼';
    }

    renderYtDetailRows(sorted, ytDetailCurrentLabel);
}

function closeYtDetail() {
    const panel = document.getElementById('yt-detail-panel');
    if (panel) panel.style.display = 'none';
    const tbody = document.getElementById('yearly-collection-body');
    if (tbody) tbody.querySelectorAll('.yt-cell-active').forEach(el => el.classList.remove('yt-cell-active'));
}

// ==========================================
// 툴팁
// ==========================================

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

// ==========================================
// 초기화
// ==========================================

function displayCurrentDate() {
    const now = new Date();
    const options = { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long' };
    const dateStr = now.toLocaleDateString('ko-KR', options);
    const el = document.getElementById('currentDate');
    if (el) el.textContent = dateStr;
}

function init() {
    displayCurrentDate();
    updateCollectionMonthDisplay();
    updateCollectionYearDisplay();
    loadCollectionData();
}

document.addEventListener('DOMContentLoaded', init);
