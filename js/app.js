(function () {
    'use strict';

    // ---------------------------------------------------------------
    // 기준 데이터 (계정과목/항목/부서/처리구분) - js/constants.js 에서 로드
    // ---------------------------------------------------------------
    var DATA = window.APP_CONSTANTS || {};
    var ACCOUNT_ITEMS = DATA.accountItems || {};
    var ACCOUNT_LIST = DATA.accountList || Object.keys(ACCOUNT_ITEMS);
    var DEPARTMENT_LIST = DATA.departmentList || [];
    var GUBUN_LIST = DATA.gubunList || [];
    var DEPT_SALES = DATA.deptSales || '영업지원팀';
    var GUBUN_CORP = DATA.gubunCorp || '법인';

    var DEBIT_LABEL_DEV = '개발투자비';
    var DEBIT_LABEL_SALES = '판매비와관리비';
    var CREDIT_ACCOUNT_LABEL = '미지급비용';
    var CREDIT_PARTY_CORP = '하나비자(카드뒷자리4개)';
    var CREDIT_PARTY_PERSONAL = '사원';

    var MAIN_HEADERS = ['NO.', '지출일', '계정과목', '항목', '상세내역', '금액'];

    // 항목별 금액 표에서 사용할, 모든 계정과목의 항목을 순서대로 펼친 목록
    var ALL_ITEMS = [];
    ACCOUNT_LIST.forEach(function (acc) {
        (ACCOUNT_ITEMS[acc] || []).forEach(function (item) {
            ALL_ITEMS.push(item);
        });
    });

    var rowSeq = 0;

    // ---------------------------------------------------------------
    // DOM 참조
    // ---------------------------------------------------------------
    var appShell = document.querySelector('.app-shell');
    var hamburgerBtn = document.getElementById('hamburgerBtn');

    var expenseTableBody = document.getElementById('expenseTableBody');
    var addRowBtn = document.getElementById('addRowBtn');
    var exportExcelBtn = document.getElementById('exportExcelBtn');
    var grandTotalEl = document.getElementById('grandTotal');

    var accountSummaryBody = document.getElementById('accountSummaryBody');
    var accountSummaryTotalEl = document.getElementById('accountSummaryTotal');
    var itemSummaryBody = document.getElementById('itemSummaryBody');
    var itemSummaryTotalEl = document.getElementById('itemSummaryTotal');

    var voucherGubunSelect = document.getElementById('voucherGubun');
    var voucherDeptSelect = document.getElementById('voucherDept');
    var voucherTableBody = document.getElementById('voucherTableBody');
    var voucherTableFoot = document.getElementById('voucherTableFoot');
    var voucherTotalKoreanEl = document.getElementById('voucherTotalKorean');

    // ---------------------------------------------------------------
    // 유틸
    // ---------------------------------------------------------------
    function formatKRW(amount) {
        var n = Number(amount) || 0;
        return '₩' + n.toLocaleString('ko-KR');
    }

    function numberToKorean(num) {
        var n = Math.floor(Number(num) || 0);
        if (n === 0) return '영';
        var digits = ['', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구'];
        var units = ['', '십', '백', '천'];
        var bigUnits = ['', '만', '억', '조'];
        function chunk4ToKorean(chunk) {
            var res = '';
            for (var i = 3; i >= 0; i--) {
                var pow = Math.pow(10, i);
                var d = Math.floor(chunk / pow);
                chunk = chunk % pow;
                if (d > 0) res += digits[d] + units[i];
            }
            return res;
        }
        var parts = [];
        var remaining = n;
        var bi = 0;
        while (remaining > 0) {
            parts.push({ val: remaining % 10000, unit: bigUnits[bi] });
            remaining = Math.floor(remaining / 10000);
            bi++;
        }
        var result = '';
        for (var i = parts.length - 1; i >= 0; i--) {
            if (parts[i].val > 0) result += chunk4ToKorean(parts[i].val) + parts[i].unit;
        }
        return result;
    }

    function escapeHtml(str) {
        if (str === null || str === undefined) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function buildOptions(list, selectedValue, placeholder) {
        var html = '';
        if (placeholder) {
            html += '<option value="">' + escapeHtml(placeholder) + '</option>';
        }
        list.forEach(function (value) {
            var selected = value === selectedValue ? ' selected' : '';
            html += '<option value="' + escapeHtml(value) + '"' + selected + '>' + escapeHtml(value) + '</option>';
        });
        return html;
    }

    // ---------------------------------------------------------------
    // 좌측 메뉴 (햄버거 토글)
    // ---------------------------------------------------------------
    hamburgerBtn.addEventListener('click', function () {
        var collapsed = appShell.classList.toggle('sidebar-collapsed');
        hamburgerBtn.setAttribute('aria-expanded', String(!collapsed));
    });

    // 메뉴 항목 클릭 -> 해당 페이지 표시 (현재는 지불승인서 1개뿐이지만 확장 가능한 구조)
    document.querySelectorAll('.menu-item').forEach(function (menuItem) {
        menuItem.addEventListener('click', function () {
            document.querySelectorAll('.menu-item').forEach(function (mi) {
                mi.classList.remove('active');
            });
            menuItem.classList.add('active');

            var targetId = menuItem.getAttribute('data-target');
            document.querySelectorAll('.menu-page').forEach(function (page) {
                page.classList.toggle('active', page.id === targetId);
            });
        });
    });

    // ---------------------------------------------------------------
    // 지출내역서 표 - 행 렌더링
    // ---------------------------------------------------------------
    function createRowElement() {
        rowSeq += 1;
        var today = new Date();
        var todayStr = today.getFullYear() + '-' +
            String(today.getMonth() + 1).padStart(2, '0') + '-' +
            String(today.getDate()).padStart(2, '0');
        var tr = document.createElement('tr');
        tr.dataset.rowId = String(rowSeq);
        tr.innerHTML =
            '<td class="no-cell"></td>' +
            '<td><input type="date" class="date-input" value="' + todayStr + '" /></td>' +
            '<td><select class="account-select">' + buildOptions(ACCOUNT_LIST, '', '선택') + '</select></td>' +
            '<td><select class="item-select" disabled><option value="">계정과목을 먼저 선택하세요</option></select></td>' +
            '<td><input type="text" class="detail-input" placeholder="상세내역을 입력하세요" /></td>' +
            '<td><input type="number" class="amount-input" min="0" step="1" placeholder="0" /></td>' +
            '<td><button type="button" class="btn-icon-danger row-delete-btn" title="이 행 삭제">✕</button></td>';
        return tr;
    }

    function addRow() {
        expenseTableBody.appendChild(createRowElement());
        renumberRows();
        recalculateAll();
    }

    function removeRow(tr) {
        if (expenseTableBody.rows.length <= 1) {
            alert('표에는 최소 1개의 행이 있어야 합니다.');
            return;
        }
        tr.parentNode.removeChild(tr);
        renumberRows();
        recalculateAll();
    }

    function renumberRows() {
        Array.prototype.forEach.call(expenseTableBody.rows, function (tr, idx) {
            tr.querySelector('.no-cell').textContent = String(idx + 1);
        });
    }

    function updateItemSelect(accountSelect, itemSelect, keepValue) {
        var account = accountSelect.value;
        var items = ACCOUNT_ITEMS[account] || [];
        if (!account) {
            itemSelect.innerHTML = '<option value="">계정과목을 먼저 선택하세요</option>';
            itemSelect.disabled = true;
            itemSelect.value = '';
            return;
        }
        var current = keepValue && items.indexOf(keepValue) !== -1 ? keepValue : '';
        itemSelect.innerHTML = buildOptions(items, current, '선택');
        itemSelect.disabled = false;
        itemSelect.value = current;
    }

    // ---------------------------------------------------------------
    // 지출내역서 표 이벤트 (이벤트 위임)
    // ---------------------------------------------------------------
    expenseTableBody.addEventListener('change', function (e) {
        var target = e.target;
        if (target.classList.contains('account-select')) {
            var tr = target.closest('tr');
            var itemSelect = tr.querySelector('.item-select');
            updateItemSelect(target, itemSelect, false);
            recalculateAll();
        } else if (target.classList.contains('item-select') ||
            target.classList.contains('date-input')) {
            recalculateAll();
        }
    });

    expenseTableBody.addEventListener('input', function (e) {
        var target = e.target;
        if (target.classList.contains('amount-input') || target.classList.contains('detail-input')) {
            recalculateAll();
        }
    });

    expenseTableBody.addEventListener('click', function (e) {
        var btn = e.target.closest('.row-delete-btn');
        if (btn) {
            removeRow(btn.closest('tr'));
        }
    });

    addRowBtn.addEventListener('click', addRow);

    // ---------------------------------------------------------------
    // 표 데이터 읽기
    // ---------------------------------------------------------------
    function getRowsData() {
        return Array.prototype.map.call(expenseTableBody.rows, function (tr) {
            return {
                date: tr.querySelector('.date-input').value,
                account: tr.querySelector('.account-select').value,
                item: tr.querySelector('.item-select').value,
                detail: tr.querySelector('.detail-input').value,
                amount: Number(tr.querySelector('.amount-input').value) || 0
            };
        });
    }

    // ---------------------------------------------------------------
    // 계정과목별 / 항목별 금액 재계산
    // ---------------------------------------------------------------
    function calcTotalsByKey(rows, key, keyList) {
        var totals = {};
        keyList.forEach(function (k) {
            totals[k] = 0;
        });
        rows.forEach(function (r) {
            if (r[key] && totals.hasOwnProperty(r[key])) {
                totals[r[key]] += r.amount;
            }
        });
        return totals;
    }

    function renderAccountSummary(rows) {
        var totals = calcTotalsByKey(rows, 'account', ACCOUNT_LIST);
        var html = '';
        var sum = 0;
        ACCOUNT_LIST.forEach(function (acc) {
            html += '<tr><td>' + escapeHtml(acc) + '</td><td>' + formatKRW(totals[acc]) + '</td></tr>';
            sum += totals[acc];
        });
        accountSummaryBody.innerHTML = html;
        accountSummaryTotalEl.textContent = formatKRW(sum);
        return totals;
    }

    function renderItemSummary(rows) {
        var totals = calcTotalsByKey(rows, 'item', ALL_ITEMS);
        var html = '';
        var sum = 0;
        ALL_ITEMS.forEach(function (item) {
            html += '<tr><td>' + escapeHtml(item) + '</td><td>' + formatKRW(totals[item]) + '</td></tr>';
            sum += totals[item];
        });
        itemSummaryBody.innerHTML = html;
        itemSummaryTotalEl.textContent = formatKRW(sum);
    }

    // ---------------------------------------------------------------
    // 회계전표 재계산
    // ---------------------------------------------------------------
    function renderVoucher(accountTotals, grandTotal) {
        var dept = voucherDeptSelect.value;
        var gubun = voucherGubunSelect.value;

        var debitAccountLabel = dept === DEPT_SALES ? DEBIT_LABEL_SALES : DEBIT_LABEL_DEV;
        var creditParty = gubun === GUBUN_CORP ? CREDIT_PARTY_CORP : CREDIT_PARTY_PERSONAL;

        var debitLines = ACCOUNT_LIST
            .map(function (acc) {
                return { account: acc, amount: accountTotals[acc] || 0 };
            })
            .filter(function (line) {
                return line.amount > 0;
            });

        var rowCount = Math.max(debitLines.length, 1);
        var bodyHtml = '';

        for (var i = 0; i < rowCount; i++) {
            var line = debitLines[i];
            bodyHtml += '<tr>';

            if (i === 0) {
                bodyHtml += '<td rowspan="' + rowCount + '">' + escapeHtml(debitAccountLabel) + '</td>';
            }

            if (line) {
                bodyHtml += '<td>' + escapeHtml(line.account) + '</td>';
                bodyHtml += '<td class="amount-cell">' + formatKRW(line.amount) + '</td>';
            } else {
                bodyHtml += '<td class="text-muted">-</td>';
                bodyHtml += '<td class="amount-cell">' + formatKRW(0) + '</td>';
            }

            if (i === 0) {
                bodyHtml += '<td rowspan="' + rowCount + '">' + escapeHtml(CREDIT_ACCOUNT_LABEL) + '</td>';
                bodyHtml += '<td rowspan="' + rowCount + '">' + escapeHtml(creditParty) + '</td>';
                bodyHtml += '<td rowspan="' + rowCount + '" class="amount-cell">' + formatKRW(grandTotal) + '</td>';
            }

            bodyHtml += '</tr>';
        }

        voucherTableBody.innerHTML = bodyHtml;
        voucherTableFoot.innerHTML =
            '<tr>' +
            '<td>합계</td><td></td><td class="amount-cell">' + formatKRW(grandTotal) + '</td>' +
            '<td>합계</td><td></td><td class="amount-cell">' + formatKRW(grandTotal) + '</td>' +
            '</tr>';

        if (voucherTotalKoreanEl) {
            voucherTotalKoreanEl.textContent = '금액(원) : ' + numberToKorean(grandTotal);
        }
    }

    // ---------------------------------------------------------------
    // 전체 재계산 (표 값이 바뀔 때마다 호출)
    // ---------------------------------------------------------------
    function recalculateAll() {
        var rows = getRowsData();
        var grandTotal = rows.reduce(function (sum, r) {
            return sum + r.amount;
        }, 0);
        grandTotalEl.textContent = formatKRW(grandTotal);

        var accountTotals = renderAccountSummary(rows);
        renderItemSummary(rows);
        renderVoucher(accountTotals, grandTotal);
    }

    voucherGubunSelect.addEventListener('change', recalculateAll);
    voucherDeptSelect.addEventListener('change', recalculateAll);

    // ---------------------------------------------------------------
    // 엑셀 다운로드 (브라우저에서 SheetJS 로 직접 생성 - 서버 불필요)
    // ---------------------------------------------------------------
    function buildExportSheetData(rows) {
        var matrix = [];
        var merges = [];
        var currencyCells = [];

        function setCell(r, c, v) {
            if (!matrix[r]) matrix[r] = ['', '', '', '', '', ''];
            matrix[r][c] = v;
        }

        function setCurrency(r, c, v) {
            setCell(r, c, v);
            currencyCells.push({ r: r, c: c });
        }

        // 제목
        setCell(0, 0, '지출내역서');
        merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } });

        // 지출내역서 표
        var headerRow = 2;
        MAIN_HEADERS.forEach(function (h, i) {
            setCell(headerRow, i, h);
        });

        var r = headerRow + 1;
        var total = 0;
        rows.forEach(function (row, idx) {
            setCell(r, 0, idx + 1);
            setCell(r, 1, row.date || '');
            setCell(r, 2, row.account || '');
            setCell(r, 3, row.item || '');
            setCell(r, 4, row.detail || '');
            setCurrency(r, 5, row.amount || 0);
            total += row.amount || 0;
            r++;
        });

        var mainTotalRow = r;
        setCell(mainTotalRow, 0, '합계');
        merges.push({ s: { r: mainTotalRow, c: 0 }, e: { r: mainTotalRow, c: 4 } });
        setCurrency(mainTotalRow, 5, total);

        return { matrix: matrix, merges: merges, currencyCells: currencyCells };
    }

    function exportToExcel() {
        if (typeof XLSX === 'undefined') {
            alert('엑셀 생성 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인한 뒤 다시 시도해주세요.');
            return;
        }

        var rows = getRowsData();

        var built = buildExportSheetData(rows);
        var ws = XLSX.utils.aoa_to_sheet(built.matrix);
        ws['!merges'] = built.merges;
        ws['!cols'] = [{ wch: 6 }, { wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 28 }, { wch: 14 }];

        var currencyFormat = '"₩"#,##0';
        built.currencyCells.forEach(function (cell) {
            var addr = XLSX.utils.encode_cell({ r: cell.r, c: cell.c });
            if (ws[addr]) {
                ws[addr].z = currencyFormat;
            }
        });

        var wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '지출내역서');

        var today = new Date();
        var y = today.getFullYear();
        var m = String(today.getMonth() + 1).padStart(2, '0');
        var d = String(today.getDate()).padStart(2, '0');
        XLSX.writeFile(wb, '지불승인서_' + y + m + d + '.xlsx');
    }

    exportExcelBtn.addEventListener('click', exportToExcel);

    // ---------------------------------------------------------------
    // 계정과목 참고 페이지
    // ---------------------------------------------------------------
    var ACCOUNT_CODE_DETAILS = DATA.accountCodeDetails || [];

    function initAccountCodePage() {
        var tbody = document.getElementById('accountCodeTableBody');
        if (!tbody || tbody.dataset.initialized) return;
        tbody.dataset.initialized = '1';

        var html = '';
        ACCOUNT_CODE_DETAILS.forEach(function (group, idx) {
            var itemCount = group.items.length;
            html += '<tr class="account-code-row" data-idx="' + idx + '">';
            html += '<td class="acc-name-cell">' + escapeHtml(group.account) + '</td>';
            html += '<td class="acc-count-cell">' + itemCount + '개</td>';
            html += '<td class="acc-toggle-cell">▼</td>';
            html += '</tr>';
            html += '<tr class="account-code-detail" data-idx="' + idx + '" style="display:none;">';
            html += '<td colspan="3"><table class="account-sub-table">';
            html += '<thead><tr><th>항목</th><th>상세내역</th></tr></thead><tbody>';
            group.items.forEach(function (d) {
                html += '<tr><td>' + escapeHtml(d.item) + '</td><td>' + escapeHtml(d.desc) + '</td></tr>';
            });
            html += '</tbody></table></td></tr>';
        });
        tbody.innerHTML = html;

        tbody.addEventListener('click', function (e) {
            var row = e.target.closest('.account-code-row');
            if (!row) return;
            var idx = row.getAttribute('data-idx');
            var detailRow = tbody.querySelector('.account-code-detail[data-idx="' + idx + '"]');
            var isOpen = detailRow.style.display !== 'none';
            detailRow.style.display = isOpen ? 'none' : '';
            row.querySelector('.acc-toggle-cell').textContent = isOpen ? '▼' : '▲';
            row.classList.toggle('open', !isOpen);
        });
    }

    // ---------------------------------------------------------------
    // 초기화
    // ---------------------------------------------------------------
    function init() {
        voucherGubunSelect.innerHTML = buildOptions(GUBUN_LIST, GUBUN_CORP, null);
        voucherDeptSelect.innerHTML = buildOptions(DEPARTMENT_LIST, '개발팀', null);

        addRow(); // 최소 1개 행으로 시작
        initAccountCodePage();
    }

    init();
})();
