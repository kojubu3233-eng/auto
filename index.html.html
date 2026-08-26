<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>스마트 발주 시스템 Pro v10.0 (홈 대시보드)</title>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        * { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; box-sizing: border-box; margin: 0; padding: 0; }
        body { background-color: #f8fafc; color: #1e293b; padding: 15px; font-size: 13px; }
        
        .panel { background: #fff; padding: 15px; border-radius: 10px; margin-bottom: 15px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); border: 1px solid #e2e8f0; }
        .header { display: flex; justify-content: space-between; align-items: center; background: linear-gradient(135deg, #1e3a8a, #3b82f6); color: white; flex-wrap: wrap; gap: 10px; }
        
        .dashboard-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; margin-bottom: 15px; }
        .card { background: white; padding: 15px; border-radius: 10px; border: 1px solid #e2e8f0; border-left: 5px solid #cbd5e1; }
        .card.blue { border-left-color: #3b82f6; }
        .card.green { border-left-color: #10b981; }
        .card.orange { border-left-color: #f97316; }
        .card h3 { font-size: 12px; color: #64748b; margin-bottom: 5px; }
        .card .value { font-size: 22px; font-weight: bold; }

        .main-container { display: grid; grid-template-columns: 1fr; gap: 15px; }
        @media (min-width: 1200px) { .main-container { grid-template-columns: 1.4fr 0.6fr; } }

        .btn { padding: 8px 12px; border: none; border-radius: 6px; font-weight: 600; cursor: pointer; color: white; display: inline-block; text-align: center; }
        .btn-primary { background: #2563eb; } .btn-success { background: #10b981; } .btn-danger { background: #ef4444; } .btn-gray { background: #64748b; }
        
        table { width: 100%; border-collapse: collapse; text-align: center; font-size: 12px; }
        th { background: #f1f5f9; padding: 8px 4px; font-weight: 600; border: 1px solid #cbd5e1; position: sticky; top: 0; }
        td { padding: 4px; border: 1px solid #cbd5e1; background: #fff; vertical-align: middle; }
        input[type="text"], input[type="number"], input[type="date"], input[type="month"], select { border: 1px solid #cbd5e1; padding: 6px; border-radius: 4px; outline: none; }
        input:focus, select:focus { border-color: #2563eb; background: #eff6ff; }
        
        .excel-table td { padding: 2px; }
        .excel-table input, .excel-table select { width: 100%; border: none; border-radius: 0; background: transparent; }
        
        .res-min { background: #f0f9ff !important; color: #0284c7; font-weight: bold; }
        .res-rec { background: #f0fdf4 !important; color: #16a34a; font-weight: bold; font-size: 14px; }
        .res-max { background: #fff7ed !important; color: #ea580c; font-weight: bold; }

        .table-wrapper { max-height: 400px; overflow-y: auto; border: 1px solid #e2e8f0; border-radius: 6px; }
        .chart-wrapper { height: 200px; }
        
        .controls-row { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; flex-wrap: wrap; gap: 10px; background: #f8fafc; padding: 10px; border-radius: 8px; border: 1px solid #e2e8f0; }
        .slider-container { display: flex; align-items: center; gap: 10px; flex: 1; min-width: 200px; }
        input[type="range"] { flex: 1; cursor: pointer; }
        .radio-group { display: flex; gap: 15px; }
        .radio-group label { cursor: pointer; font-weight: 600; display: flex; align-items: center; gap: 4px; color: #334155; }
        
        .hide-col { display: none !important; }
        .volatility-list { display: grid; grid-template-columns: repeat(auto-fill, minmax(130px, 1fr)); gap: 8px; margin-top: 10px; }
        .vol-item { background: #f8fafc; border: 1px solid #e2e8f0; padding: 8px; border-radius: 6px; text-align: center; }
        .vol-item .cat-name { font-weight: 600; font-size: 11px; color: #475569; margin-bottom: 4px; }
        .vol-item .std-val { font-size: 15px; font-weight: bold; }
        .std-safe { color: #10b981; } .std-warn { color: #f59e0b; } .std-danger { color: #ef4444; }

        .rec-badge { background: #dcfce7; color: #166534; padding: 4px 8px; border-radius: 12px; font-weight: bold; font-size: 12px; margin-left: 10px; border: 1px solid #bbf7d0; cursor: pointer;}
        .rec-badge:hover { background: #bbf7d0; }
        .day-badge { padding: 2px 6px; border-radius: 4px; font-size: 11px; font-weight: bold; margin-left: 2px; }
        .day-work { background: #eff6ff; color: #2563eb; border: 1px solid #bfdbfe; }
        .day-holi { background: #fef2f2; color: #ef4444; border: 1px solid #fecaca; }
        .day-sat { background: #ecfeff; color: #0e7490; border: 1px solid #a5f3fc; }

        .tab-bar { display: flex; gap: 8px; margin-bottom: 15px; flex-wrap: wrap; }
        .tab-btn { padding: 10px 18px; border: 1px solid #cbd5e1; background: #fff; color: #475569; font-weight: 600; border-radius: 8px 8px 0 0; cursor: pointer; font-size: 13px; }
        .tab-btn.active { background: #1e3a8a; color: #fff; border-color: #1e3a8a; }

        .ac-wrap { position: relative; }
        .ac-list { position: absolute; z-index: 50; top: 100%; left: 0; right: 0; background: #fff; border: 1px solid #cbd5e1; border-top: none; max-height: 180px; overflow-y: auto; display: none; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .ac-list div { padding: 6px 8px; font-size: 12px; text-align: left; cursor: pointer; }
        .ac-list div:hover { background: #eff6ff; }

        /* [신규] 홈 대시보드 - 당월 납품계획 표를 모바일에서 카드형으로 전환 */
        .responsive-table td { text-align: center; }
        @media (max-width: 720px) {
            body { padding: 8px; font-size: 12px; }
            .panel { padding: 10px; }
            table.responsive-table thead { display: none; }
            table.responsive-table, table.responsive-table tbody, table.responsive-table tr, table.responsive-table td { display: block; width: 100%; }
            table.responsive-table tr { margin-bottom: 8px; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden; }
            table.responsive-table td { text-align: right; padding: 6px 10px; border: none; border-bottom: 1px solid #f1f5f9; position: relative; padding-left: 42%; min-height: 16px; }
            table.responsive-table td:last-child { border-bottom: none; }
            table.responsive-table td::before { content: attr(data-label); position: absolute; left: 10px; text-align: left; font-weight: 700; color: #64748b; }
            .controls-row { flex-direction: column; align-items: stretch; }
            .dashboard-grid { grid-template-columns: 1fr 1fr; }
        }
        .mobile-cards { display: flex; flex-direction: column; gap: 8px; }
        .mobile-card { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px 12px; display: flex; justify-content: space-between; align-items: center; gap: 10px; }
        .mobile-card .mc-name { font-weight: 700; font-size: 13px; }
        .mobile-card .mc-sub { font-size: 11px; color: #64748b; margin-top: 2px; }
        .mobile-card .mc-right { text-align: right; white-space: nowrap; }
        .mobile-card.upcoming { border-left: 4px solid #ea580c; }
        .mobile-card.past { border-left: 4px solid #dc2626; }

        .grid-select-col { width: 26px; }
        select#gridYear, select#gridMonth { font-weight: bold; }
    </style>
</head>
<body>

    <div class="panel header">
        <h1 style="font-size: 18px;">🗄️ 스마트 발주 시스템 Pro v10.0</h1>
        <div style="display:flex; gap:8px; flex-wrap:wrap;">
            <button class="btn btn-gray" onclick="renormalizeExistingUnits()" title="기존에 누적된 데이터에도 부대명 정규화(공백/슬래시 통일 등)를 소급 적용합니다">🔧 부대명 재정규화</button>
            <button class="btn btn-gray" style="background:#7c3aed;" onclick="reclassifyExistingHolidays()" title="공공데이터포털 특일정보 API로 조회한 실제 공휴일을 기존 누적 데이터에도 소급 적용합니다. 평일(토요일 포함)/일요일/공휴일 3가지로 정확히 재분류됩니다.">🗓️ 휴일 재분류(공휴일 API)</button>
            <button class="btn btn-danger" onclick="clearData()">전체 데이터 초기화</button>
        </div>
    </div>

    <details class="panel" style="margin-bottom:10px; padding:10px 15px; background:#f8fafc;">
        <summary style="cursor:pointer; font-weight:bold; color:#334155; font-size:13px;">⚙️ 구글시트 연동 설정 <span id="gsheetStatus" style="font-size:11px; font-weight:normal; color:#64748b; margin-left:6px;">미설정</span></summary>
        <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap; margin-top:10px;">
            <label style="font-weight:bold; color:#166534; font-size:12px; white-space:nowrap;">🔗 웹앱 URL:</label>
            <input type="text" id="gsheetWebAppUrl" placeholder="https://script.google.com/macros/s/.../exec" style="flex:1; min-width:280px; font-size:12px;" onchange="saveGSheetUrl()">
        </div>
        <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap; margin-top:8px;">
            <label style="font-weight:bold; color:#7c3aed; font-size:12px; white-space:nowrap;" title="비워두면 전체 탭에서 자동 검색">📦 재고수불대장 시트명(선택):</label>
            <input type="text" id="gsheetLedgerSheet" placeholder="비워두면 전체 탭 자동 검색" value="" style="width:260px; font-size:12px;" onchange="saveGSheetSheetNames()">
        </div>
        <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap; margin-top:8px;">
            <button class="btn btn-gray" style="background:#0369a1; font-size:12px; padding:6px 10px;" onclick="syncUnitProductInfo()">🔄 부대·품목 정보 동기화</button>
            <span id="unitProductSyncStatus" style="font-size:11px; color:#64748b;">미동기화</span>
        </div>
        <div style="font-size:11px; color:#64748b; margin-top:8px;">🚫 이상치 제외: "데이터" 시트 A열(고유ID)에 <b>제외</b>라고 입력해두면 그 행은 통계에서 자동 제외됩니다. (또는 아래 원본 로그에서 직접 체크)</div>
    </details>

    <!-- [신규 v10] 탭 3개: 홈(기본) / 예정량·확정량 입력 / 재고 상세 -->
    <div class="tab-bar">
        <button class="tab-btn active" id="tabBtnHome" onclick="switchTab('home')">🏠 홈</button>
        <button class="tab-btn" id="tabBtnPredict" onclick="switchTab('predict')">🔮 예정량/확정량 입력</button>
        <button class="tab-btn" id="tabBtnInventory" onclick="switchTab('inventory')">📦 재고 상세</button>
    </div>

    <!-- ============================================================ -->
    <!-- [신규] 홈 탭: 접속시 가장 먼저 보이는 화면.
         1) 당월 납품계획(구글시트 '데이터' 탭 기준, 엑셀 출고계획 형식으로 표시)
         2) 재고수불대장의 납품기한임박/납품기한경과 제품
         모바일에서는 표가 카드형으로 자동 전환된다. -->
    <!-- ============================================================ -->
    <div id="tabContentHome">
        <div class="panel">
            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px; margin-bottom:10px;">
                <h2 style="font-size:15px; color:#1e3a8a;">📋 당월 납품계획</h2>
                <div style="display:flex; gap:8px;">
                    <button class="btn btn-success" onclick="saveHomeConfirmedQuantities()" title="아래 표의 '확정량' 칸에 입력/수정한 값을 구글시트 '데이터' 탭에 반영합니다">💾 확정량 저장</button>
                    <button class="btn btn-primary" onclick="refreshHomeTab()">🔄 새로고침</button>
                </div>
            </div>
            <div id="homeStatusMsg" style="font-size:12px; color:#64748b; margin-bottom:8px;">새로고침을 눌러 구글시트 데이터를 불러오세요.</div>
            <div style="font-size:11px; color:#a16207; background:#fef9c3; border:1px solid #fde68a; border-radius:6px; padding:6px 10px; margin-bottom:8px;">✏️ '확정량' 칸에 직접 입력 후 우측 상단 "💾 확정량 저장"을 누르면 구글시트에 반영됩니다. 값은 <b>kg 기준</b>으로 입력해주세요(품목정보에 포장단위가 등록돼 있으면 개수도 자동 환산되어 함께 저장됩니다).</div>
            <div class="controls-row">
                <input type="month" id="homeMonth" onchange="renderHomeDeliveryTable()">
                <select id="homeDateBasis" onchange="renderHomeDeliveryTable()">
                    <option value="del">납품일 기준</option>
                    <option value="meal">급식일 기준</option>
                </select>
                <select id="homeFilterCat" onchange="renderHomeDeliveryTable()"><option value="ALL">구분 전체</option></select>
                <select id="homeFilterUnit" onchange="renderHomeDeliveryTable()"><option value="ALL">부대 전체</option></select>
                <input type="text" id="homeFilterName" placeholder="품명 검색..." onkeyup="renderHomeDeliveryTable()" style="flex:1; min-width:120px;">
            </div>
            <div class="table-wrapper" style="max-height:520px;">
                <table class="responsive-table" id="homeDeliveryTable">
                    <thead>
                        <tr>
                            <th>납품일</th><th>급식일</th><th>구분</th><th>품명</th><th>부대</th>
                            <th>예정량(kg)</th><th>확정량(kg) ✏️</th><th>오차율</th><th>매출계획</th><th>매출</th>
                        </tr>
                    </thead>
                    <tbody id="homeDeliveryBody"><tr><td colspan="10">새로고침을 눌러 불러오세요.</td></tr></tbody>
                </table>
            </div>
        </div>

        <div class="panel">
            <h2 style="font-size:15px; color:#ea580c; margin-bottom:10px;">⏰ 납품기한임박</h2>
            <div id="homeExpiryUpcoming" class="mobile-cards"><div style="color:#94a3b8; font-size:12px;">새로고침을 눌러 불러오세요.</div></div>
        </div>

        <div class="panel">
            <h2 style="font-size:15px; color:#dc2626; margin-bottom:10px;">🚨 납품기한 경과</h2>
            <div id="homeExpiryPast" class="mobile-cards"><div style="color:#94a3b8; font-size:12px;">새로고침을 눌러 불러오세요.</div></div>
        </div>
    </div>
    <!-- /tabContentHome -->

    <div id="tabContentPredict" style="display:none;">
    <div style="margin-bottom: 10px; display: flex; align-items: center; gap: 10px;">
        <label style="font-weight: bold; color: #1e3a8a; font-size: 14px;">📊 대시보드 기준 부대:</label>
        <select id="dashUnit" onchange="syncUnitFilters(this.value)" style="width: 150px; font-weight: bold;"></select>
    </div>

    <div class="dashboard-grid">
        <div class="card blue"><h3>누적 데이터 분석 건수</h3><div class="value" id="statTotalCount">0 건</div></div>
        <div class="card blue"><h3>해당 부대 평균 오차율</h3><div class="value" id="statAvgError">0 %</div></div>
        <div class="card blue"><h3>최대 변동 품목군</h3><div class="value" id="statMaxCat">-</div></div>
    </div>

    <div class="main-container">
        <div class="panel">
            <h2 style="font-size: 15px; color:#1e3a8a; margin-bottom: 10px;">🔮 엑셀 복붙형 다중 예측기</h2>

            <!-- [신규] 년/월 일괄 설정 -->
            <div class="controls-row">
                <div class="slider-container" style="flex:0 0 auto; min-width:260px;">
                    <label style="font-weight:bold; color:#1e3a8a; font-size:12px; white-space:nowrap;">📅 일괄 년/월:</label>
                    <select id="gridYear" onchange="applyYearMonthToAllRows()"></select>
                    <select id="gridMonth" onchange="applyYearMonthToAllRows()"></select>
                    <span style="font-size:11px; color:#64748b;">선택 후 각 행에서 '일'만 바꾸면 됩니다</span>
                </div>
            </div>

            <div class="controls-row">
                <div class="slider-container" style="flex:0 0 auto; min-width:220px;">
                    <label style="font-weight:bold; color:#0369a1; font-size:12px; white-space:nowrap;">🏷️ 부대 일괄입력(전체 행):</label>
                    <input type="text" id="bulkUnitInput" list="unitNameList" placeholder="여기에 부대 입력하면 아래 전체 행에 적용" style="flex:1;" onchange="applyUnitToAllRows(this.value)">
                </div>
                <div class="slider-container" style="flex:0 0 auto; min-width:220px;">
                    <label style="font-weight:bold; color:#16a34a; font-size:12px; white-space:nowrap;">🍽 식수인원(부대 기준):</label>
                    <input type="number" id="globalHeadcount" placeholder="부대 선택 시 자동입력" style="width:100px;" oninput="this.dataset.userEdited='1'; calculateMultiRows();">
                    <span id="globalHeadcountInfo" style="font-size:11px; color:#64748b;"></span>
                </div>
            </div>

            <!-- [신규] 선택한 행에만 부대 일괄적용 -->
            <div class="controls-row" style="background:#fef9c3; border-color:#fde68a;">
                <div class="slider-container" style="flex:0 0 auto; min-width:260px;">
                    <label style="font-weight:bold; color:#a16207; font-size:12px; white-space:nowrap;">✅ 체크된 행에만 부대 일괄적용:</label>
                    <input type="text" id="selectedUnitInput" list="unitNameList" placeholder="부대명 입력" style="width:160px;">
                    <button class="btn btn-gray" style="background:#a16207;" onclick="applyUnitToSelectedRows()">선택행 적용</button>
                </div>
            </div>

            <div class="controls-row">
                <div class="radio-group">
                    <label><input type="radio" name="calcMode" value="total" checked onchange="toggleCalcMode()"> ⚖️ 총 소요량</label>
                    <label><input type="radio" name="calcMode" value="person" onchange="toggleCalcMode()"> 👥 기준량(g) × 인원</label>
                </div>

                <div class="radio-group" title="입력 그리드의 날짜 칸이 급식일 기준인지 납품일 기준인지 선택합니다. 과거 데이터와의 요일(평일/휴일) 매칭에 사용됩니다.">
                    <label><input type="radio" name="dateBasis" value="meal" checked onchange="toggleDateBasis()"> 🍚 급식일 기준</label>
                    <label><input type="radio" name="dateBasis" value="delivery" onchange="toggleDateBasis()"> 🚚 납품일 기준</label>
                </div>

                <div class="slider-container" style="flex-wrap: wrap;">
                    <span style="font-weight:bold; color:#16a34a; font-size:12px;" title="더 이상 수동으로 배율을 조정하지 않습니다. 품목별 과거 데이터의 표본수·변동성을 바탕으로 결품0%를 지키는 선에서 재고가 최대한 0에 수렴하도록 시스템이 매번 자동으로 안전재고를 계산합니다.">🤖 안전재고 자동계산 적용됨 (품목별 변동성 기반)</span>
                </div>
                <button class="btn btn-success" onclick="calculateMultiRows()">🚀 일괄 산출</button>
            </div>

            <div style="overflow-x: auto; border: 1px solid #cbd5e1;">
                <table class="excel-table">
                    <thead>
                        <tr>
                            <th class="grid-select-col"><input type="checkbox" id="gridSelectAll" onchange="toggleSelectAllRows(this)"></th>
                            <th id="dateColHeader" style="width:8%;">일(day)</th>
                            <th style="width:9%;">부대 (입력)</th>
                            <th style="width:13%;">품명 (붙여넣기)</th>
                            <th style="width:8%;">구분(자동)</th>
                            <th class="col-total" style="width:7%;">총소요량</th>
                            <th class="col-person hide-col" style="width:9%;">기준량(g)</th>
                            <th style="width:6%; background:#e2e8f0; color:#334155;">단위</th>
                            <th style="width:6%; background:#fffbeb;">포장규격</th>
                            <th style="width:7%; background:#f0f9ff; color:#0284c7;">최소(팩)</th>
                            <th style="width:8%; background:#f0fdf4; color:#16a34a;">★ 권장(팩)</th>
                            <th style="width:7%; background:#fff7ed; color:#ea580c;">최대(팩)</th>
                            <th style="width:8%; background:#ede9fe; color:#6d28d9;" title="구글시트에 저장할 때 이 셋 중 어느 값을 '예정량'으로 보낼지 선택합니다. 기본은 권장입니다.">저장기준</th>
                        </tr>
                    </thead>
                    <tbody id="excelGridBody"></tbody>
                </table>
            </div>
            <datalist id="unitNameList"></datalist>
            <div style="margin-top:10px; display:flex; gap:10px; flex-wrap:wrap; align-items:center;">
                <button class="btn btn-gray" onclick="addGridRow()">+ 5행 추가</button>
                <button class="btn btn-gray" style="background:#e2e8f0; color:#333;" onclick="exportPredictionExcel()">📥 예측결과 다운로드</button>
                <button class="btn btn-gray" style="background:#0f9d58; color:#fff;" onclick="saveGridToGoogleSheet()" title="이 그리드에 입력한 행을 상단에 설정한 '저장(신규 예정량) 시트'(기본 DB_Master)에 새 예정량 행으로 추가합니다. 각 행의 '저장기준' 선택에 따라 최소/권장/최대 중 하나가 예정량으로 들어갑니다.">📤 구글시트에 저장</button>
                <button class="btn btn-gray" style="background:#7c3aed; color:#fff;" onclick="saveGridToInventoryLedger()" title="재고수불대장에 예정 입고를 등록합니다. A업체·B제품 자동 매칭, C입고일(=납품일-리드타임)·D수량·E납품일·I부대까지 자동 기록되며(빨간 글씨+연노랑 배경), F소계(확정량)만 실제 납품 후 직접 입력하시면 됩니다.">📦 재고수불대장 등록</button>
                <span id="leadDaysWrap" style="font-size:12px; color:#475569; display:flex; align-items:center; gap:4px;" title="재고수불대장의 '입고일'은 항상 납품일보다 이 일수만큼 이릅니다(창고 입고 후 납품). 급식일 기준 입력이면 부대별정보 요일 패턴으로 먼저 납품일을 역산한 뒤 여기서 뺍니다.">리드타임(입고일 = 납품일 -
                    <input type="number" id="ledgerLeadDays" value="2" min="0" style="width:45px; padding:4px;">
                    일)
                </span>
            </div>
        </div>

        <div style="display: flex; flex-direction: column; gap: 15px;">
            <details class="panel" style="margin-bottom:0;" open>
                <summary style="cursor:pointer; font-size: 14px; font-weight:bold; color:#334155;">🔥 부대별 카테고리 변동성</summary>
                <div class="volatility-list" id="volatilityList" style="margin-top:10px;"></div>
            </details>

            <details class="panel" style="margin-bottom:0;">
                <summary style="cursor:pointer; font-size: 14px; font-weight:bold; color:#334155;">📉 구분별 평균 오차율 추이</summary>
                <div class="chart-wrapper" style="margin-top:10px;"><canvas id="dashboardChart"></canvas></div>
            </details>

            <details class="panel">
                <summary style="cursor:pointer; font-size: 14px; font-weight:bold; color:#334155;">📋 원본 데이터 로그 (최근 100건)</summary>
                <div style="display:flex; gap:5px; margin:10px 0; flex-wrap:wrap;">
                    <select id="logFilterUnit" onchange="syncUnitFilters(this.value)"><option value="ALL">부대 전체</option></select>
                    <select id="logFilterCat" onchange="renderLogTable()"><option value="ALL">구분 전체</option></select>
                    <select id="logFilterDayType" onchange="renderLogTable()">
                        <option value="ALL">평/휴일 전체</option>
                        <option value="WORK">평일만(월~금)</option>
                        <option value="SAT">토요일만</option>
                        <option value="HOLI">일요일+공휴일만</option>
                        <option value="UNKNOWN">미기재/복수병기만</option>
                    </select>
                    <input type="text" id="logFilterName" placeholder="품명 검색..." onkeyup="renderLogTable()">
                </div>
                <div class="table-wrapper">
                    <table id="logTable">
                        <thead>
                            <tr>
                                <th title="체크하면 이 기록을 통계 계산(안전재고 산출)에서 제외합니다">제외</th>
                                <th>납품일</th>
                                <th>급식일</th>
                                <th>부대</th>
                                <th>구분</th>
                                <th>품명</th>
                                <th>예정</th>
                                <th>확정</th>
                                <th>오차</th>
                            </tr>
                        </thead>
                        <tbody id="logTableBody"><tr><td colspan="9">데이터를 업로드해 주세요.</td></tr></tbody>
                    </table>
                </div>
            </details>
        </div>
    </div>
    </div><!-- /tabContentPredict -->

    <div id="tabContentInventory" style="display:none;">
        <div class="panel">
            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; margin-bottom:10px;">
                <h2 style="font-size: 15px; color:#1e3a8a;">📦 재고 현황 (재고수불대장 실시간 조회)</h2>
                <div style="display:flex; align-items:center; gap:8px;">
                    <span id="inventoryStatusMsg" style="font-size:12px; color:#64748b;">아직 불러오지 않았습니다.</span>
                    <button class="btn btn-primary" onclick="loadInventoryStatus()">🔄 재고 새로고침</button>
                </div>
            </div>
            <div style="margin-bottom:10px;">
                <input type="text" id="inventorySearch" placeholder="품명 검색..." onkeyup="filterInventoryTable()" style="width:260px;">
            </div>
            <div class="table-wrapper">
                <table id="inventoryStatusTable">
                    <thead>
                        <tr>
                            <th>제품</th>
                            <th>분류(탭)</th>
                            <th>업체</th>
                            <th>현재 재고</th>
                            <th>입고예정(미반영)</th>
                            <th>최근 기록일</th>
                        </tr>
                    </thead>
                    <tbody id="inventoryStatusBody"><tr><td colspan="6">우측 상단 '재고 새로고침'을 눌러 불러오세요.</td></tr></tbody>
                </table>
            </div>
        </div>

        <div class="panel" style="margin-top:15px;">
            <h2 style="font-size: 15px; color:#7c3aed; margin-bottom: 10px;">🧂 자사양념육 (납품계획 로그)</h2>
            <p style="font-size:11px; color:#64748b; margin-bottom:10px;">이 탭은 현재고 개념이 없고, 부대별 납품계획만 누적 기록되는 방식입니다. 품목별 가장 최근 계획을 보여줍니다.</p>
            <div class="table-wrapper">
                <table id="selfSeasonedTable">
                    <thead>
                        <tr>
                            <th>품목</th>
                            <th>최근 일자</th>
                            <th>최근 소계</th>
                            <th>본품</th>
                            <th>샘플</th>
                            <th>최근 부대</th>
                            <th>비고</th>
                            <th>누적 계획건수</th>
                        </tr>
                    </thead>
                    <tbody id="selfSeasonedBody"><tr><td colspan="8">재고 새로고침을 눌러 불러오세요.</td></tr></tbody>
                </table>
            </div>
        </div>

        <div class="panel" style="margin-top:15px;">
            <h2 style="font-size: 15px; color:#ea580c; margin-bottom: 10px;">⏰ 납품기한임박</h2>
            <div class="table-wrapper">
                <table id="expiryUpcomingTable">
                    <thead>
                        <tr>
                            <th>제품명</th>
                            <th>납품기한</th>
                            <th>재고</th>
                            <th>마지막 납품</th>
                        </tr>
                    </thead>
                    <tbody id="expiryUpcomingBody"><tr><td colspan="4">재고 새로고침을 눌러 불러오세요.</td></tr></tbody>
                </table>
            </div>
        </div>

        <div class="panel" style="margin-top:15px;">
            <h2 style="font-size: 15px; color:#dc2626; margin-bottom: 10px;">🚨 납품기한 경과</h2>
            <div class="table-wrapper">
                <table id="expiryPastTable">
                    <thead>
                        <tr>
                            <th>품목</th>
                            <th>소비기한</th>
                            <th>수량</th>
                            <th>매입단가(vat별도)</th>
                            <th>합계</th>
                            <th>비고</th>
                        </tr>
                    </thead>
                    <tbody id="expiryPastBody"><tr><td colspan="6">재고 새로고침을 눌러 불러오세요.</td></tr></tbody>
                </table>
            </div>
        </div>
    </div>

    <script>
        let dbData = []; let chartIdx = null; let specMemory = {}; let unitMemory = {};
        let indexedDBInstance = null;
        window.unitInfoMap = {}; // { 부대명: { deliveryDays:['월','수','금'], headcount:'4500~5500' } }
        window.productInfoMap = {}; // { 제품명: { category, pack } }
        window.holidayMap = {}; // { 'yyyy-MM-dd': '공휴일명' }
        window.__holidayYearsLoaded = new Set();
        window.__holidayYearsFailed = new Set();

        function initIndexedDB(callback) {
            const request = indexedDB.open("SmartOrderBigDB", 1);
            request.onupgradeneeded = function(e) {
                const db = e.target.result;
                if (!db.objectStoreNames.contains("kv_store")) {
                    db.createObjectStore("kv_store");
                }
            };
            request.onsuccess = function(e) {
                indexedDBInstance = e.target.result;
                callback();
            };
            request.onerror = function() {
                alert("대용량 데이터베이스 엔진 시동에 실패했습니다.");
            };
        }

        function saveToDB(key, data, callback) {
            if (!indexedDBInstance) return;
            const tx = indexedDBInstance.transaction("kv_store", "readwrite");
            const store = tx.objectStore("kv_store");
            store.put(data, key);
            if (callback) tx.oncomplete = () => callback();
        }

        function loadFromDB(key, callback) {
            if (!indexedDBInstance) return callback(null);
            const tx = indexedDBInstance.transaction("kv_store", "readonly");
            const store = tx.objectStore("kv_store");
            const req = store.get(key);
            req.onsuccess = function() { callback(req.result); };
            req.onerror = function() { callback(null); };
        }

        // [신규] 년/월 드롭다운 채우기 (현재 연도 기준 -1 ~ +2)
        function initYearMonthSelectors() {
            const now = new Date();
            const ySel = document.getElementById('gridYear');
            const mSel = document.getElementById('gridMonth');
            const thisYear = now.getFullYear();
            for (let y = thisYear - 1; y <= thisYear + 2; y++) ySel.add(new Option(y + '년', y));
            ySel.value = thisYear;
            for (let m = 1; m <= 12; m++) mSel.add(new Option(m + '월', m));
            mSel.value = now.getMonth() + 1;
            // 홈 탭 월 선택기본값 = 이번 달
            const hm = document.getElementById('homeMonth');
            if (hm) hm.value = `${thisYear}-${String(now.getMonth()+1).padStart(2,'0')}`;
        }

        window.onload = function() {
            initYearMonthSelectors();
            for(let i=0; i<10; i++) addGridRow(true);
            
            initIndexedDB(function() {
                loadFromDB('dbData', function(loadedDb) {
                    if (loadedDb && loadedDb.length > 0) { dbData = loadedDb; refreshAll(); }
                });
                loadFromDB('specMemory', function(loadedSpec) { if (loadedSpec) specMemory = loadedSpec; });
                loadFromDB('unitMemory', function(loadedUnit) { if (loadedUnit) unitMemory = loadedUnit; });
                loadFromDB('holidayMap', function(v) { if (v) window.holidayMap = v; });
                loadFromDB('holidayYears', function(v) { if (v) window.__holidayYearsLoaded = new Set(v); });
                loadFromDB('gsheetWebAppUrl', function(loadedUrl) {
                    if (loadedUrl) {
                        document.getElementById('gsheetWebAppUrl').value = loadedUrl;
                        document.getElementById('gsheetStatus').innerText = '✅ 설정됨';
                        document.getElementById('gsheetStatus').style.color = '#16a34a';
                        const ty = new Date().getFullYear();
                        syncHolidays([ty - 1, ty, ty + 1]).then(() => {
                            calculateMultiRows();
                            refreshHomeTab(); // 접속 시 홈 대시보드 자동 로드
                        });
                        // [신규] 품명/부대 자동완성이 처음부터 바로 뜨도록 부대·품목 정보도 접속 시 자동 동기화
                        syncUnitProductInfo();
                    }
                });
                loadFromDB('gsheetLedgerSheet', function(v) { document.getElementById('gsheetLedgerSheet').value = v || ''; });
                loadFromDB('unitInfoMap', function(v) { if (v) { window.unitInfoMap = v; updateSyncStatusLabel(); } });
                loadFromDB('productInfoMap', function(v) { if (v) { window.productInfoMap = v; } });
            });

            document.getElementById('excelGridBody').addEventListener('paste', handleTablePaste);
        };

        function isMultiDateText(val) {
            if(!val) return false;
            const str = String(val);
            return /일\s*[,、]/.test(str) || /\d+\s*[,、]\s*\d+\s*일/.test(str);
        }

        function toIsoDate(dateStr) {
            if (!dateStr || dateStr === '-') return null;
            const d = new Date(dateStr);
            if (isNaN(d.getTime())) return null;
            return d.toISOString().split('T')[0];
        }

        function classifyDay(dateStr) {
            if(!dateStr || dateStr === '-') return 'EMPTY';
            if(isMultiDateText(dateStr)) return 'MULTI';
            const d = new Date(dateStr);
            if(isNaN(d.getTime())) return 'MULTI';
            const iso = toIsoDate(dateStr);
            if (iso && window.holidayMap && window.holidayMap[iso]) return 'HOLI';
            const day = d.getDay();
            if (day === 0) return 'HOLI';
            if (day === 6) return 'SAT';
            return 'WORK';
        }

        function parseExcelDate(val) {
            if(!val) return '-';
            if(typeof val === 'number' || !isNaN(Number(val))) {
                const d = new Date((Number(val) - 25569) * 86400 * 1000);
                return d.toISOString().split('T')[0];
            }
            let str = String(val).trim();
            const isoMatch = str.match(/^(\d{4}-\d{2}-\d{2})T/);
            if (isoMatch) return isoMatch[1];
            str = str.replace(/\//g, '-').replace(/\./g, '-');
            return str;
        }

        function reconstructKoreanPartialDate(raw, fallbackDateWithYear) {
            const m = String(raw || '').match(/(\d{1,2})\s*월\s*(\d{1,2})\s*일/);
            if(!m) return null;
            const month = m[1].padStart(2, '0'), day = m[2].padStart(2, '0');
            const fy = String(fallbackDateWithYear || '').match(/^(\d{4})-/);
            if(!fy) return null;
            return `${fy[1]}-${month}-${day}`;
        }

        function toggleCalcMode() {
            const mode = document.querySelector('input[name="calcMode"]:checked').value;
            const tCols = document.querySelectorAll('.col-total'); const pCols = document.querySelectorAll('.col-person');
            if(mode === 'total') { tCols.forEach(el => el.classList.remove('hide-col')); pCols.forEach(el => el.classList.add('hide-col')); } 
            else { tCols.forEach(el => el.classList.add('hide-col')); pCols.forEach(el => el.classList.remove('hide-col')); }
        }

        function getDateBasis() {
            return document.querySelector('input[name="dateBasis"]:checked').value;
        }
        function toggleDateBasis() {
            const basis = getDateBasis();
            document.getElementById('dateColHeader').innerText = '일(day)' + (basis === 'delivery' ? ' · 납품일기준' : ' · 급식일기준');
            calculateMultiRows();
        }

        function syncUnitFilters(selectedUnit) {
            document.getElementById('dashUnit').value = selectedUnit;
            document.getElementById('logFilterUnit').value = selectedUnit; updateDashboard();
        }

        // [신규] 상단 년/월 선택값을 모든 행에 일괄 적용(일자는 유지, 연/월만 갱신)
        function applyYearMonthToAllRows() {
            const y = document.getElementById('gridYear').value;
            const m = String(document.getElementById('gridMonth').value).padStart(2, '0');
            for (let tr of document.getElementById('excelGridBody').children) {
                const dateInput = tr.querySelector('.grid-date');
                let day = '01';
                const cur = dateInput.value;
                if (cur && /^\d{4}-\d{2}-\d{2}$/.test(cur)) day = cur.slice(8, 10);
                dateInput.value = `${y}-${m}-${day}`;
            }
            calculateMultiRows();
        }

        function currentGridDateDefault() {
            const y = document.getElementById('gridYear') ? document.getElementById('gridYear').value : new Date().getFullYear();
            const m = document.getElementById('gridMonth') ? String(document.getElementById('gridMonth').value).padStart(2, '0') : String(new Date().getMonth()+1).padStart(2,'0');
            const now = new Date();
            let day = String(now.getDate()).padStart(2, '0');
            // 선택된 연/월이 오늘과 다르면 1일을 기본값으로
            if (Number(y) !== now.getFullYear() || Number(m) !== now.getMonth() + 1) day = '01';
            return `${y}-${m}-${day}`;
        }

        async function applyUnitToAllRows(unitName) {
            unitName = (unitName || '').trim();
            if (!unitName) return;
            await ensureUnitProductInfoLoaded();
            const headcountInput = document.getElementById('globalHeadcount');
            if (headcountInput) delete headcountInput.dataset.userEdited;
            for (let tr of document.getElementById('excelGridBody').children) {
                const unitInput = tr.querySelector('.grid-unit-name');
                unitInput.value = unitName;
                handleUnitChange(unitInput);
            }
            const dashSel = document.getElementById('dashUnit');
            if (dashSel) {
                if (![...dashSel.options].some(o => o.value === unitName)) dashSel.add(new Option(unitName, unitName));
                syncUnitFilters(unitName);
            }
        }

        // [신규] 체크된(선택된) 행에만 부대 일괄적용
        async function applyUnitToSelectedRows() {
            const unitName = (document.getElementById('selectedUnitInput').value || '').trim();
            if (!unitName) { alert('적용할 부대명을 입력해주세요.'); return; }
            await ensureUnitProductInfoLoaded();
            let count = 0;
            for (let tr of document.getElementById('excelGridBody').children) {
                const chk = tr.querySelector('.grid-select');
                if (chk && chk.checked) {
                    const unitInput = tr.querySelector('.grid-unit-name');
                    unitInput.value = unitName;
                    handleUnitChange(unitInput);
                    count++;
                }
            }
            if (count === 0) alert('체크된 행이 없습니다. 각 행 맨 앞의 체크박스를 선택해주세요.');
        }

        function toggleSelectAllRows(headerChk) {
            document.querySelectorAll('#excelGridBody .grid-select').forEach(chk => chk.checked = headerChk.checked);
        }

        function addGridRow(isSingle=false) {
            const tbody = document.getElementById('excelGridBody');
            const mode = document.querySelector('input[name="calcMode"]:checked').value;
            const tClass = mode === 'total' ? '' : 'hide-col'; const pClass = mode === 'person' ? '' : 'hide-col';
            const lastUnit = tbody.lastElementChild ? (tbody.lastElementChild.querySelector('.grid-unit-name')?.value || '') : '';
            const defaultDate = currentGridDateDefault();

            for(let i=0; i<(isSingle?1:5); i++) {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td class="grid-select-col"><input type="checkbox" class="grid-select"></td>
                    <td>
                        <div style="display:flex; align-items:center; justify-content:center; gap:2px;">
                            <input type="date" class="grid-date" style="padding:4px; font-size:11px;" value="${defaultDate}" onchange="calculateMultiRows()" onkeydown="handleDateKey(event, this)">
                            <span class="row-day-badge day-badge" style="display:none; padding:2px 4px;"></span>
                        </div>
                    </td>
                    <td><input type="text" class="grid-unit-name" list="unitNameList" value="${lastUnit}" placeholder="부대 입력" onchange="handleUnitChange(this)"></td>
                    <td class="ac-wrap">
                        <input type="text" class="grid-name" autocomplete="off"
                            oninput="showProductSuggestions(this)"
                            onblur="handleNameBlur(this)"
                            onkeydown="handleNameKeydown(event, this)">
                        <div class="ac-list"></div>
                    </td>
                    <td><input type="text" class="grid-cat" placeholder="자동" oninput="this.dataset.autoFilled=''"></td>
                    <td class="col-total ${tClass}"><input type="number" class="grid-req-tot"></td>
                    <td class="col-person ${pClass}"><input type="number" class="grid-req-per" placeholder="예: 160(g)"></td>
                    <td style="background:#f1f5f9;">
                        <select class="grid-unit">
                            <option value="kg">kg</option>
                            <option value="ea">EA(팩)</option>
                        </select>
                    </td>
                    <td style="background:#fffbeb;"><input type="number" class="grid-spec" value="1.0"></td>
                    <td class="res-min">-</td><td class="res-rec">-</td><td class="res-max">-</td>
                    <td style="background:#f5f3ff;">
                        <select class="grid-save-basis" title="구글시트 저장 시 사용할 값">
                            <option value="min">최소</option>
                            <option value="rec" selected>★권장</option>
                            <option value="max">최대</option>
                        </select>
                    </td>
                `;
                tbody.appendChild(tr);
            }
            calculateMultiRows();
        }

        function handleDateKey(e, el) {
            if (e.ctrlKey && (e.key === 'd' || e.key === 'D')) {
                e.preventDefault();
                const tr = el.closest('tr');
                const prevTr = tr.previousElementSibling;
                if (prevTr) {
                    const prevDate = prevTr.querySelector('.grid-date').value;
                    if (prevDate) {
                        el.value = prevDate;
                        calculateMultiRows();
                    }
                }
            }
        }

        function getProductCandidates() {
            const names = new Set();
            dbData.forEach(d => { if (d.name) names.add(d.name); });
            Object.keys(window.productInfoMap || {}).forEach(n => names.add(n));
            return [...names].sort();
        }

        function showProductSuggestions(inputEl) {
            const td = inputEl.closest('td');
            const list = td.querySelector('.ac-list');
            const q = inputEl.value.trim();
            if (!q) { list.style.display = 'none'; list.innerHTML = ''; return; }
            const matches = getProductCandidates().filter(n => n.includes(q)).slice(0, 20);
            if (matches.length === 0) { list.style.display = 'none'; list.innerHTML = ''; return; }
            list.innerHTML = matches.map(n =>
                `<div data-name="${n.replace(/"/g,'&quot;')}" onmousedown="selectProductSuggestion(this, this.dataset.name)">${n}</div>`
            ).join('');
            list.style.display = 'block';
        }

        function selectProductSuggestion(el, name) {
            const td = el.closest('td');
            const input = td.querySelector('.grid-name');
            input.value = name;
            td.querySelector('.ac-list').style.display = 'none';
            autoFillMemory(input);
        }

        function handleNameBlur(inputEl) {
            setTimeout(() => {
                const list = inputEl.closest('td').querySelector('.ac-list');
                if (list) list.style.display = 'none';
            }, 150);
            autoFillMemory(inputEl);
        }

        function handleNameKeydown(e, inputEl) {
            if (e.key === 'Escape') {
                const list = inputEl.closest('td').querySelector('.ac-list');
                if (list) list.style.display = 'none';
            }
        }

        function autoFillMemory(inputEl) {
            const name = inputEl.value.trim();
            if(name) {
                const tr = inputEl.closest('tr');
                if(specMemory[name]) {
                    tr.querySelector('.grid-spec').value = specMemory[name];
                } else if (window.productInfoMap && window.productInfoMap[name]) {
                    tr.querySelector('.grid-spec').value = window.productInfoMap[name].pack;
                }
                if(unitMemory[name]) tr.querySelector('.grid-unit').value = unitMemory[name];

                const catInput = tr.querySelector('.grid-cat');
                if (!catInput.value.trim() || catInput.dataset.autoFilled === '1') {
                    let catValue = '';
                    if (dbData.length > 0) {
                        const found = dbData.find(d => d.name === name);
                        if (found) catValue = found.cat;
                    }
                    if (!catValue && window.productInfoMap && window.productInfoMap[name]) {
                        catValue = window.productInfoMap[name].category || '';
                    }
                    if (catValue) { catInput.value = catValue; catInput.dataset.autoFilled = '1'; }
                    else { catInput.value = ''; catInput.dataset.autoFilled = ''; }
                }
                calculateMultiRows();
            }
        }

        function handleTablePaste(e) {
            e.preventDefault();
            const text = (e.clipboardData || window.clipboardData).getData('Text');
            const rows = text.split('\n').map(r => r.split('\t'));
            const activeEl = document.activeElement; if (!activeEl || activeEl.tagName !== 'INPUT') return;
            const trs = document.getElementById('excelGridBody').children;
            let currRow = Array.from(trs).indexOf(activeEl.closest('tr'));
            
            for (let i = 0; i < rows.length; i++) {
                if (rows[i].length <= 1 && rows[i][0] === "") continue;
                if (currRow >= trs.length) addGridRow(true);
                
                // 체크박스(grid-select)는 붙여넣기 매핑에서 제외
                const visibleInputs = Array.from(trs[currRow].querySelectorAll('input, select'))
                    .filter(inp => !inp.closest('td').classList.contains('hide-col') && !inp.classList.contains('grid-select'));
                
                for(let j=0; j < rows[i].length && j < visibleInputs.length; j++) {
                    if(rows[i][j] !== undefined) {
                        let val = rows[i][j].trim();
                        if(visibleInputs[j].classList.contains('grid-date')) {
                            let parsed = parseExcelDate(val);
                            if(parsed !== '-' && parsed.length >= 8) visibleInputs[j].value = parsed;
                        } else if(visibleInputs[j].tagName === 'SELECT') {
                            const opt = Array.from(visibleInputs[j].options).find(o => o.value === val || o.text === val);
                            if(opt) visibleInputs[j].value = opt.value;
                        } else {
                            visibleInputs[j].value = val;
                        }
                    }
                }
                autoFillMemory(trs[currRow].querySelector('.grid-name'));
                currRow++;
            }
            calculateMultiRows();
        }

        const UNIT_ALIAS_MAP = {
            '7사단 8여단': '7사단/8여단',
            '7사단 5여단': '7사단/5여단',
            '7사단 3여단': '7사단/3여단',
            '8사단73여단': '8사단/73여단',
            '계룡대 근무지원단': '계룡대근무지원단',
            '2해병': '2해병대',
        };
        function normalizeUnitName(raw) {
            if(!raw) return raw;
            let s = String(raw).trim().replace(/\s+/g, ' ');
            return UNIT_ALIAS_MAP[s] || s;
        }

        function pickQtyColumn(headers, keyword) {
            const candidates = headers.map((h,i)=>({h,i})).filter(o => o.h.includes(keyword) && !o.h.includes('순번'));
            if(candidates.length === 0) return -1;
            const kgCol = candidates.find(o => o.h.toLowerCase().includes('kg'));
            if(kgCol) return kgCol.i;
            const eaCol = candidates.find(o => o.h.includes('개'));
            if(eaCol) { console.warn(`⚠️ "${keyword}" 항목에 kg 컬럼이 없어 개수 기준 컬럼(${eaCol.h})을 사용합니다.`); return eaCol.i; }
            return candidates[0].i;
        }

        function ingestRows(json, sourceLabel) {
            if(!json || json.length === 0) { alert('가져올 데이터가 없습니다.'); return; }
            let hIdx = 0, maxScore = 0;
            for(let i=0; i<Math.min(json.length, 10); i++) {
                const rowStr = (json[i]||[]).join(''); let score = 0;
                if(rowStr.includes('품명')) score++; if(rowStr.includes('예정')) score++; if(rowStr.includes('확정')) score++;
                if(score > maxScore) { maxScore = score; hIdx = i; }
            }
            const headers = json[hIdx].map(h => String(h).replace(/\s/g, ''));
            const map = { 
                uid: headers.findIndex(h => h.includes('고유ID') || h.includes('고유')),
                delDate: headers.findIndex(h => h.includes('납품')), 
                mealDate: headers.findIndex(h => h.includes('급식')), 
                date: headers.findIndex(h => h.includes('일자')), 
                name: headers.findIndex(h => h.includes('품명')), 
                cat: headers.findIndex(h => h.includes('구분')), 
                unit: headers.findIndex(h => h.includes('부대')), 
                plan: pickQtyColumn(headers, '예정'), 
                conf: pickQtyColumn(headers, '확정'),
                // [신규] 홈 대시보드에서 엑셀 출고계획 형식(단가/매출)까지 함께 보여주기 위해 추가 캡처
                price: headers.findIndex(h => h.includes('단가')),
                planSales: headers.findIndex(h => h.includes('예정매출')),
                confSales: headers.findIndex(h => h.includes('확정매출'))
            };

            const skip = { bothEmpty: 0, duplicate: 0, planMissingKept: 0, confMissingKept: 0, multiDateKept: 0, markedExcluded: 0 };
            const existingKeys = new Set(dbData.map(d => `${d.delDate}|${d.mealDate}|${d.unit}|${d.name}|${d.plan}|${d.conf}`));
            const existingUnits = new Set(dbData.map(d => d.unit));
            const newUnitsSeen = new Set();

            const newList = [];
            for(let i = hIdx + 1; i < json.length; i++) {
                const row = json[i]; if(!row || !row[map.name]) continue;

                const uidCell = map.uid !== -1 ? row[map.uid] : '';
                const isMarkedExcluded = String(uidCell || '').trim() === '제외';
                if (isMarkedExcluded) skip.markedExcluded++;

                const planRaw = map.plan !== -1 ? row[map.plan] : undefined;
                const confRaw = map.conf !== -1 ? row[map.conf] : undefined;
                const planMissing = (planRaw === undefined || planRaw === null || planRaw === '' || isNaN(parseFloat(planRaw)) || parseFloat(planRaw) === 0);
                const confMissing = (confRaw === undefined || confRaw === null || confRaw === '');
                const p = planMissing ? null : Math.round(parseFloat(planRaw));
                const c = confMissing ? null : Math.round(parseFloat(confRaw) || 0);

                if(p === null && c === null) { skip.bothEmpty++; continue; }

                let dDate = map.delDate !== -1 ? parseExcelDate(row[map.delDate]) : (map.date !== -1 ? parseExcelDate(row[map.date]) : '-');
                let mDate = map.mealDate !== -1 ? parseExcelDate(row[map.mealDate]) : (map.date !== -1 ? parseExcelDate(row[map.date]) : '-');
                const reconstructedDDate = reconstructKoreanPartialDate(row[map.delDate], mDate);
                if(reconstructedDDate) dDate = reconstructedDDate;

                const dayCat = classifyDay(mDate);
                const isHoli = (dayCat === 'WORK' || dayCat === 'SAT') ? false : (dayCat === 'HOLI') ? true : null;
                const delDayCat = classifyDay(dDate);
                const isDelHoli = (delDayCat === 'WORK' || delDayCat === 'SAT') ? false : (delDayCat === 'HOLI') ? true : null;
                const holidayNameVal = (dayCat === 'HOLI') ? (window.holidayMap[toIsoDate(mDate)] || '') : '';
                if(dayCat === 'MULTI') skip.multiDateKept++;

                const nameVal = String(row[map.name]).trim();
                const catVal = map.cat!==-1 ? String(row[map.cat]).trim().replace(/\s+/g,' ') : '기타';
                const unitValRaw = map.unit!==-1 ? String(row[map.unit]).trim() : '공통';
                const unitVal = normalizeUnitName(unitValRaw);
                if(!existingUnits.has(unitVal)) newUnitsSeen.add(unitVal);

                const key = `${dDate}|${mDate}|${unitVal}|${nameVal}|${p}|${c}`;
                if(existingKeys.has(key)) { skip.duplicate++; continue; }
                existingKeys.add(key);

                if(p === null) skip.planMissingKept++;
                if(c === null) skip.confMissingKept++;

                const priceVal = map.price !== -1 && row[map.price] !== undefined && row[map.price] !== '' ? parseFloat(String(row[map.price]).replace(/,/g,'')) : null;
                const planSalesVal = map.planSales !== -1 && row[map.planSales] !== undefined && row[map.planSales] !== '' ? parseFloat(String(row[map.planSales]).replace(/,/g,'')) : null;
                const confSalesVal = map.confSales !== -1 && row[map.confSales] !== undefined && row[map.confSales] !== '' ? parseFloat(String(row[map.confSales]).replace(/,/g,'')) : null;

                newList.push({ 
                    id: key,
                    delDate: dDate, mealDate: mDate, isHoliday: isHoli, dayType: dayCat,
                    isDelHoliday: isDelHoli, delDayType: delDayCat,
                    dayCategory: dayCat, delDayCategory: delDayCat, holidayName: holidayNameVal,
                    name: nameVal, cat: catVal, unit: unitVal,
                    plan: p, conf: c,
                    price: (priceVal !== null && !isNaN(priceVal)) ? priceVal : null,
                    planSales: (planSalesVal !== null && !isNaN(planSalesVal)) ? planSalesVal : null,
                    confSales: (confSalesVal !== null && !isNaN(confSalesVal)) ? confSalesVal : null,
                    rate: (p !== null && c !== null) ? ((c - p) / p) * 100 : null,
                    excludedFromStats: isMarkedExcluded
                });
            }
            dbData = dbData.concat(newList);
            saveToDB('dbData', dbData, function() {
                refreshAll();
                refreshHomeTab(false); // [신규] 데이터 갱신 시 홈 화면도 함께 재계산 (원격 재조회는 안 함)
                let msg = `✅ ${sourceLabel || '업로드'} 완료! 이번 ${newList.length.toLocaleString()}건 추가 (누적 ${dbData.length.toLocaleString()}건)\n\n`;
                msg += `- 완전 공백행 제외: ${skip.bothEmpty}건\n`;
                msg += `- 중복행 제외(재업로드 등): ${skip.duplicate}건\n`;
                msg += `- 예정량 없이 유지(재포장 등): ${skip.planMissingKept}건 → 오차율 계산엔 미사용, 로그에서만 확인\n`;
                msg += `- 확정량 없이 유지(급식일별 미분리 등): ${skip.confMissingKept}건 → 오차율 계산엔 미사용, 로그에서만 확인\n`;
                msg += `- 급식일 복수병기/판독불가: ${skip.multiDateKept}건 → 평일/휴일 통계에서만 제외\n`;
                if (skip.markedExcluded > 0) msg += `- 🚫 고유ID열에 "제외" 표시된 행: ${skip.markedExcluded}건 → 통계 계산(안전재고 산출)에서 제외, 로그에는 그대로 표시\n`;
                if(newUnitsSeen.size > 0) {
                    msg += `\n⚠️ 처음 보는 부대명 ${newUnitsSeen.size}건 (신규 부대이거나 오타일 수 있으니 확인 권장):\n- ` + [...newUnitsSeen].join('\n- ');
                }
                alert(msg);
            });
        }

        function saveGSheetUrl() {
            const url = document.getElementById('gsheetWebAppUrl').value.trim();
            saveToDB('gsheetWebAppUrl', url, function() {
                const status = document.getElementById('gsheetStatus');
                if(url) { status.innerText = '✅ 설정됨'; status.style.color = '#16a34a'; }
                else { status.innerText = '미설정'; status.style.color = '#64748b'; }
            });
        }

        const HISTORY_SHEET_NAME = '데이터';
        const WRITE_SHEET_NAME = 'DB_Master';
        function saveGSheetSheetNames() {
            const ledger = document.getElementById('gsheetLedgerSheet').value.trim();
            saveToDB('gsheetLedgerSheet', ledger);
        }

        function getGSheetUrl() {
            const url = document.getElementById('gsheetWebAppUrl').value.trim();
            if(!url) { alert("먼저 상단에 구글시트 웹앱(Apps Script) URL을 입력해주세요. 배포 방법은 안내 문서를 참고하세요."); return null; }
            return url;
        }

        function getGSheetUrlSilent() {
            const url = document.getElementById('gsheetWebAppUrl').value.trim();
            return url || null;
        }

        async function syncHolidays(years, force = false) {
            const url = getGSheetUrlSilent(); if (!url) return { ok: false, errors: [] };
            const yearStrs = [...new Set(years.map(String))];
            const missing = force ? yearStrs : yearStrs.filter(y => !window.__holidayYearsLoaded.has(y));
            if (missing.length === 0) return { ok: true, errors: [] };

            try {
                let reqUrl = url + (url.includes('?') ? '&' : '?') + 'action=getHolidays&years=' + missing.join(',');
                if (force) reqUrl += '&force=1';
                const res = await fetch(reqUrl);
                const rawText = await res.text();
                const json = JSON.parse(rawText);
                if (json.error) throw new Error(json.error);

                Object.assign(window.holidayMap, json.holidays || {});
                const errYears = new Set((json.errors || []).map(e => String(e.year)));
                missing.forEach(y => { if (!errYears.has(y)) window.__holidayYearsLoaded.add(y); });
                if (force) missing.forEach(y => window.__holidayYearsFailed.delete(y));
                saveToDB('holidayMap', window.holidayMap);
                saveToDB('holidayYears', [...window.__holidayYearsLoaded]);
                return { ok: errYears.size === 0, errors: json.errors || [] };
            } catch (err) {
                console.error('공휴일 동기화 실패:', err);
                return { ok: false, errors: [{ year: missing.join(','), message: err.message }] };
            }
        }

        async function reclassifyExistingHolidays() {
            if (dbData.length === 0) { alert('재분류할 데이터가 없습니다.'); return; }
            const url = getGSheetUrl(); if (!url) return;

            const years = new Set();
            dbData.forEach(d => {
                const my = String(d.mealDate || '').match(/^(\d{4})-/); if (my) years.add(my[1]);
                const dy = String(d.delDate || '').match(/^(\d{4})-/); if (dy) years.add(dy[1]);
            });
            if (years.size === 0) { alert('데이터에서 연도를 찾을 수 없습니다.'); return; }

            const btn = (typeof event !== 'undefined' && event.target) ? event.target : null;
            const originalText = btn ? btn.innerText : null;
            if (btn) { btn.disabled = true; btn.innerText = '⏳ 공휴일 조회 중...'; }

            const syncResult = await syncHolidays([...years], true);

            if (btn) { btn.disabled = false; btn.innerText = originalText; }

            if (syncResult.errors && syncResult.errors.length > 0) {
                const errMsgs = syncResult.errors.map(e => `- ${e.year}년: ${e.message}`).join('\n');
                const proceed = confirm(`⚠️ 일부 연도 공휴일 조회에 실패했습니다:\n${errMsgs}\n\n실패한 연도는 이번엔 반영되지 않지만, 다음에 다시 시도하면 재조회됩니다(임의로 '평일'로 고정되지 않음).\n성공한 연도만이라도 재분류를 계속할까요?`);
                if (!proceed) return;
            } else if (!syncResult.ok) {
                alert('공휴일 정보를 불러오지 못했습니다. 웹앱 URL/배포 상태를 확인해주세요.');
                return;
            }

            let changed = 0;
            dbData.forEach(d => {
                const newCat = classifyDay(d.mealDate);
                const newDelCat = classifyDay(d.delDate);
                const newIsHoli = (newCat === 'WORK' || newCat === 'SAT') ? false : (newCat === 'HOLI') ? true : null;
                const newIsDelHoli = (newDelCat === 'WORK' || newDelCat === 'SAT') ? false : (newDelCat === 'HOLI') ? true : null;
                const newHolidayName = (newCat === 'HOLI') ? (window.holidayMap[toIsoDate(d.mealDate)] || '') : '';

                if (d.dayCategory !== newCat || d.delDayCategory !== newDelCat || d.isHoliday !== newIsHoli || d.isDelHoliday !== newIsDelHoli) {
                    d.dayCategory = newCat; d.delDayCategory = newDelCat;
                    d.isHoliday = newIsHoli; d.isDelHoliday = newIsDelHoli;
                    d.dayType = newCat; d.delDayType = newDelCat;
                    d.holidayName = newHolidayName;
                    changed++;
                }
            });

            if (changed === 0) { alert('이미 최신 공휴일 기준으로 분류되어 있어 변경된 행이 없습니다.'); return; }
            saveToDB('dbData', dbData, function() {
                refreshAll();
                alert(`✅ 휴일 재분류 완료: ${changed.toLocaleString()}건의 평일/토요일/휴일(일요일+공휴일) 분류가 갱신되었습니다.\n(조회한 연도: ${[...years].sort().join(', ')})`);
            });
        }

        function extractRowsArray(json) {
            if (Array.isArray(json)) return json;
            if (json && typeof json === 'object') {
                for (const key of ['values', 'data', 'rows', 'result', 'DB_Master']) {
                    if (Array.isArray(json[key])) return json[key];
                }
            }
            return null;
        }

        function updateSyncStatusLabel() {
            const el = document.getElementById('unitProductSyncStatus');
            const uCount = Object.keys(window.unitInfoMap || {}).length;
            const pCount = Object.keys(window.productInfoMap || {}).length;
            if (uCount === 0 && pCount === 0) { el.innerText = '미동기화 — "자동화시스템"의 부대별정보/품목정보 시트를 불러옵니다'; el.style.color = '#64748b'; return; }
            el.innerText = `✅ 부대 ${uCount}건, 품목 ${pCount}건 동기화됨`;
            el.style.color = '#16a34a';
        }

        async function syncUnitProductInfo() {
            const url = getGSheetUrl(); if (!url) return;
            const statusEl = document.getElementById('unitProductSyncStatus');
            statusEl.innerText = '⏳ 동기화 중...'; statusEl.style.color = '#64748b';
            try {
                const [uRes, pRes] = await Promise.all([
                    fetch(url + (url.includes('?') ? '&' : '?') + 'action=getUnitInfo'),
                    fetch(url + (url.includes('?') ? '&' : '?') + 'action=getProductInfo')
                ]);
                const uJson = JSON.parse(await uRes.text());
                const pJson = JSON.parse(await pRes.text());
                if (uJson.error) throw new Error('부대별정보: ' + uJson.error);
                if (pJson.error) throw new Error('품목정보: ' + pJson.error);

                const unitMap = {};
                (uJson.units || []).forEach(u => { unitMap[u.name] = { deliveryDays: u.deliveryDays, headcount: u.headcount }; });
                const productMap = {};
                (pJson.products || []).forEach(p => { productMap[p.name] = { category: p.category, pack: p.pack }; });

                window.unitInfoMap = unitMap;
                window.productInfoMap = productMap;
                saveToDB('unitInfoMap', unitMap);
                saveToDB('productInfoMap', productMap);
                updateSyncStatusLabel();
            } catch (err) {
                statusEl.innerText = '❌ 동기화 실패: ' + err.message;
                statusEl.style.color = '#ef4444';
            }
        }

        const KOREAN_DAY_TO_JS = { '일': 0, '월': 1, '화': 2, '수': 3, '목': 4, '금': 5, '토': 6 };

        function getUnitDeliveryDayNums(unitName) {
            const info = (window.unitInfoMap || {})[unitName];
            if (!info || !info.deliveryDays || info.deliveryDays.length === 0) return null;
            const nums = info.deliveryDays.map(d => KOREAN_DAY_TO_JS[d]).filter(v => v !== undefined);
            return nums.length > 0 ? nums : null;
        }

        function shiftToPatternDay(dateStr, patternDays, direction) {
            if (!dateStr || !patternDays) return null;
            const d = new Date(dateStr);
            if (isNaN(d.getTime())) return null;
            for (let step = 1; step <= 8; step++) {
                const cand = new Date(d);
                cand.setDate(cand.getDate() + (direction === 'next' ? step : -step));
                if (patternDays.includes(cand.getDay())) return cand.toISOString().split('T')[0];
            }
            return null;
        }

        function computeCounterpartDate(unitName, dateStr, basis) {
            const pattern = getUnitDeliveryDayNums(unitName);
            if (!pattern) return null;
            return shiftToPatternDay(dateStr, pattern, basis === 'delivery' ? 'next' : 'prev');
        }

        function guessHeadcountNumber(raw) {
            if (!raw) return null;
            const str = String(raw).trim();
            if (/^\d+$/.test(str)) return Number(str);
            const rangeMatch = str.match(/^(\d+)\s*~\s*(\d+)$/);
            if (rangeMatch) return Math.round((Number(rangeMatch[1]) + Number(rangeMatch[2])) / 2);
            return null;
        }

        function handleUnitChange(selectEl) {
            checkUnitTypo(selectEl);
            const unitName = selectEl.value.trim();
            const headcountInput = document.getElementById('globalHeadcount');
            if (headcountInput && headcountInput.dataset.userEdited !== '1') {
                const info = window.unitInfoMap[unitName];
                const guess = info ? guessHeadcountNumber(info.headcount) : null;
                if (guess !== null) headcountInput.value = guess;
                const infoSpan = document.getElementById('globalHeadcountInfo');
                if (infoSpan) infoSpan.innerText = (info && info.headcount) ? `(원본: ${info.headcount}명)` : '';
            }
            calculateMultiRows();
        }

        function levenshteinDistance(s1, s2) {
            s1 = String(s1 || ''); s2 = String(s2 || '');
            const m = s1.length, n = s2.length;
            if (m === 0) return n;
            if (n === 0) return m;
            const dp = Array.from({length: m + 1}, () => new Array(n + 1).fill(0));
            for (let i = 0; i <= m; i++) dp[i][0] = i;
            for (let j = 0; j <= n; j++) dp[0][j] = j;
            for (let i = 1; i <= m; i++) {
                for (let j = 1; j <= n; j++) {
                    const cost = s1[i-1] === s2[j-1] ? 0 : 1;
                    dp[i][j] = Math.min(dp[i-1][j] + 1, dp[i][j-1] + 1, dp[i-1][j-1] + cost);
                }
            }
            return dp[m][n];
        }

        function getKnownUnitNames() {
            const set = new Set();
            Object.keys(window.unitInfoMap || {}).forEach(n => set.add(n));
            if (set.size === 0) dbData.forEach(d => { if (d.unit) set.add(d.unit); });
            return [...set];
        }

        function findClosestUnitMatch(input) {
            const known = getKnownUnitNames();
            if (!input || known.length === 0) return null;
            let best = null, bestDist = Infinity;
            for (const name of known) {
                if (name === input) continue;
                const dist = levenshteinDistance(input, name);
                if (dist < bestDist) { bestDist = dist; best = name; }
            }
            if (!best) return null;
            const threshold = Math.max(1, Math.min(2, Math.floor(Math.max(input.length, best.length) * 0.3)));
            return bestDist <= threshold ? { name: best, dist: bestDist } : null;
        }

        function checkUnitTypo(inputEl) {
            const raw = inputEl.value.trim();
            const td = inputEl.closest('td');
            let hintEl = td ? td.querySelector('.unit-typo-hint') : null;
            if (td && !hintEl) {
                hintEl = document.createElement('div');
                hintEl.className = 'unit-typo-hint';
                hintEl.style.cssText = 'font-size:10px; line-height:1.3; margin-top:1px; cursor:pointer; display:none;';
                td.appendChild(hintEl);
            }
            if (!raw) { inputEl.style.borderColor = ''; inputEl.title=''; if(hintEl) hintEl.style.display='none'; return; }

            const known = getKnownUnitNames();
            const normalized = normalizeUnitName(raw);
            const exact = known.length === 0 || known.includes(raw) || known.includes(normalized);

            if (exact) {
                inputEl.style.borderColor = '';
                inputEl.title = '';
                if (hintEl) hintEl.style.display = 'none';
                return;
            }

            const match = findClosestUnitMatch(raw);
            if (match) {
                inputEl.style.borderColor = '#ef4444';
                inputEl.title = `⚠️ 오타 의심: "${match.name}"을(를) 의도하신 건가요?`;
                if (hintEl) {
                    hintEl.style.display = 'block';
                    hintEl.style.color = '#ef4444';
                    hintEl.innerHTML = `⚠️ "${match.name}"?`;
                    hintEl.onclick = function() {
                        inputEl.value = match.name;
                        handleUnitChange(inputEl);
                    };
                }
            } else {
                inputEl.style.borderColor = '#f59e0b';
                inputEl.title = '"부대별정보" 탭에 없는 새 부대명입니다. 오타가 아니라면 그대로 두세요.';
                if (hintEl) {
                    hintEl.style.display = 'block';
                    hintEl.style.color = '#b45309';
                    hintEl.innerHTML = '🆕 신규 부대명';
                    hintEl.onclick = null;
                }
            }
        }

        // ============================================================
        // [신규 v10] 탭 전환 (홈 / 예정량입력 / 재고상세)
        // ============================================================
        function switchTab(name) {
            document.getElementById('tabContentHome').style.display = (name === 'home') ? '' : 'none';
            document.getElementById('tabContentPredict').style.display = (name === 'predict') ? '' : 'none';
            document.getElementById('tabContentInventory').style.display = (name === 'inventory') ? '' : 'none';
            document.getElementById('tabBtnHome').classList.toggle('active', name === 'home');
            document.getElementById('tabBtnPredict').classList.toggle('active', name === 'predict');
            document.getElementById('tabBtnInventory').classList.toggle('active', name === 'inventory');
            if (name === 'inventory' && !window.__inventoryLoadedOnce) {
                window.__inventoryLoadedOnce = true;
                loadInventoryStatus();
            }
        }

        const DELIVERY_LOG_SHEETS = ['범용제품'];
        const SELF_SEASONED_SHEETS = ['자사양념육'];
        const EXPIRY_UPCOMING_SHEETS = ['납품기한임박'];
        const EXPIRY_PAST_SHEETS = ['납품기한 경과'];
        function getSheetKindClient(sheetName) {
            if (EXPIRY_UPCOMING_SHEETS.indexOf(sheetName) !== -1) return 'expiryUpcoming';
            if (EXPIRY_PAST_SHEETS.indexOf(sheetName) !== -1) return 'expiryPast';
            if (SELF_SEASONED_SHEETS.indexOf(sheetName) !== -1) return 'selfSeasoned';
            if (DELIVERY_LOG_SHEETS.indexOf(sheetName) !== -1) return 'deliveryLog';
            return 'standard';
        }

        async function loadInventoryStatus() {
            const url = getGSheetUrl(); if(!url) return;
            const statusEl = document.getElementById('inventoryStatusMsg');
            statusEl.innerText = '⏳ 불러오는 중...'; statusEl.style.color = '#64748b';
            let rawText = '';
            try {
                const res = await fetch(url + (url.includes('?') ? '&' : '?') + 'action=readLedgerAll');
                rawText = await res.text();
                if(!res.ok) throw new Error(`HTTP ${res.status} — ${rawText.slice(0,200)}`);
                let json;
                try { json = JSON.parse(rawText); }
                catch(e) { throw new Error(`응답이 JSON이 아닙니다: ${rawText.slice(0,200)}`); }
                if (json && json.error) throw new Error(json.error);
                const sheetsObj = json && json.sheets;
                if (!sheetsObj || typeof sheetsObj !== 'object') throw new Error(`탭별 데이터를 찾지 못했습니다: ${JSON.stringify(json).slice(0,200)}`);

                renderInventoryStatus(sheetsObj);
                const tabCount = Object.keys(sheetsObj).length;
                statusEl.innerText = `✅ ${new Date().toLocaleTimeString('ko-KR')} 기준 (탭 ${tabCount}개 통합)`;
                statusEl.style.color = '#16a34a';
            } catch(err) {
                console.error('재고 현황 불러오기 실패:', err, rawText);
                statusEl.innerText = '❌ 불러오기 실패: ' + err.message;
                statusEl.style.color = '#ef4444';
                document.getElementById('inventoryStatusBody').innerHTML = `<tr><td colspan="6">불러오기 실패했습니다. 웹앱 URL과 배포 상태를 확인해주세요.</td></tr>`;
            }
        }

        function renderInventoryStatus(sheetsObj) {
            const productMap = {}; const order = [];
            let selfSeasonedList = [], expiryUpcomingList = [], expiryPastList = [];

            Object.keys(sheetsObj).forEach(sheetName => {
                const rows = sheetsObj[sheetName] || [];
                const kind = getSheetKindClient(sheetName);

                if (kind === 'selfSeasoned') { selfSeasonedList = selfSeasonedList.concat(parseSelfSeasonedSheet(rows)); return; }
                if (kind === 'expiryUpcoming') { expiryUpcomingList = expiryUpcomingList.concat(parseExpiryUpcomingSheet(rows)); return; }
                if (kind === 'expiryPast') { expiryPastList = expiryPastList.concat(parseExpiryPastSheet(rows)); return; }

                if (kind === 'deliveryLog') {
                    let currentProduct = '';
                    for (let i = 1; i < rows.length; i++) {
                        const r = rows[i]; if (!r) continue;
                        const col0 = String(r[0] || '').trim();
                        if (col0 === '합계') continue;
                        if (col0) currentProduct = col0;
                        if (!currentProduct) continue;
                        const date = r[1], qty = r[2];
                        if (!date && !qty) continue;
                        const key = sheetName + '|' + currentProduct;
                        if (!productMap[key]) { productMap[key] = { name: currentProduct, category: sheetName, company: '', stock: null, pendingIn: 0, lastDate: '', stockUnavailable: true }; order.push(key); }
                        if (date) productMap[key].lastDate = String(date);
                    }
                    return;
                }

                let dataStart = 2;
                for (let i = 0; i < Math.min(rows.length, 5); i++) {
                    const rowStr = (rows[i]||[]).join('');
                    if (rowStr.includes('제품') && rowStr.includes('재고')) { dataStart = i + 1; break; }
                }

                for (let i = dataStart; i < rows.length; i++) {
                    const r = rows[i]; if(!r) continue;
                    const company = r[0], product = String(r[1]||'').trim();
                    if (!product) continue;
                    const inDate = r[2], inQty = r[3];
                    const outDate = r[4], stockVal = r[9];

                    const key = sheetName + '|' + product;
                    if (!productMap[key]) { productMap[key] = { name: product, category: sheetName, company: '', stock: null, pendingIn: 0, lastDate: '' }; order.push(key); }
                    const p = productMap[key];
                    if (company) p.company = company;
                    const stockNum = Number(stockVal);
                    if (stockVal !== '' && stockVal !== null && stockVal !== undefined && !isNaN(stockNum)) {
                        p.stock = stockNum;
                    }
                    if (inQty && (stockVal === '' || stockVal === null || stockVal === undefined)) {
                        p.pendingIn += Number(inQty) || 0;
                    }
                    const latestDate = outDate || inDate;
                    if (latestDate) p.lastDate = String(latestDate);
                }
            });

            window.__inventoryProductMap = productMap;
            window.__inventoryOrder = order;
            filterInventoryTable();

            window.__selfSeasonedList = selfSeasonedList;
            window.__expiryUpcomingList = expiryUpcomingList;
            window.__expiryPastList = expiryPastList;
            renderSelfSeasonedTable();
            renderExpiryUpcomingTable();
            renderExpiryPastTable();
            renderHomeExpiryCards(); // [신규] 홈 탭 임박/경과 카드도 함께 갱신
        }

        function parseSelfSeasonedSheet(rows) {
            let dataStart = 2;
            for (let i = 0; i < Math.min(rows.length, 5); i++) {
                const rowStr = (rows[i] || []).join('');
                if (rowStr.includes('품목') && rowStr.includes('일자')) { dataStart = i + 1; break; }
            }
            const map = {}; const order = [];
            let currentProduct = '';
            for (let i = dataStart; i < rows.length; i++) {
                const r = rows[i]; if (!r) continue;
                const col0 = String(r[0] || '').trim();
                if (col0) currentProduct = col0;
                if (!currentProduct) continue;
                const date = r[1], subtotal = r[2], main = r[3], sample = r[4], unit = r[5], note = r[6];
                if (!date && (subtotal === '' || subtotal === undefined || subtotal === null)) continue;
                if (!map[currentProduct]) { map[currentProduct] = { name: currentProduct, lastDate: '', lastSubtotal: null, lastMain: null, lastSample: null, lastUnit: '', note: '', count: 0 }; order.push(currentProduct); }
                const p = map[currentProduct];
                if (date) p.lastDate = String(date);
                if (subtotal !== '' && subtotal !== undefined && subtotal !== null) p.lastSubtotal = Number(subtotal);
                if (main !== '' && main !== undefined && main !== null) p.lastMain = Number(main);
                if (sample !== '' && sample !== undefined && sample !== null) p.lastSample = Number(sample);
                if (unit) p.lastUnit = String(unit);
                if (note && note !== '-') p.note = String(note);
                p.count++;
            }
            return order.map(name => map[name]);
        }

        function renderSelfSeasonedTable() {
            const list = window.__selfSeasonedList || [];
            const tbody = document.getElementById('selfSeasonedBody');
            if (list.length === 0) { tbody.innerHTML = '<tr><td colspan="8">표시할 품목이 없습니다.</td></tr>'; return; }
            tbody.innerHTML = list.map(p => `<tr>
                <td style="text-align:left; font-weight:600;">${p.name}</td>
                <td>${p.lastDate || '-'}</td>
                <td style="font-weight:bold; color:#7c3aed;">${p.lastSubtotal !== null ? p.lastSubtotal.toLocaleString() : '-'}</td>
                <td>${p.lastMain !== null ? p.lastMain.toLocaleString() : '-'}</td>
                <td>${p.lastSample !== null ? p.lastSample.toLocaleString() : '-'}</td>
                <td>${p.lastUnit || '-'}</td>
                <td style="color:#64748b; font-size:11px;">${p.note || '-'}</td>
                <td style="color:#94a3b8;">${p.count}건</td>
            </tr>`).join('');
        }

        function parseExpiryUpcomingSheet(rows) {
            let dataStart = 2;
            for (let i = 0; i < Math.min(rows.length, 5); i++) {
                const rowStr = (rows[i] || []).join('');
                if (rowStr.includes('제품명') && rowStr.includes('재고')) { dataStart = i + 1; break; }
            }
            const list = [];
            for (let i = dataStart; i < rows.length; i++) {
                const r = rows[i]; if (!r) continue;
                const name = String(r[0] || '').trim(); if (!name) continue;
                const stockVal = r[2];
                list.push({
                    name,
                    deadline: r[1] || '',
                    stock: (stockVal !== '' && stockVal !== undefined && stockVal !== null && !isNaN(Number(stockVal))) ? Number(stockVal) : null,
                    lastDelivery: r[3] || ''
                });
            }
            return list;
        }

        function renderExpiryUpcomingTable() {
            const list = window.__expiryUpcomingList || [];
            const tbody = document.getElementById('expiryUpcomingBody');
            if (list.length === 0) { tbody.innerHTML = '<tr><td colspan="4">표시할 품목이 없습니다.</td></tr>'; return; }
            tbody.innerHTML = list.map(p => `<tr>
                <td style="text-align:left; font-weight:600;">${p.name}</td>
                <td style="color:#ea580c; font-weight:bold;">${p.deadline || '-'}</td>
                <td>${p.stock !== null ? p.stock.toLocaleString() : '<span style="color:#94a3b8;">기록없음</span>'}</td>
                <td style="color:#64748b; font-size:11px; text-align:left;">${p.lastDelivery || '-'}</td>
            </tr>`).join('');
        }

        function parseExpiryPastSheet(rows) {
            let dataStart = 1;
            for (let i = 0; i < Math.min(rows.length, 5); i++) {
                const rowStr = (rows[i] || []).join('');
                if (rowStr.includes('품목') && rowStr.includes('소비기한')) { dataStart = i + 1; break; }
            }
            const list = [];
            for (let i = dataStart; i < rows.length; i++) {
                const r = rows[i]; if (!r) continue;
                const name = String(r[0] || '').trim();
                if (!name || name.includes('총합계') || name.includes('합계')) continue;
                const qtyVal = r[2], priceVal = r[3], totalVal = r[4];
                list.push({
                    name,
                    expiry: r[1] || '',
                    qty: (qtyVal !== '' && qtyVal !== undefined && qtyVal !== null && !isNaN(Number(qtyVal))) ? Number(qtyVal) : null,
                    unitPrice: (priceVal !== '' && priceVal !== undefined && priceVal !== null && !isNaN(Number(priceVal))) ? Number(priceVal) : null,
                    total: (totalVal !== '' && totalVal !== undefined && totalVal !== null && !isNaN(Number(totalVal))) ? Number(totalVal) : null,
                    note: r[5] || ''
                });
            }
            return list;
        }

        function renderExpiryPastTable() {
            const list = window.__expiryPastList || [];
            const tbody = document.getElementById('expiryPastBody');
            if (list.length === 0) { tbody.innerHTML = '<tr><td colspan="6">표시할 품목이 없습니다.</td></tr>'; return; }
            tbody.innerHTML = list.map(p => `<tr>
                <td style="text-align:left; font-weight:600;">${p.name}</td>
                <td style="color:#dc2626; font-weight:bold;">${p.expiry || '-'}</td>
                <td>${p.qty !== null ? p.qty.toLocaleString() : '-'}</td>
                <td>${p.unitPrice !== null ? p.unitPrice.toLocaleString() : '-'}</td>
                <td style="font-weight:bold;">${p.total !== null ? p.total.toLocaleString() : '-'}</td>
                <td style="color:#64748b; font-size:11px; text-align:left;">${p.note || '-'}</td>
            </tr>`).join('');
        }

        function filterInventoryTable() {
            const order = window.__inventoryOrder || [];
            const productMap = window.__inventoryProductMap || {};
            const kw = (document.getElementById('inventorySearch').value || '').trim().toLowerCase();
            const tbody = document.getElementById('inventoryStatusBody');

            const filtered = order.filter(key => {
                const p = productMap[key];
                return !kw || p.name.toLowerCase().includes(kw) || p.category.toLowerCase().includes(kw);
            });
            if (filtered.length === 0) { tbody.innerHTML = '<tr><td colspan="6">표시할 품목이 없습니다.</td></tr>'; return; }

            tbody.innerHTML = filtered.map(key => {
                const p = productMap[key];
                const stockTxt = p.stockUnavailable
                    ? '<span style="color:#94a3b8;">재고 미관리(납품이력만)</span>'
                    : (p.stock === null ? '<span style="color:#94a3b8;">기록없음</span>' : p.stock.toLocaleString());
                const stockColor = (!p.stockUnavailable && p.stock !== null && p.stock <= 0) ? '#ef4444' : '#16a34a';
                return `<tr>
                    <td style="text-align:left; font-weight:600;">${p.name}</td>
                    <td style="color:#64748b;">${p.category}</td>
                    <td>${p.company || '-'}</td>
                    <td style="font-weight:bold; color:${stockColor};">${stockTxt}</td>
                    <td style="color:#7c3aed;">${p.pendingIn > 0 ? '+' + p.pendingIn.toLocaleString() : '-'}</td>
                    <td style="color:#64748b; font-size:11px;">${p.lastDate || '-'}</td>
                </tr>`;
            }).join('');
        }

        // ============================================================
        // [신규] 홈 대시보드
        // ============================================================

        // 새로고침 버튼: 구글시트 '데이터' 탭 + 재고수불대장을 모두 다시 불러온다.
        // silentRemote=false 이면 원격 재조회 없이 이미 메모리에 있는 dbData만으로 다시 그린다.
        async function refreshHomeTab(silentRemote = true) {
            const statusEl = document.getElementById('homeStatusMsg');
            if (!getGSheetUrlSilent()) {
                if (statusEl) { statusEl.innerText = '⚠️ 상단 "구글시트 연동 설정"에서 웹앱 URL을 먼저 입력해주세요.'; statusEl.style.color = '#b45309'; }
                renderHomeDeliveryTable();
                return;
            }
            if (silentRemote) {
                if (statusEl) { statusEl.innerText = '⏳ 구글시트에서 불러오는 중...'; statusEl.style.color = '#64748b'; }
                try {
                    await loadFromGoogleSheetSilent();
                    await loadInventoryStatus();
                    if (statusEl) { statusEl.innerText = `✅ ${new Date().toLocaleTimeString('ko-KR')} 기준으로 갱신됨`; statusEl.style.color = '#16a34a'; }
                } catch (err) {
                    if (statusEl) { statusEl.innerText = '❌ 불러오기 실패: ' + err.message; statusEl.style.color = '#ef4444'; }
                }
            }
            populateHomeFilterOptions();
            renderHomeDeliveryTable();
            renderHomeExpiryCards();
        }

        // loadFromGoogleSheet()와 동일 로직이지만 alert 없이 조용히 처리(홈 자동새로고침용)
        async function loadFromGoogleSheetSilent() {
            const url = getGSheetUrlSilent(); if (!url) return;
            const res = await fetch(url + (url.includes('?') ? '&' : '?') + 'action=read&sheet=' + encodeURIComponent(HISTORY_SHEET_NAME));
            const rawText = await res.text();
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            const json = JSON.parse(rawText);
            const rowsArr = extractRowsArray(json);
            if (!rowsArr) throw new Error('데이터 형식을 인식하지 못했습니다.');
            ingestRowsSilent(rowsArr);
        }

        // ingestRows()와 같지만 alert 없이 조용히 반영(홈 자동새로고침용)
        function ingestRowsSilent(json) {
            if (!json || json.length === 0) return;
            let hIdx = 0, maxScore = 0;
            for(let i=0; i<Math.min(json.length, 10); i++) {
                const rowStr = (json[i]||[]).join(''); let score = 0;
                if(rowStr.includes('품명')) score++; if(rowStr.includes('예정')) score++; if(rowStr.includes('확정')) score++;
                if(score > maxScore) { maxScore = score; hIdx = i; }
            }
            const headers = json[hIdx].map(h => String(h).replace(/\s/g, ''));
            const map = {
                uid: headers.findIndex(h => h.includes('고유ID') || h.includes('고유')),
                delDate: headers.findIndex(h => h.includes('납품')),
                mealDate: headers.findIndex(h => h.includes('급식')),
                date: headers.findIndex(h => h.includes('일자')),
                name: headers.findIndex(h => h.includes('품명')),
                cat: headers.findIndex(h => h.includes('구분')),
                unit: headers.findIndex(h => h.includes('부대')),
                plan: pickQtyColumn(headers, '예정'),
                conf: pickQtyColumn(headers, '확정'),
                price: headers.findIndex(h => h.includes('단가')),
                planSales: headers.findIndex(h => h.includes('예정매출')),
                confSales: headers.findIndex(h => h.includes('확정매출'))
            };
            const existingKeys = new Set(dbData.map(d => `${d.delDate}|${d.mealDate}|${d.unit}|${d.name}|${d.plan}|${d.conf}`));
            const newList = [];
            for(let i = hIdx + 1; i < json.length; i++) {
                const row = json[i]; if(!row || !row[map.name]) continue;
                const uidCell = map.uid !== -1 ? row[map.uid] : '';
                const isMarkedExcluded = String(uidCell || '').trim() === '제외';
                const planRaw = map.plan !== -1 ? row[map.plan] : undefined;
                const confRaw = map.conf !== -1 ? row[map.conf] : undefined;
                const planMissing = (planRaw === undefined || planRaw === null || planRaw === '' || isNaN(parseFloat(planRaw)) || parseFloat(planRaw) === 0);
                const confMissing = (confRaw === undefined || confRaw === null || confRaw === '');
                const p = planMissing ? null : Math.round(parseFloat(planRaw));
                const c = confMissing ? null : Math.round(parseFloat(confRaw) || 0);
                if(p === null && c === null) continue;

                let dDate = map.delDate !== -1 ? parseExcelDate(row[map.delDate]) : (map.date !== -1 ? parseExcelDate(row[map.date]) : '-');
                let mDate = map.mealDate !== -1 ? parseExcelDate(row[map.mealDate]) : (map.date !== -1 ? parseExcelDate(row[map.date]) : '-');
                const reconstructedDDate = reconstructKoreanPartialDate(row[map.delDate], mDate);
                if(reconstructedDDate) dDate = reconstructedDDate;

                const dayCat = classifyDay(mDate);
                const isHoli = (dayCat === 'WORK' || dayCat === 'SAT') ? false : (dayCat === 'HOLI') ? true : null;
                const delDayCat = classifyDay(dDate);
                const isDelHoli = (delDayCat === 'WORK' || delDayCat === 'SAT') ? false : (delDayCat === 'HOLI') ? true : null;
                const holidayNameVal = (dayCat === 'HOLI') ? (window.holidayMap[toIsoDate(mDate)] || '') : '';

                const nameVal = String(row[map.name]).trim();
                const catVal = map.cat!==-1 ? String(row[map.cat]).trim().replace(/\s+/g,' ') : '기타';
                const unitValRaw = map.unit!==-1 ? String(row[map.unit]).trim() : '공통';
                const unitVal = normalizeUnitName(unitValRaw);

                const key = `${dDate}|${mDate}|${unitVal}|${nameVal}|${p}|${c}`;
                if(existingKeys.has(key)) continue;
                existingKeys.add(key);

                const priceVal = map.price !== -1 && row[map.price] !== undefined && row[map.price] !== '' ? parseFloat(String(row[map.price]).replace(/,/g,'')) : null;
                const planSalesVal = map.planSales !== -1 && row[map.planSales] !== undefined && row[map.planSales] !== '' ? parseFloat(String(row[map.planSales]).replace(/,/g,'')) : null;
                const confSalesVal = map.confSales !== -1 && row[map.confSales] !== undefined && row[map.confSales] !== '' ? parseFloat(String(row[map.confSales]).replace(/,/g,'')) : null;

                newList.push({
                    id: key,
                    delDate: dDate, mealDate: mDate, isHoliday: isHoli, dayType: dayCat,
                    isDelHoliday: isDelHoli, delDayType: delDayCat,
                    dayCategory: dayCat, delDayCategory: delDayCat, holidayName: holidayNameVal,
                    name: nameVal, cat: catVal, unit: unitVal,
                    plan: p, conf: c,
                    price: (priceVal !== null && !isNaN(priceVal)) ? priceVal : null,
                    planSales: (planSalesVal !== null && !isNaN(planSalesVal)) ? planSalesVal : null,
                    confSales: (confSalesVal !== null && !isNaN(confSalesVal)) ? confSalesVal : null,
                    rate: (p !== null && c !== null) ? ((c - p) / p) * 100 : null,
                    excludedFromStats: isMarkedExcluded
                });
            }
            if (newList.length > 0) {
                dbData = dbData.concat(newList);
                saveToDB('dbData', dbData);
                refreshAll();
            }
        }

        function populateHomeFilterOptions() {
            const cats = [...new Set(dbData.map(d => d.cat))].sort();
            const units = [...new Set(dbData.map(d => d.unit))].sort();
            const fillSel = (id, arr, label) => {
                const el = document.getElementById(id);
                const cur = el.value;
                el.innerHTML = `<option value="ALL">${label} 전체</option>`;
                arr.forEach(v => el.add(new Option(v, v)));
                if ([...el.options].some(o => o.value === cur)) el.value = cur;
            };
            fillSel('homeFilterCat', cats, '구분');
            fillSel('homeFilterUnit', units, '부대');
        }

        // [신규] 당월 납품계획 표 렌더링 — 엑셀 출고계획 형식(납품일/구분/품명/부대/예정/확정/매출)
        function renderHomeDeliveryTable() {
            const tbody = document.getElementById('homeDeliveryBody');
            if (!tbody) return;
            if (dbData.length === 0) { tbody.innerHTML = '<tr><td colspan="10">새로고침을 눌러 불러오세요.</td></tr>'; return; }

            const monthVal = document.getElementById('homeMonth').value; // 'yyyy-mm'
            const basis = document.getElementById('homeDateBasis').value; // 'del' | 'meal'
            const cat = document.getElementById('homeFilterCat').value;
            const unit = document.getElementById('homeFilterUnit').value;
            const nameKw = (document.getElementById('homeFilterName').value || '').trim().toLowerCase();

            let filtered = dbData.filter(d => {
                const dateField = basis === 'meal' ? d.mealDate : d.delDate;
                const passMonth = !monthVal || (dateField && String(dateField).startsWith(monthVal));
                const passCat = cat === 'ALL' || d.cat === cat;
                const passUnit = unit === 'ALL' || d.unit === unit;
                const passName = !nameKw || (d.name || '').toLowerCase().includes(nameKw);
                return passMonth && passCat && passUnit && passName;
            });

            filtered.sort((a, b) => {
                const da = basis === 'meal' ? a.mealDate : a.delDate;
                const db_ = basis === 'meal' ? b.mealDate : b.delDate;
                return String(da).localeCompare(String(db_));
            });

            if (filtered.length === 0) { tbody.innerHTML = '<tr><td colspan="10">해당 조건의 납품계획이 없습니다.</td></tr>'; window.__homeTableRows = []; return; }

            const shown = filtered.slice(0, 500);
            window.__homeTableRows = shown; // [신규] 확정량 입력칸이 참조할 원본 레코드 배열(인덱스로 매칭)

            tbody.innerHTML = shown.map((d, idx) => {
                const rateTxt = d.rate === null ? '<span style="color:#94a3b8;">-</span>' : `${d.rate>0?'+':''}${Math.round(d.rate)}%`;
                const rateColor = d.rate === null ? '#94a3b8' : (d.rate>0?'#ea580c':'#2563eb');
                const planTxt = d.plan === null ? '<span style="color:#94a3b8;">-</span>' : Math.round(d.plan).toLocaleString();
                const planSalesTxt = d.planSales !== null && d.planSales !== undefined ? Math.round(d.planSales).toLocaleString() : '-';
                const confSalesTxt = d.confSales !== null && d.confSales !== undefined ? Math.round(d.confSales).toLocaleString() : '-';
                const confVal = (d.conf !== null && d.conf !== undefined) ? d.conf : '';
                return `<tr>
                    <td data-label="납품일">${d.delDate || '-'}</td>
                    <td data-label="급식일">${d.mealDate || '-'}</td>
                    <td data-label="구분">${d.cat || '-'}</td>
                    <td data-label="품명" style="font-weight:600;">${d.name}</td>
                    <td data-label="부대">${d.unit || '-'}</td>
                    <td data-label="예정량(kg)">${planTxt}</td>
                    <td data-label="확정량(kg)"><input type="number" step="any" class="home-conf-input" data-idx="${idx}" value="${confVal}" placeholder="미확정" style="width:82px; padding:4px; text-align:right; border:1px solid #cbd5e1; border-radius:4px;" onchange="markHomeConfDirty(this)"></td>
                    <td data-label="오차율" style="font-weight:bold; color:${rateColor};">${rateTxt}</td>
                    <td data-label="매출계획">${planSalesTxt}</td>
                    <td data-label="매출">${confSalesTxt}</td>
                </tr>`;
            }).join('') + (filtered.length > 500 ? `<tr><td colspan="10" style="color:#94a3b8;">…외 ${(filtered.length-500).toLocaleString()}건 더 있음 (필터로 좁혀보세요)</td></tr>` : '');
        }

        // [신규] 확정량 칸을 수정하면 시각적으로 "변경됨" 표시(저장 전까지)
        function markHomeConfDirty(inputEl) {
            inputEl.style.background = '#fef9c3';
            inputEl.style.borderColor = '#f59e0b';
            inputEl.dataset.dirty = '1';
        }

        // [신규] 홈 탭에서 직접 입력한 확정량을 구글시트 '데이터' 탭에 반영.
        // ※ 이 기능이 동작하려면 Apps Script 웹앱에 action=updateConfirmed 처리가 추가돼 있어야 합니다.
        async function saveHomeConfirmedQuantities() {
            const url = getGSheetUrl(); if (!url) return;
            const rows = window.__homeTableRows || [];
            const inputs = document.querySelectorAll('.home-conf-input[data-dirty="1"]');
            if (inputs.length === 0) { alert('변경된 확정량이 없습니다. 칸에 값을 입력한 뒤 다시 눌러주세요.'); return; }

            const updates = [];
            inputs.forEach(inp => {
                const idx = Number(inp.dataset.idx);
                const rec = rows[idx];
                if (!rec) return;
                const raw = inp.value.trim();
                const newVal = raw === '' ? null : parseFloat(raw);
                if (raw !== '' && isNaN(newVal)) return;
                const packInfo = window.productInfoMap && window.productInfoMap[rec.name];
                const confEa = (newVal !== null && packInfo && packInfo.pack) ? Math.round(newVal / packInfo.pack) : '';
                updates.push({ rec, inp, delDate: rec.delDate, mealDate: rec.mealDate, unit: rec.unit, name: rec.name, confQtyKg: newVal === null ? '' : newVal, confQty: confEa });
            });
            if (updates.length === 0) { alert('반영할 값이 없습니다.'); return; }
            if (!confirm(`${updates.length}건의 확정량을 구글시트 '${HISTORY_SHEET_NAME}' 탭에 반영할까요?`)) return;

            try {
                const res = await fetch(url, {
                    method: 'POST',
                    headers: {'Content-Type': 'text/plain;charset=utf-8'},
                    body: JSON.stringify({
                        action: 'updateConfirmed',
                        sheet: HISTORY_SHEET_NAME,
                        updates: updates.map(u => ({ delDate: u.delDate, mealDate: u.mealDate, unit: u.unit, name: u.name, confQty: u.confQty, confQtyKg: u.confQtyKg }))
                    })
                });
                if (!res.ok) throw new Error(`HTTP ${res.status}`);
                const result = await res.json();
                if (result.error) throw new Error(result.error);

                updates.forEach(u => {
                    u.rec.conf = (u.confQtyKg === '' ? null : u.confQtyKg);
                    u.rec.rate = (u.rec.plan !== null && u.rec.plan !== 0 && u.rec.conf !== null) ? ((u.rec.conf - u.rec.plan) / u.rec.plan) * 100 : null;
                    u.inp.style.background = ''; u.inp.style.borderColor = '#cbd5e1'; delete u.inp.dataset.dirty;
                });
                saveToDB('dbData', dbData, function() { renderHomeDeliveryTable(); updateDashboard(); });

                const notMatched = (result.results || []).filter(r => !r.updated);
                let msg = `✅ 확정량 ${updates.length}건 저장 완료!`;
                if (notMatched.length > 0) {
                    msg += `\n\n⚠️ 시트에서 일치하는 행을 찾지 못한 항목 ${notMatched.length}건(구글시트 원본이 이미 바뀌었을 수 있어요, 새로고침 후 확인해주세요):\n- ` + notMatched.map(r => `${r.name}(${r.unit}, ${r.delDate})`).join('\n- ');
                }
                alert(msg);
            } catch (err) {
                alert("확정량 저장에 실패했습니다: " + err.message + "\n\n※ 이 기능은 Apps Script 웹앱에 'updateConfirmed' 액션이 추가되어 있어야 동작합니다. 아직 추가하지 않으셨다면 안내드린 스크립트를 배포에 넣어주세요.");
            }
        }

        // [신규] 홈 탭 - 납품기한임박/경과 카드 렌더링 (모바일 최적화 카드형)
        function renderHomeExpiryCards() {
            const upEl = document.getElementById('homeExpiryUpcoming');
            const pastEl = document.getElementById('homeExpiryPast');
            if (!upEl || !pastEl) return;

            const upList = window.__expiryUpcomingList || [];
            const pastList = window.__expiryPastList || [];

            upEl.innerHTML = upList.length === 0
                ? '<div style="color:#94a3b8; font-size:12px;">표시할 품목이 없습니다.</div>'
                : upList.map(p => `
                    <div class="mobile-card upcoming">
                        <div>
                            <div class="mc-name">${p.name}</div>
                            <div class="mc-sub">마지막 납품: ${p.lastDelivery || '-'}</div>
                        </div>
                        <div class="mc-right">
                            <div style="font-weight:bold; color:#ea580c;">${p.deadline || '-'}</div>
                            <div class="mc-sub">재고 ${p.stock !== null ? p.stock.toLocaleString() : '기록없음'}</div>
                        </div>
                    </div>`).join('');

            pastEl.innerHTML = pastList.length === 0
                ? '<div style="color:#94a3b8; font-size:12px;">표시할 품목이 없습니다.</div>'
                : pastList.map(p => `
                    <div class="mobile-card past">
                        <div>
                            <div class="mc-name">${p.name}</div>
                            <div class="mc-sub">수량 ${p.qty !== null ? p.qty.toLocaleString() : '-'} · 단가 ${p.unitPrice !== null ? p.unitPrice.toLocaleString() : '-'}</div>
                        </div>
                        <div class="mc-right">
                            <div style="font-weight:bold; color:#dc2626;">${p.expiry || '-'}</div>
                            <div class="mc-sub">합계 ${p.total !== null ? p.total.toLocaleString() : '-'}</div>
                        </div>
                    </div>`).join('');
        }

        async function loadFromGoogleSheet() {
            const url = getGSheetUrl(); if(!url) return;
            const historySheet = HISTORY_SHEET_NAME;
            let rawText = '';
            try {
                const res = await fetch(url + (url.includes('?') ? '&' : '?') + 'action=read&sheet=' + encodeURIComponent(historySheet));
                rawText = await res.text();
                if(!res.ok) throw new Error(`HTTP ${res.status} — 응답: ${rawText.slice(0,300)}`);

                let json;
                try {
                    json = JSON.parse(rawText);
                } catch(parseErr) {
                    throw new Error(`응답이 JSON이 아닙니다 (배포 권한 문제 가능성 높음). 응답 시작 부분: ${rawText.slice(0,200)}`);
                }

                const rowsArr = extractRowsArray(json);
                if (!rowsArr) {
                    console.warn('구글시트 응답(원본):', json);
                    throw new Error(`응답은 받았지만 표 형태의 배열을 찾지 못했습니다. 실제 응답 구조: ${JSON.stringify(json).slice(0,300)}`);
                }
                console.log(`구글시트에서 ${rowsArr.length}행 수신, 헤더행 포함 미리보기:`, rowsArr.slice(0,3));
                ingestRows(rowsArr, '구글시트 불러오기');
            } catch(err) {
                console.error('구글시트 불러오기 실패:', err, rawText);
                alert("구글시트를 불러오지 못했습니다: " + err.message + "\n\n확인할 것:\n1) 웹앱 URL이 최신 배포 URL인지 (배포를 새로 할 때마다 URL이 바뀔 수 있음)\n2) 배포 시 '액세스 권한'이 '전체' 또는 '링크가 있는 모든 사용자'인지\n3) '" + historySheet + "' 라는 이름의 시트가 실제로 존재하는지 (탭 이름 오타 확인)\n4) 해당 시트에 예정량·확정량이 채워진 데이터가 실제로 있는지\n\n브라우저 개발자도구(F12) > Console 탭에 원본 응답이 로그로 남습니다.");
            }
        }

        function makeUniqueId(seq) {
            return `WEB${Date.now()}${String(seq).padStart(2,'0')}`;
        }

        async function ensureUnitProductInfoLoaded() {
            if (Object.keys(window.unitInfoMap || {}).length > 0) return;
            await syncUnitProductInfo();
        }

        async function saveGridToGoogleSheet() {
            const url = getGSheetUrl(); if(!url) return;
            await ensureUnitProductInfoLoaded();
            const writeSheet = WRITE_SHEET_NAME;
            const dateBasis = getDateBasis();
            const mode = document.querySelector('input[name="calcMode"]:checked').value;
            const rowsToSend = [];
            const missingPatternRows = [];
            let seq = 0;
            for(let tr of document.getElementById('excelGridBody').children) {
                const name = tr.querySelector('.grid-name').value.trim(); if(!name) continue;
                const dateVal = tr.querySelector('.grid-date').value;
                const unitVal = tr.querySelector('.grid-unit-name').value.trim();
                const catVal = tr.querySelector('.grid-cat').value.trim();
                const unitType = tr.querySelector('.grid-unit').value;
                const globalHeadcountVal = parseFloat(document.getElementById('globalHeadcount')?.value) || 0;
                let req = mode === 'total' ? parseFloat(tr.querySelector('.grid-req-tot').value) || 0 : (parseFloat(tr.querySelector('.grid-req-per').value)||0) * globalHeadcountVal;
                if(req <= 0) continue;

                const basis = tr.querySelector('.grid-save-basis').value;
                let qty = basis === 'min' ? tr.dataset.minVal : basis === 'max' ? tr.dataset.maxVal : tr.dataset.recVal;
                if (qty === undefined) { calculateMultiRows(); qty = basis === 'min' ? tr.dataset.minVal : basis === 'max' ? tr.dataset.maxVal : tr.dataset.recVal; }
                qty = Number(qty) || 0;
                if (qty <= 0) continue;

                seq++;
                const id = makeUniqueId(seq);
                const counterpartDate = (unitVal && dateVal) ? (computeCounterpartDate(unitVal, dateVal, dateBasis) || '') : '';
                let delDate = dateBasis === 'delivery' ? dateVal : counterpartDate;
                let mealDate = dateBasis === 'meal' ? dateVal : counterpartDate;
                const missingPattern = unitVal && !counterpartDate;

                const packInfo = window.productInfoMap && window.productInfoMap[name];
                const qtyKg = (packInfo && packInfo.pack) ? Math.round(qty * packInfo.pack * 100) / 100 : '';

                rowsToSend.push({ values: [id, delDate, catVal, name, qty, qtyKg, '', '', mealDate, unitVal] });
                if (missingPattern) missingPatternRows.push(`${unitVal || '(부대미입력)'} - ${name}`);
            }
            if(rowsToSend.length === 0) { alert('전송할 행이 없습니다. 품명과 소요량을 입력해주세요.'); return; }
            let confirmMsg = `구글시트 '${writeSheet}' 시트에 ${rowsToSend.length}건을 새 행으로 추가합니다.\n(각 행은 지정한 저장기준: 최소/권장/최대 중 선택된 값으로 예정량이 채워집니다)`;
            if (missingPatternRows.length > 0) {
                confirmMsg += `\n\n⚠️ 아래 ${missingPatternRows.length}건은 부대별정보에 납품요일 패턴이 없어 납품일을 계산하지 못했습니다(빈 칸으로 저장됨):\n- ` + missingPatternRows.join('\n- ');
            }
            confirmMsg += `\n계속할까요?`;
            if(!confirm(confirmMsg)) return;
            try {
                const res = await fetch(url, {
                    method: 'POST',
                    headers: {'Content-Type': 'text/plain;charset=utf-8'},
                    body: JSON.stringify({ action: 'append', sheet: writeSheet, rows: rowsToSend })
                });
                if(!res.ok) throw new Error(`HTTP ${res.status}`);
                const result = await res.json();
                if (result.error) throw new Error(result.error);
                alert(`✅ 구글시트 '${writeSheet}'에 ${result.added ?? rowsToSend.length}건 저장 완료!`);
            } catch(err) {
                alert("구글시트 저장에 실패했습니다: " + err.message);
            }
        }

        async function saveGridToInventoryLedger() {
            const url = getGSheetUrl(); if(!url) return;
            await ensureUnitProductInfoLoaded();
            const ledgerSheetHint = document.getElementById('gsheetLedgerSheet').value.trim();
            const dateBasis = getDateBasis();
            const mode = document.querySelector('input[name="calcMode"]:checked').value;
            const leadDays = parseInt(document.getElementById('ledgerLeadDays').value, 10) || 0;
            const rowsToSend = [];
            const missingDeliveryRows = [];

            for (let tr of document.getElementById('excelGridBody').children) {
                const name = tr.querySelector('.grid-name').value.trim(); if(!name) continue;
                const dateVal = tr.querySelector('.grid-date').value;
                const unitName = tr.querySelector('.grid-unit-name').value.trim();
                const globalHeadcountVal = parseFloat(document.getElementById('globalHeadcount')?.value) || 0;
                let req = mode === 'total' ? parseFloat(tr.querySelector('.grid-req-tot').value) || 0 : (parseFloat(tr.querySelector('.grid-req-per').value)||0) * globalHeadcountVal;
                if (req <= 0) continue;

                const basis = tr.querySelector('.grid-save-basis').value;
                let qty = basis === 'min' ? tr.dataset.minVal : basis === 'max' ? tr.dataset.maxVal : tr.dataset.recVal;
                if (qty === undefined) { calculateMultiRows(); qty = basis === 'min' ? tr.dataset.minVal : basis === 'max' ? tr.dataset.maxVal : tr.dataset.recVal; }
                qty = Number(qty) || 0;
                if (qty <= 0) continue;

                const deliveryDateForCalc = dateBasis === 'delivery'
                    ? dateVal
                    : ((unitName && dateVal) ? computeCounterpartDate(unitName, dateVal, dateBasis) : null);

                let incomingDate = '';
                if (deliveryDateForCalc) {
                    const d = new Date(deliveryDateForCalc);
                    if (!isNaN(d.getTime())) {
                        d.setDate(d.getDate() - leadDays);
                        incomingDate = d.toISOString().split('T')[0];
                    }
                }
                if (!incomingDate) missingDeliveryRows.push(`${unitName || '(부대미입력)'} - ${name}`);

                rowsToSend.push({ name, date: incomingDate, deliveryDate: deliveryDateForCalc || '', qty, unit: unitName });
            }

            if (rowsToSend.length === 0) { alert('전송할 행이 없습니다. 품명과 소요량을 입력해주세요.'); return; }
            const targetDesc = ledgerSheetHint ? `'${ledgerSheetHint}' 탭 우선, 없으면 다른 탭 자동 탐색` : '전체 탭 자동 탐색';
            let confirmMsg = `재고수불대장(${targetDesc})에 ${rowsToSend.length}건을 예정 입고로 등록합니다.\n(A업체·B제품 자동 매칭, C입고일=납품일-리드타임 ${leadDays}일, D수량=선택한 저장기준 값, E납품일도 함께 채워집니다. F소계(확정량)는 실제 납품 후 직접 입력하는 칸이라 비워둡니다.)`;
            if (missingDeliveryRows.length > 0) {
                confirmMsg += `\n\n⚠️ 아래 ${missingDeliveryRows.length}건은 납품일을 계산하지 못해 입고일이 빈 채로 등록됩니다(부대별정보에 해당 부대 요일 패턴이 없음 — 직접 채워주세요):\n- ` + missingDeliveryRows.join('\n- ');
            }
            confirmMsg += `\n계속할까요?`;
            if (!confirm(confirmMsg)) return;

            try {
                const res = await fetch(url, {
                    method: 'POST',
                    headers: {'Content-Type': 'text/plain;charset=utf-8'},
                    body: JSON.stringify({ action: 'insertInventoryOrder', sheet: ledgerSheetHint, rows: rowsToSend })
                });
                if(!res.ok) throw new Error(`HTTP ${res.status}`);
                const result = await res.json();
                if (result.error) throw new Error(result.error);

                const items = result.results || [];
                const okItems = items.filter(r => r.ok);
                const newProducts = items.filter(r => r.ok && r.needsReview).map(r => r.name);
                const failed = items.filter(r => !r.ok);

                let msg = `✅ 재고수불대장에 ${okItems.length}건 등록 완료!`;
                if (okItems.length > 0) {
                    msg += `\n` + okItems.map(r => `- ${r.name} → [${r.sheet}] 탭`).join('\n');
                }
                if (newProducts.length > 0) {
                    msg += `\n\n⚠️ 어느 탭에도 없던 신규 품목이라 업체(A열)를 자동으로 채우지 못했습니다. 직접 확인해주세요:\n- ` + newProducts.join('\n- ');
                }
                if (failed.length > 0) {
                    msg += `\n\n❌ 등록 실패:\n- ` + failed.map(r => `${r.name || '(품명없음)'}: ${r.reason}`).join('\n- ');
                }
                alert(msg);
                if (window.__inventoryLoadedOnce) loadInventoryStatus();
            } catch(err) {
                alert("재고수불대장 등록에 실패했습니다: " + err.message);
            }
        }

        function reconcileSplitPlanGroups() {
            const groups = {};
            dbData.forEach(d => {
                const key = `${d.delDate}|${d.unit}|${d.name}`;
                if (!groups[key]) groups[key] = [];
                groups[key].push(d);
            });
            Object.values(groups).forEach(group => {
                if (group.length < 2) return;
                const withPlan = group.filter(d => d.plan !== null);
                const withConf = group.filter(d => d.conf !== null);
                if (withPlan.length < 2 || withConf.length === 0 || withConf.length === withPlan.length) return;
                const totalPlan = withPlan.reduce((a,d) => a + d.plan, 0);
                const totalConf = withConf.reduce((a,d) => a + d.conf, 0);
                if (totalPlan <= 0) return;
                const aggregateRate = ((totalConf - totalPlan) / totalPlan) * 100;
                withConf.forEach(d => { d.rate = aggregateRate; d.splitGroupAdjusted = true; d.splitGroupTotalPlan = totalPlan; });
            });
        }

        function refreshAll() {
            if(dbData.length===0) return;
            reconcileSplitPlanGroups();
            const units = [...new Set(dbData.map(d=>d.unit))].sort(); const cats = [...new Set(dbData.map(d=>d.cat))].sort();
            const fillSel = (id, arr, label) => { const el = document.getElementById(id); el.innerHTML = `<option value="ALL">${label} 전체</option>`; arr.forEach(v => el.add(new Option(v,v))); };
            fillSel('dashUnit', units, '부대'); fillSel('logFilterUnit', units, '부대'); fillSel('logFilterCat', cats, '구분');
            const dl = document.getElementById('unitNameList');
            if(dl) { dl.innerHTML = ''; units.forEach(u => dl.appendChild(new Option(u, u))); }
            updateDashboard();
            populateHomeFilterOptions();
        }

        function updateDashboard() {
            if(dbData.length===0) return;
            const targetUnit = document.getElementById('dashUnit').value;
            const dashData = targetUnit === 'ALL' ? dbData : dbData.filter(d => d.unit === targetUnit);
            
            const dashRated = dashData.filter(d => d.rate !== null && !d.excludedFromStats);
            document.getElementById('statTotalCount').innerText = dbData.length.toLocaleString() + " 건";
            document.getElementById('statAvgError').innerText = dashRated.length > 0 ? Math.round(dashRated.reduce((a,c)=>a+Math.abs(c.rate),0)/dashRated.length) + " %" : "0 %";
            
            const cats = [...new Set(dashRated.map(d=>d.cat))].sort();
            const volDiv = document.getElementById('volatilityList');
            let html = ''; let maxVol = -1; let maxCat = '-';

            if(dashRated.length === 0) {
                volDiv.innerHTML = '<div style="padding:10px; color:#64748b;">해당 부대의 오차율 계산 가능 데이터가 없습니다.</div>';
                document.getElementById('statMaxCat').innerText = "-";
            } else {
                cats.forEach(c => {
                    const rates = dashRated.filter(d => d.cat === c).map(d => d.rate); if(rates.length === 0) return;
                    const mean = rates.reduce((a,b)=>a+b,0) / rates.length;
                    const std = Math.sqrt(rates.reduce((a,b)=>a+Math.pow(b-mean,2),0) / rates.length);
                    if(std > maxVol) { maxVol = std; maxCat = c; }
                    let colorClass = std >= 20 ? 'std-danger' : (std >= 10 ? 'std-warn' : 'std-safe');
                    html += `<div class="vol-item"><div class="cat-name">${c}</div><div class="std-val ${colorClass}">${std.toFixed(1)}</div></div>`;
                });
                volDiv.innerHTML = html; document.getElementById('statMaxCat').innerText = `${maxCat} (${Math.round(maxVol)} 변동)`;
            }
            renderLogTable(); renderChart(dashRated); calculateMultiRows();
        }

        const SAFETY_STOCK_REC_Z = 1.28;
        const SAFETY_STOCK_MAX_Z = 2.33;
        function autoSafetyZ(sampleSize, baseZ) {
            if (sampleSize <= 0) return baseZ;
            return baseZ * Math.sqrt(1 + 1 / sampleSize);
        }

        function calculateMultiRows() {
            const trs = document.getElementById('excelGridBody').children;
            const mode = document.querySelector('input[name="calcMode"]:checked').value;
            const dateBasis = getDateBasis();
            const dayCatField = (dateBasis === 'delivery') ? 'delDayCategory' : 'dayCategory';
            const holidayField = (dateBasis === 'delivery') ? 'isDelHoliday' : 'isHoliday';

            for (let tr of trs) {
                let rowDateInput = tr.querySelector('.grid-date');
                let dateStr = rowDateInput.value;
                
                let rowDayCategory = 'EMPTY';
                const rowBadge = tr.querySelector('.row-day-badge');
                
                if(dateStr) { 
                    rowDayCategory = classifyDay(dateStr);
                    rowBadge.style.display = 'inline-block';
                    const badgeClassMap = { WORK: 'day-work', SAT: 'day-sat', HOLI: 'day-holi' };
                    const badgeTextMap = { WORK: '평', SAT: '토', HOLI: '휴' };
                    rowBadge.className = 'row-day-badge day-badge ' + (badgeClassMap[rowDayCategory] || 'day-work');
                    rowBadge.innerText = badgeTextMap[rowDayCategory] || '?';
                    const iso = toIsoDate(dateStr);
                    rowBadge.title = (rowDayCategory === 'HOLI' && iso && window.holidayMap[iso]) ? window.holidayMap[iso] : '';
                    const gridYear = dateStr.slice(0,4);
                    if (document.getElementById('gsheetWebAppUrl').value.trim() && !window.__holidayYearsLoaded.has(gridYear) && !window.__holidayYearsFailed.has(gridYear)) {
                        syncHolidays([gridYear]).then(result => {
                            if (result.ok) calculateMultiRows();
                            else window.__holidayYearsFailed.add(gridYear);
                        });
                    }
                } else {
                    rowBadge.style.display = 'none';
                }

                const name = tr.querySelector('.grid-name').value.trim();
                const targetUnit = (tr.querySelector('.grid-unit-name')?.value || '').trim();
                let cat = tr.querySelector('.grid-cat').value.trim();
                const spec = parseFloat(tr.querySelector('.grid-spec').value) || 1.0; 
                const unitType = tr.querySelector('.grid-unit').value;
                
                const globalHeadcountVal = parseFloat(document.getElementById('globalHeadcount')?.value) || 0;
                let req = mode === 'total' ? parseFloat(tr.querySelector('.grid-req-tot').value) || 0 : (parseFloat(tr.querySelector('.grid-req-per').value)||0) * globalHeadcountVal;
                if (!name || req <= 0) {
                    tr.querySelector('.res-min').innerText = '-'; tr.querySelector('.res-rec').innerText = '-'; tr.querySelector('.res-max').innerText = '-';
                    delete tr.dataset.minVal; delete tr.dataset.recVal; delete tr.dataset.maxVal;
                    continue;
                }

                specMemory[name] = spec; unitMemory[name] = unitType;
                const calcDivisor = (unitType === 'ea') ? 1.0 : spec;

                let filtered = targetUnit ? dbData.filter(d => d.unit === targetUnit && d.name === name) : dbData.filter(d => d.name === name);
                let usedCategoryFallback = false;

                if (filtered.length === 0 && targetUnit) {
                    if (!cat && window.productInfoMap && window.productInfoMap[name]) {
                        cat = window.productInfoMap[name].category || '';
                        if (cat) tr.querySelector('.grid-cat').value = cat;
                    }
                    if (cat) {
                        const catFallback = dbData.filter(d => d.unit === targetUnit && d.cat === cat);
                        if (catFallback.length > 0) {
                            filtered = catFallback;
                            usedCategoryFallback = true;
                        }
                    }
                }

                if (dateStr && filtered.length > 0) {
                    const dayFiltered = filtered.filter(d =>
                        d[dayCatField] !== undefined ? d[dayCatField] === rowDayCategory
                        : d[holidayField] === (rowDayCategory === 'SAT' || rowDayCategory === 'HOLI')
                    );
                    if (dayFiltered.length > 0) {
                        filtered = dayFiltered;
                    }
                }

                if (!cat && filtered.length > 0) { cat = filtered[0].cat; tr.querySelector('.grid-cat').value = cat; }

                let rated = filtered.filter(d => d.rate !== null && !d.excludedFromStats);

                let mean = 0, itemStd = 0, maxRate = 0;
                let maxRateSourceRecord = null;
                if (rated.length > 0) {
                    const rates = rated.map(d => d.rate);
                    mean = rates.reduce((a,b)=>a+b,0) / rates.length;
                    itemStd = Math.sqrt(rates.reduce((a,b) => a + Math.pow(b-mean,2), 0) / rates.length) || 0;
                    maxRate = Math.max(...rates, 0);
                    if (maxRate > 0) maxRateSourceRecord = rated.find(d => d.rate === maxRate) || null;
                }

                let appliedStd = itemStd;
                let sampleSizeForZ = rated.length;

                const recZ = autoSafetyZ(sampleSizeForZ, SAFETY_STOCK_REC_Z);
                const maxZ = autoSafetyZ(sampleSizeForZ, SAFETY_STOCK_MAX_Z);

                const minRatePct = mean;
                const recRatePct = mean + (recZ * appliedStd);
                const maxRateStatPct = mean + (maxZ * appliedStd);

                let finalMinPct = minRatePct;
                let finalRecPct = Math.max(finalMinPct, recRatePct);
                let finalMaxPct = Math.max(finalRecPct, maxRateStatPct, maxRate);

                const isExtreme = maxRate >= 200 || maxRate <= -80;
                const minCell = tr.querySelector('.res-min');
                const recCell = tr.querySelector('.res-rec');
                const maxCell = tr.querySelector('.res-max');

                const minVal = Math.round((req * (1 + (finalMinPct/100))) / calcDivisor);
                const recVal = Math.round((req * (1 + (finalRecPct/100))) / calcDivisor);
                const maxVal = Math.round((req * (1 + (finalMaxPct/100))) / calcDivisor);

                tr.dataset.minVal = minVal; tr.dataset.recVal = recVal; tr.dataset.maxVal = maxVal;

                const fallbackBadge = usedCategoryFallback ? `<span style="font-size:9px; color:#a16207; display:block; font-weight:normal;">📋 "${cat}" 구분 대체</span>` : '';
                const fallbackNote = usedCategoryFallback ? `\n📋 "${name}" 자체 이력이 없어, ${targetUnit}의 "${cat}" 구분 전체 이력(${rated.length}건)으로 대체 산출됨. (다른 부대 데이터는 사용하지 않음)` : '';

                minCell.innerHTML = fallbackBadge + minVal.toLocaleString();
                minCell.title = `과거 평균 오차율(${mean.toFixed(1)}%)만 반영한 최소 준비량 (기준 실적 ${rated.length}건)` + fallbackNote;

                recCell.innerHTML = fallbackBadge + recVal.toLocaleString();
                recCell.title = `평균(${mean.toFixed(1)}%) + 완만한 안전마진(z≈${recZ.toFixed(2)}, 표본 ${sampleSizeForZ}건 기준)` + fallbackNote;

                const sourceInfoText = maxRateSourceRecord
                    ? `\n실제 사례: ${maxRateSourceRecord.name}(${maxRateSourceRecord.unit}, ${maxRateSourceRecord.mealDate || maxRateSourceRecord.delDate || ''}) 예정 ${Math.round(maxRateSourceRecord.plan).toLocaleString()} → 확정 ${Math.round(maxRateSourceRecord.conf).toLocaleString()}`
                    : '';
                maxCell.innerHTML = isExtreme
                    ? `<span title="예정 대비 ${Math.round(maxRate)}% 편차가 있어 최대값에 반영됨. 입력 오류일 수 있으니 원본 로그를 확인해보세요.${sourceInfoText}${fallbackNote}" style="cursor:help;">⚠️ ${maxVal.toLocaleString()}<br><span style="font-size:9px; font-weight:normal; color:#c2410c;">과거 최고편차 ${Math.round(maxRate)}%</span>${fallbackBadge}</span>`
                    : fallbackBadge + maxVal.toLocaleString();
                maxCell.title = `평균(${mean.toFixed(1)}%) + 넉넉한 안전마진(z≈${maxZ.toFixed(2)})과 과거 실측 최고 편차(${Math.round(maxRate)}%) 중 큰 값 사용`
                    + (isExtreme ? `\n⚠️ 예정 대비 ${Math.round(maxRate)}% 편차가 있어 최대값에 반영됨. 입력 오류일 수 있으니 원본 로그를 확인해보세요.${sourceInfoText}` : '')
                    + fallbackNote;
            }

            saveToDB('specMemory', specMemory);
            saveToDB('unitMemory', unitMemory);
        }

        function renderLogTable() {
            const u = document.getElementById('logFilterUnit').value; 
            const c = document.getElementById('logFilterCat').value;
            const dt = document.getElementById('logFilterDayType').value;
            const n = document.getElementById('logFilterName').value.toLowerCase();
            
            const filtered = dbData.filter(d => {
                const passU = (u==='ALL' || d.unit===u);
                const passC = (c==='ALL' || d.cat===c);
                const passN = (n==='' || d.name.toLowerCase().includes(n));
                const rowCat = d.dayCategory !== undefined ? d.dayCategory
                    : (d.isHoliday === false ? 'WORK' : d.isHoliday === true ? 'HOLI' : null);
                const passDt = (dt==='ALL') || (dt==='WORK' && rowCat==='WORK') || (dt==='SAT' && rowCat==='SAT') || (dt==='HOLI' && rowCat==='HOLI') || (dt==='UNKNOWN' && rowCat===null);
                return passU && passC && passN && passDt;
            });

            document.getElementById('logTableBody').innerHTML = filtered.slice(0,100).map(d => {
                let dayBadge;
                const rowCat = d.dayCategory !== undefined ? d.dayCategory
                    : (d.isHoliday === false ? 'WORK' : d.isHoliday === true ? 'HOLI' : null);
                if (rowCat === 'HOLI') dayBadge = `<span class="day-badge day-holi" title="${d.holidayName || '일요일/공휴일'}">휴</span>`;
                else if (rowCat === 'SAT') dayBadge = `<span class="day-badge day-sat">토</span>`;
                else if (rowCat === 'WORK') dayBadge = `<span class="day-badge day-work">평</span>`;
                else dayBadge = `<span class="day-badge" style="background:#f1f5f9;color:#64748b;border:1px solid #e2e8f0;">${d.dayType==='MULTI'?'복수/미상':'미기재'}</span>`;
                const planTxt = d.plan === null ? '<span style="color:#94a3b8;">없음</span>' : Math.round(d.plan).toLocaleString();
                const confTxt = d.conf === null ? '<span style="color:#94a3b8;">없음</span>' : Math.round(d.conf).toLocaleString();
                const rateTxt = d.rate === null ? '<span style="color:#94a3b8;">계산불가</span>' : `${d.rate>0?'+':''}${Math.round(d.rate)}%`;
                const rateColor = d.rate === null ? '#94a3b8' : (d.rate>0?'#ea580c':'#2563eb');
                const excludedBadge = d.excludedFromStats ? ' <span style="font-size:10px; background:#e2e8f0; color:#475569; padding:1px 5px; border-radius:3px;" title="통계(안전재고 산출)에서 제외됨 — 시트 고유ID열의 &quot;제외&quot; 표시 또는 로그에서 직접 체크">🚫제외</span>' : '';
                const adjustedBadge = d.splitGroupAdjusted ? ` <span style="font-size:10px; background:#dbeafe; color:#1d4ed8; padding:1px 5px; border-radius:3px;" title="같은 납품일·부대·품명으로 예정량이 여러 급식일에 나뉘고 확정량은 합쳐서 온 경우라, 그룹 예정량 합계(${Math.round(d.splitGroupTotalPlan||0).toLocaleString()}) 대비로 오차율을 재계산함">🔗합산보정</span>` : '';
                const rowStyle = d.excludedFromStats ? ' style="opacity:0.55;"' : '';
                return `
                <tr${rowStyle}>
                    <td><input type="checkbox" ${d.excludedFromStats ? 'checked' : ''} onchange="toggleRecordExclusion('${d.id}')" title="이 기록을 통계에서 제외/포함"></td>
                    <td>${d.delDate}</td>
                    <td>${d.mealDate} ${dayBadge}</td>
                    <td>${d.unit}</td><td>${d.cat}</td>
                    <td style="text-align:left;font-weight:600;">${d.name}</td>
                    <td>${planTxt}</td>
                    <td>${confTxt}</td>
                    <td style="font-weight:bold; color:${rateColor}">${rateTxt}${excludedBadge}${adjustedBadge}</td>
                </tr>
            `;}).join('');
        }

        function toggleRecordExclusion(id) {
            const record = dbData.find(d => d.id === id);
            if (!record) return;
            record.excludedFromStats = !record.excludedFromStats;
            saveToDB('dbData', dbData, function() {
                renderLogTable();
                updateDashboard();
            });
        }

        function renderChart(dashData) {
            const ctx = document.getElementById('dashboardChart').getContext('2d'); if(chartIdx) chartIdx.destroy();
            if(!dashData || dashData.length === 0) return;
            const catMap = {}; dashData.forEach(d => { if(!catMap[d.cat]) catMap[d.cat] = []; catMap[d.cat].push(d.rate); });
            const labels = [], values = []; for(let c in catMap) { labels.push(c); values.push(catMap[c].reduce((a,b)=>a+b,0)/catMap[c].length); }
            chartIdx = new Chart(ctx, { type: 'bar', data: { labels: labels.slice(0, 10), datasets: [{ label: '구분별 평균 오차율(%)', data: values.slice(0, 10), backgroundColor: values.slice(0,10).map(v=>v>0?'#fb923c':'#60a5fa') }] }, options: { responsive: true, maintainAspectRatio: false } });
        }

        function exportPredictionExcel() {
            const mode = document.querySelector('input[name="calcMode"]:checked').value;
            const dateBasisLabel = (getDateBasis() === 'delivery') ? '납품일' : '급식일';
            const rows = [[dateBasisLabel, "부대", "품명", "구분", mode==='total'?"총소요량":"기준량(g)", mode==='total'?"-":"식수인원", "단위", "포장규격", "최소준비", "권장안전", "최대리스크"]];
            for(let tr of document.getElementById('excelGridBody').children) {
                const name = tr.querySelector('.grid-name').value; if(!name) continue;
                
                let rowDate = tr.querySelector('.grid-date').value;
                let rowUnit = tr.querySelector('.grid-unit-name').value;
                
                rows.push([ rowDate, rowUnit, name, tr.querySelector('.grid-cat').value, mode==='total' ? tr.querySelector('.grid-req-tot').value : tr.querySelector('.grid-req-per').value, mode==='total' ? "-" : (document.getElementById('globalHeadcount')?.value || ''), tr.querySelector('.grid-unit').value, tr.querySelector('.grid-spec').value, tr.querySelector('.res-min').innerText.replace(/,/g, ''), tr.querySelector('.res-rec').innerText.replace(/,/g, ''), tr.querySelector('.res-max').innerText.replace(/,/g, '') ]);
            }
            const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "예측결과"); XLSX.writeFile(wb, "발주예측결과_V10.xlsx");
        }

        function renormalizeExistingUnits() {
            if(dbData.length === 0) { alert('정리할 데이터가 없습니다.'); return; }
            let changed = 0;
            dbData.forEach(d => {
                const norm = normalizeUnitName(d.unit);
                if(norm !== d.unit) { d.unit = norm; changed++; }
            });
            if(changed === 0) { alert('이미 정규화되어 있어 변경된 행이 없습니다.'); return; }
            saveToDB('dbData', dbData, function() {
                refreshAll();
                alert(`✅ 부대명 재정규화 완료: ${changed.toLocaleString()}건의 부대명 표기를 통일했습니다.`);
            });
        }

        function clearData() { 
            if(confirm("모든 원본 누적 데이터를 완전히 초기화하시겠습니까?")) { 
                if(!indexedDBInstance) return;
                const tx = indexedDBInstance.transaction("kv_store", "readwrite");
                tx.objectStore("kv_store").clear();
                tx.oncomplete = function() { alert("초기화되었습니다."); location.reload(); };
            } 
        }
    </script>
</body>
</html>
