const fs = require('fs');
const path = require('path');
const XLSX_NODE = require('xlsx');

const pdfPath = path.join(__dirname, 'assets', '채권신고 취하서_양식.pdf');
const outputPath = path.join(__dirname, 'index.html');

if (!fs.existsSync(pdfPath)) {
  console.error('원본 PDF 파일이 없습니다:', pdfPath);
  process.exit(1);
}

// PDF 파일을 Base64 문자열로 변환
const pdfBase64 = fs.readFileSync(pdfPath).toString('base64');

// 시부인표 엑셀 파일을 빌드 시점에 JSON으로 파싱하여 내장
const EXCEL_PATH_NODE = require('path').join(__dirname, 'data', '파산채권 시부인표 데이터.xlsx');
let embeddedDbJson = 'null';
let embeddedDbCount = 0;
if (fs.existsSync(EXCEL_PATH_NODE)) {
  try {
    const wb = XLSX_NODE.readFile(EXCEL_PATH_NODE);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX_NODE.utils.sheet_to_json(ws, { header: 1 });
    const dbData = [];
    for (let i = 1; i < rows.length; i++) {
      const rowClaimNo = String(rows[i][1] || '').trim();
      const rowNameVal = String(rows[i][2] || '').trim();
      if (rowNameVal && rowClaimNo) dbData.push({ claimNo: rowClaimNo, name: rowNameVal });
    }
    embeddedDbJson = JSON.stringify(dbData);
    embeddedDbCount = dbData.length;
    console.log('시부인표 데이터 자동 내장 완료:', embeddedDbCount, '건');
  } catch (e) {
    console.warn('엑셀 파싱 실패 (수동 첨부 모드):', e.message);
  }
} else {
  console.warn('엑셀 파일 없음 → 수동 첨부 모드로 빌드:', EXCEL_PATH_NODE);
}

const htmlContent = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>채권신고 취하서 작성 프로그램</title>
  
  <!-- 라이브러리 CDN 호출 -->
  <script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
  <script src="https://unpkg.com/pdf-lib/dist/pdf-lib.min.js"></script>
  <script src="https://unpkg.com/@pdf-lib/fontkit/dist/fontkit.umd.js"></script>
  
  <style>
    :root {
      --primary: #2563eb;
      --primary-hover: #1d4ed8;
      --bg: #f1f5f9;
      --panel-bg: #ffffff;
      --text: #0f172a;
      --text-light: #475569;
      --border: #cbd5e1;
      --success: #10b981;
      --error: #ef4444;
    }

    body {
      font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif;
      margin: 0;
      padding: 0;
      background-color: var(--bg);
      color: var(--text);
      display: flex;
      height: 100vh;
      overflow: hidden;
    }

    /* 좌측 컨트롤 영역 */
    .left-panel {
      width: 450px;
      background-color: var(--panel-bg);
      border-right: 1px solid var(--border);
      display: flex;
      flex-direction: column;
      height: 100%;
      box-shadow: 4px 0 10px rgba(0,0,0,0.05);
      z-index: 10;
    }

    .panel-header {
      padding: 20px;
      border-bottom: 1px solid #f1f5f9;
      background-color: #fafbfd;
    }

    .panel-header h1 {
      margin: 0;
      font-size: 18px;
      color: #0f172a;
      font-weight: 700;
    }
    
    .panel-header p {
      margin: 4px 0 0 0;
      font-size: 12px;
      color: var(--text-light);
    }

    .panel-content {
      padding: 20px;
      flex: 1;
      overflow-y: auto;
    }

    .section-box {
      background: #f8fafc;
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 14px;
      margin-bottom: 16px;
    }

    .section-title {
      font-size: 13px;
      font-weight: bold;
      margin-bottom: 10px;
      color: #1e293b;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }

    .form-group {
      margin-bottom: 14px;
    }

    label {
      display: block;
      font-size: 12px;
      font-weight: 600;
      margin-bottom: 4px;
      color: #475569;
    }

    input[type="text"], input[type="tel"], select {
      width: 100%;
      padding: 8px 10px;
      border: 1px solid var(--border);
      border-radius: 6px;
      font-size: 13px;
      box-sizing: border-box;
      background-color: #fff;
    }

    input[type="text"]:focus, input[type="tel"]:focus, select:focus {
      outline: none;
      border-color: var(--primary);
    }

    .btn {
      width: 100%;
      padding: 10px;
      border: none;
      border-radius: 6px;
      font-size: 13px;
      font-weight: bold;
      cursor: pointer;
      transition: background 0.2s;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
    }

    .btn-primary {
      background: var(--primary);
      color: white;
    }

    .btn-primary:hover {
      background: var(--primary-hover);
    }

    .btn-success {
      background: var(--success);
      color: white;
    }

    .btn-success:hover {
      background: #059669;
    }

    .btn-secondary {
      background: #64748b;
      color: white;
      margin-top: 8px;
    }

    .btn:disabled {
      background: #cbd5e1;
      color: #94a3b8;
      cursor: not-allowed;
    }

    .status-badge {
      font-size: 11px;
      padding: 2px 6px;
      border-radius: 4px;
      font-weight: bold;
    }
    .status-badge.red { background: #fee2e2; color: var(--error); }
    .status-badge.green { background: #ecfdf5; color: var(--success); }

    .message {
      padding: 10px;
      border-radius: 6px;
      font-size: 12px;
      margin-top: 8px;
      display: none;
    }
    .message.error { background: #fee2e2; color: var(--error); display: block; border: 1px solid #fecaca; }
    .message.success { background: #ecfdf5; color: #065f46; display: block; border: 1px solid #a7f3d0; }

    /* 접수현황 리스트 */
    .history-box {
      margin-top: 20px;
      border-top: 1px solid var(--border);
      padding-top: 16px;
    }

    .history-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 11px;
      margin-top: 8px;
    }

    .history-table th, .history-table td {
      border: 1px solid var(--border);
      padding: 6px;
      text-align: left;
    }

    .history-table th {
      background: #f8fafc;
    }

    /* 우측 PDF 미리보기 영역 */
    .preview-panel {
      flex: 1;
      height: 100%;
      display: flex;
      flex-direction: column;
      background-color: #475569;
    }

    .preview-header {
      background-color: #1e293b;
      padding: 12px 20px;
      color: white;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }

    .preview-header h2 {
      margin: 0;
      font-size: 14px;
      font-weight: 600;
    }

    .preview-body {
      flex: 1;
      position: relative;
      display: flex;
      justify-content: center;
      align-items: center;
      padding: 4px;
    }

    iframe {
      width: 100%;
      height: 100%;
      border: none;
      background-color: #ffffff;
      border-radius: 4px;
      box-shadow: 0 4px 20px rgba(0,0,0,0.25);
    }

    .loading-overlay {
      position: absolute;
      top: 0; left: 0; right: 0; bottom: 0;
      background: rgba(30, 41, 59, 0.85);
      display: flex;
      flex-direction: column;
      justify-content: center;
      align-items: center;
      color: white;
      font-size: 14px;
      z-index: 20;
      gap: 10px;
    }

    /* 좌표 조정 버튼 스타일 */
    .btn-coord {
      width: 22px;
      height: 22px;
      padding: 0;
      border: 1px solid var(--border);
      background: #f1f5f9;
      border-radius: 4px;
      cursor: pointer;
      font-size: 12px;
      font-weight: bold;
      display: flex;
      align-items: center;
      justify-content: center;
      user-select: none;
    }
    .btn-coord:hover {
      background: #e2e8f0;
    }
  </style>
</head>
<body>

  <!-- 좌측 컨트롤 패널 -->
  <div class="left-panel">
    <div class="panel-header">
      <h1>채권신고 취하서 타이핑 프로그램</h1>
      <p>더블클릭만으로 실행되는 오프라인 단독형 프로그램입니다.</p>
    </div>

    <div class="panel-content">
      <!-- 1단계: 시부인표 데이터 연동 (빌드 시 자동 내장) -->
      <div class="section-box">
        <div class="section-title">
          1단계. 시부인표 데이터 연동
          <span id="dbStatus" class="status-badge red">확인 중</span>
        </div>
        <div id="excelInfo" style="font-size: 12px; color: var(--success); display: none; font-weight: bold;">
          ✅ 데이터 자동 연동 완료! (총 <span id="dbCount">0</span>명)
        </div>
        <div id="excelUploadGroup" style="display:none;">
          <label for="excelFile" style="font-size:12px;">엑셀 파일 수동 등록 (.xlsx)</label>
          <input type="file" id="excelFile" accept=".xlsx" onchange="loadExcel(this)">
          <div style="font-size: 11px; color: var(--text-light); margin-top: 4px;">
            * '파산채권 시부인표 데이터.xlsx' 파일을 등록해 주세요.
          </div>
        </div>
      </div>

      <!-- 2단계: 성함 조회 -->
      <div class="section-box" id="searchSection" style="opacity: 0.5; pointer-events: none;">
        <div class="section-title">2단계. 채권자 검색</div>
        <div class="form-group">
          <label for="nameInput">채권자 성함 (또는 법인명)</label>
          <input type="text" id="nameInput" placeholder="이름 입력 후 엔터 또는 검색 클릭" onkeydown="if(event.key==='Enter') lookupName()">
        </div>
        <button type="button" class="btn btn-primary" id="btnLookup" onclick="lookupName()">성함으로 채권번호 조회</button>
        <div id="lookupMsg" class="message"></div>
      </div>

      <!-- 3단계: 상세 정보 입력 및 인쇄/다운로드 -->
      <div class="section-box" id="step2" style="display: none;">
        <div class="section-title">3단계. 정보 입력 및 저장</div>
        <div id="multiClaimMsg" style="display:none; color:#dc2626; font-size:12px; font-weight:bold; margin-bottom:8px; padding:6px 8px; background:#fee2e2; border-radius:4px; border:1px solid #fecaca;"></div>
        <div class="form-group">
          <label for="claimSelect">확인된 채권번호 선택</label>
          <select id="claimSelect" onchange="onClaimSelectChange()"></select>
        </div>

        <div class="form-group">
          <label for="claimNoInput">채권번호 (수정/직접입력 가능)</label>
          <input type="text" id="claimNoInput" oninput="queuePreview()" required>
        </div>

        <div class="form-group">
          <label for="unitNo">구분 (동/호수)</label>
          <input type="text" id="unitNo" placeholder="예: 101동 202호" oninput="queuePreview()" required>
        </div>

        <div class="form-group">
          <label for="phone">연락처</label>
          <input type="tel" id="phone" placeholder="예: 010-1234-5678" oninput="queuePreview()" required>
        </div>

        <div class="form-group" style="margin-top: 16px;">
          <button type="button" class="btn btn-success" id="btnSubmit" onclick="saveAndPrint()">
            🖨️ 즉시 인쇄
          </button>
        </div>
        <div id="submitMsg" class="message"></div>
      </div>

      <!-- 좌표 조정 모드 -->
      <div class="section-box" style="margin-top: 16px;">
        <details>
          <summary style="font-size: 13px; font-weight: bold; cursor: pointer; color: #1e293b; user-select: none;">
            🔧 출력 글자 위치 미세조정 (상하좌우 위치)
          </summary>
          <div style="font-size: 11px; color: var(--text-light); margin: 8px 0; line-height: 1.4;">
            * 각 항목의 글자 출력 위치(가로 X, 세로 Y)를 조정할 수 있습니다. (PDF-Lib 기준: 왼쪽 아래가 0, 0 이며 값이 커질수록 오른쪽/위로 이동합니다.)
          </div>
          <table style="width: 100%; font-size: 12px; border-collapse: collapse; border-top: 1px solid var(--border);">
            <thead>
              <tr style="background: #f8fafc; border-bottom: 1px solid var(--border);">
                <th style="text-align: left; padding: 6px;">항목</th>
                <th style="padding: 6px; width: 90px; text-align: center;">가로 위치(X)</th>
                <th style="padding: 6px; width: 90px; text-align: center;">세로 위치(Y)</th>
              </tr>
            </thead>
            <tbody>
              <!-- 성명 -->
              <tr style="border-bottom: 1px dashed #e2e8f0;">
                <td style="padding: 6px; font-weight: 600; color: #475569;">성명</td>
                <td style="padding: 6px; text-align: center;">
                  <div style="display:flex; align-items:center; justify-content:center; gap:2px;">
                    <button type="button" class="btn-coord" onclick="adjustCoord('nameX', -2)">-</button>
                    <input type="number" id="coord_nameX" style="width:36px; text-align:center; font-size:11px; padding:2px;" onchange="saveCoords()">
                    <button type="button" class="btn-coord" onclick="adjustCoord('nameX', 2)">+</button>
                  </div>
                </td>
                <td style="padding: 6px; text-align: center;">
                  <div style="display:flex; align-items:center; justify-content:center; gap:2px;">
                    <button type="button" class="btn-coord" onclick="adjustCoord('nameY', -2)">-</button>
                    <input type="number" id="coord_nameY" style="width:36px; text-align:center; font-size:11px; padding:2px;" onchange="saveCoords()">
                    <button type="button" class="btn-coord" onclick="adjustCoord('nameY', 2)">+</button>
                  </div>
                </td>
              </tr>
              <!-- 동호수 -->
              <tr style="border-bottom: 1px dashed #e2e8f0;">
                <td style="padding: 6px; font-weight: 600; color: #475569;">동호수</td>
                <td style="padding: 6px; text-align: center;">
                  <div style="display:flex; align-items:center; justify-content:center; gap:2px;">
                    <button type="button" class="btn-coord" onclick="adjustCoord('unitX', -2)">-</button>
                    <input type="number" id="coord_unitX" style="width:36px; text-align:center; font-size:11px; padding:2px;" onchange="saveCoords()">
                    <button type="button" class="btn-coord" onclick="adjustCoord('unitX', 2)">+</button>
                  </div>
                </td>
                <td style="padding: 6px; text-align: center;">
                  <div style="display:flex; align-items:center; justify-content:center; gap:2px;">
                    <button type="button" class="btn-coord" onclick="adjustCoord('unitY', -2)">-</button>
                    <input type="number" id="coord_unitY" style="width:36px; text-align:center; font-size:11px; padding:2px;" onchange="saveCoords()">
                    <button type="button" class="btn-coord" onclick="adjustCoord('unitY', 2)">+</button>
                  </div>
                </td>
              </tr>
              <!-- 채권번호 -->
              <tr style="border-bottom: 1px dashed #e2e8f0;">
                <td style="padding: 6px; font-weight: 600; color: #475569;">채권번호</td>
                <td style="padding: 6px; text-align: center;">
                  <div style="display:flex; align-items:center; justify-content:center; gap:2px;">
                    <button type="button" class="btn-coord" onclick="adjustCoord('claimX', -2)">-</button>
                    <input type="number" id="coord_claimX" style="width:36px; text-align:center; font-size:11px; padding:2px;" onchange="saveCoords()">
                    <button type="button" class="btn-coord" onclick="adjustCoord('claimX', 2)">+</button>
                  </div>
                </td>
                <td style="padding: 6px; text-align: center;">
                  <div style="display:flex; align-items:center; justify-content:center; gap:2px;">
                    <button type="button" class="btn-coord" onclick="adjustCoord('claimY', -2)">-</button>
                    <input type="number" id="coord_claimY" style="width:36px; text-align:center; font-size:11px; padding:2px;" onchange="saveCoords()">
                    <button type="button" class="btn-coord" onclick="adjustCoord('claimY', 2)">+</button>
                  </div>
                </td>
              </tr>
              <!-- 연락처 -->
              <tr style="border-bottom: 1px dashed #e2e8f0;">
                <td style="padding: 6px; font-weight: 600; color: #475569;">연락처</td>
                <td style="padding: 6px; text-align: center;">
                  <div style="display:flex; align-items:center; justify-content:center; gap:2px;">
                    <button type="button" class="btn-coord" onclick="adjustCoord('phoneX', -2)">-</button>
                    <input type="number" id="coord_phoneX" style="width:36px; text-align:center; font-size:11px; padding:2px;" onchange="saveCoords()">
                    <button type="button" class="btn-coord" onclick="adjustCoord('phoneX', 2)">+</button>
                  </div>
                </td>
                <td style="padding: 6px; text-align: center;">
                  <div style="display:flex; align-items:center; justify-content:center; gap:2px;">
                    <button type="button" class="btn-coord" onclick="adjustCoord('phoneY', -2)">-</button>
                    <input type="number" id="coord_phoneY" style="width:36px; text-align:center; font-size:11px; padding:2px;" onchange="saveCoords()">
                    <button type="button" class="btn-coord" onclick="adjustCoord('phoneY', 2)">+</button>
                  </div>
                </td>
              </tr>
              <!-- 날짜(월) -->
              <tr style="border-bottom: 1px dashed #e2e8f0;">
                <td style="padding: 6px; font-weight: 600; color: #475569;">날짜(월)</td>
                <td style="padding: 6px; text-align: center;">
                  <div style="display:flex; align-items:center; justify-content:center; gap:2px;">
                    <button type="button" class="btn-coord" onclick="adjustCoord('monthX', -2)">-</button>
                    <input type="number" id="coord_monthX" style="width:36px; text-align:center; font-size:11px; padding:2px;" onchange="saveCoords()">
                    <button type="button" class="btn-coord" onclick="adjustCoord('monthX', 2)">+</button>
                  </div>
                </td>
                <td style="padding: 6px; text-align: center;">
                  <div style="display:flex; align-items:center; justify-content:center; gap:2px;">
                    <button type="button" class="btn-coord" onclick="adjustCoord('monthY', -2)">-</button>
                    <input type="number" id="coord_monthY" style="width:36px; text-align:center; font-size:11px; padding:2px;" onchange="saveCoords()">
                    <button type="button" class="btn-coord" onclick="adjustCoord('monthY', 2)">+</button>
                  </div>
                </td>
              </tr>
              <!-- 날짜(일) -->
              <tr style="border-bottom: 1px dashed #e2e8f0;">
                <td style="padding: 6px; font-weight: 600; color: #475569;">날짜(일)</td>
                <td style="padding: 6px; text-align: center;">
                  <div style="display:flex; align-items:center; justify-content:center; gap:2px;">
                    <button type="button" class="btn-coord" onclick="adjustCoord('dayX', -2)">-</button>
                    <input type="number" id="coord_dayX" style="width:36px; text-align:center; font-size:11px; padding:2px;" onchange="saveCoords()">
                    <button type="button" class="btn-coord" onclick="adjustCoord('dayX', 2)">+</button>
                  </div>
                </td>
                <td style="padding: 6px; text-align: center;">
                  <div style="display:flex; align-items:center; justify-content:center; gap:2px;">
                    <button type="button" class="btn-coord" onclick="adjustCoord('dayY', -2)">-</button>
                    <input type="number" id="coord_dayY" style="width:36px; text-align:center; font-size:11px; padding:2px;" onchange="saveCoords()">
                    <button type="button" class="btn-coord" onclick="adjustCoord('dayY', 2)">+</button>
                  </div>
                </td>
              </tr>
              <!-- 서명자성명 -->
              <tr style="border-bottom: 1px solid var(--border);">
                <td style="padding: 6px; font-weight: 600; color: #475569;">서명란 이름</td>
                <td style="padding: 6px; text-align: center;">
                  <div style="display:flex; align-items:center; justify-content:center; gap:2px;">
                    <button type="button" class="btn-coord" onclick="adjustCoord('sigX', -2)">-</button>
                    <input type="number" id="coord_sigX" style="width:36px; text-align:center; font-size:11px; padding:2px;" onchange="saveCoords()">
                    <button type="button" class="btn-coord" onclick="adjustCoord('sigX', 2)">+</button>
                  </div>
                </td>
                <td style="padding: 6px; text-align: center;">
                  <div style="display:flex; align-items:center; justify-content:center; gap:2px;">
                    <button type="button" class="btn-coord" onclick="adjustCoord('sigY', -2)">-</button>
                    <input type="number" id="coord_sigY" style="width:36px; text-align:center; font-size:11px; padding:2px;" onchange="saveCoords()">
                    <button type="button" class="btn-coord" onclick="adjustCoord('sigY', 2)">+</button>
                  </div>
                </td>
              </tr>
            </tbody>
          </table>
          <div style="margin-top: 10px;">
            <button type="button" class="btn btn-secondary" style="margin:0; padding:6px; font-size:11px;" onclick="resetCoords()">기본값으로 초기화</button>
          </div>
        </details>
      </div>

    </div>
  </div>

  <!-- 우측 미리보기 패널 -->
  <div class="preview-panel">
    <div class="preview-header">
      <h2>실시간 PDF 타이핑 미리보기</h2>
      <div style="font-size: 11px; color: #94a3b8;">* 한글 나눔고딕 폰트 자동 적용 완료</div>
    </div>
    <div class="preview-body">
      <div id="loadingOverlay" class="loading-overlay">
        <div>시스템 리소스 초기화 중...</div>
        <div style="font-size: 12px; color: #94a3b8;">(최초 실행 시 한글 폰트를 온라인에서 다운로드합니다)</div>
      </div>
      <iframe id="pdfPreview"></iframe>
    </div>
  </div>

  <script>
    // 원본 PDF 템플릿 바이트 (빌더에 의해 Base64 코드로 치환됨)
    const EMBEDDED_PDF_BASE64 = "${pdfBase64}";
    const EMBEDDED_DB = ${embeddedDbJson}; // 빌드 시 자동 내장된 시부인표 데이터
    
    let templateBytes = null;
    let fontBytes = null;
    let currentPdfBytes = null;
    let previewDebounce = null;
    let creditorsDatabase = [];
    let submitHistory = [];

    // 기본 좌표 및 조정 설정
    const DEFAULT_COORDS = {
      nameX: 376, nameY: 552,
      unitX: 376, unitY: 505,
      claimX: 376, claimY: 458,
      phoneX: 376, phoneY: 411,
      monthX: 300, monthY: 248,
      dayX: 340, dayY: 248,
      sigX: 297, sigY: 213
    };

    let coords = { ...DEFAULT_COORDS };
    const APP_VERSION = '2.0';

    function checkAppVersion() {
      const savedVersion = localStorage.getItem('app_version');
      if (savedVersion !== APP_VERSION) {
        // 버전이 다르면 기존 좌표 설정을 덮어씌워서 강제로 기본값으로 리셋
        localStorage.setItem('draw_coords', JSON.stringify(DEFAULT_COORDS));
        localStorage.setItem('app_version', APP_VERSION);
      }
    }

    // 초기 파일 로드 및 세팅
    window.onload = async function() {
      const loader = document.getElementById('loadingOverlay');
      try {
        // 1. 내장된 Base64 PDF 로드
        const binaryString = atob(EMBEDDED_PDF_BASE64);
        const len = binaryString.length;
        templateBytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
          templateBytes[i] = binaryString.charCodeAt(i);
        }

        // 2. 구글 나눔고딕 한글 폰트 다운로드 (CORS 허용된 CDN 주소)
        const fontRes = await fetch('https://raw.githubusercontent.com/google/fonts/main/ofl/nanumgothic/NanumGothic-Regular.ttf');
        if (!fontRes.ok) throw new Error('한글 폰트 파일 다운로드 실패');
        fontBytes = await fontRes.arrayBuffer();

        loader.style.display = 'none';
        
        // 버전 체크 및 필요시 좌표 자동 리셋
        checkAppVersion();
        
        // 내장 DB가 있으면 자동 로드, 없으면 로컬스토리지에서 로드
        if (EMBEDDED_DB && EMBEDDED_DB.length > 0) {
          creditorsDatabase = EMBEDDED_DB;
          showExcelStatus(EMBEDDED_DB.length);
        } else {
          loadSavedExcelData();
        }
        // 저장된 글자 위치(좌표) 로드
        loadSavedCoords();
        
        updatePreview();
      } catch (err) {
        loader.innerHTML = '<div style="color:red; font-weight:bold;">초기화 오류가 발생했습니다. 인터넷 연결 상태를 확인하고 새로고침 해주세요.</div>';
        console.error(err);
      }
    };

    // 로컬 브라우저에 저장되어 있는 시부인표 데이터가 있는지 로드
    function loadSavedExcelData() {
      const saved = localStorage.getItem('creditor_db_saved');
      if (saved) {
        try {
          creditorsDatabase = JSON.parse(saved);
          showExcelStatus(creditorsDatabase.length);
        } catch (e) {}
      }
    }

    // 로컬 스토리지 접수 이력 로드
    function loadHistoryFromStorage() {
      const saved = localStorage.getItem('today_submissions');
      if (saved) {
        try {
          submitHistory = JSON.parse(saved);
          updateHistoryTable();
        } catch (e) {}
      }
    }

    // 엑셀 파일 로딩 및 파싱 (xlsx 라이브러리 사용)
    function loadExcel(input) {
      const file = input.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

          creditorsDatabase = [];
          // 2번째 행부터 읽음 (1행은 헤더)
          for (let i = 1; i < rows.length; i++) {
            const rowClaimNo = String(rows[i][1] || '').trim(); // B열: 접수번호
            const rowNameVal = String(rows[i][2] || '').trim(); // C열: 채권자명

            if (rowNameVal && rowClaimNo) {
              creditorsDatabase.push({
                claimNo: rowClaimNo,
                name: rowNameVal
              });
            }
          }

          if (creditorsDatabase.length > 0) {
            // 브라우저 영구 기억 세팅 (새로고침해도 파일 재등록 불필요하게 처리)
            localStorage.setItem('creditor_db_saved', JSON.stringify(creditorsDatabase));
            showExcelStatus(creditorsDatabase.length);
          } else {
            alert('엑셀 파일에서 유효한 채권자 데이터를 찾지 못했습니다.');
          }
        } catch (err) {
          alert('엑셀 파일을 파싱하는 데 실패했습니다: ' + err.toString());
        }
      };
      reader.readAsArrayBuffer(file);
    }

    function showExcelStatus(count) {
      document.getElementById('dbStatus').className = 'status-badge green';
      document.getElementById('dbStatus').innerText = '연동 완료';
      document.getElementById('excelUploadGroup').style.display = 'none';
      document.getElementById('dbCount').innerText = count;
      document.getElementById('excelInfo').style.display = 'block';
      
      // 검색창 활성화
      document.getElementById('searchSection').style.opacity = '1';
      document.getElementById('searchSection').style.pointerEvents = 'auto';
    }

    // 성함 검색 및 매칭 찾기
    function lookupName() {
      const name = document.getElementById('nameInput').value.trim();
      const msg = document.getElementById('lookupMsg');

      if (!name) {
        alert('조회할 이름을 입력해 주세요.');
        return;
      }

      const searchName = name.replace(/\s+/g, '');
      if (searchName.length < 2) {
        alert('이름을 2글자 이상 입력해 주세요.');
        return;
      }

      const matches = [];
      creditorsDatabase.forEach(item => {
        const rowNameNorm = item.name.replace(/\s+/g, '');
        // 1. 완전 일치 또는 공백 제거 일치
        let isMatch = (rowNameNorm === searchName);

        // 2. 부분 일치 (예: 검색어 "곽현수"가 DB의 "곽현수(국일종합상사)"에 포함되는지 확인)
        if (!isMatch && rowNameNorm.includes(searchName)) {
          isMatch = true;
        }

        // 3. 공동 채권자 처리 (쉼표 구분)
        if (!isMatch && item.name.includes(',')) {
          const parts = item.name.split(',');
          for (let p of parts) {
            const cleanPart = p.replace(/\s+/g, '').trim();
            if (cleanPart === searchName || cleanPart.includes(searchName)) {
              isMatch = true;
              break;
            }
          }
        }

        if (isMatch) {
          matches.push(item);
        }
      });

      if (matches.length > 0) {
        msg.style.display = 'none';
        const select = document.getElementById('claimSelect');
        select.innerHTML = '';

        // 다건 검색 시 빨간 메시지 표시
        const multiMsg = document.getElementById('multiClaimMsg');
        if (matches.length > 1) {
          multiMsg.textContent = '⚠️ 총 ' + matches.length + '건의 채권번호가 검색되었습니다. 아래에서 선택하세요.';
          multiMsg.style.display = 'block';
        } else {
          multiMsg.style.display = 'none';
        }

        // 만약 매칭이 여러 개면 "모두 합치기" 옵션을 맨 위에 추가
        if (matches.length > 1) {
          const allClaimNos = matches.map(m => m.claimNo).join(', ');
          const opt = document.createElement('option');
          opt.value = allClaimNos;
          opt.setAttribute('data-name', matches[0].name); // 대표 이름 사용
          opt.text = '★ 모든 채권번호 합치기 (' + allClaimNos + ')';
          select.appendChild(opt);
        }

        matches.forEach(m => {
          const opt = document.createElement('option');
          opt.value = m.claimNo;
          opt.setAttribute('data-name', m.name);
          opt.text = m.claimNo + ' [' + m.name + ']';
          select.appendChild(opt);
        });

        // 3단계 정보창 열기
        document.getElementById('step2').style.display = 'block';
        
        // 첫 번째 옵션 선택 후 claimNoInput에 값 동기화
        const input = document.getElementById('claimNoInput');
        if (input && select.options.length > 0) {
          input.value = select.options[0].value;
        }

        // 이전 동호수 및 연락처 입력값 초기화
        document.getElementById('unitNo').value = '';
        document.getElementById('phone').value = '';

        updatePreview();
      } else {
        msg.innerHTML = '일치하는 채권자를 찾을 수 없습니다.<br><small style="color:gray;">엑셀에 기재된 정확한 이름으로 검색해 주세요.</small>';
        msg.className = 'message error';
      }
    }

    function queuePreview() {
      clearTimeout(previewDebounce);
      previewDebounce = setTimeout(updatePreview, 300);
    }

    function onClaimSelectChange() {
      const select = document.getElementById('claimSelect');
      const input = document.getElementById('claimNoInput');
      if (select && input) {
        input.value = select.value;
      }
      queuePreview();
    }

    // PDF 그리기 로직 (브라우저 자체 가동)
    async function updatePreview() {
      if (!templateBytes || !fontBytes) return;

      try {
        const name = document.getElementById('nameInput').value.trim();
        const claimSelect = document.getElementById('claimSelect');
        const claimNoInput = document.getElementById('claimNoInput');
        const claimNo = claimNoInput ? claimNoInput.value.trim() : '';
        const selectedOpt = claimSelect && claimSelect.options[claimSelect.selectedIndex];
        const officialName = selectedOpt ? selectedOpt.getAttribute('data-name') : name;
        const unitNo = document.getElementById('unitNo').value.trim();
        const phone = document.getElementById('phone').value.trim();

        const pdfDoc = await PDFLib.PDFDocument.load(templateBytes);
        pdfDoc.registerFontkit(window.fontkit);
        const customFont = await pdfDoc.embedFont(fontBytes);
        const firstPage = pdfDoc.getPages()[0];

        // maxAvailableWidth를 전달받아 글씨 크기를 자동으로 줄여주는 그리기 함수
        const draw = (text, x, y, size = 12, center = false, maxAvailableWidth = null) => {
          if (!text) return;
          let currentSize = size;
          if (maxAvailableWidth) {
            const textWidth = customFont.widthOfTextAtSize(text, currentSize);
            if (textWidth > maxAvailableWidth) {
              currentSize = Math.max(8, currentSize * (maxAvailableWidth / textWidth));
            }
          }
          let drawX = x;
          if (center) {
            const textWidth = customFont.widthOfTextAtSize(text, currentSize);
            drawX = x - (textWidth / 2);
          }
          firstPage.drawText(text, {
            x: drawX, y: y, size: currentSize,
            font: customFont,
            color: PDFLib.rgb(0, 0, 0)
          });
        };

        // PDF 오버레이 쓰기 (테이블 셀 최대 가로 크기를 300pt로 제한하여 자동 축소 적용)
        draw(officialName, coords.nameX, coords.nameY, 16, true, 300);
        
        const fullUnit = unitNo ? "안산유통상가 " + unitNo : '';
        draw(fullUnit, coords.unitX, coords.unitY, 16, true, 300);
        
        draw(claimNo, coords.claimX, coords.claimY, 16, true, 300);
        draw(phone, coords.phoneX, coords.phoneY, 16, true, 300);

        const today = new Date();
        draw(String(today.getMonth() + 1), coords.monthX, coords.monthY, 13, true); // 월 (기존 16 -> 13)
        draw(String(today.getDate()), coords.dayX, coords.dayY, 13, true);    // 일 (기존 16 -> 13)

        // 이름이 길어 서명하는 부분의 '(인 또는 서명)' 글씨와 겹치지 않도록 글자 크기 동적 조절
        let sigFontSize = 16;
        if (officialName) {
          const textWidth = customFont.widthOfTextAtSize(officialName, sigFontSize);
          const maxSigWidth = 160; // 서명란의 최대 허용 가로 너비 (sigX 297 기준 160pt가 안전)
          if (textWidth > maxSigWidth) {
            sigFontSize = Math.max(11, sigFontSize * (maxSigWidth / textWidth)); // 최소 크기도 11로 올려 가독성 확보
          }
        }
        draw(officialName, coords.sigX, coords.sigY, sigFontSize, true); // 위 채권자 서명자명 (center: true)

        currentPdfBytes = await pdfDoc.save();
        const pdfBlob = new Blob([currentPdfBytes], { type: 'application/pdf' });
        const pdfUrl = URL.createObjectURL(pdfBlob);
        // #toolbar=0&navpanes=0&scrollbar=0&zoom=65 옵션으로 PDF를 크게 꽉 채워서 표시
        document.getElementById('pdfPreview').src = pdfUrl + '#toolbar=0&navpanes=0&scrollbar=0&zoom=65';
      } catch (err) {
        console.error('PDF 갱신 오류:', err);
      }
    }

    // PDF 다운로드 및 인쇄 창 출력
    function saveAndPrint() {
      const nameInputVal = document.getElementById('nameInput').value.trim();
      const claimSelect = document.getElementById('claimSelect');
      const selectedOpt = claimSelect && claimSelect.options[claimSelect.selectedIndex];
      const officialName = selectedOpt ? selectedOpt.getAttribute('data-name') : nameInputVal;
      const claimNoInput = document.getElementById('claimNoInput');
      const claimNo = claimNoInput ? claimNoInput.value.trim() : '';
      const unitNo = document.getElementById('unitNo').value.trim();
      const phone = document.getElementById('phone').value.trim();

      if (!unitNo || !phone) {
        alert('동호수와 연락처를 입력해 주세요.');
        return;
      }

      // 1. 접수 이력에 누적 저장
      const submission = {
        time: new Date().toLocaleString('ko-KR'),
        name: officialName,
        unitNo,
        claimNo,
        phone
      };
      
      submitHistory.unshift(submission);
      localStorage.setItem('today_submissions', JSON.stringify(submitHistory));
      updateHistoryTable();

      // 2. 새 창(탭)에 PDF를 열고 onload 시점에 인쇄 창 자동 실행
      const pdfBlob = new Blob([currentPdfBytes], { type: 'application/pdf' });
      const downloadUrl = URL.createObjectURL(pdfBlob);
      const printWin = window.open(downloadUrl, '_blank');
      if (printWin) {
        printWin.onload = function() {
          setTimeout(function() {
            printWin.focus();
            printWin.print();
          }, 500);
        };
      }

      const msg = document.getElementById('submitMsg');
      msg.innerHTML = '<strong>접수 및 인쇄 작업 기동 완료!</strong><br>인쇄 팝업이 활성화되었습니다.';
      msg.className = 'message success';
    }

    // 접수 내역 테이블 동적 업데이트
    function updateHistoryTable() {
      const tbody = document.getElementById('historyBody');
      if (submitHistory.length === 0) {
        tbody.innerHTML = '<tr><td colspan="3" style="text-align: center; color: var(--text-light);">접수 내역이 없습니다.</td></tr>';
        return;
      }

      tbody.innerHTML = '';
      submitHistory.forEach(item => {
        const tr = document.createElement('tr');
        tr.innerHTML = '<td><strong>' + item.name + '</strong></td>' +
                       '<td>' + item.unitNo + '</td>' +
                       '<td><code>' + item.claimNo + '</code></td>';
        tbody.appendChild(tr);
      });
    }

    // 접수 내역을 엑셀로 내보내기 (CSV)
    function exportHistoryToCSV() {
      if (submitHistory.length === 0) {
        alert('내보낼 접수 내역이 없습니다.');
        return;
      }

      const csvRows = ['\\ufeff접수시간,성명,동호수,채권번호,연락처'];
      submitHistory.forEach(h => {
        csvRows.push('"' + h.time + '","' + h.name + '","' + h.unitNo + '","' + h.claimNo + '","' + h.phone + '"');
      });

      const csvContent = csvRows.join('\\n');
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      
      const a = document.createElement('a');
      a.href = url;
      a.download = '채권취하서_접수현황_' + new Date().toISOString().slice(0, 10) + '.csv';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }

    function resetForm() {
      document.getElementById('step2').style.display = 'none';
      document.getElementById('nameInput').value = '';
      document.getElementById('unitNo').value = '';
      document.getElementById('phone').value = '';
      const claimNoInput = document.getElementById('claimNoInput');
      if (claimNoInput) claimNoInput.value = '';
      document.getElementById('submitMsg').style.display = 'none';
      document.getElementById('lookupMsg').style.display = 'none';
      updatePreview();
    }

    // 좌표 미세조정 함수 구현
    function loadSavedCoords() {
      const saved = localStorage.getItem('draw_coords');
      if (saved) {
        try {
          coords = { ...DEFAULT_COORDS, ...JSON.parse(saved) };
        } catch (e) {
          coords = { ...DEFAULT_COORDS };
        }
      } else {
        coords = { ...DEFAULT_COORDS };
      }
      updateCoordInputs();
    }

    function updateCoordInputs() {
      for (const key in coords) {
        const input = document.getElementById('coord_' + key);
        if (input) {
          input.value = coords[key];
        }
      }
    }

    function adjustCoord(key, amount) {
      if (coords[key] !== undefined) {
        coords[key] = parseInt(coords[key]) + amount;
        const input = document.getElementById('coord_' + key);
        if (input) {
          input.value = coords[key];
        }
        saveCoords();
      }
    }

    function saveCoords() {
      for (const key in coords) {
        const input = document.getElementById('coord_' + key);
        if (input) {
          coords[key] = parseInt(input.value) || DEFAULT_COORDS[key];
        }
      }
      localStorage.setItem('draw_coords', JSON.stringify(coords));
      queuePreview();
    }

    function resetCoords() {
      coords = { ...DEFAULT_COORDS };
      updateCoordInputs();
      localStorage.setItem('draw_coords', JSON.stringify(coords));
      queuePreview();
    }
  </script>
</body>
</html>`;

fs.writeFileSync(outputPath, htmlContent);
console.log('HTML 빌드 성공! 파일 생성 위치:', outputPath);
