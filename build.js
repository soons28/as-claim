const fs = require('fs');
const path = require('path');
const XLSX_NODE = require('xlsx');

const pdfPath = path.join(__dirname, 'assets', '채권신고 취하서_양식.pdf');
const htmlSourcePath = path.join(__dirname, 'public', 'index.html');
const outputPath = path.join(__dirname, 'index.html');

if (!fs.existsSync(pdfPath)) {
  console.error('원본 PDF 파일이 없습니다:', pdfPath);
  process.exit(1);
}

if (!fs.existsSync(htmlSourcePath)) {
  console.error('HTML 소스 파일이 없습니다:', htmlSourcePath);
  process.exit(1);
}

// 1. PDF 파일을 Base64 문자열로 변환
console.log('PDF 템플릿 변환 중...');
const pdfBase64 = fs.readFileSync(pdfPath).toString('base64');

// 2. 시부인표 엑셀 파일을 빌드 시점에 JSON으로 파싱하여 내장
const EXCEL_PATH_NODE = path.join(__dirname, 'data', '파산채권 시부인표 데이터.xlsx');
let embeddedDbJson = 'null';
let embeddedDbCount = 0;

if (fs.existsSync(EXCEL_PATH_NODE)) {
  try {
    const wb = XLSX_NODE.readFile(EXCEL_PATH_NODE);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX_NODE.utils.sheet_to_json(ws, { header: 1 });
    const dbData = [];
    
    // 2번째 행부터 탐색 (1행은 헤더)
    for (let i = 1; i < rows.length; i++) {
      const rowClaimNo = String(rows[i][1] || '').trim(); // B열: 접수번호
      const rowNameVal = String(rows[i][2] || '').trim(); // C열: 채권자명
      if (rowNameVal && rowClaimNo) {
        dbData.push({ claimNo: rowClaimNo, name: rowNameVal });
      }
    }
    
    embeddedDbJson = JSON.stringify(dbData);
    embeddedDbCount = dbData.length;
    console.log(`시부인표 데이터 자동 내장 성공! (총 ${embeddedDbCount}건)`);
  } catch (e) {
    console.warn('엑셀 파일 파싱 실패 (수동 첨부 모드로 빌드됩니다):', e.message);
  }
} else {
  console.warn('엑셀 파일 없음 → 수동 첨부 모드로 빌드됩니다. 경로:', EXCEL_PATH_NODE);
}

// 3. HTML 템플릿을 읽어서 플레이스홀더 치환
console.log('HTML 소스 코드 읽는 중...');
let htmlContent = fs.readFileSync(htmlSourcePath, 'utf8');

// 플레이스홀더 문자열 치환
htmlContent = htmlContent.replace('"/* EMBEDDED_PDF_BASE64_PLACEHOLDER */"', `"${pdfBase64}"`);
htmlContent = htmlContent.replace('/* EMBEDDED_DB_PLACEHOLDER */', embeddedDbJson);

// 4. 빌드 결과 출력
fs.writeFileSync(outputPath, htmlContent, 'utf8');
console.log('====================================================');
console.log('🎉 HTML 빌드가 완료되었습니다!');
console.log(`💾 빌드 파일 위치: ${outputPath}`);
console.log(`🔗 오프라인 상태에서도 위 HTML 파일을 더블클릭하면 실행됩니다.`);
console.log('====================================================');
