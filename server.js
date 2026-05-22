const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const app = express();
const PORT = 3000;

// JSON 요청 본문 파싱 크기 제한 늘리기 (PDF base64 전송용)
app.use(express.json({ limit: '20mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const EXCEL_PATH = path.join(__dirname, 'data', '파산채권 시부인표 데이터.xlsx');
const PDF_TEMPLATE_PATH = path.join(__dirname, 'assets', '채권신고 취하서_양식.pdf');
const FONT_PATH = 'C:\\Windows\\Fonts\\malgun.ttf';
const OUTPUT_DIR = path.join(__dirname, 'output');
const CSV_PATH = path.join(OUTPUT_DIR, '접수현황.csv');

// 시부인표 전체 데이터 목록 조회 API
app.get('/api/list', (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_PATH)) {
      return res.status(500).json({ success: false, message: '시부인표 엑셀 파일을 찾을 수 없습니다.' });
    }

    const workbook = XLSX.readFile(EXCEL_PATH);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const dbData = [];
    // 2번째 행부터 탐색 (1행은 헤더)
    for (let i = 1; i < data.length; i++) {
      const rowClaimNo = String(data[i][1] || '').trim(); // B열: 접수번호
      const rowNameVal = String(data[i][2] || '').trim(); // C열: 채권자명
      if (rowNameVal && rowClaimNo) {
        dbData.push({ claimNo: rowClaimNo, name: rowNameVal });
      }
    }

    return res.json({ success: true, list: dbData });
  } catch (error) {
    return res.status(500).json({ success: false, message: error.toString() });
  }
});

// 1. 성함으로 시부인표 데이터 조회 API
app.get('/api/search', (req, res) => {
  try {
    const { name } = req.query;
    if (!name || name.trim().length < 2) {
      return res.status(400).json({ success: false, message: '이름을 2글자 이상 정확히 입력해 주세요.' });
    }

    if (!fs.existsSync(EXCEL_PATH)) {
      return res.status(500).json({ success: false, message: '시부인표 엑셀 파일을 찾을 수 없습니다.' });
    }

    const workbook = XLSX.readFile(EXCEL_PATH);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const searchName = name.replace(/\s+/g, '').trim();
    const matches = [];

    // 2번째 행부터 탐색 (1행은 헤더)
    for (let i = 1; i < data.length; i++) {
      const rowClaimNo = String(data[i][1] || '').trim(); // B열: 접수번호
      const rowNameVal = String(data[i][2] || '').trim(); // C열: 채권자명

      if (!rowNameVal || !rowClaimNo) continue;

      const rowNameNorm = rowNameVal.replace(/\s+/g, '');
      let isMatch = (rowNameNorm === searchName);

      // 공동 채권자 구분 처리 (쉼표 구분)
      if (!isMatch && rowNameVal.includes(',')) {
        const parts = rowNameVal.split(',');
        for (let p of parts) {
          if (p.replace(/\s+/g, '').trim() === searchName) {
            isMatch = true;
            break;
          }
        }
      }

      if (isMatch) {
        matches.push({
          claimNo: rowClaimNo,
          name: rowNameVal
        });
      }
    }

    if (matches.length === 0) {
      return res.json({ success: false, message: '일치하는 채권자 정보를 찾을 수 없습니다.' });
    }

    return res.json({ success: true, matches });
  } catch (error) {
    return res.status(500).json({ success: false, message: error.toString() });
  }
});

// 2. 원본 PDF 템플릿 파일 제공 API
app.get('/api/template', (req, res) => {
  if (fs.existsSync(PDF_TEMPLATE_PATH)) {
    res.sendFile(PDF_TEMPLATE_PATH);
  } else {
    res.status(404).send('PDF 템플릿 파일을 찾을 수 없습니다.');
  }
});

// 3. 한글 맑은 고딕 폰트 제공 API
app.get('/api/font', (req, res) => {
  if (fs.existsSync(FONT_PATH)) {
    res.sendFile(FONT_PATH);
  } else {
    res.status(404).send('맑은 고딕 폰트를 찾을 수 없습니다.');
  }
});

// 4. 작성 완료 및 로컬 PDF 저장 API
app.post('/api/submit', (req, res) => {
  try {
    const { name, unitNo, claimNo, phone, date, pdfBytes } = req.body;

    if (!name || !unitNo || !claimNo || !pdfBytes) {
      return res.status(400).json({ success: false, message: '필수 데이터가 누락되었습니다.' });
    }

    // 4.1. PDF 파일 로컬 저장
    const buffer = Buffer.from(pdfBytes, 'base64');
    const safeName = name.replace(/[\/\\:\*\?"<>\|]/g, ''); // 파일명 안전 처리
    const safeUnit = unitNo.replace(/[\/\\:\*\?"<>\|]/g, '');
    const pdfFilename = `채권신고 취하서_${safeName}_${safeUnit}.pdf`;
    const pdfSavePath = path.join(OUTPUT_DIR, pdfFilename);

    fs.writeFileSync(pdfSavePath, buffer);

    // 4.2. 접수현황 CSV 누적 저장 (엑셀 오픈용 BOM 추가)
    const isNew = !fs.existsSync(CSV_PATH);
    const csvContent = [];
    
    if (isNew) {
      // UTF-8 BOM 추가하여 엑셀에서 한글 안 깨지게 처리
      csvContent.push('\ufeff접수시간,성명,동호수,채권번호,연락처,저장파일명');
    }
    
    const nowStr = new Date().toLocaleString('ko-KR');
    csvContent.push(`"${nowStr}","${name}","${unitNo}","${claimNo}","${phone}","${pdfFilename}"`);
    
    fs.appendFileSync(CSV_PATH, csvContent.join('\n') + '\n', 'utf8');

    console.log(`[접수 완료] PDF 저장: ${pdfFilename} / 대장 기록 완료`);
    return res.json({ success: true, filename: pdfFilename });
  } catch (error) {
    console.error(error);
    return res.status(500).json({ success: false, message: error.toString() });
  }
});

// 서버 기동
app.listen(PORT, () => {
  console.log(`====================================================`);
  console.log(`🚀 채권신고 취하서 로컬 타이핑 시스템이 시작되었습니다.`);
  console.log(`👉 브라우저를 열고 다음 주소로 접속해 주세요:`);
  console.log(`   http://localhost:${PORT}`);
  console.log(`====================================================`);
});
