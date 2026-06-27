// =====================================================
// Google Apps Script - 설문 응답 수신 웹앱
// =====================================================
// 사용법:
// 1. Google Sheets 새 파일 열기
// 2. 상단 메뉴 → 확장 프로그램 → Apps Script
// 3. 이 코드 전체를 붙여넣기
// 4. SHEET_ID를 실제 스프레드시트 ID로 교체
// 5. 배포 → 새 배포 → 웹앱 → 액세스: 모든 사용자 → 배포
// 6. 배포 URL을 복사해서 HTML 파일의 SCRIPT_URL에 붙여넣기
// =====================================================

const SHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // ← 여기에 스프레드시트 ID 입력
const SHEET_NAME = '설문응답';

const HEADERS = [
  '제출일시',
  '성함',
  '연락처',
  '이용서비스',
  '이용횟수',
  // 에너지바
  '에너지바_전체만족도',
  '에너지바_맛과식감',
  '에너지바_재료품질',
  '에너지바_영양성분',
  '에너지바_포장',
  '에너지바_배송',
  '에너지바_다양성',
  '에너지바_구매우선순위',
  '에너지바_가격인식',
  '에너지바_신제품요청',
  // 교육
  '교육_전체만족도',
  '교육_강좌내용',
  '교육_강사역량',
  '교육_일정',
  '교육_교재',
  '교육_환경',
  '교육_수강료',
  '교육_강사소통',
  '교육_수업방식',
  '교육_희망강좌',
  // 고객경험
  '유입경로',
  'CS응대만족도',
  '불편사항',
  // NPS
  'NPS점수',
  'NPS이유',
  '재구매의향',
  // 자유의견
  '자유의견',
];

function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);

    // 시트가 없으면 생성 후 헤더 추가
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
      // 헤더 스타일
      const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setBackground('#4a7c59');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const data = JSON.parse(e.postData.contents);

    const row = [
      new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' }),
      data.name || '',
      data.contact || '',
      [].concat(data.service_type || []).join(', '),
      data.frequency || '',
      // 에너지바
      data.bar_overall || '',
      data.bar_taste || '',
      data.bar_ingredient || '',
      data.bar_nutrition || '',
      data.bar_packaging || '',
      data.bar_delivery || '',
      data.bar_variety || '',
      [].concat(data.bar_priority || []).join(', '),
      data.bar_price || '',
      data.bar_request || '',
      // 교육
      data.edu_overall || '',
      data.edu_content || '',
      data.edu_instructor || '',
      data.edu_schedule || '',
      data.edu_material || '',
      data.edu_venue || '',
      data.edu_price || '',
      data.edu_communication || '',
      data.edu_format || '',
      data.edu_request || '',
      // 고객경험
      [].concat(data.channel || []).join(', '),
      data.cs_rating || '',
      data.pain_point || '',
      // NPS
      data.nps || '',
      data.nps_reason || '',
      data.repurchase || '',
      // 자유의견
      data.free_comment || '',
    ];

    sheet.appendRow(row);

    // 새 행 배경색 교차 적용
    const lastRow = sheet.getLastRow();
    if (lastRow % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, HEADERS.length).setBackground('#f4f6f0');
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// CORS 대응용 GET 핸들러
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ result: 'ok', message: '설문 서버 정상 작동 중' }))
    .setMimeType(ContentService.MimeType.JSON);
}
