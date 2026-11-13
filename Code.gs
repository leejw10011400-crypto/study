/**
 * Google Apps Script 웹 앱 백엔드 로직
 * 대상 스프레드시트 ID: 1RbmVL0ah5L6UK00_UAPvSWX-5-Sx4dfXHs44FyYTsi0
 */

const SPREADSHEET_ID = '1RbmVL0ah5L6UK00_UAPvSWX-5-Sx4dfXHs44FyYTsi0';
const SHEET_NAME = '공부기록';

/**
 * 웹 앱 접속 시 HTML 페이지를 렌더링합니다.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('오늘의 공부 기록장')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 스프레드시트를 가져오고, 헤더가 없으면 자동으로 생성합니다.
 */
function getTargetSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  
  if (sheet.getLastRow() === 0) {
    const headers = [
      '날짜',
      '공부 시간대 (구간별)',
      '오늘 총 공부량',
      '오늘 공부한 내용',
      '내일 공부할 계획',
      '기록 일시'
    ];
    sheet.appendRow(headers);
    
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4F46E5')
               .setFontColor('#FFFFFF')
               .setFontWeight('bold')
               .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * 클라이언트에서 전송한 공부 기록 데이터를 시트에 저장합니다.
 */
function saveStudyLog(data) {
  try {
    const sheet = getTargetSheet();
    
    // 여러 공부 시간대 구간(타임) 정리
    const timeRangesText = Array.isArray(data.timeRanges) && data.timeRanges.length > 0
      ? data.timeRanges.map((range, idx) => `${idx + 1}타임: ${range.start} ~ ${range.end}`).join('\n')
      : '선택 없음';

    // 오늘 공부한 내용 정리
    const studiedItemsText = Array.isArray(data.studiedItems) 
      ? data.studiedItems.filter(item => item.trim() !== '').map((item, idx) => `${idx + 1}. ${item.trim()}`).join('\n')
      : String(data.studiedItems || '');

    const timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    
    const rowData = [
      data.date,             
      timeRangesText,        
      data.durationText,     
      studiedItemsText,      
      data.tomorrowPlan,     
      timestamp              
    ];
    
    sheet.appendRow(rowData);
    
    const lastRow = sheet.getLastRow();
    const rowRange = sheet.getRange(lastRow, 1, 1, rowData.length);
    rowRange.setWrap(true).setVerticalAlignment('middle');
    sheet.getRange(lastRow, 1, 1, 3).setHorizontalAlignment('center');
    
    return {
      success: true,
      message: '공부 기록이 스프레드시트에 성공적으로 저장되었습니다!'
    };
    
  } catch (error) {
    Logger.log('Error saving study log: ' + error.toString());
    return {
      success: false,
      message: '저장 중 오류가 발생했습니다: ' + error.message
    };
  }
}

/**
 * 최근 기록을 불러옵니다.
 */
function getRecentLogs() {
  try {
    const sheet = getTargetSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return [];
    }
    
    const numRows = Math.min(5, lastRow - 1);
    const startRow = lastRow - numRows + 1;
    const values = sheet.getRange(startRow, 1, numRows, 5).getValues();
    
    return values.reverse().map(row => {
      let dateStr = row[0];
      if (row[0] instanceof Date) {
        dateStr = Utilities.formatDate(row[0], 'Asia/Seoul', 'yyyy-MM-dd');
      }
      return {
        date: dateStr,
        timeRanges: String(row[1] || ''),
        duration: String(row[2] || ''),
        studied: String(row[3] || ''),
        tomorrow: String(row[4] || '')
      };
    });
  } catch (error) {
    Logger.log('Error fetching logs: ' + error.toString());
    return [];
  }
}
