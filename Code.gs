function doGet(e) {
  var mondayDate = PropertiesService.getScriptProperties().getProperty("MONDAY_DATE") || null;
  var sundayDate = PropertiesService.getScriptProperties().getProperty("SUNDAY_DATE") || null;
  
  // HTML 템플릿에 dateRange 값을 전달
  var template = HtmlService.createTemplateFromFile("index");
  template.mondayDate = mondayDate;
  template.sundayDate = sundayDate;

  return template.evaluate()
                 .setTitle('내사무실 야근확인 페이지')
                 .addMetaTag('viewport', 'width=device-width, maximum-scale=1, initial-scale=1, user-scalable=no'); // HTML 파일을 반환
}

/**
 * loadDataByName. index key is name.
 * @param { name } search UserName if Exist.
 * @return { name , day : { bOverTime, iTransportCost } } / null
 */
function getWeeklyDataByName(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("비용정산DB_Web");
  const data = sheet.getDataRange().getValues();
  
  // 데이터에서 해당 이름을 찾기
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === name) {
      // 이름에 해당하는 데이터가 있을 경우 반환
      return {
        name: data[i][0],
        monday: { overtime: data[i][1], transportCost: data[i][2] },
        tuesday: { overtime: data[i][3], transportCost: data[i][4] },
        wednesday: { overtime: data[i][5], transportCost: data[i][6] },
        thursday: { overtime: data[i][7], transportCost: data[i][8] },
        friday: { overtime: data[i][9], transportCost: data[i][10] },
        saturday: { overtime: data[i][11], transportCost: data[i][12] },
        sunday: { overtime: data[i][13], transportCost: data[i][14] }
      };
    }
  }
  
  // 해당 이름이 없으면 null 반환
  return null;
}


/**
 * loadDataByName. index key is name.
 * @param { name } search UserName if Exist.
 * @return { name , day : { bOverTime, iTransportCost } } / null
 */
function getAllWeeklyData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("비용정산DB_Web");
  const data = sheet.getDataRange().getValues();
  let result = {};
  
  // 각 행을 순회하며 데이터를 변환
  for (let i = 0; i < data.length; i++) {
    const name = data[i][0];
    if (!name) continue; // 이름이 없으면 건너뜁니다.
    
    // 각 요일의 데이터를 원하는 형식으로 저장 (여기서는 월~일을 mon, tue, wed, thu, fri, sat, sun으로 사용)
    result[name] = {
      monday: { overtime: data[i][1], transportCost: data[i][2] },
      tuesday: { overtime: data[i][3], transportCost: data[i][4] },
      wednesday: { overtime: data[i][5], transportCost: data[i][6] },
      thursday: { overtime: data[i][7], transportCost: data[i][8] },
      friday: { overtime: data[i][9], transportCost: data[i][10] },
      saturday: { overtime: data[i][11], transportCost: data[i][12] },
      sunday: { overtime: data[i][13], transportCost: data[i][14] }
    };
  }
  
  return result;
}

/**
 * saveDataByName. index key is name.
 * @param { name } search UserName if Exist.
 * @if UserName isn't exist in DBSheet, create and save OverTime Data.
 * @return { popup_Message }
 */
function saveWeeklyData(name, weeklyData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("비용정산DB_Web");

  // 이름에 해당하는 데이터가 이미 있는지 확인
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;

  // 이름이 이미 존재하는지 확인
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === name) {
      rowIndex = i + 1; // 해당하는 행을 찾으면, 그 행 번호를 기록
      break;
    }
  }

  if (rowIndex === -1) {
    // 새로운 이름일 경우, 데이터 추가
    rowIndex = sheet.getLastRow() + 1;
    sheet.getRange(rowIndex, 1).setValue(name);
  }

  // 이름에 맞는 데이터를 저장 (7일 데이터 저장)
  for (let i = 0; i < weeklyData.length; i++) {
    sheet.getRange(rowIndex, i * 2 + 2).setValue(weeklyData[i].overtime); // 야근 여부
    sheet.getRange(rowIndex, i * 2 + 3).setValue(weeklyData[i].transportCost); // 교통비
  }

  return "데이터가 저장되었습니다!"; // 성공 메시지 반환
}

// 해당트리거 스위치는 스트립트 앱의 <트리거>에 의해 실행됨 
// 트리거 실행시간은 <트리거>탭에서 수정가능하며 트리거 실행이후에는 다음주 데이터만 등록가능
// (결재올릴 야근정보의 수정이 필요할경우, <비용정산DB> 시트의 내용을 수정하여 결재해야함)
function adminTriggerSwitch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getSheetByName("비용정산DB_Web");
  var prevWeekSheet = ss.getSheetByName("비용정산DB");

  // 1. 결재 시트 초기화 (비용정산 DB)
  var lastRow = prevWeekSheet.getLastRow();
  var lastColumn = prevWeekSheet.getLastColumn();
  prevWeekSheet.getRange(4, 1, lastRow, lastColumn).clearContent();

  // 2. Web시트 데이터를 결재 시트로 복사 (비용정산DB_Web -> 비용정산DB)
  var data = activeSheet.getDataRange().getValues();
  if (data.length > 0 && data[0].length > 0) {
    prevWeekSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }

  // 3. 현재 사용중인 시트 데이터를 삭제 (필요 시 백업추가필요)
  var lastRow = activeSheet.getLastRow();
  var lastColumn = activeSheet.getLastColumn();
  activeSheet.getRange(4, 1, lastRow, lastColumn).clearContent();

  // 4. 차주 날짜 등록: 트리거 실행 날짜가 포함된 주의 월요일부터 일요일까지를 Q1에 등록
  var triggerDate = new Date();  // 트리거 실행 날짜
  var day = triggerDate.getDay(); // 0(일) ~ 6(토)
  var monday = new Date(triggerDate);
  
  // 만약 실행일이 일요일(0)인 경우, 다음날부터
  // 월 ~ 토인경우 해당주 월요일
  if(day === 0) {
    monday.setDate(triggerDate.getDate() + 1);
  } else {
    monday.setDate(triggerDate.getDate() - (day - 1));
  }
  
  var sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  
  // 날짜를 "YYYY/MM/DD" 형식으로 변환하는 함수
  function formatDate(date) {
    var year = date.getFullYear();
    var month = date.getMonth() + 1;
    if (month < 10) month = "0" + month;
    var d = date.getDate();
    if (d < 10) d = "0" + d;
    return year + "/" + month + "/" + d;
  }
  
  var formattedMonday = formatDate(monday);
  var formattedSunday = formatDate(sunday);

  // 날짜형식 : "(YYYY/MM/DD ~ YYYY/MM/DD)"
  var dateRangeString = "(" + formattedMonday + " ~ " + formattedSunday + ")";
  
  // "비용정산DB_Web" 시트의 Q1 셀에 날짜 범위 저장
  activeSheet.getRange("Q1").setValue(dateRangeString);
  PropertiesService.getScriptProperties().setProperty("MONDAY_DATE", formattedMonday);
  PropertiesService.getScriptProperties().setProperty("SUNDAY_DATE", formattedSunday);
}