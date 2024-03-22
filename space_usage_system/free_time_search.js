// import "google-apps-script";

function doGet(p) {
  let param = p.parameter;
  let startDateTime = new Date(param.startDateTime);
  let endDateTime = new Date(param.endDateTime);


  // 尋找該時間段所有日曆事件
  let freeSpace = [];
  allCalendarId_array.forEach((calendarID) => {
    let calendar = CalendarApp.getCalendarById(calendarID);
    let events = calendar.getEvents(startDateTime, endDateTime);
    if (events.length == 0) {
      freeSpace.push(calendar.getName());
    }
  });

  return ContentService.createTextOutput(JSON.stringify(freeSpace))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 教室清單(合併兩行)
 */
function mergeNewRow() {
  let ss3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("空堂查詢");
  for (let i = 4; i < 51; i++) {
    ss3.getRange(i, 1, 1, 2).merge();
  }
}

/**
 * 尋找空堂教室
 */
function searchFreeSapce() {
  let ss3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("空堂查詢");
  let oldLastRow = ss3.getLastRow();
  if (oldLastRow > 3) {
    ss3.deleteRows(4, oldLastRow - 3);
    ss3.insertRows(20, oldLastRow - 3);
    mergeNewRow();
  }
  let startDateTime = ss3.getRange(1, 2).getValue();
  let endDateTime = ss3.getRange(2, 2).getValue();
  Logger.log(startDateTime);
  Logger.log(endDateTime);
  // 尋找該時間段所有日曆事件
  allCalendarId_array.forEach((calendarID) => {
    let calendar = CalendarApp.getCalendarById(calendarID);
    let events = calendar.getEvents(startDateTime, endDateTime);
    let lastRow = ss3.getLastRow();
    if (events.length == 0) {
      ss3.getRange(lastRow + 1, 1).setValue(calendar.getName());
    }
  });
}

function setTime() {
  let ss3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("空堂查詢");
  ss3
    .getRange(1, 2)
    .setValue(new Date(new Date("2022/4/20").getTime() + 15 * 60 * 60 * 1000));
  ss3
    .getRange(2, 2)
    .setValue(new Date(new Date("2022/4/20").getTime() + 16 * 60 * 60 * 1000));
}
