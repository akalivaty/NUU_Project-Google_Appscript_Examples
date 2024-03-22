function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("日曆功能")
    .addItem("建立日曆事件", "select_create")
    .addItem('取消事件', 'select_cancel_events')
    .addSeparator()
    .addItem("共用所有教室日曆", "share_calendar")
    .addItem("訂閱所有教室日曆", "subscribe_calendar")
    .addSeparator()
    .addItem('更新週空堂表', 'updateFreeTime')
    .addToUi();
}

/**
 * 共用日曆，授予權限
 */
function share_calendar() {
  let userEmail = SpreadsheetApp.getUi().prompt('輸入Email').getResponseText();
  var resource = {
    'scope': {
      'type': 'user',
      'value': userEmail
    },
    'role': 'owner'
  };

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('正在共用給 ' + userEmail);
    let shared = [];
    for (var i = 0; i < allCalendarId_array.length; i++) {
      Calendar.Acl.insert(resource, allCalendarId_array[i]);
      shared.push(CalendarApp.getCalendarById(allCalendarId_array[i]).getName());
    }
    let text = '已共用日曆\n';
    for (let i = 0; i < shared.length; i++) {
      text = text + shared[i] + '\n';
    }
    SpreadsheetApp.getUi().alert(text);
  } catch (e) {
    Logger.log(e.message);
    let alertText = '已擁有權限 或 ' + e.message;
    throw alertText;
  }
}

/**
 * 訂閱日曆
 */
function subscribe_calendar() {
  let subscribed = [];
  for (let i = 0; i < allCalendarId_array.length; i++) {
    var calendar = CalendarApp.subscribeToCalendar(allCalendarId_array[i]);
    subscribed.push(calendar.getName());
  }
  let text = '已訂閱日曆\n';
  for (let i = 0; i < subscribed.length; i++) {
    text = text + subscribed[i] + '\n';
  }
  SpreadsheetApp.getUi().alert(text);
}

/**
 * 選取範圍建立事件
 */
function select_create() {
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('表單回應 1');     // 取得工作表
  let data = ss.getDataRange().getValues();                                       // 取得資料
  let selectRange = SpreadsheetApp.getActiveRangeList();                          // 取得選取範圍
  let listLength = selectRange.getRanges().length;                                // 取得個選取範圍列數
  let eventCount = 0;                                                             // 事件計數
  for (let i = 0; i < listLength; i++, eventCount++) {
    let rows = selectRange.getRanges()[i].getNumRows();
    let startRow = selectRange.getRanges()[i].getRow();
    createCalendarEvent(data[startRow - 1]);
    let changeRange = ss.getRange(startRow, 1, 1, 14);                            // 取得該列範圍
    changeRange.setBackgroundRGB(248, 204, 204);                                  // 設定背景為粉紅色
    let html = get_html_content(startRow - 1, 'replyMail');
    MailApp.sendEmail(data[startRow - 1][12], '您的空間使用申請已登記完成', '', { htmlBody: html });

    // 當選取範圍有複數列
    if (rows > 1) {
      for (let j = 0; j < rows - 1; j++) {
        createCalendarEvent(data[startRow + j]);
        let changeRange = ss.getRange(startRow + j + 1, 1, 1, 14);    // 取得該列範圍
        changeRange.setBackgroundRGB(248, 204, 204);                  // 設定背景為粉紅色
        try {
          let html = get_html_content(startRow + j, 'replyMail');
          MailApp.sendEmail(data[startRow + j][12], '您的空間使用申請已登記完成', '', { htmlBody: html });
        } catch (e) {
          Logger.log(e);
          // throw '申請人填寫的Email無效, 未寄出通知';
        } finally {
          eventCount++;
        }
      }
    }
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('已建立 ' + eventCount + ' 個日曆事件');
}

/**
 * 取得回信內容
 * @param {number} row - data[row]
 * @param {string} mailType - *.html
 * @returns {html}
 */
function get_html_content(row, mialType) {
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('表單回應 1');     // 取得工作表
  let data = ss.getDataRange().getValues();                                       // 取得資料
  var template = HtmlService.createTemplateFromFile(mialType);
  template.eventSerialNum = data[row][15];
  template.space = data[row][3];
  template.startTime = date_time_regex(data[row][4], data[row][6]);
  template.endTime = date_time_regex(data[row][5], data[row][7]);
  template.purpose = data[row][8];
  template.eventName = data[row][9];
  return template.evaluate().getContent();
}

/**
 * 時間轉換為字串
 * @param {string} date
 * @param {string} time
 * @returns {string}
 */
function date_time_regex(date, time) {
  let year = addZero(date.getFullYear());
  let month = addZero(date.getMonth() + 1);
  let day = addZero(date.getDate());
  let hour = addZero(time.getHours());
  let minute = addZero(time.getMinutes());
  let second = addZero(time.getSeconds());
  return year + '/' + month + '/' + day + ' - ' + hour + ':' + minute + ':' + second;
}

/**
 * 時間轉換為字串保持兩位數
 * @param {number} number
 * @returns {string}
 */
function addZero(number) {
  if (number < 10) {
    return '0' + number.toString();
  }
  else {
    return number.toString();
  }
}

/**
 * 檢查時間合法性
 */
function check_valid_time() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('表單回應 1');            // 取得工作表
  let data = sheet.getDataRange().getValues();                                              // 取得資料
  let lastRow = sheet.getLastRow();                                                         // 取得最後一列數
  let lastData = data[lastRow - 1];
  let newStartDateTime = convertTimeFormat(data[lastRow - 1][4], data[lastRow - 1][6]);
  let newEndDateTime = convertTimeFormat(data[lastRow - 1][5], data[lastRow - 1][7]);

  // 確認新資料時間是否合法
  if (newStartDateTime.getTime() > newEndDateTime.getTime()) {
    let html = get_html_content(lastRow - 1, 'timeInvalidMail');
    MailApp.sendEmail(lastData[12], '您的空間申請時段錯誤，請確認時段後再重新申請', '', { htmlBody: html });
    sheet.deleteRow(lastRow);
  }
  else {
    let calendar = get_calendar_by_ID(lastData[3].split('(')[0]);
    let startDateTime = convertTimeFormat(lastData[4], lastData[6]);
    let endDateTime = convertTimeFormat(lastData[5], lastData[7]);
    let events = calendar.getEvents(startDateTime, endDateTime);
    if (events.length != 0) {
      let html = get_html_content(lastRow - 1, 'timeInvalidMail');
      MailApp.sendEmail(lastData[12], '您的空間申請時段已被預約，請確認時段後再重新申請', '', { htmlBody: html });
      sheet.deleteRow(lastRow);
    }
  }
}

/**
 * 週期性刪除過期資料 (需設trigger)
 */
function delData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('表單回應 1');    // 取得工作表
  var data = sheet.getDataRange().getValues();                                      // 取得資料
  var lastRow = sheet.getLastRow();                                                 // 取得最後一列數
  var currentDate = new Date();                                                     // 建立日期變數
  for (var i = lastRow - 1; i > 0; i--) // 從最新資料開始刪除
    if (data[i][4] < currentDate) {                                                 // 若【預計使用起始日期】<【現在日期】
      sheet.deleteRow(i + 1);                                                       // 刪除一整列
      Logger.log('Delete state time : ' + data[i][4]);
    }
}

/**
 * 給定日期與時間，轉換為Date物件
 * @param {string} date
 * @param {string} time
 * @returns {Date}
 */
function convertTimeFormat(date, time) {
  var timeHour = time.getHours() * 1000 * 60 * 60;
  var timeMinutes = time.getMinutes() * 1000 * 60;
  return new Date(date.getTime() + timeHour + timeMinutes);
}

/**
 * 建立一筆日曆事件
 * @param {string[]} newEvent
 */
function createCalendarEvent(newEvent) {
  var timestamp = Utilities.formatDate(newEvent[0], "GMT+8", "yyyy/MM/dd - HH:mm:ss");
  var person = newEvent[1];
  var tel = newEvent[2];
  var space = newEvent[3].split('(')[0];
  var startDate = newEvent[4];
  var endDate = newEvent[5];
  var startTime = newEvent[6];
  var endTime = newEvent[7];
  var purpose = newEvent[8];
  var title = newEvent[9];
  var mail = newEvent[11];

  var startDateTime = convertTimeFormat(startDate, startTime);
  var endDateTime = convertTimeFormat(endDate, endTime);
  Logger.log(startDateTime);
  Logger.log(endDateTime);

  try {
    switch (space) {
      case "B2-101階梯教室":
        var calendar = CalendarApp.getCalendarById("CALENDAR_ID");
        break;
      case "B2-105創客空間":
        var calendar = CalendarApp.getCalendarById("CALENDAR_ID");
        break;
      case "B2-201講義教室":
        var calendar = CalendarApp.getCalendarById("CALENDAR_ID");
        break;
      case "B2-202講義教室":
        var calendar = CalendarApp.getCalendarById("CALENDAR_ID");
        break;
      case "B2-203講義教室":
        var calendar = CalendarApp.getCalendarById("CALENDAR_ID");
        break;
      case "B2-204講義教室":
        var calendar = CalendarApp.getCalendarById("CALENDAR_ID");
        break;
      case "B2-205講義教室":
        var calendar = CalendarApp.getCalendarById("CALENDAR_ID");
        break;
      case "B2-206講義教室":
        var calendar = CalendarApp.getCalendarById("CALENDAR_ID");
        break;
      case "B2-211講義教室":
        var calendar = CalendarApp.getCalendarById("CALENDAR_ID");
        break;
      case "B2-213講義教室":
        var calendar = CalendarApp.getCalendarById("CALENDAR_ID");
        break;
      case "B2-214講義教室":
        var calendar = CalendarApp.getCalendarById("CALENDAR_ID");
        // Logger.log(calendar);
        break;
      case "B2-215講義教室":
        var calendar = CalendarApp.getCalendarById("CALENDAR_ID");
        break;
      case "B2-216講義教室":
        var calendar = CalendarApp.getCalendarById("CALENDAR_ID");
        break;
      case "B2-302研討室":
        var calendar = CalendarApp.getCalendarById("CALENDAR_ID");
        break;
      case "B2-309研討室":
        var calendar = CalendarApp.getCalendarById("CALENDAR_ID");
        break;
      case "B2-313系會議室":
        var calendar = CalendarApp.getCalendarById("CALENDAR_ID");
        break;
      default:
        break;
    }
    var description = purpose + "\n\n" + person + '\n' + tel + '\n' + mail + '\n' + timestamp;
    var event = calendar.createEvent(title, startDateTime, endDateTime, { description: description }); // 建立日曆事件
    Logger.log(space + " " + event.getDescription() + " event created successfully. ");
  } catch (e) {
    Logger.log(e);
    throw '請先訂閱日曆';
  }
}

/**
 * 表單提交後，預設填入序列號
 * 
 */
function defaultFill() {
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('表單回應 1');
  let lastRow = ss.getLastRow();
  let dateTime = new Date();
  let year = addZero(dateTime.getFullYear()).split('0')[1];
  let month = addZero(dateTime.getMonth() + 1);
  let day = addZero(dateTime.getDate());
  let hour = addZero(dateTime.getHours());
  let minute = addZero(dateTime.getMinutes());
  let second = addZero(dateTime.getSeconds());
  let eventSerialNum = year + month + day + hour + minute + second;
  ss.getRange(lastRow, 14).setValue(eventSerialNum);

}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

function createCalendar() {
  var calendar = CalendarApp.createCalendar("B2-101", { timeZone: "Asia/Taipei" });
  Logger.log('Created the calendar "%s", with the ID "%s".', calendar.getName(), calendar.getId());

  calendar = CalendarApp.createCalendar("B2-105", { timeZone: "Asia/Taipei" });
  Logger.log('Created the calendar "%s", with the ID "%s".', calendar.getName(), calendar.getId());

  var space = "B2-20";
  for (var i = 1; i <= 6; i++) {
    calendar = CalendarApp.createCalendar(space + i, { timeZone: "Asia/Taipei" });
    Logger.log('Created the calendar "%s", with the ID "%s".', calendar.getName(), calendar.getId());
  }

  space = "B2-2";
  for (var i = 13; i <= 16; i++) {
    calendar = CalendarApp.createCalendar(space + i, { timeZone: "Asia/Taipei" });
    Logger.log('Created the calendar "%s", with the ID "%s".', calendar.getName(), calendar.getId());
  }

  calendar = CalendarApp.createCalendar("B2-302", { timeZone: "Asia/Taipei" });
  Logger.log('Created the calendar "%s", with the ID "%s".', calendar.getName(), calendar.getId());

  calendar = CalendarApp.createCalendar("B2-309", { timeZone: "Asia/Taipei" });
  Logger.log('Created the calendar "%s", with the ID "%s".', calendar.getName(), calendar.getId());

  calendar = CalendarApp.createCalendar("B2-313", { timeZone: "Asia/Taipei" });
  Logger.log('Created the calendar "%s", with the ID "%s".', calendar.getName(), calendar.getId());
}
