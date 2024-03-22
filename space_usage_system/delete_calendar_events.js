function auto_delete_event() {
  let ss1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("表單回應 1"); // 取得工作表
  let data1 = ss1.getDataRange().getValues(); // 取得資料
  let lastRow1 = ss1.getLastRow();
  let ss2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("取消申請"); // 取得工作表
  let data2 = ss2.getDataRange().getValues(); // 取得資料
  let lastRow2 = ss2.getLastRow();

  for (let i = 1; i < lastRow2; i++) {
    for (let j = 1; j < lastRow1; j++) {
      if (data2[i][1] == data1[j][13]) {
        let calendar = get_calendar_by_ID(data1[j][3].split("(")[0]);
        let startDateTime = timeRegex(data1[j][4], data1[j][6]);
        let endDateTime = timeRegex(data1[j][5], data1[j][7]);
        deleteFromCalendar(calendar, startDateTime, endDateTime);
        try {
          let mailContent =
            "【使用空間】" +
            data1[j][3] +
            "\n【預計使用開始時間】" +
            date_time_regex(data1[j][4], data1[j][6]) +
            "\n【預計使用結束時間】" +
            date_time_regex(data1[j][5], data1[j][7]) +
            "\n【用途說明】" +
            data1[j][8] +
            "\n【活動名稱】" +
            data1[j][9];
          MailApp.sendEmail(
            data1[j][12],
            "您的空間使用申請已取消",
            mailContent
          );
          ss1.deleteRow(j + 1);
          ss2.deleteRow(i + 1);
        } catch (e) {
          Logger.log(e);
        }
      }
    }
  }
}

/**
 * 功能: 表單選擇取消
 */
function select_cancel_events() {
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("表單回應 1"); // 取得工作表
  let data = ss.getDataRange().getValues(); // 取得資料
  let selectRange = SpreadsheetApp.getActiveRangeList(); // 取得選取範圍
  let rangeListLength = selectRange.getRanges().length; // 取得個選取範圍列數
  let eventCount = 0; // 事件計數
  for (let i = 0; i < rangeListLength; i++, eventCount++) {
    let startRow = selectRange.getRanges()[i].getRow();

    let calendar = get_calendar_by_ID(data[startRow - 1][3].split("(")[0]);
    let startDateTime = timeRegex(data[startRow - 1][4], data[startRow - 1][6]);
    let endDateTime = timeRegex(data[startRow - 1][5], data[startRow - 1][7]);
    deleteFromCalendar(calendar, startDateTime, endDateTime);
    try {
      let mailContent =
        "【使用空間】" +
        data[startRow - 1][3] +
        "\n【預計使用開始時間】" +
        date_time_regex(data[startRow - 1][4], data[startRow - 1][6]) +
        "\n【預計使用結束時間】" +
        date_time_regex(data[startRow - 1][5], data[startRow - 1][7]) +
        "\n【用途說明】" +
        data[startRow - 1][8] +
        "\n【活動名稱】" +
        data[startRow - 1][9];

      MailApp.sendEmail(
        data[startRow - 1][12],
        "您的空間使用申請已取消",
        mailContent
      );
      ss.deleteRow(startRow);
    } catch (e) {
      Logger.log(e);
      throw "申請人填寫的Email無效, 未寄出通知";
    }

    // 若為多選
    let rows = selectRange.getRanges()[i].getNumRows();
    if (rows > 1) {
      for (let j = 0; j < rows - 1; j++) {
        calendar = get_calendar_by_ID(data[startRow + j][3].split("(")[0]);
        startDateTime = timeRegex(data[startRow + j][4], data[startRow + j][6]);
        endDateTime = timeRegex(data[startRow + j][5], data[startRow + j][7]);
        deleteFromCalendar(calendar, startDateTime, endDateTime);
        try {
          let mailContent =
            "【使用空間】" +
            data[startRow + j][3] +
            "\n【預計使用開始時間】" +
            date_time_regex(data[startRow + j][4], data[startRow + j][6]) +
            "\n【預計使用結束時間】" +
            date_time_regex(data[startRow + j][5], data[startRow + j][7]) +
            "\n【用途說明】" +
            data[startRow + j][8] +
            "\n【活動名稱】" +
            data[startRow + j][9];

          MailApp.sendEmail(
            data[startRow + j][12],
            "您的空間使用申請已取消",
            mailContent
          );
          ss.deleteRow(startRow + j + 1);
          eventCount++;
        } catch (e) {
          Logger.log(e);
          throw "申請人填寫的Email無效, 未寄出通知";
        }
      }
    }
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "已取消 " + eventCount + " 個申請"
  );
}

/**
 * 從給定的日曆 ID 刪除指定時間事件
 * @param {String} calendarID
 * @param {Number} startDateTime
 * @param {Number} endDateTime
 */
function deleteFromCalendar(calendarID, startDateTime, endDateTime) {
  Logger.log("calendar: " + calendarID.getName());
  Logger.log("startDateTime: " + startDateTime);
  Logger.log("endDateTime: " + endDateTime);
  let events = calendarID.getEvents(startDateTime, endDateTime);
  if (events.length != 1) {
    throw "此時段 沒有/有兩個以上 空間申請，請確認並手動刪除";
  }
  events[0].deleteEvent();
}

/**
 * 按空間名稱傳回日曆 ID
 * @param {String} space
 * @returns {String}
 */
function get_calendar_by_ID(space) {
  switch (space) {
    case "B2-101階梯教室":
      return CalendarApp.getCalendarById(
        "CALENDAR_ID"
      );
    case "B2-105創客空間":
      return CalendarApp.getCalendarById(
        "CALENDAR_ID"
      );
    case "B2-201講義教室":
      return CalendarApp.getCalendarById(
        "CALENDAR_ID"
      );
    case "B2-202講義教室":
      return CalendarApp.getCalendarById(
        "CALENDAR_ID"
      );
    case "B2-203講義教室":
      return CalendarApp.getCalendarById(
        "CALENDAR_ID"
      );
    case "B2-204講義教室":
      return CalendarApp.getCalendarById(
        "CALENDAR_ID"
      );
    case "B2-205講義教室":
      return CalendarApp.getCalendarById(
        "CALENDAR_ID"
      );
    case "B2-206講義教室":
      return CalendarApp.getCalendarById(
        "CALENDAR_ID"
      );
    case "B2-211講義教室":
      return CalendarApp.getCalendarById(
        "CALENDAR_ID"
      );
    case "B2-213講義教室":
      return CalendarApp.getCalendarById(
        "CALENDAR_ID"
      );
    case "B2-214講義教室":
      return CalendarApp.getCalendarById(
        "CALENDAR_ID"
      );
    case "B2-215講義教室":
      return CalendarApp.getCalendarById(
        "CALENDAR_ID"
      );
    case "B2-216講義教室":
      return CalendarApp.getCalendarById(
        "CALENDAR_ID"
      );
    case "B2-302研討室":
      return CalendarApp.getCalendarById(
        "CALENDAR_ID"
      );
    case "B2-309研討室":
      return CalendarApp.getCalendarById(
        "c_vu3injmrgg0ood0bnqmpi1i338@group.calendar.google.com"
      );
    case "B2-313系會議室":
      return CalendarApp.getCalendarById(
        "c_vmptj016on2231q1odh69r8f3o@group.calendar.google.com"
      );
    default:
      break;
  }
}

/**
 * 時間正規化
 * @param {Number} day
 * @param {Number} time
 * @returns {Date}
 */
function timeRegex(day, time) {
  let hour = time.getHours();
  let minute = time.getMinutes();
  let second = time.getSeconds();
  return new Date(
    day.getTime() + hour * 60 * 60 * 1000 + minute * 60 * 1000 + second * 1000
  );
}
