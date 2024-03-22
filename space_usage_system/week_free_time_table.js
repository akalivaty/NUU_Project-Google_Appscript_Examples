
function auto_update_free_time() {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("週空堂表");
  ss.getRange(1, 9).setValue('本週');
  updateFreeTime();
}

function updateFreeTime() {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("週空堂表");
  // 清除空堂欄位舊資料
  const gridRange = ss.getRange(5, 2, 11, 40);
  gridRange.clearContent();
  // 根據sheet中下拉選單指定該週次星期一
  const targetWeek = ss.getRange(1, 6).getValue();
  let now = new Date();
  switch (targetWeek) {
    case "本週":
      now = new Date(now.getFullYear(), now.getMonth(), now.getDate() - (now.getDay() - 1));
      break;
    case "下週":
      now = new Date(now.getFullYear(), now.getMonth(), now.getDate() - (now.getDay() - 1) + 7);
      break;
    case "下下週":
      now = new Date(now.getFullYear(), now.getMonth(), now.getDate() - (now.getDay() - 1) + 14);
      break;
    default:
      break;
  }
  // 變更日期標頭
  ss.getRange(2, 2).setValue(now);
  ss.getRange(2, 10).setValue(new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1));
  ss.getRange(2, 18).setValue(new Date(now.getFullYear(), now.getMonth(), now.getDate() + 2));
  ss.getRange(2, 26).setValue(new Date(now.getFullYear(), now.getMonth(), now.getDate() + 3));
  ss.getRange(2, 34).setValue(new Date(now.getFullYear(), now.getMonth(), now.getDate() + 4));

  let allSpaceWeekGrid = [];
  // 11個教室
  for (let i = 2; i < 13; i++) {
    let calendar = CalendarApp.getCalendarById(allCalendarId_array[i]);
    let weekGrid = [];
    // 星期一 ~ 星期五
    for (let j = 1; j <= 5; j++) {
      let dayGrid = [];
      let events = calendar.getEventsForDay(now);
      events.forEach((event) => {
        let eventStartTime = event.getStartTime().getHours() * 60 + event.getStartTime().getMinutes();
        let eventEndTime = event.getEndTime().getHours() * 60 + event.getEndTime().getMinutes();
        let lessonLength = dayGrid.length;
        for (let key in lessonTime) {
          let lessonPtr = parseInt(key) + lessonLength;
          // 超過第8節
          if (lessonPtr > 8) {
            break;
          }
          let lessonStartTime = lessonTime[lessonPtr].startInMinutes;
          let lessonEndTime = lessonTime[lessonPtr].endInMinutes;
          // 節次開始時間 比 事件開始時間 早50分鐘以上
          if (lessonStartTime - eventStartTime < -50) {
            dayGrid.push("O");
          } else {
            // 節次結束時間 比 事件結束時間 早
            if (lessonEndTime - eventEndTime <= 0) {
              dayGrid.push("");
            } else {
              break;
            }
          }
        }
      });
      if (dayGrid.length < 8) {
        for (let k = dayGrid.length; k < 8; k++) {
          dayGrid.push("O");
        }
      }
      // console.log(`dayGrid: ${dayGrid}`);
      weekGrid.push(...dayGrid);
      dayGrid.length = 0;
      now = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
    }
    allSpaceWeekGrid.push(weekGrid);
    now = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 5);
  }
  gridRange.setValues(allSpaceWeekGrid);
}

const lessonTime = {
  1: { startHour: 8, startMinute: 0, startInMinutes: 480, endHour: 8, endMinute: 50, endInMinutes: 530 },
  2: { startHour: 9, startMinute: 0, startInMinutes: 540, endHour: 9, endMinute: 50, endInMinutes: 590 },
  3: { startHour: 10, startMinute: 10, startInMinutes: 610, endHour: 11, endMinute: 0, endInMinutes: 660 },
  4: { startHour: 11, startMinute: 10, startInMinutes: 670, endHour: 12, endMinute: 0, endInMinutes: 720 },
  5: { startHour: 13, startMinute: 0, startInMinutes: 780, endHour: 13, endMinute: 50, endInMinutes: 830 },
  6: { startHour: 14, startMinute: 0, startInMinutes: 840, endHour: 14, endMinute: 50, endInMinutes: 890 },
  7: { startHour: 15, startMinute: 10, startInMinutes: 910, endHour: 16, endMinute: 0, endInMinutes: 960 },
  8: { startHour: 16, startMinute: 10, startInMinutes: 970, endHour: 17, endMinute: 0, endInMinutes: 1020 },
};
