function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Email')
    .addItem('寄送派工通知', 'sendMail')
    .addItem('寄送重複案件合併通知&刪除此筆資料', 'report_duplicated_problem')
    .addToUi();
}

/**
 * 設定處理人後寄送信通知對方
 */
function sendMail() {
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('報修紀錄');
  let data = ss.getDataRange().getValues();
  let selectRange = ss.getActiveRangeList(); // 選取範圍
  let listLength = selectRange.getRanges().length; // 選取範圍個數(連續選取算一個)
  let eventCount = 0;
  try {
    for (let i = 0; i < listLength; i++) {
      let rows = selectRange.getRanges()[i].getNumRows(); // 第i個選取範圍中的列數
      let startRow = selectRange.getRanges()[i].getRow(); // 第i個選取範圍的第一列(表單中的列數)
      let handler = data[startRow - 1][6];
      switch (handler) {
        case 'handler1':
          handler = 'MAIL_ADDRESS1';
          break;
        case 'handler2':
          handler = 'MAIL_ADDRESS2';
          break;
        case 'handler3':
          handler = 'MAIL_ADDRESS3';
          break;
        default:
          break;
      }

      // 建立信件html範本
      let template = HtmlService.createTemplateFromFile('replyResult');
      template.caseNum = data[startRow - 1][10];
      template.time = Utilities.formatDate(data[startRow - 1][0], 'GMT+08:00', 'yyy/MM/dd - HH:mm:ss');
      template.applyer = data[startRow - 1][2];
      template.email = data[startRow - 1][1];
      template.phone = data[startRow - 1][3];
      template.problem = data[startRow - 1][4];
      let html = template.evaluate().getContent();

      let handleResult = data[startRow - 1][9];
      // 【處理結果】為空
      if (handleResult == '') {
        // 計算分配案件耗時
        let preChangeTime = data[startRow - 1][8];
        let newChangeTime = new Date();
        let timeDiffInMiniSec = newChangeTime.getTime() - preChangeTime.getTime();
        let hours = Math.round(timeDiffInMiniSec / 1000 / 60 / 60);
        let minutes = Math.round(timeDiffInMiniSec / 1000 / 60 % 60);
        ss.getRange(startRow, 12).setValue(hours + ' 小時 ' + minutes + ' 分鐘'); // 填入分配案件耗時
        ss.getRange(startRow, 12).setHorizontalAlignment('center'); // 文字置中
        ss.getRange(startRow, 8).setValue('處理中'); // 更新處理狀態
        ss.getRange(startRow, 9).setValue(newChangeTime); // 更新處理狀態變更時間
        let targetMail = handler;
        MailApp.sendEmail(targetMail, '報修處理通知', '', { htmlBody: html });  // 寄信通知處理人
      }
      else {
        // 計算維修過程耗時
        let preChangeTime = data[startRow - 1][0];
        let newChangeTime = new Date();
        let timeDiffInMiniSec = newChangeTime.getTime() - preChangeTime.getTime();
        let days = Math.round(timeDiffInMiniSec / 1000 / 60 / 60 / 24);
        let hours = Math.round(timeDiffInMiniSec / 1000 / 60 / 60 % 24);
        let minutes = Math.round(timeDiffInMiniSec / 1000 / 60 % 60);
        ss.getRange(startRow, 13).setValue(days + ' 天 ' + hours + ' 小時 ' + minutes + ' 分鐘'); // 填入維修過程耗時
        ss.getRange(startRow, 13).setHorizontalAlignment('center'); // 文字置中
        ss.getRange(startRow, 8).setValue('已解決'); // 更新處理狀態
        ss.getRange(startRow, 9).setValue(newChangeTime); // 更新處理狀態變更時間
        let targetMail = data[startRow - 1][1];
        let htmlBody = '<p>【問題描述】</p>' + data[startRow - 1][4] + '<p>【處理結果】</p>' + handleResult;
        MailApp.sendEmail(targetMail, '您的報修申請已修復完成', '', { htmlBody: htmlBody });  // 寄信通知報修申請人
      }
      eventCount++;

      // 當第i個選取範圍有複數列
      if (rows > 1) {
        for (let j = 0; j < rows - 1; j++) {
          handler = data[startRow + j][6];
          switch (handler) {
            case 'handler1':
              handler = 'MAIL_ADDRESS1';
              break;
            case 'handler2':
              handler = 'MAIL_ADDRESS2';
              break;
            case 'handler3':
              handler = 'MAIL_ADDRESS3';
              break;
            default:
              break;
          }

          // 建立信件html範本
          template.caseNum = data[startRow + j][10];
          template.time = Utilities.formatDate(data[startRow + j][0], 'GMT+08:00', 'yyy/MM/dd - HH:mm:ss');
          template.applyer = data[startRow + j][2];
          template.email = data[startRow + j][1];
          template.phone = data[startRow + j][3];
          template.problem = data[startRow + j][4];
          html = template.evaluate().getContent();

          handleResult = data[startRow + j][9];
          // 【處理結果】為空
          if (handleResult == '') {
            // 計算分配案件耗時
            preChangeTime = data[startRow + j][8];
            newChangeTime = new Date();
            timeDiffInMiniSec = newChangeTime.getTime() - preChangeTime.getTime();
            hours = Math.round(timeDiffInMiniSec / 1000 / 60 / 60);
            minutes = Math.round(timeDiffInMiniSec / 1000 / 60 % 60);
            ss.getRange(startRow + 1 + j, 12).setValue(hours + ' 小時 ' + minutes + ' 分鐘'); // 填入分配案件耗時
            ss.getRange(startRow + 1 + j, 12).setHorizontalAlignment('center'); // 文字置中
            ss.getRange(startRow + 1 + j, 8).setValue('處理中'); // 更新處理狀態
            ss.getRange(startRow + 1 + j, 9).setValue(newChangeTime); // 更新處理狀態變更時間
            targetMail = handler;
            MailApp.sendEmail(targetMail, '報修處理通知', '', { htmlBody: html });  // 寄信通知處理人
          }
          else {
            // 計算維修過程耗時
            preChangeTime = data[startRow + j][0];
            newChangeTime = new Date();
            timeDiffInMiniSec = newChangeTime.getTime() - preChangeTime.getTime();
            days = Math.round(timeDiffInMiniSec / 1000 / 60 / 60 / 24);
            hours = Math.round(timeDiffInMiniSec / 1000 / 60 / 60 % 24);
            minutes = Math.round(timeDiffInMiniSec / 1000 / 60 % 60);
            ss.getRange(startRow + 1 + j, 13).setValue(days + ' 天 ' + hours + ' 小時 ' + minutes + ' 分鐘'); // 填入維修過程耗時
            ss.getRange(startRow + 1 + j, 13).setHorizontalAlignment('center'); // 文字置中
            ss.getRange(startRow + 1 + j, 8).setValue('已解決'); // 更新處理狀態
            ss.getRange(startRow + 1 + j, 9).setValue(newChangeTime); // 更新處理狀態變更時間
            targetMail = data[startRow + j][1];
            let htmlBody = '<p>【問題描述】</p>' + data[startRow + j][4] + '<p>【處理結果】</p>' + handleResult;
            MailApp.sendEmail(targetMail, '您的報修申請已修復完成', '', { htmlBody: htmlBody });  // 寄信通知報修申請人
          }
          eventCount++;
        }
      }
    }
  } catch (e) {
    if (e.name == 'TypeError') {
      throw '超出選取範圍';
    }
    else {
      throw e.message;
    }
  } finally {
    SpreadsheetApp.getActiveSpreadsheet().toast('已發送 ' + eventCount + ' 封 mail');
  }
}

/**
 * 表單送出後填入【處理狀態】、【處理人】、【處理變更時間】、【案件編號】，寄送報修申請通知
 * ! 須設定trigger
 * @param {Array} rawData From form trigger
 */
function preProcess(rawData) {
  Logger.log(rawData.namedValues);
  let data = rawData.namedValues;

  // 若表單為【報修紀錄】
  if (data['問題描述'] != null) {
    Logger.log('Access 【報修紀錄】 sheet.')

    let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('報修紀錄');
    let lastRow = ss.getLastRow();
    let lastColumn = ss.getLastColumn();
    // 【問題描述】不為空
    if (ss.getRange(lastRow, 5).getValue() != '') {
      let changeTime = new Date();
      ss.getRange(lastRow, 7).setValue('待分配');
      ss.getRange(lastRow, 8).setValue('待處理');
      ss.getRange(lastRow, 9).setValue(changeTime);
      ss.getRange(lastRow, 1).setNumberFormat('yyy/mm/dd - hh:mm:ss');
      ss.getRange(lastRow, 9).setNumberFormat('yyy/mm/dd - hh:mm:ss');
      let caseNum = timeRegex(changeTime);
      ss.getRange(lastRow, 11).setValue(caseNum);
    }

    let lastRowData = ss.getRange(lastRow, 1, 1, lastColumn).getValues();

    let centerMail = 'MAIL_ADDRESS1';

    // 建立信件html範本
    let template = HtmlService.createTemplateFromFile('centerMail');
    template.caseNum = lastRowData[0][10];
    template.time = Utilities.formatDate(lastRowData[0][0], 'GMT+08:00', 'yyy/MM/dd - HH:mm:ss');
    template.applyer = lastRowData[0][2];
    template.email = lastRowData[0][1];
    template.phone = lastRowData[0][3];
    template.problem = lastRowData[0][4];
    let html = template.evaluate().getContent();
    MailApp.sendEmail(centerMail, '報修申請通知', '', { htmlBody: html });
  }
  // 表單為【處理結果回填】
  if (data['問題描述'] == null) {
    Logger.log('Access【處理結果回填】 sheet.')
    refill();
  }
}

/**
 * 時間正規化
 * @param {Date} time
 * @returns {string}
 */
function timeRegex(time) {
  return time.getFullYear().toString().substring(2) + addZero(time.getMonth() + 1) + addZero(time.getDate()) + addZero(time.getHours()) + addZero(time.getMinutes()) + addZero(time.getSeconds());
}

/**
 * 數字小於10加0並轉為字串
 * @param {number} i
 * @retruns {string}
 */
function addZero(i) {
  if (i < 10) {
    return '0' + i.toString();
  } else {
    return i.toString();
  }
}

/**
 * 回填處理結果
 */
function refill() {
  let ss2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('處理結果回填');
  let ss2Data = ss2.getDataRange().getValues();
  let ss2LastRow = ss2.getLastRow();
  let ss1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('報修紀錄');
  let ss1LastRow = ss1.getLastRow();
  let ss1Data = ss1.getDataRange().getValues();

  if (ss2LastRow > 1) {
    for (let ss2Row = 1; ss2Row < ss2LastRow; ss2Row++) {
      for (let ss1Row = 1; ss1Row < ss1LastRow; ss1Row++) {
        // 案件編號相同
        if (ss1Data[ss1Row][10] == ss2Data[ss2Row][1]) {
          let handleStatus = ss2Data[ss2Row][3];
          // 處理狀態: 處理中
          if (handleStatus == '處理中') {
            ss1.getRange(ss1Row + 1, 10).setValue(ss2Data[ss2Row][2]);  // 填入處理結果
            ss1.getRange(ss1Row + 1, 9).setValue(new Date());           // 更新處理狀態變更時間
            ss2.deleteRow(ss2Row + 1);
          }
          // 處理狀態: 已解決
          if (handleStatus == '已解決') {
            ss1.getRange(ss1Row + 1, 8).setValue(handleStatus);         // 填入處理狀態
            ss1.getRange(ss1Row + 1, 10).setValue(ss2Data[ss2Row][2]);  // 填入處理結果
            ss2.deleteRow(ss2Row + 1);

            // 計算維修過程耗時
            let preChangeTime = ss1Data[ss1Row][0];
            let newChangeTime = new Date();
            let timeDiffInMiniSec = newChangeTime.getTime() - preChangeTime.getTime();
            let days = Math.round(timeDiffInMiniSec / 1000 / 60 / 60 / 24);
            let hours = Math.round(timeDiffInMiniSec / 1000 / 60 / 60 % 24);
            let minutes = Math.round(timeDiffInMiniSec / 1000 / 60 % 60);
            ss1.getRange(ss1Row + 1, 13).setValue(days + ' 天 ' + hours + ' 小時 ' + minutes + ' 分鐘'); // 填入維修過程耗時
            ss1.getRange(ss1Row + 1, 13).setHorizontalAlignment('center'); // 文字置中
            ss1.getRange(ss1Row + 1, 9).setValue(newChangeTime); // 更新處理狀態變更時間

            // 寄信通知報修申請人
            let targetMail = ss1Data[ss1Row][1];
            let problem = ss1Data[ss1Row][4];
            let handleResult = ss2Data[ss2Row][2];
            let htmlBody = '<p>【問題描述】</p>' + problem + '<p>【處理狀態】</p>' + handleStatus + '<p>【處理結果】</p>' + handleResult;
            MailApp.sendEmail(targetMail, '您的報修申請已修復完成', '', { htmlBody: htmlBody });
          }
        }
      }
    }
    ss2.deleteRows(2, ss2LastRow);
  }
}

function report_duplicated_problem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('報修紀錄');
  const data = ss.getDataRange().getValues();
  const selectRange = ss.getActiveRangeList();
  const listLength = selectRange.getRanges().length;
  let arr_toDeleteRows = [];
  try {
    for (let i = 0; i < listLength; i++) {
      // 取得選取範圍
      const rows = selectRange.getRanges()[i].getNumRows();
      const startRow = selectRange.getRanges()[i].getRow();
      // 建立信件html範本
      const template = HtmlService.createTemplateFromFile('duplicatedCase');
      template.caseNum = data[startRow - 1][10];
      template.time = Utilities.formatDate(data[startRow - 1][0], 'GMT+08:00', 'yyy/MM/dd - HH:mm:ss');
      template.applyer = data[startRow - 1][2];
      template.email = data[startRow - 1][1];
      template.phone = data[startRow - 1][3];
      template.problem = data[startRow - 1][4];
      let html = template.evaluate().getContent();
      let targetMail = data[startRow - 1][1];
      MailApp.sendMail(targetMail, '報修申請合併歸檔通知', '', { htmlBody: html });
      arr_toDeleteRows.push(startRow);

      if(rows > 1) {
        for (let j = 0; j < rows - 1; j++) {
          // 建立信件html範本
          template.caseNum = data[startRow + j][10];
          template.time = Utilities.formatDate(data[startRow + j][0], 'GMT+08:00', 'yyy/MM/dd - HH:mm:ss');
          template.applyer = data[startRow + j][2];
          template.email = data[startRow + j][1];
          template.phone = data[startRow + j][3];
          template.problem = data[startRow + j][4];
          html = template.evaluate().getContent();
          targetMail = data[startRow + j][1];
          MailApp.sendMail(targetMail, '報修申請合併歸檔通知', '', { htmlBody: html });
          arr_toDeleteRows.push(startRow + j + 1);
        }
      }
    }
  } catch (e) {
    if (e.name == 'TypeError') {
      throw '超出選取範圍';
    }
    else {
      throw e.message;
    }
  } finally {
    const eventCount = arr_toDeleteRows.reverse().reduce((acc, row) => {
      ss.deleteRow(row);
      return acc + 1;
    }, 0);
    SpreadsheetApp.getActiveSpreadsheet().toast('已發送 ' + eventCount + ' 封 mail');
  }
}

/**
 * html引入js與css之用
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
