// import "google-apps-script";
/**
 * 學期結束移入歷史報修紀錄表單
 * ! 須設 trigger
 */
function moveToHistory() {
  const month = new Date().getMonth() + 1;
  // 7月和 2月執行
  if (month == 7 || month == 2) {
    const ssRecord = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("報修紀錄");
    const arr_ssRecordData = ssRecord.getRange(2, 1, ssRecord.getLastRow() - 1, ssRecord.getLastColumn()).getValues();
    let arr_ssRecordObj = arr_ssRecordData.map(rowData => {
      return {
        '時間戳記': rowData[0],
        'Email': rowData[1],
        '申請人': rowData[2],
        '連絡電話': rowData[3],
        '問題描述': rowData[4],
        '圖片上傳': rowData[5],
        '處理人': rowData[6],
        '處理狀態': rowData[7],
        '處理狀態變更時間': rowData[8],
        '處理結果': rowData[9],
        '案件編號': rowData[10],
        '分配案件耗時': rowData[11],
        '維修過程耗時': rowData[12],
      };
    });

    // 檢查時間戳記
    const now = new Date();
    let arr_oldDataObj = [];
    let arr_oldDataStatus = []; // [[idx, isDateExpired, isStatusFinished]]
    for (let idx = 0; idx < arr_ssRecordObj.length; idx++) {
      let isDateExpired = arr_ssRecordObj[idx]["時間戳記"].getTime() - now.getTime() < 0 ? true : false;
      let isStatusFinished = arr_ssRecordObj[idx]["處理狀態"] == "已解決" ? true : false;
      arr_oldDataStatus.push([idx, isDateExpired, isStatusFinished]);
      // 資料已過期且已解決
      if (isDateExpired && isStatusFinished) {
        arr_oldDataObj.push(arr_ssRecordObj[idx]);
      }
      // 資料未過期
      if (isDateExpired == false) {
        break;
      }
    }

    // 刪除已解決的過期資料
    for (let i = arr_oldDataStatus.length - 1; i >= 0; i--) {
      if (arr_oldDataStatus[1] && arr_oldDataStatus[2]) {
        ssRecord.deleteRow(i + 2);
      }
    }

    // 物件 -> 陣列
    const arr_oldData = arr_oldDataObj.map(obj => {
      let tempArr = [];
      for (let key in obj) {
        tempArr.push(obj[key]);
      }
      return tempArr;
    });

    // 寫入【歷史報修紀錄】試算表
    const ssHistory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('歷史報修紀錄');
    ssHistory.getRange(ssHistory.getLastRow() + 1, 1, arr_oldData.length, arr_oldData[0].length).setValues(arr_oldData);
  }
}
