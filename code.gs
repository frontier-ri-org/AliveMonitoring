/** @OnlyCurrentDoc */


/**
* 死活監視を行うためのハンドラ
*/
function aliveMonitoringHandler(){
  
  // 死活監視の対象ページをSpreadSheetから取得する
  const aliveMonitoringTargetsSheetName = getSettingVlue("aliveMonitoringTargetsSheetName");
  if ( aliveMonitoringTargetsSheetName === null ) {
    throw new Error("設定項目（シート名：" + settingSheetName + " ）に、項目：aliveMonitoringTargetsSheetName が設定されていません"); 
  }
  const aliveMonitoringTargetsSheet = getSheet(aliveMonitoringTargetsSheetName);
  const lastColumnNumber = aliveMonitoringTargetsSheet.getLastColumn();  // データが終わる列番号を取得
  const lastRowNumber = aliveMonitoringTargetsSheet.getLastRow();  // データが終わる列番号を取得
  
  // シートが空白の場合は、null を返す
  if (lastRowNumber === 0 || lastRowNumber === 0) {
    throw new Error("シートに値が設定されていません");
  }
  
  const aliveMonitoringTargets = aliveMonitoringTargetsSheet.getRange(1, 1, lastRowNumber, lastColumnNumber).getValues();
  let serviceNameColumnIndex = null;
  let urlColumnIndex = null;
  for (let col=0; col < lastColumnNumber; col++) {
    let value = aliveMonitoringTargets[0][col]; 
    // 列名の確認
    if (value === "serviceName") {
      serviceNameColumnIndex = col;
    } else if (value === "url"){
      urlColumnIndex = col;
    }
  }
  
  // key列又は、value列が見つからなかった場合はエラーを返す
  if (serviceNameColumnIndex === null || urlColumnIndex === null) {
    throw new Error("適切な列名が設定されていません");
  }
  
  let failureServiceList = [["サービス名", "URL", "停止検知時刻"]]
  const alertThresholdTime = getSettingVlue("alertThresholdTime");
  if ( alertThresholdTime === null ) {
    throw new Error("設定項目（シート名：" + "settings" + " ）に、項目：alertThresholdTime が設定されていません"); 
  }
  const reAlertThresholdTime = getSettingVlue("re-alertThresholdTime");
  if ( reAlertThresholdTime === null ) {
    throw new Error("設定項目（シート名：" + "settings" + " ）に、項目：re-alertThresholdTime が設定されていません"); 
  }
  let isSendAlertEmail = false;
  for (let row=1; row < lastRowNumber; row++) {
    const url = aliveMonitoringTargets[row][urlColumnIndex];
    
    // 古いログを削除する
    deleteOldLogs(url);
    
    // 日時を取得する
    const date =  Utilities.formatDate( (new Date()), 'Asia/Tokyo', 'yyyy/MM/dd hh:mm:ss');
    //const date = new Date();
    
    // ステータスコードを確認
    const statusCode = fetchStatus(url);
    
    // ログデータを作成
    const logData = {"date": date, "statusCode": statusCode, "url": url, "isFailure": false, "isSendAlertMail": false};
    
    console.log(url);
    
    // ステータスコードがエラーを示すものかどうかを判定する
    if ( isFailureCode(statusCode) ) {
      logData["isFailure"] = true;  // ログデータを更新
      isSendAlertEmail = false;
      // アラートメールを送信するかどうかを判定する
      const currentContinualFailureTime = getCurrentContinualFailureTime(url) + 1;
      const currentContinualFailureTimeFromBeforeAlert = getCurrentContinualFailureTimeFromBeforeAlert(url) + 1;
      if(currentContinualFailureTime >= alertThresholdTime && currentContinualFailureTimeFromBeforeAlert === currentContinualFailureTime){
        isSendAlertEmail = true;
      }else if (currentContinualFailureTime >= alertThresholdTime && currentContinualFailureTimeFromBeforeAlert >= reAlertThresholdTime) {
        isSendAlertEmail = true;
      }
      
      // アラートメールの表を作成
      if (isSendAlertEmail) {
        failureServiceList.push([aliveMonitoringTargets[row][serviceNameColumnIndex], url, getCurrentContinualFailureDate(url)]); 
      }
      logData["isSendAlertMail"] = isSendAlertEmail;
    }
    
    // ログを記録する
    insertLog(url, logData);
  }
  
  // アラートメールを送信する
  if(isSendAlertEmail){
    sendAlertMail(generateAlertMailMessage(failureServiceList));
  }
}

/**
* 古いログを消す
* ログは上の行から下に行くほど古くなるという前提
* シートが存在しない場合は、エラーが発生する
*/
function deleteOldLogs(url){
  const domain = extractDomain(url);  // ドメイン名からLog用シート名を取得
  let logSheetNamePrefix = getSettingVlue("logSheetNamePrefix");
  if ( logSheetNamePrefix == null ) {
    throw new Error("設定項目（シート名：" + settingSheetName + " ）に、項目：logSheetNamePrefix が設定されていません"); 
  }
  const sheetName = logSheetNamePrefix + domain;
  
  const maxLogRecordDays = getSettingVlue("maxLogRecordDays");
  if ( maxLogRecordDays === null ) {
    throw new Error("設定項目（シート名：" + "settings" + " ）に、項目：maxLogRecordDays が設定されていません"); 
  }
  
  // 日時を取得する
  const thresholdDate = new Date();
  thresholdDate.setDate(thresholdDate.getDay() - maxLogRecordDays);
  
  // シートオブジェクトを取得する
  const logSheet = getSheet(sheetName);
  
  const lastColumnNumber = logSheet.getLastColumn();  // データが終わる列番号を取得
  const lastRowNumber = logSheet.getLastRow();  // データが終わる列番号を取得
  
  // シートが空白の場合
  if (lastRowNumber === 0 || lastRowNumber === 0) {
    return;
  }
  
  const aliveMonitoringLog = logSheet.getRange(1, 1, lastRowNumber, lastColumnNumber).getValues();
  let dateColumnIndex = null;
  for (let col=0; col < lastColumnNumber; col++) {
    let value = aliveMonitoringLog[0][col]; 
    // 列名の確認
    if (value === "date") {
      dateColumnIndex = col;
      break;
    }
  }
  
  if (dateColumnIndex === null) {
    throw new Error("ログシート（シート名：" + sheetName + "）のフォーマットが不適切です");
  }
  
  // 先頭行は列名のため、読み飛ばす
  for (let row=1; row < lastRowNumber; row++ ) {
    let loggedTime =  new Date(aliveMonitoringLog[row][dateColumnIndex]);
    if ( loggedTime - thresholdDate < 0 ) {
      logSheet.deleteRows(row+1, lastRowNumber - row);
      console.log("Delete logs");
      break;
    }
  }
}



/**
* 直近の連続したエラー（isFailure）の回数を返す
*
*/
function getCurrentContinualFailureTime(url){
  // スプレッドシートオブジェクトを取得する
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const domain = extractDomain(url);  // ドメイン名からLog用シート名を取得
  let logSheetNamePrefix = getSettingVlue("logSheetNamePrefix");
  if ( logSheetNamePrefix == null ) {
    throw new Error("設定項目（シート名：" + settingSheetName + " ）に、項目：logSheetNamePrefix が設定されていません"); 
  }
  const sheetName = logSheetNamePrefix + domain;
  
  // シートオブジェクトを取得する
  const logSheet = getSheet(sheetName);
  
  const lastColumnNumber = logSheet.getLastColumn();  // データが終わる列番号を取得
  const lastRowNumber = logSheet.getLastRow();  // データが終わる列番号を取得
  
  // シートが空白の場合は、0 を返す
  if (lastRowNumber === 0 || lastRowNumber === 0) {
    return 0;
  }
  
  // 前提として、ログデータは上から順に新しいデータである
  const aliveMonitoringLog = logSheet.getRange(1, 1, lastRowNumber, lastColumnNumber).getValues();
  let isFailureColumnIndex = null;
  for (let col=0; col < lastColumnNumber; col++) {
    let value = aliveMonitoringLog[0][col]; 
    // 列名の確認
    if (value === "isFailure") {
      isFailureColumnIndex = col;
      break;
    }
  }
  
  if (isFailureColumnIndex === null) {
    throw new Error("ログシート（シート名：" + sheetName + "）のフォーマットが不適切です");
  }
  
  let counter = 0;
  for (let row=1; row < lastRowNumber && aliveMonitoringLog[row][isFailureColumnIndex]; row++) {
    counter++;
  }
  
  return counter;
}

/**
* 直近の連続したエラー（isFailure）の回数または、アラートメール送信以降の連続したエラーの回数のうち小さい方を返す
*
*/
function getCurrentContinualFailureTimeFromBeforeAlert(url){
  // スプレッドシートオブジェクトを取得する
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const domain = extractDomain(url);  // ドメイン名からLog用シート名を取得
  let logSheetNamePrefix = getSettingVlue("logSheetNamePrefix");
  if ( logSheetNamePrefix == null ) {
    throw new Error("設定項目（シート名：" + settingSheetName + " ）に、項目：logSheetNamePrefix が設定されていません"); 
  }
  const sheetName = logSheetNamePrefix + domain;
  
  // シートオブジェクトを取得する
  const logSheet = getSheet(sheetName);
  
  const lastColumnNumber = logSheet.getLastColumn();  // データが終わる列番号を取得
  const lastRowNumber = logSheet.getLastRow();  // データが終わる列番号を取得
  
  // シートが空白の場合は、0 を返す
  if (lastRowNumber === 0 || lastRowNumber === 0) {
    return 0;
  }
  
  // 前提として、ログデータは上から順に新しいデータである
  const aliveMonitoringLog = logSheet.getRange(1, 1, lastRowNumber, lastColumnNumber).getValues();
  let isFailureColumnIndex = null;
  let isSendAlertMailIndex = null;
  for (let col=0; col < lastColumnNumber; col++) {
    let value = aliveMonitoringLog[0][col]; 
    // 列名の確認
    if (value === "isFailure") {
      isFailureColumnIndex = col;
    } else if (value == "isSendAlertMail") {
      isSendAlertMailIndex = col;
    }
    
    if (isFailureColumnIndex != null && isSendAlertMailIndex != null) {
      break;
    }
  }
  
  if (isFailureColumnIndex === null || isSendAlertMailIndex === null) {
    throw new Error("ログシート（シート名：" + sheetName + "）のフォーマットが不適切です");
  }
  
  let counter = 0;
  for (let row=1; row < lastRowNumber && aliveMonitoringLog[row][isFailureColumnIndex]; row++) {
    if (aliveMonitoringLog[row][isSendAlertMailIndex]) {
      break;
    }
    counter++;
  }
  
  return counter;
}

/**
* 直近の連続したエラー（isFailure）を最初に検知した時刻を返す
*
*/
function getCurrentContinualFailureDate(url){
  // スプレッドシートオブジェクトを取得する
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const domain = extractDomain(url);  // ドメイン名からLog用シート名を取得
  let logSheetNamePrefix = getSettingVlue("logSheetNamePrefix");
  if ( logSheetNamePrefix == null ) {
    throw new Error("設定項目（シート名：" + settingSheetName + " ）に、項目：logSheetNamePrefix が設定されていません"); 
  }
  const sheetName = logSheetNamePrefix + domain;
  
  // シートオブジェクトを取得する
  const logSheet = getSheet(sheetName);
  
  const lastColumnNumber = logSheet.getLastColumn();  // データが終わる列番号を取得
  const lastRowNumber = logSheet.getLastRow();  // データが終わる列番号を取得
  
  // シートが空白の場合は、0 を返す
  if (lastRowNumber === 0 || lastRowNumber === 0) {
    return 0;
  }
  
  // 前提として、ログデータは上から順に新しいデータである
  const aliveMonitoringLog = logSheet.getRange(1, 1, lastRowNumber, lastColumnNumber).getValues();
  let isFailureColumnIndex = null;
  let dateColumnIndex = null;
  for (let col=0; col < lastColumnNumber; col++) {
    let value = aliveMonitoringLog[0][col]; 
    // 列名の確認
    if (value === "isFailure") {
      isFailureColumnIndex = col;
    } else if (value == "date") {
      dateColumnIndex = col;
    }
    
    if (isFailureColumnIndex != null && dateColumnIndex != null) {
      break;
    }
  }
  
  if (isFailureColumnIndex === null || dateColumnIndex === null) {
    throw new Error("ログシート（シート名：" + sheetName + "）のフォーマットが不適切です");
  }
  
  let row;
  for (row=1; row < lastRowNumber && aliveMonitoringLog[row][isFailureColumnIndex]; row++) {}
  
  return aliveMonitoringLog[row-1][dateColumnIndex];
}


/**
* アラートメール(HTML)を送信する
*
*/
function sendAlertMail(messageBody, settingSheetName="settings") {
  // スプレッドシートオブジェクトを取得する
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mailListSheetName = getSettingVlue("mailListSheetName", settingSheetName);
  let mailAddressColumnNumber = null;
  
  // シートが存在しない場合は、エラーを返す
  if ( !(isExistSheet(mailListSheetName)) ) {
    throw new Error("シート名: " + mailListSheetName + "  は存在しません");
  }
  const sheet = getSheet(mailListSheetName);
  const lastColumnNumber = sheet.getLastColumn();  // データが終わる列番号を取得
  const lastRowNumber = sheet.getLastRow();  // データが終わる列番号を取得
  
  // シートが空白の場合は、null を返す
  if (lastRowNumber === 0 || lastRowNumber === 0) {
    throw new Error("シートに値が設定されていません");
  }
  
  const sheetValue = sheet.getRange(1, 1, lastRowNumber, lastColumnNumber).getValues();
  
  // メールアドレスが記述された列（列名：e-mail）を検索し、存在しない場合はエラーを返す
  let mailAddressColumnIndex = null;
  let key = "e-mail";
  for (let col=0; col < lastColumnNumber; col++) {
    let value = sheetValue[0][col];
    // 列名の確認
    if (value === key) {
      mailAddressColumnIndex = col;
    }
  }
  if( mailAddressColumnIndex === null ) {
    throw new Error("指定された設定項目(" + key + ")が見つかりませんでした"); 
  }
  
  const alertMailFrom = getSettingVlue("alertMailFrom", settingSheetName);
  if ( alertMailFrom === null ) {
    throw new Error("設定項目（シート名：" + settingSheetName + " ）に、項目：alertMailFrom が設定されていません"); 
  }
  
  const alertMailTitle = getSettingVlue("alertMailTitle", settingSheetName);
  if ( alertMailTitle === null ) {
    throw new Error("設定項目（シート名：" + settingSheetName + " ）に、項目：alertMailTitle が設定されていません"); 
  }
  
  // メールの送信
  for(let row=1; row < lastRowNumber; row++){
    GmailApp.sendEmail(
      sheetValue[row][mailAddressColumnIndex], //宛先
      alertMailTitle, //件名
      messageBody.textReport, //本文
      {
        from: alertMailFrom, //送り元
        htmlBody: messageBody.htmlReport
      }
    );
  }
  
}


/**
* アラートメールのメッセージを作成する
* @param {Object} failureServiceList - 以下のような形式にすること
*                                      [["列名01", "列名02", "列名03"],
*                                       ["サービス名01", "URL01", "停止検知時刻01"],
*                                       ["サービス名02", "URL02", "停止検知時刻02"]]
*/
function generateAlertMailMessage(failureServiceList, settingSheetName="settings"){
  const alertMailTitle = getSettingVlue("alertMailTitle", settingSheetName);
  if ( alertMailTitle === null ) {
    throw new Error("設定項目（シート名：" + settingSheetName + " ）に、項目：alertMailTitle が設定されていません"); 
  }
  
  let htmlReport = "";
  let textReport = "";  // メールクライアントがHTMLに対応していなかった際の保険
  
  // HTMLレポートのhead部分の作成
  const today = new Date();
  htmlReport += '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8" /><title>';
  htmlReport += alertMailTitle;
  htmlReport += '</title><meta name="viewport" content="width=device-width, initial-scale=1.0"/></head>';
  
  const alertMailMessage = getSettingVlue("alertMailMessage", settingSheetName);
  if ( alertMailTitle === null ) {
    throw new Error("設定項目（シート名：" + settingSheetName + " ）に、項目：alertMailMessage が設定されていません"); 
  }
  // メッセージ部の作成
  htmlReport += "<body><p>";
  htmlReport += alertMailMessage;
  htmlReport += "</p>";
  textReport += alertMailMessage;
  textReport += "\n";
  
  // レポートのデータ部分の作成  
  htmlReport += "<table border='1'>";
  for(let row=0; row<failureServiceList.length; row++){
    let htmlReportRow = "<tr>";
    let textReportRow = "| ";
    for(let col=0; col<failureServiceList[row].length; col++){
      let cellData = String(failureServiceList[row][col]);
      
      // 空白埋めをするための処理
      let cellDataPadding;
      if(cellData.length < 15){
        let padding = "";
        for(i=0; i<(15 - cellData.length); i++){
          padding += " ";
        }
        cellDataPadding = padding + cellData;
      }
      htmlReportRow += ("<td>" + cellData + "</td>");
      textReportRow += (cellDataPadding + " |");
    }
    htmlReportRow += "</tr>";
    htmlReport += htmlReportRow;
    
    textReportRow += "\n";
    textReport += textReportRow;
  }
  htmlReport += "</table>";
  
  
  // 本文補足メッセージ部の作成
  const alertMailMessageSupplement = getSettingVlue("alertMailMessageSupplement", settingSheetName);
  if ( alertMailMessageSupplement === null ) {
    throw new Error("設定項目（シート名：" + settingSheetName + " ）に、項目：alertMailMessageSupplement が設定されていません"); 
  }
  
  htmlReport += "<p>";
  htmlReport += alertMailMessageSupplement;
  htmlReport += "</p>";
  htmlReport += "</body></html>";
  
  textReport += "\n\n";
  textReport += alertMailMessageSupplement;
  
  return {"htmlReport": htmlReport, "textReport": textReport};
}


/**
* スプレッドシートにLogを記録する
* @param {string} url
* @param {Object} logData
*/
function insertLog (url, logData, settingSheetName="settings") {
  // スプレッドシートオブジェクトを取得する
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const domain = extractDomain(url);  // ドメイン名からLog用シート名を取得
  let logSheetNamePrefix = getSettingVlue("logSheetNamePrefix");
  if ( logSheetNamePrefix == null ) {
    throw new Error("設定項目（シート名：" + settingSheetName + " ）に、項目：logSheetNamePrefix が設定されていません"); 
  }
  const sheetName = logSheetNamePrefix + domain;
  
  // シートオブジェクトを取得する
  const logSheet = getSheet(sheetName);
  
  // ログを記録
  insertRow(logSheet, logData);
}


/**
* 設定シートの設定を取得する
* 設定シートの列名には "item" と "key" が必ず存在する必要がある
* @param {string} key - 検索したい設定項目
* @param {string} sheetName - 検索対象となるシート名
* @return {string}
*/
function getSettingVlue (key, sheetName="settings") {
  // シートが存在しない場合は、null を返す
  if ( !(isExistSheet(sheetName)) ) {
    throw new Error("シート名: " + sheetName + "  は存在しません");
  }
  
  const sheet = getSheet(sheetName);  // シートオブジェクトをを取得
  const lastColumnNumber = sheet.getLastColumn();  // データが終わる列番号を取得
  const lastRowNumber = sheet.getLastRow();  // データが終わる列番号を取得
  
  // シートが空白の場合は、null を返す
  if (lastRowNumber === 0 || lastRowNumber === 0) {
    console.warn("シートに値が設定されていません");
    return null; 
  }
  
  // key（item）列と、value列の列番号を検索
  let keyColumnNumber = null;
  let valueColumnNumber = null;
  for (let col=1; col <= lastColumnNumber; col++) {
    let value = sheet.getRange(1, col).getValue(); 
    // 列名の確認
    if (value === "item") {
      keyColumnNumber = col;
    } else if (value === "value") {
      valueColumnNumber = col;
    }
  }
  
  // key列又は、value列が見つからなかった場合はnullを返す
  if (keyColumnNumber === null || valueColumnNumber === null) {
    console.warn("適切な列名が設定されていません");
    return null; 
  }
  
  // 設定を検索
  for (let row=2; row <= lastRowNumber; row++) {
    let value = sheet.getRange(row, keyColumnNumber).getValue();
    if ( value === key ) {
      return sheet.getRange(row, valueColumnNumber).getValue();
    }
  }
  
  //設定が見つからなかった場合、null を返す
  console.warn("指定された設定項目(" + key + ")が見つかりませんでした");
  return null;
}



/**
* シートに値を行ごと一番上に挿入する
* カラム名が存在しない場合は、新規に列を作成する
* @param {Object} sheet - sheetオブジェクト
* @param {Object} value - 辞書型のデータ。入れ子には対応しない。
*                   {"colName": value}
*/
function insertRow(sheet, value) {
  let lastColumnNumber = sheet.getLastColumn();  // データが終わる列番号を取得
  
  sheet.insertRowAfter(1);
  
  // シートにすでにある行名を取得して、値を挿入
  for(let col=1; col <= lastColumnNumber; col++){
    let key = String(sheet.getRange(1, col).getValue());  // 列名を取得
    if (key in value) {
      sheet.getRange(2, col).setValue(value[key]);
      delete value[key];
    }
  }
  // シートに設定されていない行名を設定して、値を挿入
  for (key in value) {
    sheet.getRange(1, ++lastColumnNumber).setValue(key);
    sheet.getRange(2, lastColumnNumber).setValue(value[key]);
    Logger.log("Inserted new row: " + key);
  }
}



/**
* URLからドメインを抽出する
* @param {string} url - ドメインを抽出したいURL
*/
function extractDomain ( url ) {
  const domain = url.match(/^https?:\/{2,}(.*?)(?:\/|\?|#|$)/)[1];
  return domain;
}


/**
* HTTPレスポンススステータスコードがエラーを示すものかどうかを判定する。
* @param {Integer} code - 確認したいシートの名前
* @return {bool} 存在する場合は true 、存在しない場合は false
*/
function isFailureCode (code) {
  // TODO: エラーコードに応じてエラーの説明も返すようにする
  return 400 <= code && code <= 599;
}

/**
* 当該URLにGETリクエストを送り、ステータスコードを返す
* @param {string} url - ステータスコードを確認したいWebページへのURL
* @return {Integer} ステータスコードを返す
*/
function fetchStatus(url) {
  let options = {
    muteHttpExceptions: true
  };
  
  let response;
  try {
    // Makes a request to fetch a URL.
    response = UrlFetchApp.fetch(url, options);
  } catch (e) {
    // DNS error, etc.
    return;
  }
  
  let code = response.getResponseCode();
  return code;
}


/**
* 指定されてた名前のシートを取得する。存在しない場合は、新規作成する。
* @param {string} name - 取得したいシート名
* @return {Object} シートオブジェクト
*/
function getSheet(name){
  // スプレッドシートオブジェクトを取得する
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 指定された名前のシートが存在する場合
  if( isExistSheet(name) ) {
    let sheet = ss.getSheetByName(name);
    if(sheet) {
      return sheet;
    }
  }
  
  // 指定された名前のシートが存在しない場合は新規に作成する
  let sheet = ss.insertSheet();
  sheet.setName(name);
  return sheet;
}


/**
* 指定されてた名前のシートが存在するかどうかを確認する。
* @param {string} name - 確認したいシートの名前
* @return {bool} 存在する場合は true 、存在しない場合は false
*/
function isExistSheet(name){
  // スプレッドシートオブジェクトを取得する
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // シート一覧を取得して、シート名を確認
  const sheetList = ss.getSheets();
  for(let i=0; i < sheetList.length; i++){
    if(sheetList[i].getName() === name){
      return true;
    }
  }
  return false;
}