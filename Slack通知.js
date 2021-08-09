function myFunction() { 
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // スプレッドシート
  var activeSheet = activeSpreadsheet.getActiveSheet(); // アクティブシート
  if(activeSheet.getName() != "シートの名前"){
    return;
  }
  var activeCell = activeSheet.getActiveCell(); // アクティブセル
 
  if(activeCell.getColumn() == 10 && activeCell.getValues() != ""){
    var newInputRow = activeCell.getRow();
    var tantousya = activeSheet.getRange(activeCell.getRow(), 9).getValues();
    var syousai = activeSheet.getRange(activeCell.getRow(), 10).getValues();
    var ageage = activeSheet.getRange(activeCell.getRow(), 16).getValues();
    // 送信するSlackのテキスト
    var slackText = tantousya + "さんが説明会に入ったよ！年齢は" + ageage + "だ！お疲れ様！\n" +　"```"  + syousai + "```";
  sendSlack(slackText);
  }
}

function sendSlack(slackText){
  // Step1で取得したWebhook URLを設定
  var webHookUrl = "webhookのURL";
  
  var jsonData =
      {
        "channel": "#通知したいチャンネル名",   // 通知したいチャンネル 
        "text" : slackText,
        "link_names" : 1,
      };
  
  var payload = JSON.stringify(jsonData);
  
  var options =
      {
        "method" : "post",
        "contentType" : "application/json",
        "payload" : payload,
      };
  
  // リクエスト
  UrlFetchApp.fetch(webHookUrl, options);
}
