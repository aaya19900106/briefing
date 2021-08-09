function RegExpForm() {
  //データを吐き出す箇所を指定
  var sheet_url = "シートのURL";
  var sheet_name = "シートの名前";
  var ss = SpreadsheetApp.openByUrl(sheet_url);
  var sheet = ss.getSheetByName(sheet_name);

  var zyouken = 'subject:【検索したいメールの懸命】 newer_than:1d'; //検索条件
  var thds = GmailApp.search(zyouken, 0, 8);//条件に合致するメールのスレッドを取得
  var messe = GmailApp.getMessagesForThreads(thds);//スレッド内のメールを取得。[スレッド番号][メッセージ番号]の2次元配列になる

  //maxまで繰り返しで文面検索し、スプシに吐き出す
  for(var i = 0; i < messe.length; i++) {
    for(var j = 0; j < messe[i].length; j++) {

      var messageId = messe[i][j].getId();

      //もし、スプレッドシートに存在したら実行しない
      if(!hasId(messageId)){

      var date = thds[i].getMessages()[j].getDate();
      var setDate = Utilities.formatDate(date, "JST","MM/dd");
      var body = messe[i][j].getPlainBody();

      var regName = new RegExp('名前: ' + '.*?' + '\r' );
      var Name = body.match(regName)[0].replace('名前: ', '').replace('\r', '');
      
      var regSend = new RegExp('メールアドレス: ' + '.*?' + '\r' );
      var Send = body.match(regSend)[0].replace('メールアドレス: ', '').replace('\r', '');
      
      var regDate = new RegExp('日時: ' + '.*?' + ' ' );
      var SetsumeiDate = body.match(regDate)[0].replace('日時: 2021/', '').replace('\r', '');

      var regAge = new RegExp('年齢: ' + '.*?' + '\r' );
      var Ageage = body.match(regAge)[0].replace('年齢: ', '').replace('歳\r', '');

      sheet.appendRow(["",setDate,SetsumeiDate,"","", Name, Send ,"","","","","","","","",Ageage,"","","","","","","","", messageId]);
      }
    }
  }

  function hasId(id){  
   //今回は1列目にメールIDを入れていくので1列目から探す
    var data = sheet.getRange(1, 25,sheet.getLastRow(),1).getValues(); //X列にメールのID掲載するよう設定
    var hasId = data.some(function(value,index,data){
   //コールバック関数
    return (value[0] === id);
  });
  return hasId;
}

  //受信日時順の昇順に並び替え
 let narabikae = sheet.getRange('B900:Z1500'); //全部並び替え対象にすると重すぎるので一部指定に。今後バグったらここが怪しそう。
  narabikae.sort({column: 3, ascending: true}); //説明会実施日が3列目のためそこで降順並び替え

}
