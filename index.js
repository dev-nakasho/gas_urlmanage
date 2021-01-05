// 必要なプロパティをmodelオブジェクトとして保持
var model = {
  // チャットワークルームID
  ROOM_ID: "ルームID",
  // URL(文字列)
  url: "",
  // APIに対するリクエスト時に渡すパラメータ(オブジェクト)
  params: {
    headers: { "X-ChatWorkToken": "APIトークン" },
  },
  setMethod: function (smethod) {
    this.params.method = smethod;
  },
  setPayload: function (opayload) {
    this.params.payload = opayload;
  },
};

// POSTイベントハンドラ
function doPost(e) {
  model.url = "https://api.chatwork.com/v2/rooms/" + model.ROOM_ID + "/messages";

  // スプレッドシートのアクティブシートを取得
  var spread_sheet = SpreadsheetApp.getActiveSheet();
  // 現在データが入力されている最後の行番号を取得
  var lr = spread_sheet.getLastRow();
  // 最後の行番号の次の行の２列目のセルを取得し、そのセルのA1表記を取得
  var notation = spread_sheet.getRange(lr + 1, 2).getA1Notation();

  // httpを文字列に含む場合のみ処理を実施
  if (JSON.parse(e.postData.contents).webhook_event.body.includes("http")) {
    var text = "";
    // 改行を含む場合(投稿に表題がつけられる場合)
    if (JSON.parse(e.postData.contents).webhook_event.body.includes("\n")) {
      // URLだけを取得する
      var text_array = JSON.parse(e.postData.contents).webhook_event.body.split("\n");
      text = text_array[text_array.length - 1];
    } else {
      text = JSON.parse(e.postData.contents).webhook_event.body;
    }
    // 最後の行番号の次の行の２列目のセルに、チャットワークAPIで取得したメッセージをセットする
    spread_sheet.getRange(lr + 1, 2).setValue(text);
    // 最後の業番号の次の行の１列目にタイトル取得のための関数をセットする
    spread_sheet.getRange(lr + 1, 1).setValue("=urlToTitle(" + notation + ")");

    // 合計件数編集
    var total_count = spread_sheet.getLastRow();
    var payload = { body: "お気に入りサイトを登録しました！\r合計" + total_count + "件" };

    // 投稿本文と通信メソッドの設定
    model.setPayload(payload);
    model.setMethod("post");

    // スプレッドシートに登録できたことを、チャットワークへ通知する
    UrlFetchApp.fetch(model.url, model.params);
  }

}

// サイトのタイトル取得
function urlToTitle (url) {
  // 指定したURLを実行したレスポンスを取得
  var response = UrlFetchApp.fetch(url);
  // 正規表現を用意
  var myRegexp = /<title>([\s\S]*?)<\/title>/i;
  // 変換処理→配列を取得
  var match = myRegexp.exec(response.getContentText());
  var title = match[1].replace(/(^\s+)|(\s+$)/g, "");
  // タイトルを返却する
  return title;
}
