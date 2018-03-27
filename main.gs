/**
 * ※1行目がヘッダ、2行目からデータとしてやってます
 * ・[メンバーマスタ]シート
 * A1 : [ID]
 * B1 : [あだ名]
 * C1 : [利用可能ポイント]
 * D1 : [累計ポイント]
 * E1 : [受取済チケット]
 * F1 : [あだ名チェック用、入力ミスとかのをいれるとこ]
 * 
 * ・[投票ログ]シート
 * A1 : [投票元(メールアドレス)]
 * B1 : [投票先(あだ名)]
 * C1 : [投票ポイント]
 * D1 : [投票時刻]
 * E1 : [あだ名チェックエラー](メンバーマスタに存在しない名称の場合に『x』が入る)
 */

/**
 * 投げ銭のスプレッドシートを取得する
 */
function getThanksGivingSheet(sheetName) {
  return getThanksGivingSpreadSheet().getSheetByName(sheetName);
}

function getThanksGivingSpreadSheet() {
  var id = "☆投げ銭用のスプレッドシートID☆";// ※実際につなぐスプレッドシートのIDを設定
  return SpreadsheetApp.openById(id);
}

/**
 * 『メンバーマスタ』のシート情報を取得する
 */
function getResultSheet() {
  return getThanksGivingSheet("メンバーマスタ");
}

/**
 * 『投票ログ』のシート情報を取得する
 */
function getLogSheet() {
  return getThanksGivingSheet("投票ログ");
}

/**
 * <画面から>
 * 初期画面を取得する
 */
function doGet() {
  // 初期表示用の画面を返却する
  var result = validateMailAddresses();
  if (result.auth) {
    // 初期表示用処理
    init(result);

    // 画面用に値を設定して画面表示
    var html = HtmlService.createTemplateFromFile('index');
    html.memberId = result.id;
    html.nickname = result.nickname;
    return html.evaluate();
  } else {
    return HtmlService.createTemplateFromFile('error').evaluate();
  }
}

function init(result) {
  if (getResultSheet().getLastRow() < 2) {
    // 登録情報がなければ登録
    appendNewMember(result.id, result.nickname);
  }
  var sheet = getResultSheet();
  var members = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < members.length; i++) {
    if (members[i][0] == result.id) {
      return;
    }
    if (members[i][1] == result.nickname) {
      // IDが設定なければ更新する
      sheet.getRange(i + 2, 1, 1, 1).setValues([[result.id]]);
      return;
    }
  }
  appendNewMember(result.id, result.nickname);
}

function getDefaultPoint() {
  return 1000;
}

/**
 * 新規追加分処理
 */
function appendNewMember(id, nickname) {
  var sheet = getResultSheet();
  sheet.appendRow([id, nickname, getDefaultPoint(), 0]);
}

/**
 * アクセスしているユーザのメールアドレスを取得
 */
function getMailAddress() {
  return Session.getActiveUser().getEmail();
}

/**
 * 有効なメールアドレスかをチェック
 */
function validateMailAddresses()
{
  // POSTデータ
  var payload = {
    "auth_type" : "owners_club",
    "email" : getMailAddress()
  }
  // POSTオプション
  var options = {
    method : "POST",
    payload : payload,
    muteHttpExceptions: true
  }
  // アクセス先
  var url = "https://script.google.com/macros/s/AKfycbwSTvj_jXyhgGFxTF1E-X5dCxODB7IrDxiOUkbTGTr3sL9ueco/exec"
  // POSTリクエスト
  var response = UrlFetchApp.fetch(url, options);
  var result = JSON.parse(response);
  return result;
}

/**
 * 「メンバーマスタ」からデータを取得
 */
function getMyData(form) {
  var masterSheet = getResultSheet();
  var range = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 4).getValues();
  for (var i = 0; i < range.length; i++) {
    if (range[i][0] == form.memberId) {
      var row = i + 2;
      return {memberId: range[i][0], nickname: range[i][1], point: range[i][2], totalPoint: range[i][3], row: row};
    }
  }
  return {memberId: "", nickname: "", point: 0, totalPoint: 0, row: 0};
  
}

/**
 * <画面から>
 * 投票を行う
 */
function doPost(form) {
  // 投票を受取り、所持ポイントを超過していないかをチェック
  // 投票元と投票先が同じでないことを確認
  // 問題なければ所持ポイントを減算
  // 投票先があだ名チェックに含まれる場合、本来の名前に置き換え
  // 登録を行う
  var sheet = getLogSheet();
  var myData = getMyData(form);
  // 所持ポイント以内であること
  if (myData.point < form.point) {
    return {result: "ERROR", message: "ポイントが不足しています。"};
  }
  // 自分への投票でないこと
  if (!isAvailablePayment(myData, form)) {
    return {result: "ERROR", message: "自分には投票できません。"};
  }
  if (isValidName(form.name)) {
    sheet.appendRow([form.nickname, form.name, form.point, new Date()]);
    updatePoint(myData, form.point);
    return {result: "SUCCESS", message: "投票しました☆"};
  }
  var name = transName(form.name);
  if (name == form.name) {
    sheet.appendRow([form.nickname, form.name, form.point, new Date(), "x"]);
  } else {
    sheet.appendRow([form.nickname, name, form.point, new Date()]);
  }
  updatePoint(myData, form.point);
  return {result: "SUCCESS", message: "投票しました☆"};
}

/**
 * 利用ポイント分を更新する
 */
function updatePoint(myData, point) {
  var sheet = getResultSheet();
  sheet.getRange(myData.row, 3, 1, 1).setValues([[myData.point - point]]);
}

/**
 * 自分への投票かをチェック
 */
function isAvailablePayment(myData, form) {
  var myName = myData.nickname;
  if (myName == form.name) {
    return false;
  }
  var tmpName = transName(form.name);
  if (tmpName == myName) {
    return false;
  } 
  return true;
}

/**
 * 有効なあだ名かをチェック
 */
function isValidName(name) {
  var masterSheet = getResultSheet();
  var names = masterSheet.getRange(2, 2, masterSheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (names[i][0] == name) {
      return true;
    }
  }
  return false;
}

/**
 * 正式名称に変換できるものについて、正式名称を取得する
 */
function transName(name) {
  var masterSheet = getResultSheet();
  var names = masterSheet.getRange(2, 2, masterSheet.getLastRow() - 1, 5).getValues();
  for (var i = 0; i < names.length; i++) {
    if (!names[i][4]) {
      continue;
    }
    var sameNames = names[i][4].split(",");
    for (var j = 0; j < sameNames.length; j++) {
      if (sameNames[j] == name) {
        return names[i][0];
      }
    }
  }
  return name;
}

/**
 * <画面から>
 * 現在の受領合計ポイントを取得する
 */
function getTotalPoint(memberId) {
  var sheet = getResultSheet();
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  for (var i = 0; i < range.length; i++) {
    if (range[i][0] == memberId) {
      return range[i][3];
    }
  }
  return 0;
}

/**
 * <画面から>
 * 現在の利用可能ポイントを取得する
 */
function getAvailablePoint(memberId) {
  var sheet = getResultSheet();
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  for (var i = 0; i < range.length; i++) {
    if (range[i][0] == memberId) {
      return range[i][2];
    }
  }
  return 0;
}

/**
 * <トリガーから>
 * 月次集計処理
 */
function monthlyCalc() {
  var resultSheet = getResultSheet();
  // 月次処理を行う
  if (resultSheet.getLastRow() > 1) {
    var resultRange = resultSheet.getRange(2, 1, resultSheet.getLastRow() - 1, 2).getValues();
    // 0. 全ての使用可能ポイントをゼロする
    for (var i = 0; i < resultRange.length; i++) {
      resultSheet.getRange(i + 2, 3, 1, 1).setValues([[0]]);
    }
  }
  // 1. ログから投票先ごとに集計する
  var logSheet = getLogSheet();
  if (logSheet.getLastRow() > 1) {
    var logRange = logSheet.getRange(2, 2, logSheet.getLastRow() - 1, 2).getValues();
    var pointList = [];
    for (i in logRange) {
      var f = true;
      for (j in pointList) {
        if (pointList[j][0] == logRange[i][0]) {
          pointList[j][1] = pointList[j][1] + logRange[i][1];
          f = false;
          break;
        }
      }
      if (f) {
        pointList[pointList.length] = [];
        pointList[pointList.length - 1][0] = logRange[i][0];
        pointList[pointList.length - 1][1] = logRange[i][1];
      }
    }
  }
  Logger.log(pointList);
  // 2.集計後マスタの累積に加算する
  // 加算先がなければ行を追加する
  var resultRange = resultSheet.getRange(2, 2, resultSheet.getLastRow() - 1, 3).getValues();
  for (i in pointList) {
    Logger.log("point" + pointList[i]);
    var f = true;
    for (j in resultRange) {
      if (pointList[i][0] == resultRange[j][0]) {
        Logger.log("result" + resultRange[j]);
        Logger.log("j: " + j);
        var total = pointList[i][1] + resultRange[j][1];
        resultSheet.getRange(j + 2, 4, 1, 1).setValues([[total]]);
        f = false;
        break;
      }
    }
    if (f) {
      resultSheet.appendRow(["", pointList[i][0], 0, pointList[i][1]]);
    }
  }
  // 3. ログシートを切り替える
  makeNewLogSheet();
  // 4. 使用可能ポイントをリセットする
  for (var i = 0; i < resultSheet.getLastRow() - 1; i++) {
    resultSheet.getRange(i + 2, 3, 1, 1).setValues([[getDefaultPoint()]]);
  }
}

function makeNewLogSheet() {
  // 集計済みデータをリネームする
  var d = new Date();
  var y = d.getFullYear();
  var m = d.getMonth();
  var d = d.getDate();
  date = new Date(y, m - 1, d);
  getLogSheet().setName("投票ログ_" + Utilities.formatDate( date, 'Asia/Tokyo', 'yyyyMM'));
  // 新しいログシートを作成する
  getThanksGivingSpreadSheet().insertSheet("投票ログ");
  getLogSheet().appendRow(["投票元", "投票先", "投票ポイント", "投票時刻", "あだ名チェックエラー"]);
}
