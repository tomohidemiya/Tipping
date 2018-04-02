/**
 * ※1行目がヘッダ、2行目からデータとしてやってます
 * ・[全体メンバマスタ]シート
 * A1 : [ID]
 * B1 : [あだ名]
 * C1 : [メールアドレス]
 * D1 : [バックアップ参加](0:未参加、1:参加中)
 * E1 : [メンターグループ]
 *
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
 * <画面から>
 * 初期画面を取得する
 */
function doGet() {
  // 初期表示用の画面を返却する
  if (validateMailAddresses()) {
    var html = HtmlService.createTemplateFromFile('index');
    html.memberId = getMemberId();
    return html.evaluate();
  } else {
    return HtmlService.createTemplateFromFile('error').evaluate();
  }
}

/**
 * メンバマスタのスプレッドシートを取得する
 */
function getMasterSheet(sheetName) {
  var id = "☆全体メンバー管理用のスプレッドシートID☆";// ※実際につなぐスプレッドシートのIDを設定
  var spreadSheet = SpreadsheetApp.openById(id);
  return spreadSheet.getSheetByName(sheetName);
}

/**
 * 投げ銭のスプレッドシートを取得する
 */
function getThanksGivingSheet(sheetName) {
  var id = "☆投げ銭用のスプレッドシートID☆";// ※実際につなぐスプレッドシートのIDを設定
  var spreadSheet = SpreadsheetApp.openById(id);
  return spreadSheet.getSheetByName(sheetName);
}

/**
 * 『全体メンバマスタ』のシート情報を取得する
 */
function getMasterMemberSheet() {
  return getMasterSheet("全体メンバマスタ");
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
 * アクセスしているユーザのメールアドレスを取得
 */
function getMailAddress() {
  return Session.getActiveUser().getEmail();
}

function getMemberId() {
  var masterSheet = getMasterMemberSheet();
  var members = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 3).getValues();
  for(var i = 0; i < members.length; i++) {
    if (members[i][2] != "" && members[i][2] == getMailAddress()) {
      return members[i][0];
    }
  }
  return "";
}

function getMyNickName() {
  var masterSheet = getMasterMemberSheet();
  var mailAddresses = masterSheet.getRange(2, 2, masterSheet.getLastRow() - 1, 2).getValues();
  for(var i = 0; i < mailAddresses.length; i++) {
    if (mailAddresses[i][1] != "" && mailAddresses[i][1] == getMailAddress()) {
      return mailAddresses[i][0];
    }
  }
  return getMailAddress();
}

/**
 * 有効なメールアドレスかをチェック
 */
function validateMailAddresses()
{
  var masterSheet = getMasterMemberSheet();
  var mailAddresses = masterSheet.getRange(2, 3, masterSheet.getLastRow() - 1, 2).getValues();
  for(var i = 0; i < mailAddresses.length; i++) {
    if (mailAddresses[i][0] == getMailAddress()) {
      return mailAddresses[i][1] == 1; 
    }
  }
  return false;
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
      return {memberId: range[i][0], point: range[i][2], totalPoint: range[i][3], row: row};
    }
  }
  return {memberId: "", point: 0, totalPoint: 0, row: 0};
  
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
  if (myData.point < form.point) {
    return {result: "ERROR", message: "ポイントが不足しています。"};
  }
  if (!isAvailablePayment(myData, form)) {
    return {result: "ERROR", message: "自分には投票できません。"};
  }
  var masterSheet = getResultSheet();
  masterSheet.getRange(myData.row, 3, 1, 1).setValues([[myData.point - form.point]]);
  if (isValidName(form.name)) {
    sheet.appendRow([getMyNickName(), form.name, form.point, new Date()]);
    return {result: "SUCCESS"};
  }
  var name = transName(form.name);
  if (name == form.name) {
    sheet.appendRow([getMyNickName(), form.name, form.point, new Date(), "x"]);
  } else {
    sheet.appendRow([getMyNickName(), name, form.point, new Date()]);
  }
  return {result: "SUCCESS"};
}

/**
 * 投票可能かをチェック
 * 1. 所持ポイント以内であること
 * 2. 自分への投票でないこと
 */
function isAvailablePayment(myData, form) {
  var resultSheet = getResultSheet();
  var masters = resultSheet.getRange(2, 1, resultSheet.getLastRow() - 1, 2).getValues();
  var masterName;
  for (var i = 0; i < masters.length; i++) {
    if (masters[i][0] == form.memberId && masters[i][1] == form.name) {
      return false;
    } else if (masters[i][0] == form.memberId) {
      masterName = masters[i][0];
    }
  }
  var tmpName = transName(form.name);
  if (tmpName == masterName) {
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
  var resultRange = resultSheet.getRange(2, 1, resultSheet.getLastRow() - 1, 2).getValues();
  // 全てメンバーを取得する
  var allMembers = getAllMemberList();
  // 登録されていないメンバーの項目を作成
  var exist = false;
  for (var i = 0; i < allMembers.length; i++) {
    exist = false;
    for (var j = 0; j < resultRange.length; j++) {
      if (allMembers[i][0] == resultRange[j][0]) {
        if (allMembers[i][1] != resultRange[j][1]) {
          // あだ名が変わってたら新しいのに変える
          resultSheet.getRange(j + 2, 2, 1, 1).setValues([[allMembers[i][1]]]);
        }
        exist = true;
        break;
      }
    }
    if (!exist) {
      resultSheet.appendRow([allMembers[i][0], allMembers[i][1]]);
    }
  }
  // 月次処理を行う
  resultRange = resultSheet.getRange(2, 1, resultSheet.getLastRow() - 1, 2).getValues();
  var logSheet = getLogSheet();
  // 0. 全ての使用可能ポイントをゼロする
  for (var i = 0; i < resultRange.length; i++) {
    resultSheet.getRange(i + 2, 3, 1, 1).setValues([[0]]);
  }
  // 1. ログから各人の累計ポイントを計算する
  var logRange = logSheet.getRange(2, 2, logSheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < resultRange.length; i++) {
    var totalPoint = 0;
    var tmpName = resultRange[i][1];
    for (var j = 0; j < logRange.length; j++) {
      if (logRange[j][0] == tmpName) {
        totalPoint += logRange[j][1];
      }
    }
    resultSheet.getRange(i + 2, 4, 1, 1).setValues([[totalPoint]]);
  }
  // 2. 全ての使用可能ポイントをリセットする
  for (var i = 0; i < resultRange.length; i++) {
    resultSheet.getRange(i + 2, 3, 1, 1).setValues([[1000]]);
  }
}

function getAllMemberList() {
  var masterSheet = getMasterMemberSheet();
  return masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 2).getValues();
}
