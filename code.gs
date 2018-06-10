// configシートの設定値を保持する変数
var config = {};

/**
* スプレッドシート起動時に実行される
*/
function onOpen(){
  // 独自メニュー追加
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var name = "スクリプト実行";
  var subMenus = [
    { name: "ユーザの一括追加", functionName: "addUsers" },
    { name: "ユーザの一括削除", functionName: "deleteUsers" }
  ];
  activeSpreadsheet.addMenu(name, subMenus);
}

/**
* シートを参照し、Backlogユーザーの追加を行う
*/
function addUsers() {
  // config設定値の読み込み
  setUpConfig_()
  // メンバーシートの読み込み
  var sheet = getMemberSheet_();

  // 登録ユーザーリスト
  var registeredUsers = [];

  // シートの各行を確認し必要であればユーザー登録を行う
  var rowSize = sheet.getLastRow();
  for (var i = 0; i < rowSize; i++) {
    var row = (i+1); // 行番号は1から始まるためi+1

    if(row == config['headerRow']){
      // ヘッダー行は除外のためスキップ
      continue;
    }
    if(!canBeRegistered_(sheet, row)){
      // アカウント作成対象外はスキップ
      continue;
    }

    // パラメータ決定
    // メールアドレス
    var inputMailAddress = sheet.getRange(row, config['mailAddressColumn']).getValue();
    // ユーザID
    var inputUserId = generateUserId_(inputMailAddress);
    // 初期パスワード
    var inputPassword = getRandomString_(config['initialPasswordLength']);
    // ハンドルネーム
    var inputName = sheet.getRange(row, config['nameColumn']).getValue();
    // 権限
    var inputRoleType= String(config['defaultRole']);

    var user = {
      'userId' : inputUserId,
      'password' : inputPassword,
      'name' : inputName,
      'mailAddress' : inputMailAddress,
      'roleType' : inputRoleType
    };

    // 登録処理
    var registeredUser = postAddUserApi_(user);
    // TODO password を結果に残すかどうか
    registeredUser['password'] = inputPassword;

    // Backlogアカウント情報を設定する
    sheet.getRange(row, config['backlogAccountStatusColumn']).setValue(config['isRegistered']);
    sheet.getRange(row, config['backlogAccountIDColumn']).setValue(registeredUser['id']);

    // 登録ユーザーリストに追加
    registeredUsers.push(registeredUser);
  }

  if(registeredUsers.length > 0){
    // 登録結果を新しいシートに反映する
    createResult_(config['registeredSheetName'], registeredUsers);
  }
}

/**
* シートを参照し、Backlogユーザーの削除を行う
*/
function deleteUsers() {
  // config設定値の読み込み
  setUpConfig_()
  // メンバーシートの読み込み
  var sheet = getMemberSheet_();

  // 登録ユーザーリスト
  var deletedUsers = [];

  // シートの各行を確認し必要であればユーザー削除を行う
  var rowSize = sheet.getLastRow();
  for (var i = 0; i < rowSize; i++) {
    var row = (i+1); // 行番号は1から始まるためi+1

    if(row == config['headerRow']){
      // ヘッダー行は除外のためスキップ
      continue;
    }
    if(!canBeDeleted_(sheet, row)){
      // アカウント削除対象外はスキップ
      continue;
    }

    // パラメータ決定
    // アカウントID
    var inputId = sheet.getRange(row, config['backlogAccountIDColumn']).getValue();

    // 削除処理
    var deletedUser = postDeleteUserApi_(inputId)

    // Backlogアカウント情報を設定する
    sheet.getRange(row, config['backlogAccountStatusColumn']).setValue(config['isDeleted']);

    // 削除ユーザーリストに追加
    deletedUsers.push(deletedUser);
  }

  if(deletedUsers.length > 0){
    // 削除結果を新しいシートに反映する
    createResult_(config['deletedSheetName'], deletedUsers);
  }
}

/**
* configシートの内容を取り込む
*/
function setUpConfig_(){
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var　configSheet = activeSpreadsheet.getSheetByName('config');

  var row = 2; // データ行開始
  var column = 1; // カラム開始
  var numRows = configSheet.getLastRow()-1; // ヘッダー行を除外した行数分取り込む
  var numColumns = 2; // 取り込む列数(設定キー, 設定値)
  var values = configSheet.getRange(row, column, numRows, numColumns).getValues();

  Logger.log('config 取り込み開始');
  for(var i =0; i < values.length; i++){
    var key = values[i][0];
    var value = values[i][1]
    config[key] = value;
    Logger.log(key + " : " + config[key]);
  }
  Logger.log('config 取り込み終了');
}

/**
* ユーザー管理対象メンバーが記載されているシートを読み込む
*/
function getMemberSheet_() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = activeSpreadsheet.getSheetByName(config['memberSheetName']);
  return memberSheet;
}

/**
* シートを参照し、アカウントを作成可能か判定する
*/
function canBeRegistered_(sheet, row) {
  // Backlogアカウント状態
  var status = sheet.getRange(row, config['backlogAccountStatusColumn']).getValue();
  return (status == config['canBeRegistered']);
}

/**
* メールアドレスをもとにユーザID文字列を生成する
* @param mailAddress メールアドレス
*/
function generateUserId_(mailAddress){
  var regex = /(.+)@/;
  // ユーザID(メールアドレスのローカルパートを設定)
  var userId = mailAddress.match(regex)[0].replace("@","");
  return userId;
}

/**
* ランダム文字列を返す
* @param len 生成桁数
*/
function getRandomString_(len){
  //使用文字の定義
  const str = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!#$%&=~/*-+";

  //ランダムな文字列の生成
  var result = "";
  for(var i = 0; i < len; i++){
    result += str.charAt(Math.floor(Math.random() * str.length));
  }
  return result;
}

/**
* Backlog API(ユーザーの追加)を呼び出す
*/
function postAddUserApi_(user) {
  // APIのURL
  const url = 'https://' + config['spaceId'] + '/api/v2/users?apiKey=' + config['apiKey'];

  // オプション
  var options = {
    'method' : 'post',
    'payload' : user
  };    

  // リクエスト送信
  var result = UrlFetchApp.fetch(url, options).getContentText();

  return JSON.parse(result);
}

/**
* API実行結果を新しいシートに反映する
*/
function createResult_(sheetName, response){
  var resultSheet = createNewSheet_(sheetName);
  var resultRows = response.length;
  var headers = Object.keys(response[0]);

  // ヘッダー行
  resultSheet.appendRow(headers);

  // データ行
  var resultColumns = headers.length;
  for(var i = 0; i < resultRows; i++){
    var data = [];
    for(var j = 0; j < resultColumns; j++){
      var user = response[i];
      var value = user[headers[j]];
      data.push(value);
    }
    resultSheet.appendRow(data);
  }
}

/**
* スプレッドシートに新しいシートを作成する
*/
function createNewSheet_(sheetName){
  // シート名が重複しないように日時情報をシート名に含める
  var now = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd_HH:mm:ss");

  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var resultSheet = activeSpreadsheet.insertSheet(sheetName + "_" + now);
  return resultSheet;
}

/**
* シートを参照し、アカウントを削除可能か判定する
*/
function canBeDeleted_(sheet, row) {
  // Backlogアカウント状態
  var status = sheet.getRange(row, config['backlogAccountStatusColumn']).getValue();
  return (status == config['canBeDeleted']);
}

/**
* Backlog API(ユーザーの削除)を呼び出す
* @param userId 削除するユーザーのID(数字)
*/
function postDeleteUserApi_(userId) {
  // APIのURL
  const url = 'https://' + config['spaceId'] + '/api/v2/users/' + userId  + '?apiKey=' + config['apiKey'];

  // オプション
  var options = {
    'method' : 'delete'
  };    

  // リクエスト送信
  var result = UrlFetchApp.fetch(url, options).getContentText();

  return JSON.parse(result);
}
