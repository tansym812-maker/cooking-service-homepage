// ===================================================
// 作り置きサポート — お問い合わせフォーム受信
// ===================================================

// ① ここにスプレッドシートIDを貼り付ける
//    （スプレッドシートのURLの /d/〜/edit の〜の部分）
var SPREADSHEET_ID = '1plyO7mxEK1iSjvYEGcFnuPt703G3frA-5XbjAhbLIJs';
var SHEET_NAME = 'お問い合わせ';

// ホームページのURL（送信完了後のリダイレクト先）
var HOMEPAGE_URL = 'https://tansym812-maker.github.io/cooking-service-homepage/';

// ② LINE通知の設定（トークンはコードに直接書かない）
//    GASエディタの「プロジェクトの設定（歯車）> スクリプト プロパティ」に以下の2つを登録する
//    ・LINE_CHANNEL_ACCESS_TOKEN … Messaging APIのチャネルアクセストークン（長期）
//    ・LINE_USER_ID              … 通知を受け取る自分のユーザーID（チャネル基本設定の「あなたのユーザーID」）

// ===================================================
// フォーム送信受信（POST）
// ===================================================
function doPost(e) {
  try {
    var params = e.parameter || {};

    var sheet = SpreadsheetApp
      .openById(SPREADSHEET_ID)
      .getSheetByName(SHEET_NAME);

    var timestamp = new Date();

    sheet.appendRow([
      timestamp,
      params['namae']   || '',
      params['email']   || '',
      params['area']    || '',
      params['plan']    || '',
      params['message'] || '',
      '未対応',
      ''
    ]);

    // LINEに通知（失敗してもスプレッドシート記録には影響させない）
    try {
      sendLineNotification(params);
    } catch (lineErr) {
      // 通知失敗は無視（記録は済んでいるため）
    }

    // 送信完了ページへリダイレクト
    var html = '<script>window.top.location.href="' + HOMEPAGE_URL + '?sent=1";</script>';
    return HtmlService.createHtmlOutput(html);

  } catch (err) {
    var errHtml = '<script>window.top.location.href="' + HOMEPAGE_URL + '?sent=error";</script>';
    return HtmlService.createHtmlOutput(errHtml);
  }
}

// ===================================================
// LINE通知（Messaging APIのプッシュメッセージ）
// ===================================================
function sendLineNotification(params) {
  var props  = PropertiesService.getScriptProperties();
  var token  = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  var userId = props.getProperty('LINE_USER_ID');

  // スクリプトプロパティが未設定なら何もしない
  if (!token || !userId) return;

  var text = '【作り置きサポート】\n'
    + '新しいお問い合わせが届きました。\n'
    + '────────────\n'
    + 'お名前：' + (params['namae'] || '未入力') + '\n'
    + 'メール：' + (params['email'] || '未入力') + '\n'
    + 'エリア：' + (params['area'] || '未入力') + '\n'
    + 'プラン：' + (params['plan'] || '未入力') + '\n'
    + '────────────\n'
    + (params['message'] || '（お問い合わせ内容なし）');

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify({
      to: userId,
      messages: [{ type: 'text', text: text }]
    })
  });
}

// LINE通知のテスト（スクリプトプロパティ設定後に1回手動実行して確認する）
function testLineNotification() {
  sendLineNotification({
    namae: 'テスト 太郎',
    email: 'test@example.com',
    area: '神戸市北区 藤原台',
    plan: '月2回 定期プラン',
    message: 'これはLINE通知のテストです。スマホにこのメッセージが届けば連携成功です。'
  });
  Logger.log('テスト通知を送信しました。LINEを確認してください。');
}

// ===================================================
// スプレッドシートの初期セットアップ（最初に1回だけ手動実行）
// ===================================================
function setup() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // ヘッダー行
  var headers = [
    '受信日時',
    'お名前',
    'メールアドレス',
    'お住まいのエリア',
    'ご希望のプラン',
    'お問い合わせ内容',
    '対応状況',
    'メモ'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダーのデザイン（グリーン背景・白文字・太字）
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4A7C59');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // 列幅
  sheet.setColumnWidth(1, 170); // 受信日時
  sheet.setColumnWidth(2, 120); // お名前
  sheet.setColumnWidth(3, 210); // メールアドレス
  sheet.setColumnWidth(4, 130); // お住まいのエリア
  sheet.setColumnWidth(5, 170); // ご希望のプラン
  sheet.setColumnWidth(6, 320); // お問い合わせ内容
  sheet.setColumnWidth(7, 100); // 対応状況
  sheet.setColumnWidth(8, 220); // メモ

  // 1行目を固定
  sheet.setFrozenRows(1);

  // 「対応状況」列にデータ入力規則（ドロップダウン）を設定
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未対応', '対応中', '完了'], true)
    .build();
  sheet.getRange(2, 7, 999, 1).setDataValidation(statusRule);

  Logger.log('セットアップ完了しました。');
}
