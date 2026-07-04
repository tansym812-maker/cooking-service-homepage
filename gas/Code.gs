// ===================================================
// 作り置きサポート — お問い合わせフォーム受信
// ===================================================

// ① ここにスプレッドシートIDを貼り付ける
//    （スプレッドシートのURLの /d/〜/edit の〜の部分）
var SPREADSHEET_ID = '1plyO7mxEK1iSjvYEGcFnuPt703G3frA-5XbjAhbLIJs';
var SHEET_NAME = 'お問い合わせ';

// ホームページのURL（送信完了後のリダイレクト先）
var HOMEPAGE_URL = 'https://tansym812-maker.github.io/cooking-service-homepage/';

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

    // 送信完了ページへリダイレクト
    var html = '<script>window.top.location.href="' + HOMEPAGE_URL + '?sent=1";</script>';
    return HtmlService.createHtmlOutput(html);

  } catch (err) {
    var errHtml = '<script>window.top.location.href="' + HOMEPAGE_URL + '?sent=error";</script>';
    return HtmlService.createHtmlOutput(errHtml);
  }
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
