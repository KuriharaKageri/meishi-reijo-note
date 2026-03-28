// ============================================================
// 名刺礼状アシスタント（汎用版）
// Google Apps Script（GAS）
//
// ── このファイルの使い方 ──────────────────────────────────────
// ① Googleドライブで新しいスプレッドシートを作成（名前は何でもOK）
// ② メニュー「拡張機能」→「Apps Script」を開く
// ③ 最初から入っているコードをすべて削除する
// ④ このファイルの内容をコピーして貼り付け、保存（Command+S）
// ⑤ 上部の関数名プルダウンで「setupSheet」を選んで ▶ 実行
//    → スプレッドシートにヘッダー行が自動作成されます
// ⑥ メニュー「デプロイ」→「新しいデプロイ」
//    設定：種類 = ウェブアプリ
//          次のユーザーとして実行 = 自分
//          アクセスできるユーザー = 全員
//    「デプロイ」を押し、表示されたURLをコピー
// ⑦ アプリの「設定」タブ → GAS デプロイURL欄に貼り付けて保存
// ============================================================

const SHEET_NAME = '名刺記録';

// ── ① シート初期セットアップ（最初に1回だけ実行） ──────────────
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // ヘッダー行の内容
  const headers = [
    '記録日時',      // A
    '氏名',          // B
    'ふりがな',      // C
    '会社名',        // D
    '部署',          // E
    '役職',          // F
    'メールアドレス', // G
    '電話番号',      // H
    '住所',          // I
    '備考',          // J（2件目以降の電話・メールなど）
    '名刺交換日',    // K
    '状況・メモ',    // L
    '差出人',        // M
  ];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  // ヘッダー書式（深い赤茶色）
  headerRange.setBackground('#6b2a2a');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment('center');

  // 列幅（各列の内容に合わせて設定）
  const colWidths = [
    140,  // 記録日時
    90,   // 氏名
    90,   // ふりがな
    180,  // 会社名
    120,  // 部署
    80,   // 役職
    200,  // メールアドレス
    130,  // 電話番号
    200,  // 住所
    200,  // 備考
    90,   // 名刺交換日
    300,  // 状況・メモ
    150,  // 差出人
  ];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // 1行目を固定（スクロールしてもヘッダーが見える）
  sheet.setFrozenRows(1);

  SpreadsheetApp.flush();
  Logger.log('セットアップ完了：シート名「' + SHEET_NAME + '」');
}

// ── ② POSTリクエスト受信（アプリからデータが送られてきたとき） ──
function doPost(e) {
  try {
    let data;
    // データの受け取り方を2パターン対応（ブラウザ環境の違い対応）
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.data) {
      data = JSON.parse(e.parameter.data);
    } else {
      throw new Error('データが受信できませんでした');
    }

    appendRecord(data);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: '記録しました' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('エラー: ' + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── ③ GETリクエスト（URLをブラウザで開いて動作確認するとき） ──
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      message: '名刺礼状アシスタント GAS 稼働中',
      sheet: SHEET_NAME
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ④ スプレッドシートに1行追加する ─────────────────────────
function appendRecord(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  // シートがなければ自動作成
  if (!sheet) {
    setupSheet();
    sheet = ss.getSheetByName(SHEET_NAME);
  }

  // 現在時刻（日本時間）
  const now = new Date();
  const jst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const timestamp = Utilities.formatDate(jst, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

  // 1行分のデータ（ヘッダーの順番と対応させる）
  const row = [
    timestamp,            // 記録日時
    data.name    || '',   // 氏名
    data.kana    || '',   // ふりがな
    data.company || '',   // 会社名
    data.dept    || '',   // 部署
    data.role    || '',   // 役職
    data.email   || '',   // メールアドレス
    data.tel     || '',   // 電話番号
    data.address || '',   // 住所
    data.note    || '',   // 備考
    data.date    || '',   // 名刺交換日
    data.memo    || '',   // 状況・メモ
    data.sender  || '',   // 差出人
  ];

  sheet.appendRow(row);

  // データ行の書式設定
  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(lastRow, 1, 1, row.length);
  dataRange.setFontSize(10);
  dataRange.setVerticalAlignment('top');
  dataRange.setWrap(true);   // セル内で折り返し表示

  // 偶数行に薄いベージュ背景（見やすさのため）
  if (lastRow % 2 === 0) {
    dataRange.setBackground('#faf6f0');
  }

  SpreadsheetApp.flush();
  Logger.log('記録追加：' + (data.name || '（名前なし）') + ' / ' + timestamp);
}

// ── ⑤ 動作テスト用（開発時に手動実行して確認する） ─────────────
function testAppendRecord() {
  const testData = {
    name:    'テスト 太郎',
    kana:    'てすと たろう',
    company: '株式会社テスト商事',
    dept:    '広報部',
    role:    '部長',
    email:   'test@example.com',
    tel:     '090-0000-0000',
    address: '東京都千代田区1-1-1',
    note:    'TEL: 03-0000-0000',
    date:    '2026-03-29',
    memo:    'テスト取材にて名刺交換。ダイヤ改正について詳しくお話しいただいた。',
    sender:  '栗原 景\nフリーランスライター',
  };

  appendRecord(testData);
  Logger.log('テスト記録を追加しました');
}
