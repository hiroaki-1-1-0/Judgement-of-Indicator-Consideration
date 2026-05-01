/**
 * 大喜利 考慮判定 Webアプリ（Google Apps Script）
 * 1) SPREADSHEET_ID に記録先スプレッドシートIDを設定
 * 2) setupSheet を最初に1回だけ実行
 * 3) ウェブアプリとしてデプロイ
 */
const SPREADSHEET_ID = 'ここに記録先スプレッドシートIDを貼り付けてください';
const SHEET_NAME = 'responses';

const METRICS = [
  '語彙的新規性の低さ',
  '不一致の解消度合い',
  '比喩の使用度',
  '連想距離の遠さ',
  'ボケの面白さ',
  '曖昧さの活用度',
  '無害な違反の度合い',
  '視点の転換度合い',
  '特殊文字の少なさ',
  'お題に対する回答の短さ',
  '文字数の少なさ'
];

// 保存値: 考慮している=1、考慮していない=-1、考慮しているか不明=0
const ALLOWED_VALUES = [1, -1, 0];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('大喜利 考慮判定')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  const headers = [
    '送信日時',
    'セッションID',
    '名前（ランサーズ表示名）',
    '項目ID',
    '項目番号',
    '大喜利のお題',
    '大喜利の回答1',
    '大喜利の回答2',
    '相対評価時の文章'
  ].concat(METRICS).concat(['ユーザーエージェント']);
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function saveResponse(payload) {
  if (!payload) throw new Error('送信データが空です。');
  const required = ['sessionId', 'workerName', 'itemId', 'itemIndex', 'topic', 'answer1', 'answer2', 'relativeText', 'ratings'];
  required.forEach(function (key) {
    if (payload[key] === undefined || payload[key] === null || String(payload[key]).trim() === '') {
      throw new Error('未入力の項目があります: ' + key);
    }
  });

  METRICS.forEach(function (metric) {
    const value = Number(payload.ratings[metric]);
    if (ALLOWED_VALUES.indexOf(value) === -1) {
      throw new Error('未選択または不正な評価があります: ' + metric);
    }
    payload.ratings[metric] = value;
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) setupSheet();
    const row = [
      new Date(),
      payload.sessionId,
      payload.workerName,
      payload.itemId,
      payload.itemIndex,
      payload.topic,
      payload.answer1,
      payload.answer2,
      payload.relativeText
    ].concat(METRICS.map(function (metric) { return payload.ratings[metric]; })).concat([payload.userAgent || '']);
    sheet.appendRow(row);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}
