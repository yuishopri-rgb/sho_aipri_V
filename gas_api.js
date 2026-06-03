// ============================================================
//  アイプリ 所持数管理 API  (Google Apps Script)
//  スプレッドシートに貼り付けて「ウェブアプリとして導入」する
// ============================================================

const HEADER_SHEET = 'ヘッダー';   // 弾一覧シート名
const SKIP_SHEETS  = ['ヘッダー']; // APIから除外するシート名

// ------------------------------------------------------------
//  GET  ?action=getDans          → 弾一覧
//  GET  ?action=getItems&dan=v1  → 指定弾のアイテム一覧+所持数
// ------------------------------------------------------------
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'getDans') {
    return jsonResponse(getDans());
  }
  if (action === 'getItems') {
    const dan = e.parameter.dan;
    return jsonResponse(getItems(dan));
  }
  return jsonResponse({ error: 'unknown action' });
}

// ------------------------------------------------------------
//  POST  { action:'saveItem', dan, itemId, アクセ, トップス, ボトムス, シューズ }
// ------------------------------------------------------------
function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  if (data.action === 'saveItem') {
    const result = saveItem(data);
    return jsonResponse(result);
  }
  return jsonResponse({ error: 'unknown action' });
}

// ------------------------------------------------------------
//  弾一覧を返す
// ------------------------------------------------------------
function getDans() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName(HEADER_SHEET);
  const rows   = sheet.getDataRange().getValues();
  const dans   = [];

  // 1行目はヘッダー（弾情報, Image, コード数, 所持数, 所持率）
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0]) continue;
    dans.push({
      id       : row[0],          // "v1弾" など
      image    : row[1],          // 代表画像URL
      total    : row[2],          // コード数
      owned    : row[3],          // 所持数
      rate     : row[4],          // 所持率
    });
  }
  return { dans };
}

// ------------------------------------------------------------
//  指定弾のアイテム一覧を返す
// ------------------------------------------------------------
function getItems(danId) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  // シート名が "v1弾" の場合と "v1" の場合どちらにも対応
  const sheet = ss.getSheetByName(danId) || ss.getSheetByName(danId.replace('弾',''));
  if (!sheet) return { error: 'sheet not found: ' + danId };

  const rows  = sheet.getDataRange().getValues();
  const items = [];

  // 1行目はヘッダー（Item ID, Name, アクセ, トップス, ボトムス, シューズ, Image）
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0]) continue;
    items.push({
      id    : String(row[0]),
      name  : row[1],
      アクセ  : Number(row[2]) || 0,
      トップス : Number(row[3]) || 0,
      ボトムス : Number(row[4]) || 0,
      シューズ : Number(row[5]) || 0,
      image : row[6] || '',
    });
  }
  return { dan: danId, items };
}

// ------------------------------------------------------------
//  所持数を書き戻す
// ------------------------------------------------------------
function saveItem(data) {
  const { dan, itemId } = data;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(dan) || ss.getSheetByName(dan.replace('弾',''));
  if (!sheet) return { error: 'sheet not found: ' + dan };

  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(itemId)) {
      // C列=アクセ(3), D列=トップス(4), E列=ボトムス(5), F列=シューズ(6)
      sheet.getRange(i + 1, 3).setValue(data['アクセ']  ?? rows[i][2]);
      sheet.getRange(i + 1, 4).setValue(data['トップス'] ?? rows[i][3]);
      sheet.getRange(i + 1, 5).setValue(data['ボトムス'] ?? rows[i][4]);
      sheet.getRange(i + 1, 6).setValue(data['シューズ'] ?? rows[i][5]);
      return { success: true, itemId };
    }
  }
  return { error: 'item not found: ' + itemId };
}

// ------------------------------------------------------------
//  ユーティリティ
// ------------------------------------------------------------
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
