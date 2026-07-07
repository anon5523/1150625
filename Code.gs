/**
 * 1150625豪雨 新舊災點共同填報 — Google Apps Script 後端
 * 綁定一個 Google 試算表使用。
 * doGet  : 回傳每個序號「最新一筆」填報結果(JSON),供網頁載入同步
 * doPost : 寫入一筆填報紀錄(append,保留完整歷程供稽核)
 */
var SHEET_NAME = '填報紀錄';
var HEADERS = ['填報時間','序號','行政區','淹水路段(地址)','判定類型',
               '非積淹水原因','新災點合併名稱','舊災點項次','舊災點名稱',
               '填報人','緯度','經度','定位方式'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doGet(e) {
  var sh = getSheet_();
  var rows = sh.getDataRange().getValues();
  rows.shift(); // 去表頭
  var latest = {};
  rows.forEach(function (r) {
    if (r[1] === '' || r[1] == null) return;
    latest[String(r[1])] = {
      ts: r[0] instanceof Date ? r[0].toISOString() : r[0],
      seq: r[1], dist: r[2], addr: r[3], type: r[4],
      reason: r[5], newName: r[6], oldIdx: r[7], oldName: r[8],
      reporter: r[9], lat: r[10], lon: r[11], locMethod: r[12]
    };
  }); // 由上而下讀取,後面(較新)自然覆蓋前面 → 每序號保留最新一筆
  return ContentService.createTextOutput(JSON.stringify(latest))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    if (!d.seq || !d.type) throw new Error('缺少必要欄位 seq / type');
    var sh = getSheet_();
    sh.appendRow([new Date(), d.seq, d.dist || '', d.addr || '', d.type,
                  d.reason || '', d.newName || '', d.oldIdx || '', d.oldName || '',
                  d.reporter || '', d.lat || '', d.lon || '', d.locMethod || '']);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
