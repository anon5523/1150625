/**
 * 1150625豪雨 新舊災點共同填報 — Google Apps Script 後端 (v4)
 * v2 變更:
 *  - 支援批次寫入({batch:[...]}),一次可寫入數百筆
 *  - 「位置更新」紀錄與「判定填報」紀錄分流:
 *      doGet 回傳 {fills:{各序號最新判定}, pos:{各序號最新座標}}
 *    → 拖曳修正的座標跨裝置同步,且不影響已填報判定
 *  - 新增「淹水高度(公分)」「權責單位」欄位
 *  - v4 新增「致災原因」「致災原因說明」「改善對策(含期程)」「改善對策說明」欄位
 * ※ 更新方式:貼上本檔後,部署 → 管理部署作業 → 鉛筆 → 版本:新版本 → 部署(網址不變)
 */
var SHEET_NAME = '填報紀錄';
var HEADERS = ['填報時間','序號','行政區','淹水路段(地址)','判定類型',
               '非積淹水原因','新災點合併名稱','舊災點項次','舊災點名稱',
               '填報人','緯度','經度','定位方式','淹水高度(公分)','權責單位',
               '致災原因','致災原因說明','改善對策(含期程)','改善對策說明'];
var JUDGE_TYPES = ['舊災點','新災點','非積淹水事件'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < HEADERS.length) {
    // 舊版表頭補上新欄位
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sh;
}

function rowToRec_(r) {
  return {
    ts: r[0] instanceof Date ? r[0].toISOString() : r[0],
    seq: r[1], dist: r[2], addr: r[3], type: r[4],
    reason: r[5], newName: r[6], oldIdx: r[7], oldName: r[8],
    reporter: r[9], lat: r[10], lon: r[11], locMethod: r[12],
    depth: r.length > 13 ? r[13] : '',
    agency: r.length > 14 ? r[14] : '',
    causes: r.length > 15 ? r[15] : '',
    causeNote: r.length > 16 ? r[16] : '',
    measures: r.length > 17 ? r[17] : '',
    measureNote: r.length > 18 ? r[18] : ''
  };
}

function doGet(e) {
  var sh = getSheet_();
  var rows = sh.getDataRange().getValues();
  rows.shift();
  var fills = {}, pos = {};
  rows.forEach(function (r) {
    if (r[1] === '' || r[1] == null) return;
    var rec = rowToRec_(r);
    var seq = String(rec.seq);
    if (JUDGE_TYPES.indexOf(rec.type) > -1) fills[seq] = rec;      // 最新判定
    if (rec.lat !== '' && rec.lon !== '' && rec.lat != null && rec.lon != null) {
      pos[seq] = { lat: rec.lat, lon: rec.lon, locMethod: rec.locMethod, ts: rec.ts }; // 最新座標(任何類型)
    }
  });
  return ContentService.createTextOutput(JSON.stringify({ fills: fills, pos: pos }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var d = JSON.parse(e.postData.contents);
    var items = d.batch ? d.batch : [d];
    if (!items.length) throw new Error('空的批次');
    var rows = items.map(function (p) {
      if (!p.seq || !p.type) throw new Error('缺少必要欄位 seq / type');
      return [new Date(), p.seq, p.dist || '', p.addr || '', p.type,
              p.reason || '', p.newName || '', p.oldIdx || '', p.oldName || '',
              p.reporter || '', p.lat != null ? p.lat : '', p.lon != null ? p.lon : '',
              p.locMethod || '', p.depth != null ? p.depth : '', p.agency || '',
              p.causes || '', p.causeNote || '', p.measures || '', p.measureNote || ''];
    });
    var sh = getSheet_();
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, n: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
