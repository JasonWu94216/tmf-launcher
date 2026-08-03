/**
 * 小台 AI 儀表盤 — 資料橋（Google Apps Script，獨立版）
 * 讀取「小台模擬交易紀錄」試算表，回傳 JSON 給手機 App；用密碼保護。
 */
var PASSWORD = "0216";
var SPREADSHEET_ID = "16H8w-d3e_DA_ZlOtlMzFpDpkM9hxEr1aqZqJT-5JKDs";
var TABS = ["每日訊號", "交易明細", "統計摘要"];

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var payload;
  if (p.pw !== PASSWORD) {
    payload = { ok: false, error: "auth" };
  } else {
    payload = { ok: true, ts: new Date().toISOString(), data: readAll() };
  }
  var json = JSON.stringify(payload);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function readAll() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var out = {};
  for (var i = 0; i < TABS.length; i++) {
    var sh = ss.getSheetByName(TABS[i]);
    out[TABS[i]] = sh ? sh.getDataRange().getDisplayValues() : [];
  }
  return out;
}
