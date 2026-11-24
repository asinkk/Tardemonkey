// 網頁資料出口：GitHub Pages 的網頁（live / trades / options）打開時來這裡抓最新資料，
// 這樣 GitHub 不用每半小時 commit。部署成「網頁應用程式」：執行身分＝我、存取權＝所有人。
// 只讀試算表、不改任何東西；回傳各分頁「顯示出來的文字」＋週選紀錄 JSON。
function doGet(e) {
  var cache = CacheService.getScriptCache(), hit = cache.get('web_json');
  var body = hit || webBuild_();
  if (!hit && body.length < 95000) cache.put('web_json', body, 60);   // 很多人同時開時，1 分鐘內共用一份
  var cb = e && e.parameter && e.parameter.callback;
  if (cb && /^[A-Za-z_$][\w$.]{0,60}$/.test(cb))
    return ContentService.createTextOutput(cb + '(' + body + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function webBuild_() {
  var ss = SpreadsheetApp.getActive();
  function rows(name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 1) return [];
    return sh.getRange(1, 1, sh.getLastRow(), Math.min(sh.getLastColumn(), 14)).getDisplayValues().map(function (r) {
      var n = r.length; while (n > 0 && r[n - 1] === '') n--;   // 去掉每列後面的空格子
      return r.slice(0, n);
    });
  }
  var opts = [];
  try {
    var sh = ss.getSheetByName('資料_週選');
    if (sh && sh.getLastRow() > 1)
      opts = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues().filter(function (r) { return r[0]; })
        .map(function (r) { return JSON.parse(r[0]); });
  } catch (err) { opts = []; }
  var live = rows('即時訊號');
  var m = String((live[1] || [])[5] || '').match(/\d{4}-\d\d-\d\d \d\d:\d\d/);
  return JSON.stringify({
    ok: true,
    updated: m ? m[0] : '',                                                       // 試算表上次算完的台北時間
    served: new Date(Date.now() + 8 * 3600e3).toISOString().slice(0, 16).replace('T', ' '),
    live: live, trades: rows('成交紀錄'), status: rows('說明與狀態'), options: opts
  });
}
