// 網頁打開時去 Google 試算表抓最新資料（試算表的 Apps Script 每 15 分鐘自己跟期交所算 v3）。
// 這樣 GitHub 不用每半小時 commit：網頁本身只有改版面時才推，資料都從試算表來。
// 抓不到（網址沒設、Google 掛了、網路擋住）就保留網頁裡原本的內容，並標示資料時間。
// 試算表比網頁裡的資料新才換上；本機排程剛推過、比試算表新的時候照用網頁的。
(function () {
  'use strict';
  // 試算表「擴充功能 → Apps Script → 部署 → 網頁應用程式」的網址（.../exec）
  var SHEET_API = 'https://script.google.com/macros/s/AKfycbxJSHwfsSu0fFvytrRzJdpAtilmmGrbQ9YllYTs3-6ALMeuJBMPwh6z9yk8qKHzX5Ph/exec';
  var REFRESH_MIN = 5;                 // 網頁開著的時候每 5 分鐘再抓一次

  var page = (document.querySelector('meta[name="tm-page"]') || {}).content || '';
  var genMeta = document.querySelector('meta[name="tm-generated"]');
  var pageGen = genMeta ? genMeta.content : '';
  var $ = function (s) { return document.querySelector(s); };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]; }); };
  var signed = function (s) {   // "+52,148 元" → 紅綠色
    return esc(s).replace(/([+\-−]\d[\d,.]*)/g, function (v) { return '<span class="' + (/^[\-−]/.test(v) ? 'dn' : /^\+0(\.0+)?$/.test(v) ? '' : 'up') + '">' + v + '</span>'; });
  };
  var cell = function (r, i) { return (r && r[i] != null ? String(r[i]) : '').trim(); };
  var blank = function (r) { return !r || !r.some(function (c) { return String(c).trim(); }); };

  function fetchSheet() {
    if (/PLACEHOLDER/.test(SHEET_API)) return Promise.reject(new Error('還沒設定試算表網址'));
    return fetch(SHEET_API + '?t=' + Date.now(), {cache: 'no-store', redirect: 'follow'})
      .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
      .then(function (d) { if (!d || !d.ok) throw new Error('資料格式不對'); return d; });
  }

  function ageText(t) {
    var m = Math.round((Date.now() - new Date(t.replace(' ', 'T') + ':00+08:00')) / 6e4);
    return m < 1 ? '剛剛' : m < 120 ? m + ' 分鐘前' : Math.round(m / 60) + ' 小時前';
  }
  function marketOpen() {
    var now = new Date(), tp = new Date(now.getTime() + (now.getTimezoneOffset() + 480) * 6e4);
    var wd = tp.getDay(), mins = tp.getHours() * 60 + tp.getMinutes();
    return (wd >= 1 && wd <= 5 && mins >= 525 && mins <= 830) || (wd >= 1 && wd <= 5 && mins >= 900) || (wd >= 2 && wd <= 6 && mins <= 300);
  }
  function staleCheck(t, who) {
    var el = $('#stale');
    if (!el || !t) return;
    var m = Math.round((Date.now() - new Date(t.replace(' ', 'T') + ':00+08:00')) / 6e4);
    el.hidden = !(marketOpen() && m > 45);
    if (!el.hidden) el.textContent = '⚠ 交易時段中，' + who + '已經 ' + ageText(t).replace('前', '') + '沒有更新，價格、停損與訊號可能過時。';
  }
  function badge(text, warn) {
    var el = $('#tm-src');
    if (!el) {
      el = document.createElement('p');
      el.id = 'tm-src';
      el.className = 'asof';
      var at = $('#hdr-asof') || $('h1');
      if (at && at.parentNode) at.parentNode.insertBefore(el, at.nextSibling); else document.body.appendChild(el);
    }
    el.textContent = text;
    el.style.color = warn ? 'var(--warn, #a86500)' : '';
  }

  // ---------- 即時訊號頁：把試算表「即時訊號」分頁的 v3 區塊排成網頁 ----------
  function findRow(rows, re, from) {
    for (var i = from || 0; i < rows.length; i++) if (re.test(cell(rows[i], 0))) return i;
    return -1;
  }
  function v3Card(rows) {
    var a = findRow(rows, /^v3・/), b = findRow(rows, /^(D2・|v3 模擬交易紀錄)/, a + 1);
    if (a < 0) return null;
    var head = rows[a], st = cell(head, 5) || cell(head, 1), sec = rows.slice(a + 1, b < 0 ? rows.length : b);
    var cls = /多/.test(st) && !/空手/.test(st) ? 'long' : /空/.test(st) && !/空手/.test(st) ? 'short' : 'flat';
    var out = [], open = false, list = false;
    var closeList = function () { if (list) { out.push('</ul>'); list = false; } };
    var closeSide = function () { closeList(); if (open) { out.push('</details>'); open = false; } };
    var sub = 0;
    sec.forEach(function (r) {
      if (blank(r)) { closeSide(); return; }
      var c0 = cell(r, 0), c1 = cell(r, 1), rest = r.slice(1).map(String).map(function (s) { return s.trim(); }).filter(Boolean);
      if (/^[✓✗]$/.test(c0)) {
        if (!list) { out.push('<ul class="ck">'); list = true; }
        out.push('<li class="' + (c0 === '✓' ? 'ok' : 'no') + '"><span class="mk">' + c0 + '</span><span>' + esc(c1) + '</span><span class="dt">' + esc(r.slice(2).map(String).filter(function (s) { return s.trim(); }).join(' ')) + '</span></li>');
        return;
      }
      if (/^做[多空]$/.test(c0)) {
        closeSide();
        out.push('<details class="side"><summary>' + esc(c0) + ' <span class="pill' + (/成立/.test(c1) && !/未/.test(c1) ? ' go' : '') + '">' + esc(c1) + '</span></summary>');
        open = true;
        return;
      }
      closeList();
      if (/^(突破|拉回)進場$/.test(c0) || /^進場條件$/.test(c0)) { out.push('<p class="gt">' + esc(c0) + '</p>'); return; }
      if (/^進場條件/.test(c0)) { closeSide(); out.push('<h3>' + esc(c0) + '</h3>'); return; }
      if (/^EMA20/.test(c0)) { out.push('<p class="lv">' + esc(c0) + '</p>'); return; }
      if (/^[●⚠]/.test(c0)) { out.push('<p class="alert' + (/^●/.test(c0) ? ' sig' : '') + '">' + esc([c0].concat(rest).join('・')) + '</p>'); return; }
      if (open) { out.push('<p class="note">' + esc([c0].concat(rest).join('・')) + '</p>'); return; }
      if (rest.length && c0.length <= 8) {   // 持倉時的「進場／目前停損／未實現」之類：標題＋數值
        out.push('<p class="sub"><b style="color:var(--fg)">' + esc(c0) + '</b>　' + signed(rest.join('・')) + '</p>');
        return;
      }
      sub++;
      out.push('<p class="sub">' + esc([c0].concat(rest).join('・')) + '</p>');
    });
    closeSide();
    var sub0 = out.length && /^<p class="sub">/.test(out[0]) ? out.shift() : '';
    return '<section class="card"><div class="ch"><div><h2>' + esc(cell(head, 0).replace(/^v3・60 分K$/, 'v3・60 分K（限時段）')) + '</h2>' + sub0 +
      '</div><span class="state ' + cls + '">' + esc(st) + '</span></div>' + out.join('') + '</section>';
  }
  function v3Record(rows) {
    var a = findRow(rows, /^v3 模擬交易紀錄/);
    if (a < 0) return null;
    var out = [], tb = [], th = null;
    for (var i = a + 1; i < rows.length; i++) {
      var r = rows[i];
      if (blank(r)) break;
      if (cell(r, 0) === '進場' && cell(r, 1) === '方向') { th = r; continue; }
      if (th && r.length >= 4) {
        var rm = cell(r, 5).split('・');
        tb.push('<tr><td>' + esc(cell(r, 0)) + '</td><td>' + esc(cell(r, 1)) + '</td><td>' + esc(cell(r, 2)) + '</td><td>' + esc(cell(r, 3)) +
          '</td><td>' + esc(cell(r, 4)) + '</td><td>' + esc((rm[0] || '').replace(/R$/, '')) + '</td><td>' + signed((rm[1] || '').replace(/ ?元$/, '')) + '</td></tr>');
        continue;
      }
      if (tb.length) { out.push('§TABLE§'); th = null; }
      out.push('<p class="sub">' + signed(cell(r, 0)).replace(/每筆明細在「成交紀錄」分頁/, '<a href="trades.html">完整歷史成交紀錄 →</a>') + '</p>');
    }
    var table = tb.length ? '<div class="scroll"><table><thead><tr><th>進場</th><th>方向</th><th>進場價</th><th>出場</th><th>出場價</th><th>R</th><th>元</th></tr></thead><tbody>' + tb.join('') + '</tbody></table></div>'
      : '<p class="note">還沒有交易。</p>';
    var html = out.join('');
    html = html.indexOf('§TABLE§') >= 0 ? html.replace('§TABLE§', table) : html.replace(/(<\/p>)/, '$1' + table);
    return '<section class="card"><h2>v3（限時段）模擬交易紀錄</h2>' + html + '</section>';
  }
  function renderLive(d) {
    var rows = d.live || [], card = v3Card(rows), rec = v3Record(rows);
    if (!card) throw new Error('即時訊號分頁找不到 v3 區塊');
    var price = cell(rows[0], 5);
    if (price && $('#price')) $('#price').textContent = price;
    if ($('#price-src') && cell(rows[2], 0)) $('#price-src').textContent = cell(rows[2], 0);
    if ($('#hdr-asof')) $('#hdr-asof').innerHTML = esc(cell(rows[1], 0).split('・')[0]) + '・Google 試算表算於台北 ' + esc(d.updated) + '（' + ageText(d.updated) + '）';
    $('#v3card').innerHTML = card;
    if (rec) $('#v3rec').innerHTML = rec;
    staleCheck(d.updated, '試算表');
  }

  // ---------- 成交紀錄頁：試算表「成交紀錄」分頁的台指期 v3 區塊 ----------
  function renderTrades(d) {
    var rows = d.trades || [], a = findRow(rows, /^台指期・/), s = findRow(rows, /^v3・/, a + 1);
    if (s < 0) throw new Error('成交紀錄分頁找不到 v3 區塊');
    var note = cell(rows[s + 1], 0), tiles = [], body = [], i = s + 2;
    var h = findRow(rows, /^階段$/, s);
    for (var k = s + 2; k < h; k += 2) {
      var t = rows[k], v = rows[k + 1] || [];
      for (var j = 0; j < t.length; j++) if (cell(t, j))
        tiles.push('<div><span>' + esc(cell(t, j)) + '</span><b>' + esc(cell(v, j)) + '</b><em>' + esc(cell(rows[k + 2], j)) + '</em></div>');
      break;
    }
    for (i = h + 1; i < rows.length && !blank(rows[i]); i++) {
      var r = rows[i], fw = cell(r, 0) === '前測', side = /^多/.test(cell(r, 2)) ? 'long' : 'short', open = !cell(r, 6);
      body.push('<tr class="' + side + (open ? ' open' : '') + '"><td><span class="ph' + (fw ? ' fw' : '') + '">' + esc(cell(r, 0)) + '</span></td>' +
        r.slice(1, 10).map(function (c) { return '<td>' + esc(c) + '</td>'; }).join('') + '<td>' + signed(cell(r, 10)) + '</td><td>' + esc(cell(r, 11)) + '</td></tr>');
    }
    var heads = (rows[h] || []).map(function (c) { return '<th>' + esc(c) + '</th>'; }).join('');
    $('#tx-v3').innerHTML = '<section class="card"><h2>v3・60 分K（限時段）</h2><p class="note">' + esc(note) + '</p><div class="tiles">' + tiles.join('') +
      '</div><div class="scroll"><table><thead><tr>' + heads + '</tr></thead><tbody>' + body.join('') + '</tbody></table></div></section>';
    if ($('#hdr-asof')) $('#hdr-asof').textContent = 'v3 60 分K（限時段）每一筆模擬交易・Google 試算表算於台北 ' + d.updated + '（' + ageText(d.updated) + '）' +
      (cell(rows[1], 0).match(/台指期持倉以現價 [\d,]+ 計/) || [''])[0].replace(/^/, '・').replace(/^・$/, '');
  }

  // ---------- 週選頁：用試算表「資料_週選」的紀錄換掉網頁裡的 ----------
  function renderOptions(d) {
    if (typeof D === 'undefined' || typeof draw !== 'function') throw new Error('週選頁還沒載入');
    var old = {};
    D.trades.forEach(function (x) { old[x.id + '|' + (x.target || D.rules.target)] = x; });
    var same = function (a, b) { return a && b && a.code === b.code && a.strike === b.strike && a.cp === b.cp; };
    var merged = (d.options || []).map(function (x) {
      var o = old[x.id + '|' + (x.target || D.rules.target)];
      if (o && o.final && same(o, x)) return o;            // 網頁裡已經用逐筆成交校正完的，照用
      if (o && same(o, x) && o.env && !x.env) x.env = o.env; // 試算表沒有價格軌跡，借舊的（之後的成交看 fills）
      return x;
    });
    // 網頁裡有、試算表沒有的（例如還沒匯入），保留
    var seen = {};
    merged.forEach(function (x) { seen[x.id + '|' + (x.target || D.rules.target)] = 1; });
    D.trades.forEach(function (x) { if (!seen[x.id + '|' + (x.target || D.rules.target)]) merged.push(x); });
    merged.sort(function (a, b) { return a.id < b.id ? -1 : a.id > b.id ? 1 : (a.target || 0) - (b.target || 0); });
    if (!merged.length) throw new Error('試算表沒有週選紀錄');
    D.trades = merged;
    D.generated = d.updated;
    var st = cell(((d.live || [])[findRow(d.live || [], /^v3・/)] || []), 5);
    D.v3 = {pos: /持?多/.test(st) && !/空手/.test(st) ? '多' : /空/.test(st) && !/空手/.test(st) ? '空' : ''};
    if ($('#hdr-asof')) $('#hdr-asof').textContent = 'v3 60 分K（限時段）出訊號時買同方向週選・Google 試算表算於台北 ' + d.updated + '（' + ageText(d.updated) + '）・試算表用期交所即時 1 分K 追蹤';
    draw();
  }

  var RENDER = {live: renderLive, trades: renderTrades, options: renderOptions};
  function run() {
    if (!RENDER[page]) return;
    fetchSheet().then(function (d) {
      if (!d.updated || (pageGen && d.updated < pageGen)) {
        badge('資料：網頁內建（台北 ' + pageGen + '，比試算表新）');
        if (page === 'live') staleCheck(pageGen, '這頁');
        return;
      }
      RENDER[page](d);
      badge('資料來源：Google 試算表（試算表自己定時跟期交所抓資料計算，這頁每 ' + REFRESH_MIN + ' 分鐘重抓一次）');
    }).catch(function (e) {
      badge('⚠ 試算表資料抓不到（' + e.message + '），顯示的是網頁內建的資料：台北 ' + (pageGen || '時間不明'), true);
      if (page === 'live') staleCheck(pageGen, '這頁');
    });
  }
  run();
  setInterval(run, REFRESH_MIN * 6e4);
})();
