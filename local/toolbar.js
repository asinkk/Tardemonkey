// 交易猩猩本機版工具列：local/server.py 會把這支插到每一頁最後面。
(function () {
  'use strict';
  var css = '#tm-local{position:fixed;left:0;right:0;bottom:0;z-index:9999;background:#1b1f24;color:#e8eaed;font:14px/1.4 system-ui,"Microsoft JhengHei",sans-serif;' +
    'padding:8px 12px;display:flex;flex-wrap:wrap;gap:6px 14px;align-items:center;box-shadow:0 -2px 8px rgba(0,0,0,.25)}' +
    '#tm-local a{color:#9ecbff;text-decoration:none}#tm-local a.on{color:#fff;font-weight:600;text-decoration:underline}' +
    '#tm-local button{font:inherit;border:0;border-radius:6px;padding:5px 12px;cursor:pointer;background:#3a4350;color:#fff}' +
    '#tm-local button.save{background:#2f7d5b}#tm-local button:disabled{opacity:.5;cursor:default}' +
    '#tm-local .msg{color:#c9d1d9;font-size:13px}#tm-local .warn{color:#ffb86b}body{padding-bottom:110px}';
  var st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);
  var bar = document.createElement('div'); bar.id = 'tm-local';
  var here = location.pathname.replace(/^\//, '') || 'live.html';
  var links = [['live.html', '即時訊號'], ['trades.html', '成交紀錄'], ['options.html', '週選'], ['strategy.html', '前測']]
    .map(function (l) { return '<a href="/' + l[0] + '"' + (here === l[0] ? ' class="on"' : '') + '>' + l[1] + '</a>'; }).join('　');
  bar.innerHTML = '<b>🦍 本機版</b><span>' + links + '</span>' +
    '<button data-job="live">立即更新</button><button data-job="daily">收今天資料</button><button data-job="save" class="save">存到 GitHub</button>' +
    '<span class="msg" id="tm-msg">讀取狀態中…</span>';
  document.body.appendChild(bar);
  var msg = bar.querySelector('#tm-msg'), btns = bar.querySelectorAll('button');
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]; }); };
  var lastGen = null, wasBusy = false, lastDirty = null;

  function refresh() {
    fetch('/__local/status', {cache: 'no-store'}).then(function (r) { return r.json(); }).then(function (s) {
      var busy = !!s.job;
      btns.forEach(function (b) { b.disabled = busy; });
      var parts = [];
      if (busy) parts.push('⏳ 正在' + esc(s.job) + '（' + esc(s.job_started) + ' 開始）…');
      parts.push('資料算於 ' + esc(s.generated || '—'));
      parts.push(s.in_session ? '交易時段中，每 ' + s.live_min + ' 分鐘自動更新' + (s.next_live ? '（下次約 ' + esc(s.next_live) + '）' : '') : '非交易時段，不自動更新');
      parts.push('今天收資料：' + (s.daily_done === new Date().toLocaleDateString('sv') ? '已完成' : '還沒（週一到週六 06:32）'));
      lastDirty = s.dirty;
      parts.push(s.dirty ? '<span class="warn">' + s.dirty + ' 個檔案還沒存到 GitHub</span>' : (s.dirty === 0 ? 'GitHub 已是最新' : ''));
      var last = Object.keys(s.last || {}).map(function (k) { return s.last[k]; }).sort(function (a, b) { return a.at < b.at ? 1 : -1; })[0];
      if (last && !busy) parts.push((last.ok ? '✓ ' : '<span class="warn">✗ ') + esc(last.msg) + '（' + esc(last.at.slice(11)) + '）' + (last.ok ? '' : '</span>'));
      msg.innerHTML = parts.filter(Boolean).join('・');
      // 算完新的一版就重新載入頁面
      if (lastGen !== null && s.generated && s.generated !== lastGen && !busy) location.reload();
      if (wasBusy && !busy && lastGen === s.generated && /strategy|options|trades/.test(here)) location.reload();
      if (lastGen === null) lastGen = s.generated;
      wasBusy = busy;
    }).catch(function () {
      msg.innerHTML = '<span class="warn">⚠ 連不到本機伺服器：「啟動交易猩猩」的黑色視窗是不是關掉了？</span>';
      btns.forEach(function (b) { b.disabled = true; });
    });
  }
  btns.forEach(function (b) {
    b.addEventListener('click', function () {
      var job = b.getAttribute('data-job');
      if (job === 'save' && !confirm('要把目前收集的資料（' + (lastDirty == null ? '?' : lastDirty) + ' 個檔案）commit 並推上 GitHub 嗎？\n\n建議一天最多存一兩次，不要太頻繁。')) return;
      fetch('/__local/run?job=' + job, {method: 'POST', headers: {'X-Tardemonkey': 'local'}})
        .then(function (r) { return r.json(); })
        .then(function (d) { if (!d.started) alert('現在正在' + (d.busy || '跑別的工作') + '，等它跑完再按。'); refresh(); });
    });
  });
  refresh();
  setInterval(refresh, 10000);
})();
