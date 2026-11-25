// 交易猩猩本機版工具列：local/server.py 會把這支插到每一頁最後面。
(function () {
  'use strict';
  var css = '#tm-local{position:fixed;left:0;right:0;bottom:0;z-index:9999;background:#1b1f24;color:#e8eaed;font:14px/1.4 system-ui,"Microsoft JhengHei",sans-serif;' +
    'padding:8px 12px;display:flex;flex-wrap:wrap;gap:6px 14px;align-items:center;box-shadow:0 -2px 8px rgba(0,0,0,.25)}' +
    '#tm-local a{color:#9ecbff;text-decoration:none}#tm-local a.on{color:#fff;font-weight:600;text-decoration:underline}' +
    '#tm-local button{font:inherit;border:0;border-radius:6px;padding:5px 12px;cursor:pointer;background:#3a4350;color:#fff}' +
    '#tm-local button.save{background:#2f7d5b}#tm-local button:disabled{opacity:.5;cursor:default}' +
    '#tm-local .msg{color:#c9d1d9;font-size:13px}#tm-local .warn{color:#ffb86b}body{padding-bottom:110px}' +
    '#tm-set{display:none;flex-basis:100%;background:#262c33;border-radius:8px;padding:10px 12px;gap:8px 18px;flex-wrap:wrap;align-items:center}' +
    '#tm-set.open{display:flex}#tm-set label{display:flex;align-items:center;gap:6px}#tm-set select{font:inherit;padding:3px 6px;border-radius:5px}' +
    '#tm-set input[type=checkbox]{width:18px;height:18px}#tm-set .note{flex-basis:100%;font-size:13px;color:#c9d1d9}' +
    '#tm-set .risk{flex-basis:100%;font-size:13px;color:#ffb86b;display:none}#tm-set .risk.on{display:block}' +
    '#tm-local button.auto-on{background:#7d5b2f}';
  var st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);
  var bar = document.createElement('div'); bar.id = 'tm-local';
  var here = location.pathname.replace(/^\//, '') || 'live.html';
  var links = [['live.html', '即時訊號'], ['trades.html', '成交紀錄'], ['options.html', '週選'], ['strategy.html', '前測']]
    .map(function (l) { return '<a href="/' + l[0] + '"' + (here === l[0] ? ' class="on"' : '') + '>' + l[1] + '</a>'; }).join('　');
  bar.innerHTML = '<b>🦍 本機版</b><span>' + links + '</span>' +
    '<button data-job="live">立即更新</button><button data-job="daily">收今天資料</button><button data-job="save" class="save">存到 GitHub</button>' +
    '<button id="tm-gear" type="button">⚙ 設定</button>' +
    '<span class="msg" id="tm-msg">讀取狀態中…</span>' +
    '<div id="tm-set">' +
    '<label>交易時段每 <select id="tm-live"></select> 分鐘抓一次資料</label>' +
    '<label><input type="checkbox" id="tm-auto"> 自動上傳到 GitHub</label>' +
    '<label>最少隔 <select id="tm-hours"></select> 小時上傳一次</label>' +
    '<button type="button" id="tm-apply" class="save">儲存設定</button><span id="tm-setmsg" class="msg"></span>' +
    '<div class="risk" id="tm-risk">⚠ 自動上傳會在 GitHub 留下很多機器 commit。原本的 txo-dashboard 就是因為自動 commit 太頻繁被 GitHub 停用。' +
    '這裡限制最少隔 2 小時、資料有變才推、一天最多 <span id="tm-max"></span> 次；建議設 6 小時以上。</div>' +
    '<div class="note">「存到 GitHub」手動按一直都可以用。抓資料頻率越短，畫面越即時，但越常連期交所。設定只存在這台電腦（local/state.json）。</div>' +
    '</div>';
  document.body.appendChild(bar);
  var msg = bar.querySelector('#tm-msg'), btns = bar.querySelectorAll('button[data-job]');
  var panel = bar.querySelector('#tm-set'), gear = bar.querySelector('#tm-gear');
  var selLive = bar.querySelector('#tm-live'), chkAuto = bar.querySelector('#tm-auto'), selHours = bar.querySelector('#tm-hours');
  var setMsg = bar.querySelector('#tm-setmsg'), risk = bar.querySelector('#tm-risk'), filled = false;
  function fillSel(sel, list, val) {
    sel.innerHTML = list.map(function (v) { return '<option value="' + v + '"' + (v === val ? ' selected' : '') + '>' + v + '</option>'; }).join('');
  }
  function syncRisk() { risk.className = 'risk' + (chkAuto.checked ? ' on' : ''); selHours.disabled = !chkAuto.checked; }
  function fillSettings(s) {
    fillSel(selLive, s.choices.live_min, s.settings.live_min);
    fillSel(selHours, s.choices.auto_save_hours, s.settings.auto_save_hours);
    chkAuto.checked = !!s.settings.auto_save;
    bar.querySelector('#tm-max').textContent = s.auto_save_max;
    syncRisk();
  }
  gear.addEventListener('click', function () { panel.classList.toggle('open'); });
  chkAuto.addEventListener('change', syncRisk);
  bar.querySelector('#tm-apply').addEventListener('click', function () {
    if (chkAuto.checked && !confirm('確定要打開自動上傳？\n\n資料有變時，最少每 ' + selHours.value + ' 小時會自己 commit 並推上 GitHub 一次。\n太頻繁的自動 commit 可能讓 GitHub 停用倉庫。')) return;
    fetch('/__local/settings', {method: 'POST', headers: {'X-Tardemonkey': 'local', 'Content-Type': 'application/json'},
      body: JSON.stringify({live_min: +selLive.value, auto_save: chkAuto.checked, auto_save_hours: +selHours.value})})
      .then(function (r) { return r.json(); })
      .then(function (d) { setMsg.textContent = (d.ok ? '✓ ' : '✗ ') + d.msg; filled = false; refresh(); })
      .catch(function () { setMsg.textContent = '✗ 儲存失敗，連不到本機伺服器'; });
  });
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]; }); };
  var lastGen = null, wasBusy = false, lastDirty = null;

  function refresh() {
    fetch('/__local/status', {cache: 'no-store'}).then(function (r) { return r.json(); }).then(function (s) {
      var busy = !!s.job;
      btns.forEach(function (b) { b.disabled = busy; });
      if (s.settings && (!filled || !panel.classList.contains('open'))) { fillSettings(s); filled = true; }
      gear.className = s.settings && s.settings.auto_save ? 'auto-on' : '';
      gear.textContent = s.settings && s.settings.auto_save ? '⚙ 設定（自動上傳：開）' : '⚙ 設定';
      var parts = [];
      if (busy) parts.push('⏳ 正在' + esc(s.job) + '（' + esc(s.job_started) + ' 開始）…');
      parts.push('資料算於 ' + esc(s.generated || '—'));
      parts.push(s.in_session ? '交易時段中，每 ' + s.live_min + ' 分鐘自動更新' + (s.next_live ? '（下次約 ' + esc(s.next_live) + '）' : '') : '非交易時段，不自動更新');
      parts.push('今天收資料：' + (s.daily_done === new Date().toLocaleDateString('sv') ? '已完成' : '還沒（週一到週六 06:32）'));
      lastDirty = s.dirty;
      if (s.settings && s.settings.auto_save) parts.push('自動上傳：開，每 ' + s.settings.auto_save_hours + ' 小時、有變才推（今天 ' + s.auto_saves_today + '/' + s.auto_save_max + ' 次' + (s.next_auto_save ? '，下次最早 ' + esc(s.next_auto_save) : '') + '）');
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
