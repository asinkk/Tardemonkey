"""週選即時訊號頁 docs/options.html。資料由 option_signal.update() 算好，這裡只負責排版。
口數試算、各批口數分配、損益都在瀏覽器裡用 JavaScript 算（改口數不用等下一次更新）。由 live.py 呼叫。
"""
import json

CSS = """
:root { --bg:#f6f7f9; --panel:#fff; --fg:#1b1f24; --muted:#5d6673; --line:#dfe3e8; --up:#c2412d; --dn:#2f7d5b;
  --accent:#2f5fd0; --warn:#a86500; --warnbg:#fff4e0; --sigbg:#e8efff; --rowl:rgba(194,65,45,.05); --rows:rgba(47,125,91,.06);
  --font:"Noto Sans TC","PingFang TC","Microsoft JhengHei",system-ui,sans-serif; }
@media (prefers-color-scheme: dark) { :root { --bg:#14171b; --panel:#1c2026; --fg:#e8ebef; --muted:#9aa3ae; --line:#2c323a;
  --up:#ef7a63; --dn:#5cc296; --accent:#7ea2ff; --warn:#f0b35a; --warnbg:#3a2c14; --sigbg:#1e2a48;
  --rowl:rgba(239,122,99,.07); --rows:rgba(92,194,150,.07); color-scheme:dark; } }
* { box-sizing:border-box; }
body { margin:0; background:var(--bg); color:var(--fg); font-family:var(--font); font-size:15px; line-height:1.5; }
main { max-width:1100px; margin:0 auto; padding:16px; }
h1 { font-size:20px; margin:0; } h2 { font-size:17px; margin:0 0 2px; } h3 { font-size:13px; color:var(--muted); font-weight:500; margin:14px 0 4px; }
.note, .asof { color:var(--muted); font-size:12.5px; margin:4px 0; }
.top { display:flex; justify-content:space-between; align-items:flex-end; gap:10px; flex-wrap:wrap; }
nav { display:flex; gap:6px; flex-wrap:wrap; margin:10px 0 0; }
nav a { font-size:13px; border:1px solid var(--line); border-radius:999px; padding:3px 10px; text-decoration:none; background:var(--panel); }
.lots { display:flex; align-items:center; gap:8px; background:var(--panel); border:1px solid var(--line); border-radius:10px; padding:8px 12px; }
.lots label { font-size:13px; color:var(--muted); }
.lots input { width:72px; font:inherit; font-size:18px; font-weight:600; padding:4px 6px; border:1px solid var(--line); border-radius:6px;
  background:var(--bg); color:var(--fg); text-align:right; }
.pick .row { display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
.pick label { font-size:13px; color:var(--muted); }
.pick select, .pick input[type=number] { font:inherit; padding:4px 6px; border:1px solid var(--line); border-radius:6px; background:var(--bg); color:var(--fg); }
.pick select { font-size:15px; font-weight:600; max-width:100%; }
.pick input[type=number] { width:64px; text-align:right; }
.pick .grid { display:grid; grid-template-columns:auto auto auto; gap:6px 12px; justify-content:start; align-items:center; margin:8px 0; font-size:13px; }
.pick .trail { font-size:13px; margin:6px 0; display:flex; gap:6px; flex-wrap:wrap; align-items:center; }
.pick input:disabled { opacity:.45; }
.card { background:var(--panel); border:1px solid var(--line); border-radius:10px; padding:14px; margin-top:12px; }
.banner { font-size:17px; font-weight:600; }
.banner.buy { background:var(--sigbg); border-color:var(--accent); }
.banner small { display:block; font-size:13px; font-weight:400; color:var(--muted); margin-top:2px; }
.alert { background:var(--warnbg); color:var(--warn); border-radius:8px; padding:6px 10px; font-size:13px; margin:8px 0 0; }
.ch { display:flex; justify-content:space-between; align-items:flex-start; gap:8px; flex-wrap:wrap; }
.state { font-weight:600; padding:3px 10px; border-radius:999px; white-space:nowrap; font-size:13px; color:#fff; }
.state.long { background:var(--up); } .state.short { background:var(--dn); }
.contract { font-size:22px; font-weight:600; margin:6px 0 0; font-variant-numeric:tabular-nums; }
.tiles { display:grid; grid-template-columns:repeat(auto-fit,minmax(130px,1fr)); gap:8px; margin:10px 0; }
.tiles div { border:1px solid var(--line); border-radius:8px; padding:8px; }
.tiles span { display:block; font-size:12px; color:var(--muted); }
.tiles b { display:block; font-size:18px; font-weight:600; font-variant-numeric:tabular-nums; }
.tiles em { display:block; font-style:normal; font-size:12px; color:var(--muted); }
.scroll { overflow-x:auto; border:1px solid var(--line); border-radius:8px; }
table { border-collapse:collapse; width:100%; font-variant-numeric:tabular-nums; font-size:13px; }
th, td { padding:6px 8px; text-align:right; border-bottom:1px solid var(--line); white-space:nowrap; }
tr:last-child td { border-bottom:0; }
th { color:var(--muted); font-weight:500; background:var(--panel); }
td:first-child, th:first-child { text-align:left; }
table.hist td:nth-child(-n+4), table.hist th:nth-child(-n+4) { text-align:left; }
tr.long td { background:var(--rowl); } tr.short td { background:var(--rows); }
tr.open td { font-weight:600; }
.ok { color:var(--accent); font-weight:600; } .wait { color:var(--muted); }
.ph { font-size:11.5px; color:var(--muted); border:1px solid var(--line); border-radius:999px; padding:0 6px; }
.ph.fw { color:#fff; background:var(--accent); border-color:var(--accent); }
.up { color:var(--up); } .dn { color:var(--dn); }
a { color:var(--accent); }
ul.rules { margin:6px 0 0; padding-left:20px; font-size:13.5px; }
footer { margin:20px 0 40px; }
"""

JS = r"""
const D = JSON.parse(document.getElementById('data').textContent), R = D.rules;
const $ = (s) => document.querySelector(s);
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const num = (v, d = 1) => Number(v).toLocaleString('zh-Hant', {minimumFractionDigits: 0, maximumFractionDigits: d});
const sgn = (v, d = 0) => (v > 0 ? '+' : '') + num(v, d);
const money = (v, d = 0) => `<span class="${v > 0 ? 'up' : v < 0 ? 'dn' : ''}">${sgn(v, d)}</span>`;
const time = (s) => new Date(s.replace(' ', 'T'));
const md = (s) => s ? s.slice(5, 16) : '';

const tickSize = (p) => p < 10 ? .1 : p < 50 ? .5 : p < 500 ? 1 : p < 1000 ? 5 : 10;
const roundTick = (p) => { const t = tickSize(p); return Math.round(Math.round(p / t) * t * 10) / 10; };

// 停利方案：legs 依「零頭優先順序」排列；w 是口數比例。tp＝限價停利倍數；trail＝漲過 arm 倍後從最高價回落 back 出場。
const TR = {trail: 1, arm: R.arm, back: R.back};
const PRESETS = {
  balanced: {name: '均衡（預設）', one: 1, legs: [{m: 3, w: .3}, {m: 5, w: .3}, {...TR, w: .2}, {m: 10, w: .2}]},
  safe: {name: '保守', legs: [{m: 2, w: .3}, {m: 3, w: .3}, {m: 5, w: .4}]},
  bold: {name: '積極', one: 1, legs: [{m: 2, w: .2}, {m: 5, w: .3}, {m: 10, w: .3}, {...TR, w: .2}]},
  all3: {name: '全部 3 倍', legs: [{m: 3, w: 1}]},
  all5: {name: '全部 5 倍', legs: [{m: 5, w: 1}]},
  all10: {name: '全部 10 倍', legs: [{m: 10, w: 1}]},
};
const clamp = (v, lo, hi, d) => { v = Number(v); return v >= lo && v <= hi ? v : d; };
function customScheme() {
  const legs = [];
  for (let i = 1; i <= 4; i++) {
    const w = clamp($('#cw' + i).value, 0, 100, 0) / 100;
    if (!(w > 0)) continue;
    if (i === 4 && $('#ctrail').checked)
      legs.push({trail: 1, arm: clamp($('#carm').value, R.env_from, 30, R.arm), back: clamp($('#cback').value, 5, 90, R.back * 100) / 100, w});
    else legs.push({m: clamp($('#cm' + i).value, R.env_from, 30, 3), w});
  }
  return {name: '自訂', legs: legs.length ? legs : [{m: 5, w: 1}]};
}
const scheme = (key) => key === 'custom' ? customScheme() : PRESETS[key] || PRESETS.balanced;

// 口數照比例分，零頭依 legs 的順序補；只有 1 口時用方案指定的那一批（均衡、積極是 5 倍）
function alloc(S, n) {
  const sum = S.legs.reduce((a, l) => a + l.w, 0), a = S.legs.map(() => 0), rest = [];
  if (n === 1 && S.one != null) { a[S.one] = 1; return a; }
  let left = n;
  S.legs.forEach((l, i) => { const q = n * l.w / sum; a[i] = Math.floor(q + 1e-9); left -= a[i]; rest.push([q - a[i], i]); });
  rest.sort((x, y) => y[0] - x[0] || x[1] - y[1]);
  for (let i = 0; i < left; i++) a[rest[i][1]]++;
  return a;
}

// 從價格軌跡 env 算：限價停利何時成交、回吐出場何時觸發
function tpFill(x, m) {
  const tg = roundTick(x.entry * m), e = (x.env || []).find((v) => v[1] >= tg);
  return {target: tg, when: e ? e[0] : null};
}
function trailExit(x, arm, back) {
  const armPx = roundTick(x.entry * arm);
  for (const [, pk, lows] of x.env || []) {
    if (pk < armPx) continue;
    const hit = lows.find((v) => v[1] <= pk * (1 - back));
    if (hit) return {armPx, t: hit[0], px: hit[1]};
  }
  return {armPx, t: null};
}

// 一筆交易在方案 S、n 口下的各批結果
function legs(x, n, S) {
  const a = alloc(S, n), pct = (b) => Math.round(b * 100);
  let out = S.legs.map((l, i) => {
    if (!l.trail) {
      const f = tpFill(x, l.m);
      return {name: `${l.m} 倍停利`, key: l.m, lots: a[i], done: !!f.when, when: f.when, exit: f.when ? f.target : x.last, how: `限價賣出 ${num(f.target)}`};
    }
    const t = trailExit(x, l.arm, l.back), armed = x.peak >= t.armPx, stop = Math.round(x.peak * (1 - l.back) * 10) / 10;
    return {name: `漲過 ${l.arm} 倍後回落 ${pct(l.back)}%`, key: 1e9, lots: a[i], done: !!t.t, when: t.t, exit: t.t ? t.px : x.last,
            how: t.t ? `跌破回落價，以 ${num(t.px)} 賣出` : armed ? `跌到 ${num(stop)} 以下就賣（最高 ${num(x.peak)}）`
              : `先等漲過 ${num(t.armPx)}，之後從最高價回落 ${pct(l.back)}% 賣出`};
  });
  out.sort((p, q) => p.key - q.key);
  let real = 0, unreal = 0;
  for (const l of out) {
    l.per = l.exit - x.entry - R.cost;
    l.sub = l.per * l.lots;
    if (l.done || x.closed) real += l.sub; else unreal += l.sub;
  }
  return {legs: out, real, unreal, total: real + unreal, cost: x.entry * n};
}
const isFresh = (x) => (time(D.generated) - time(x.id)) / 6e4 <= 60 && x.peak < x.entry * R.env_from;

function cname(x) {
  const mo = Number(x.code.slice(4, 6)), k = x.code.slice(6);
  const kind = !k ? '月選' : k[0] === 'W' ? `${k}（週三）` : `${k}（週五）`;
  return `${mo} 月 ${kind} ${num(x.strike, 0)} ${x.cp === 'C' ? '買權' : '賣權'}`;
}

function card(x, n, S) {
  const g = legs(x, n, S), now = new Date(), exp = time(x.exp_t), hrs = (exp - now) / 36e5;
  const side = x.side > 0, fresh = isFresh(x);
  const q = x.quote, alts = (x.alts || []).map((a) => `${num(a.strike, 0)}（${num(a.px)}）`).join('、');
  const rows = g.legs.map((l) => `<tr><td>${l.name}</td><td>${l.lots} 口</td><td>${l.how}</td>
    <td>${l.done ? `<span class="ok">已成交</span> ${md(l.when)}` : x.closed ? '沒成交，放到結算' : '<span class="wait">等待中</span>'}</td>
    <td>${l.lots ? money(l.per, 1) : ''}</td><td>${l.lots ? money(l.sub * R.point) : ''}</td></tr>`).join('');
  return `<section class="card">
  <div class="ch"><div><h2>${fresh ? '買哪一檔' : '持倉追蹤'}</h2>
    <p class="note">v3 ${side ? '多' : '空'}單・${esc(x.kind)}・訊號 ${esc(x.id)}・${esc(x.phase)}</p></div>
    <span class="state ${side ? 'long' : 'short'}">${side ? '買買權 CALL' : '買賣權 PUT'}</span></div>
  <p class="contract">${cname(x)}</p>
  <p class="note">到期 ${esc(x.exp_t.slice(0, 16))}${hrs > 0 ? `（還有 ${hrs >= 48 ? num(hrs / 24, 1) + ' 天' : num(hrs, 1) + ' 小時'}）` : '（已到期）'}
    ${alts ? `・相鄰履約價：${alts}` : ''}</p>
  ${x.warn ? `<p class="alert">⚠ ${esc(x.warn)}</p>` : ''}
  <div class="tiles">
    <div><span>買進參考價</span><b>${num(x.entry)}</b><em>${esc(x.src)} ${md(x.entry_at)}${x.lag > 10 ? `・晚 ${x.lag} 分` : ''}</em></div>
    <div><span>${n} 口成本</span><b>${num(g.cost * R.point, 0)}</b><em>元（${num(g.cost)} 點）</em></div>
    <div><span>目前價</span><b>${num(x.last)}</b><em>${q && q.bid && q.ask ? `買 ${num(q.bid)}／賣 ${num(q.ask)}` : md(x.last_t)}</em></div>
    <div><span>最高到過</span><b>${num(x.peak / x.entry, 1)} 倍</b><em>${num(x.peak)}・${md(x.peak_t)}</em></div>
    <div><span>已實現</span><b>${money(g.real * R.point)}</b><em>元</em></div>
    <div><span>未實現</span><b>${money(g.unreal * R.point)}</b><em>元（以目前價計）</em></div>
  </div>
  <h3>${n} 口停利掛單表・${esc(S.name)}方案（不設停損，最多賠掉權利金）</h3>
  <div class="scroll"><table><thead><tr><th>批次</th><th>口數</th><th>怎麼掛</th><th>狀態</th><th>每口損益（點）</th><th>小計（元）</th></tr></thead>
  <tbody>${rows}</tbody></table></div>
  <p class="note">掛單價是用上面的買進參考價算的。你的實際成交價不同時，請用自己的成交價乘以表上的倍數來掛。
  ${n > 10 ? '口數較多：夜盤這個價位的履約價一分鐘常只成交幾口，建議分幾分鐘買，或分散到相鄰履約價。' : ''}</p>
</section>`;
}

function history(n, S) {
  const T = D.trades;
  if (!T.length) return '<p class="note">還沒有交易。</p>';
  let tot = 0, cost = 0, win = 0, closed = 0, fwd = 0;
  const body = [...T].reverse().map((x) => {
    const g = legs(x, n, S);
    tot += g.total; cost += g.cost;
    if (x.closed) { closed++; if (g.total > 0) win++; }
    if (x.phase === '前測') fwd += g.total;
    const cells = g.legs.map((l) => `<td>${!l.lots ? '<span class="wait">0 口</span>' : l.done ? `<span class="ok">✓</span> ${num(l.exit)}・${md(l.when)}` : '<span class="wait">—</span>'}</td>`).join('');
    return `<tr class="${x.side > 0 ? 'long' : 'short'}${x.closed ? '' : ' open'}">
      <td><span class="ph${x.phase === '前測' ? ' fw' : ''}">${esc(x.phase)}</span></td><td>${esc(x.id)}</td><td>${x.side > 0 ? '多' : '空'}・${esc(x.kind)}</td>
      <td>${cname(x)}</td><td>${num(x.entry)}</td><td>${num(x.peak)}（${num(x.peak / x.entry, 1)} 倍）</td>${cells}
      <td>${x.closed ? num(x.last) : '持倉中 ' + num(x.last)}</td><td>${money(g.total, 1)}</td><td>${money(g.total * R.point)}</td></tr>`;
  }).join('');
  const heads = legs(T[0], n, S).legs.map((l) => `<th>${l.name}</th>`).join('');
  return `<div class="tiles">
    <div><span>合計損益（${n} 口・${esc(S.name)}）</span><b>${money(tot * R.point)}</b><em>元（${sgn(tot, 1)} 點）</em></div>
    <div><span>投入權利金合計</span><b>${num(cost * R.point, 0)}</b><em>元・報酬 ${cost ? sgn(tot / cost * 100) : 0}%</em></div>
    <div><span>已結束</span><b>${closed} 筆</b><em>賺 ${win} 筆</em></div>
    <div><span>其中前測</span><b>${money(fwd * R.point)}</b><em>元（${esc(D.forward_start)} 起的訊號）</em></div></div>
  <div class="scroll"><table class="hist"><thead><tr><th>階段</th><th>訊號時間</th><th>方向</th><th>合約</th><th>買進價</th><th>最高價</th>
    ${heads}<th>結算／現價</th><th>損益（點）</th><th>損益（元）</th></tr></thead>
  <tbody>${body}</tbody></table></div>`;
}
const total = (n, S) => D.trades.reduce((a, x) => a + legs(x, n, S).total, 0);

function allocText(n, S) {
  const a = alloc(S, n);
  return `${n} 口的分法：` + S.legs.map((l, i) => [l.trail ? 1e9 : l.m, `${l.trail ? '回吐出場' : l.m + ' 倍'} ${a[i]} 口`])
    .sort((p, q) => p[0] - q[0]).map((v) => v[1]).join('、');
}

const store = (k, v) => { try { localStorage.setItem(k, v); } catch (e) {} };
const load = (k) => { try { return localStorage.getItem(k); } catch (e) { return null; } };
const CUSTOM_IDS = ['cm1', 'cw1', 'cm2', 'cw2', 'cm3', 'cw3', 'cm4', 'cw4', 'carm', 'cback'];

function draw() {
  let n = Math.round(Number($('#lots').value));
  if (!(n >= 1)) n = 1;
  if (n > 99) n = 99;
  const key = $('#scheme').value, S = scheme(key);
  store('lots', n); store('scheme', key);
  store('custom', JSON.stringify(Object.fromEntries(CUSTOM_IDS.map((i) => [i, $('#' + i).value]).concat([['ctrail', $('#ctrail').checked]]))));
  $('#custom').hidden = key !== 'custom';
  for (const i of ['cm4']) $('#' + i).disabled = $('#ctrail').checked;
  for (const i of ['carm', 'cback']) $('#' + i).disabled = !$('#ctrail').checked;
  for (const o of $('#scheme').options) {
    const s = scheme(o.value);
    o.textContent = `${s.name}　${sgn(total(n, s) * R.point)} 元`;
  }
  $('#schemenote').textContent = allocText(n, S) + '。選單上的金額是下面歷史紀錄在這個口數下的合計損益。';
  const open = D.trades.filter((x) => !x.closed).reverse();
  const fresh = open.find(isFresh);
  const v = D.v3;
  let b;
  if (fresh) b = `<div class="card banner buy">新訊號：買進 ${n} 口 ${cname(fresh)}，約 ${num(fresh.entry)} 點
      <small>訊號 ${esc(fresh.id)}・買進後照下面的表掛停利單・${allocText(n, S)}</small></div>`;
  else if (open.length) b = `<div class="card banner">持倉中：${open.map(cname).join('、')}<small>照下面的表等停利成交，不設停損・${allocText(n, S)}</small></div>`;
  else b = `<div class="card banner">等待訊號<small>${v.pos ? `v3 目前持有${esc(v.pos)}單，但對應的週選已到期或已結束；` : ''}v3 出現新的進場訊號時，這裡會顯示要買哪一檔・${allocText(n, S)}</small></div>`;
  $('#banner').innerHTML = b + D.notes.map((s) => `<p class="alert">⚠ ${esc(s)}</p>`).join('');
  $('#open').innerHTML = open.map((x) => card(x, n, S)).join('');
  $('#hist').innerHTML = history(n, S);
}
$('#scheme').innerHTML = Object.keys(PRESETS).concat(['custom']).map((k) => `<option value="${k}"></option>`).join('');
if (load('lots')) $('#lots').value = load('lots');
try {
  const c = JSON.parse(load('custom') || '{}');
  for (const i of CUSTOM_IDS) if (c[i] != null) $('#' + i).value = c[i];
  if (c.ctrail != null) $('#ctrail').checked = c.ctrail;
} catch (e) {}
if (load('scheme') && (PRESETS[load('scheme')] || load('scheme') === 'custom')) $('#scheme').value = load('scheme');
for (const el of document.querySelectorAll('#lots, #scheme, #custom input')) el.addEventListener('input', draw);
draw();
"""


def render(d, now):
    r = d["rules"]
    data = json.dumps(d, ensure_ascii=False, default=str).replace("</", "<\\/")
    return f"""<!doctype html>
<html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="refresh" content="600">
<title>週選即時訊號</title>
<style>{CSS}</style></head><body><main>
<div class="top"><div><h1>週選即時訊號</h1>
<p class="asof">v3 60 分K（限時段）出訊號時買同方向週選・台北 {now:%Y-%m-%d %H:%M} 更新・逐筆成交資料到 {d["ticks_to"] or "（還沒有）"}</p></div>
<div class="lots"><label for="lots">口數</label><input id="lots" type="number" min="1" max="99" step="1" value="10" inputmode="numeric"></div></div>
<section class="card pick"><div class="row"><label for="scheme">停利方案</label><select id="scheme"></select></div>
<div id="custom" hidden>
<p class="note">最多四批，各填停利倍數（{r["env_from"]:g}–30）和佔口數的百分比；比例不必剛好加到 100，會自動換算。比例填 0 表示不用這一批。</p>
<div class="grid">
<span>第 1 批</span><label>倍數 <input id="cm1" type="number" min="{r["env_from"]:g}" max="30" step="0.5" value="3"></label><label>比例 % <input id="cw1" type="number" min="0" max="100" step="5" value="30"></label>
<span>第 2 批</span><label>倍數 <input id="cm2" type="number" min="{r["env_from"]:g}" max="30" step="0.5" value="5"></label><label>比例 % <input id="cw2" type="number" min="0" max="100" step="5" value="30"></label>
<span>第 3 批</span><label>倍數 <input id="cm3" type="number" min="{r["env_from"]:g}" max="30" step="0.5" value="10"></label><label>比例 % <input id="cw3" type="number" min="0" max="100" step="5" value="20"></label>
<span>第 4 批</span><label>倍數 <input id="cm4" type="number" min="{r["env_from"]:g}" max="30" step="0.5" value="15"></label><label>比例 % <input id="cw4" type="number" min="0" max="100" step="5" value="20"></label>
</div>
<p class="trail"><label><input id="ctrail" type="checkbox" checked> 第 4 批改用回吐出場：</label>
<label>漲過 <input id="carm" type="number" min="{r["env_from"]:g}" max="30" step="0.5" value="{r["arm"]:g}"> 倍後，</label>
<label>從最高價回落 <input id="cback" type="number" min="5" max="90" step="5" value="{r["back"] * 100:g}"> % 賣出</label></p>
</div>
<p class="note" id="schemenote"></p>
<p class="note">方案只影響這一頁的試算和掛單建議，買哪一檔不變。高倍數方案的回測金額只靠少數幾筆大行情：「全部 10 倍」9 筆裡有 4 筆全賠。</p></section>
<nav style="margin-bottom:0"><a href="live.html">← 即時訊號（期貨）</a><a href="trades.html">期貨成交紀錄</a></nav>
<div id="banner"></div>
<div id="open"></div>
<section class="card"><h2>歷史紀錄</h2>
<p class="note">每一筆 v3 進場訊號對應的週選交易，金額照上面填的口數計算。「回測」是規則定案前用逐筆成交回推的，「前測」是 {d["forward_start"]} 之後實際出現的訊號。</p>
<div id="hist"></div></section>
<section class="card"><h2>規則</h2>
<ul class="rules">
<li>v3 出多單訊號買買權、空單訊號買賣權，在訊號 K 收完、下一根 60 分K 開盤時買進。</li>
<li>買到期日最近的台指週選（週三、週五到期，第三個週三是月選）；當天就到期的話改買下一個到期的。</li>
<li>履約價選當時權利金最接近 {r["target"]} 點的那一檔。</li>
<li>分批停利（預設的均衡方案）：30% 口數掛 3 倍、30% 掛 5 倍、20% 掛 10 倍，剩下 20% 等漲過 {r["arm"]} 倍後，從最高價回落 {r["back"]:.0%} 賣出。沒成交的放到結算。上方可以換成其他方案或自訂倍數。</li>
<li>不設停損：回測裡有好幾筆先跌到剩兩到五成才噴出，設停損會被洗掉。最大虧損就是付出的權利金。</li>
</ul>
<p class="note">損益每口來回扣 {r["cost"]:g} 點成本，1 點 {r["point"]} 元。停利是限價單，以掛單價計；回吐出場以觸發當下的成交價計；到期沒賣的以到期前最後成交價計。
當天的盤先用期交所即時 1 分K 判斷（約每半小時更新一次），隔天逐筆成交資料進來後會重算校正，所以前一天的成交狀態偶爾會被修正。
買進參考價是網頁更新當下的即時報價，會比訊號成立晚幾分鐘。</p>
<p class="alert">⚠ 這套規則只用 2026/10/1–11/13 的 9 筆訊號回測過，而且那段是台指大跌、波動極大的行情，買賣權特別有利。倍數是大方向，不代表之後也會這樣。</p></section>
<footer><p class="note"><a href="live.html">即時訊號</a>・<a href="trades.html">歷史成交紀錄</a>・<a href="strategy.html">v3 前測（每日）</a></p></footer>
</main>
<script id="data" type="application/json">{data}</script>
<script>{JS}</script>
</body></html>
"""
