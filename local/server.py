"""交易猩猩本機版：在自己的電腦抓期交所資料、算 v3 訊號，用瀏覽器看（不用 GitHub）。

雙擊 repo 根目錄的「啟動交易猩猩.bat」就會：
  1. 在 http://127.0.0.1:8765 開一個只有這台電腦看得到的網頁伺服器，顯示 docs/ 裡的網頁
     （即時訊號、成交紀錄、週選、前測），每頁下方多一條本機工具列。
  2. 交易時段（日盤 08:45–13:45、夜盤 15:00–05:00）每 N 分鐘自己跑一次 strategy/live.py
     （N 在工具列「⚙ 設定」調，預設 10）。
  3. 週一到週六 06:32 收資料：期交所 1 分K、台指選擇權逐筆（期交所只留 30 個交易日）、v3 前測。
     06:32 電腦沒開的話，之後一打開就補收。
  4. 預設不會自己 commit / push。要存到 GitHub 時按工具列的「存到 GitHub」，或雙擊「存到GitHub.bat」。
     「⚙ 設定」可以打開自動上傳：最少隔 2 小時、資料有變才推、一天最多 AUTO_SAVE_MAX_PER_DAY 次。
  設定存在 local/state.json（不會推上 GitHub）。

只用 Python 標準函式庫。指令：
  python local/server.py            啟動（預設會開瀏覽器）
  python local/server.py --save     只把目前資料存到 GitHub，然後結束
  python local/server.py --no-browser --port 8765
"""
import argparse, datetime as dt, http.server, json, os, subprocess, sys, threading, time, urllib.parse, webbrowser

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
DOCS = os.path.join(REPO, "docs")
LOGS = os.path.join(HERE, "logs")
STATE = os.path.join(HERE, "state.json")
LIVE_MIN_CHOICES = [5, 10, 15, 20, 30, 60]   # 交易時段多久算一次即時訊號（分鐘）；最短 5 分鐘，免得太常敲期交所
AUTO_SAVE_HOURS_CHOICES = [2, 3, 4, 6, 12, 24]  # 自動上傳最短間隔（小時）；最短 2 小時
AUTO_SAVE_MAX_PER_DAY = 6     # 自動上傳一天最多幾次（不管間隔設多短）
DEFAULTS = {"live_min": 10, "auto_save": False, "auto_save_hours": 6}
DAILY_AT = (6, 32)            # 每天收資料的時間（週一到週六）
SAVE_PATHS = ["history", "docs"]

lock = threading.Lock()       # 同時只跑一個工作
state_lock = threading.Lock()
status = {"live_ts": 0.0, "job": "", "job_started": "", "last": {}, "next_live": "", "dirty": None, "dirty_at": 0}


def now():
    return dt.datetime.now()


def log(msg):
    os.makedirs(LOGS, exist_ok=True)
    line = f"{now():%Y-%m-%d %H:%M:%S} {msg}"
    print(line, flush=True)
    with open(os.path.join(LOGS, f"{now():%Y-%m}.log"), "a", encoding="utf-8") as f:
        f.write(line + "\n")


def load_state():
    try:
        return json.load(open(STATE, encoding="utf-8"))
    except Exception:
        return {}


def save_state(**kw):
    with state_lock:
        s = load_state()
        s.update(kw)
        tmp = STATE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(s, f, ensure_ascii=False, indent=1)
        os.replace(tmp, STATE)


def settings():
    """目前的設定（不合法的值一律退回預設）。"""
    s = load_state()
    out = dict(DEFAULTS)
    if s.get("live_min") in LIVE_MIN_CHOICES:
        out["live_min"] = s["live_min"]
    if isinstance(s.get("auto_save"), bool):
        out["auto_save"] = s["auto_save"]
    if s.get("auto_save_hours") in AUTO_SAVE_HOURS_CHOICES:
        out["auto_save_hours"] = s["auto_save_hours"]
    return out


def update_settings(d):
    kw = {}
    try:
        if "live_min" in d:
            v = int(d["live_min"])
            if v not in LIVE_MIN_CHOICES:
                return False, f"更新頻率只能選 {LIVE_MIN_CHOICES} 分鐘"
            kw["live_min"] = v
        if "auto_save_hours" in d:
            v = int(d["auto_save_hours"])
            if v not in AUTO_SAVE_HOURS_CHOICES:
                return False, f"自動上傳間隔只能選 {AUTO_SAVE_HOURS_CHOICES} 小時"
            kw["auto_save_hours"] = v
        if "auto_save" in d:
            kw["auto_save"] = d["auto_save"] is True
    except (TypeError, ValueError):
        return False, "設定格式不對"
    if not kw:
        return False, "沒有要改的設定"
    old = settings()
    save_state(**kw)
    new = settings()
    changes = []
    if new["live_min"] != old["live_min"]:
        changes.append(f"更新頻率 {old['live_min']} → {new['live_min']} 分鐘")
    if new["auto_save"] != old["auto_save"]:
        changes.append("自動上傳 " + ("打開" if new["auto_save"] else "關閉"))
    if new["auto_save_hours"] != old["auto_save_hours"]:
        changes.append(f"自動上傳間隔 {old['auto_save_hours']} → {new['auto_save_hours']} 小時")
    if changes:
        log("設定變更：" + "、".join(changes))
    return True, "已儲存設定" + ("：" + "、".join(changes) if changes else "（沒有變動）")


def run(argv, cwd=REPO, timeout=1800):
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    try:
        p = subprocess.run(argv, cwd=cwd, capture_output=True, timeout=timeout, env=env,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception as e:
        log(f"  {argv[0]} 執行失敗：{e}")
        return 1, str(e)
    out = (p.stdout + p.stderr).decode("utf-8", errors="replace").strip()
    for l in out.splitlines()[-15:]:
        log("  " + l)
    return p.returncode, out


def py(*args, cwd=REPO):
    return run([sys.executable, *args], cwd=cwd)


# ---------------- 工作 ----------------
def job_live():
    status["live_ts"] = time.time()
    rc, out = py("strategy/live.py", "--out", "docs")
    return rc == 0, "即時訊號已更新" if rc == 0 else "live.py 失敗：" + out[-300:]


def job_daily():
    msgs = []
    rc, out = py("strategy/collect_ticks.py", "history/tx_1min")
    if rc != 0:
        return False, "收台指期 1 分K 失敗（等一下會再試）：" + out[-300:]
    msgs.append("1 分K ✓")
    rc, out = py("strategy/collect_option_ticks.py", "history/txo_ticks")
    msgs.append("選擇權逐筆 ✓" if rc == 0 else "選擇權逐筆失敗（明天會補）")
    rc, out = py("forward.py", "../history/tx_1min", "../docs/strategy.html", "../history/forward_trades.csv",
                 cwd=os.path.join(REPO, "strategy"))
    msgs.append("前測 ✓" if rc == 0 else "前測失敗")
    ok, m = job_live()
    msgs.append("即時訊號 ✓" if ok else "即時訊號失敗")
    save_state(daily_done=f"{now():%Y-%m-%d}")
    return True, "收資料完成：" + "、".join(msgs)


def git(*args, timeout=300):
    return run(["git", *args], timeout=timeout)


def dirty_files():
    rc, out = run(["git", "status", "--porcelain", "--", *SAVE_PATHS], timeout=60)
    return None if rc != 0 else len([l for l in out.splitlines() if l.strip()])


def ensure_main():
    """資料只存到 main。資料夾停在別的分支（例如本機版程式的 PR 還沒合併）時：已合併就切回 main，沒合併就不存。"""
    rc, br = git("rev-parse", "--abbrev-ref", "HEAD")
    if rc != 0 or br.strip() == "main":
        return None
    git("fetch", "-q", "origin", "main")
    rc, _ = git("merge-base", "--is-ancestor", "HEAD", "origin/main")
    if rc != 0:
        return f"這個資料夾目前在「{br.strip()}」分支（本機版程式還沒合併進 main），請先到 GitHub 合併那個 PR 再存"
    rc, out = git("checkout", "-q", "main")
    if rc != 0:
        return "切回 main 失敗：" + out[-300:]
    git("merge", "-q", "--ff-only", "origin/main")
    return None


def job_save(auto=False):
    err = ensure_main()
    if err:
        return False, err
    git("add", "-A", "--", *SAVE_PATHS)
    rc, _ = git("diff", "--cached", "--quiet")
    if rc == 0:
        return True, "沒有新的資料要存，GitHub 上已經是最新"
    msg = f"本機資料 {now():%Y-%m-%d %H:%M}（{'自動存' if auto else '手動存'}）"
    rc, out = git("-c", "user.name=txo-bot", "-c", "user.email=txo-bot@users.noreply.github.com", "commit", "-q", "-m", msg)
    if rc != 0:
        return False, "commit 失敗：" + out[-300:]
    # GitHub 上有新的程式或網頁就先接上；產生的網頁／json 衝突時用本機剛算的（rebase 時 -X theirs 指自己的 commit）
    rc, out = git("pull", "-q", "--rebase", "-X", "theirs", "origin", "main")
    if rc != 0:
        git("rebase", "--abort")
        return False, "跟 GitHub 對齊失敗（資料已在本機 commit，沒有遺失，可以晚點再按一次）：" + out[-300:]
    rc, out = git("push", "-q", "origin", "HEAD:main")
    if rc != 0:
        return False, "推上 GitHub 失敗（資料已在本機 commit，沒有遺失，可以晚點再按一次）：" + out[-300:]
    save_state(saved_at=f"{now():%Y-%m-%d %H:%M}", saved_ts=time.time())
    if auto:
        c = load_state().get("auto_saves") or {}
        today = f"{now():%Y-%m-%d}"
        save_state(auto_saves={"date": today, "n": (c.get("n", 0) if c.get("date") == today else 0) + 1})
    return True, "已存到 GitHub：" + msg


def job_autosave():
    save_state(auto_try_ts=time.time())     # 失敗也要等下一個間隔才再試，不會一直重推
    return job_save(auto=True)


JOBS = {"live": ("更新即時訊號", job_live), "daily": ("收期交所資料", job_daily), "save": ("存到 GitHub", job_save),
        "autosave": ("自動存到 GitHub", job_autosave)}
MANUAL_JOBS = ("live", "daily", "save")


def do(name, why=""):
    """跑一個工作；已經有工作在跑就回 False。"""
    if not lock.acquire(blocking=False):
        return False
    title, fn = JOBS[name]
    status.update(job=title, job_started=f"{now():%H:%M:%S}")
    log(f"開始：{title}{('（' + why + '）') if why else ''}")
    t0 = time.time()
    try:
        ok, msg = fn()
    except Exception as e:
        ok, msg = False, f"{title}出錯：{e}"
    finally:
        status.update(job="", job_started="", dirty_at=0)
        lock.release()
    status["last"][name] = {"ok": ok, "msg": msg, "at": f"{now():%Y-%m-%d %H:%M}", "sec": round(time.time() - t0)}
    log(("完成：" if ok else "失敗：") + msg + f"（{time.time() - t0:.0f} 秒）")
    return True


def start(name, why=""):
    if lock.locked():
        return False
    threading.Thread(target=do, args=(name, why), daemon=True).start()
    return True


# ---------------- 排程 ----------------
def in_session(t):
    wd, m = t.weekday(), t.hour * 60 + t.minute     # 週一 = 0
    day = wd <= 4 and 8 * 60 + 45 <= m <= 13 * 60 + 50
    night = (wd <= 4 and m >= 15 * 60) or (1 <= wd <= 5 and m <= 5 * 60 + 6)
    return day or night


def auto_saves_today(t):
    c = load_state().get("auto_saves") or {}
    return c.get("n", 0) if c.get("date") == f"{t:%Y-%m-%d}" else 0


def autosave_due(t, cfg):
    """自動上傳：有打開、距離上次上傳（手動或自動）超過間隔、今天還沒到上限、而且資料真的有變才推。"""
    if not cfg["auto_save"]:
        return False
    st = load_state()
    since = time.time() - max(st.get("auto_try_ts", 0), st.get("saved_ts", 0))
    if since < cfg["auto_save_hours"] * 3600:
        return False
    if auto_saves_today(t) >= AUTO_SAVE_MAX_PER_DAY:
        return False
    if status["dirty_at"] < time.time() - 60:
        status["dirty"], status["dirty_at"] = dirty_files(), time.time()
    return bool(status["dirty"])


def daily_due(t):
    return t.weekday() <= 5 and (t.hour, t.minute) >= DAILY_AT and load_state().get("daily_done") != f"{t:%Y-%m-%d}"


def scheduler():
    time.sleep(2)
    while True:
        t = now()
        try:
            cfg = settings()
            if daily_due(t):
                if do("daily", "每天 06:32 收資料"):
                    if not status["last"]["daily"]["ok"]:
                        time.sleep(600)          # 期交所還沒好之類的，10 分鐘後再試
            elif in_session(t) and time.time() - status["live_ts"] >= cfg["live_min"] * 60 - 5:
                do("live", "交易時段定時")
            elif autosave_due(t, cfg):
                do("autosave", f"每 {cfg['auto_save_hours']} 小時、資料有變")
            nxt = ""
            if in_session(t):
                nxt = (dt.datetime.fromtimestamp(max(status["live_ts"] + cfg["live_min"] * 60, time.time()))).strftime("%H:%M")
            status["next_live"] = nxt
        except Exception as e:
            log(f"排程出錯：{e}")
        time.sleep(20)


# ---------------- 網頁 ----------------
TOOLBAR_TAG = '<script src="/__local/toolbar.js"></script>'
SHEET_STUB = b"/* local: \xe6\x9c\xac\xe6\xa9\x9f\xe7\x89\x88\xe4\xb8\x8d\xe8\xae\x80 Google \xe8\xa9\xa6\xe7\xae\x97\xe8\xa1\xa8 */\n"
TYPES = {".html": "text/html; charset=utf-8", ".js": "application/javascript; charset=utf-8", ".json": "application/json; charset=utf-8",
         ".css": "text/css; charset=utf-8", ".png": "image/png", ".svg": "image/svg+xml", ".ico": "image/x-icon"}


def status_json():
    st = load_state()
    cfg = settings()
    if status["dirty_at"] < time.time() - 60 and not lock.locked():
        status["dirty"], status["dirty_at"] = dirty_files(), time.time()
    gen = ""
    try:
        gen = json.load(open(os.path.join(DOCS, "live.json"), encoding="utf-8")).get("generated", "")
    except Exception:
        pass
    return {"job": status["job"], "job_started": status["job_started"], "last": status["last"], "next_live": status["next_live"],
            "in_session": in_session(now()), "generated": gen, "daily_done": st.get("daily_done", ""),
            "saved_at": st.get("saved_at", ""), "dirty": status["dirty"], "live_min": cfg["live_min"],
            "settings": cfg, "choices": {"live_min": LIVE_MIN_CHOICES, "auto_save_hours": AUTO_SAVE_HOURS_CHOICES},
            "auto_save_max": AUTO_SAVE_MAX_PER_DAY, "auto_saves_today": auto_saves_today(now()), "next_auto_save": next_auto_save(cfg, st)}


def next_auto_save(cfg, st):
    if not cfg["auto_save"]:
        return ""
    if auto_saves_today(now()) >= AUTO_SAVE_MAX_PER_DAY:
        return "明天"
    last = max(st.get("auto_try_ts", 0), st.get("saved_ts", 0))
    return dt.datetime.fromtimestamp(max(last + cfg["auto_save_hours"] * 3600, time.time())).strftime("%m/%d %H:%M")


class Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def send(self, code, body, ctype="text/plain; charset=utf-8"):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path
        if path in ("/", "/index.html"):
            self.send_response(302)
            self.send_header("Location", "/live.html")
            self.end_headers()
            return
        if path == "/__local/status":
            return self.send(200, json.dumps(status_json(), ensure_ascii=False), TYPES[".json"])
        if path == "/__local/toolbar.js":
            return self.send(200, open(os.path.join(HERE, "toolbar.js"), "rb").read(), TYPES[".js"])
        if path == "/sheet.js":
            return self.send(200, SHEET_STUB, TYPES[".js"])
        f = os.path.normpath(os.path.join(DOCS, urllib.parse.unquote(path).lstrip("/")))
        if not f.startswith(DOCS + os.sep) or not os.path.isfile(f):
            return self.send(404, "找不到這一頁。還沒產生的話，按工具列的「立即更新」或「收今天資料」。")
        ext = os.path.splitext(f)[1].lower()
        body = open(f, "rb").read()
        if ext == ".html":
            txt = body.decode("utf-8", errors="replace")
            i = txt.rfind("</body>")
            txt = txt[:i] + TOOLBAR_TAG + txt[i:] if i >= 0 else txt + TOOLBAR_TAG
            body = txt.encode("utf-8")
        self.send(200, body, TYPES.get(ext, "application/octet-stream"))

    def do_POST(self):
        # 只接受本機工具列發的（自訂標頭讓別的網站沒辦法偷按）
        if self.headers.get("X-Tardemonkey") != "local":
            return self.send(403, "forbidden")
        path = urllib.parse.urlparse(self.path).path
        if path == "/__local/settings":
            try:
                n = int(self.headers.get("Content-Length") or 0)
                d = json.loads(self.rfile.read(min(n, 10000)).decode("utf-8") or "{}")
                ok, m = update_settings(d if isinstance(d, dict) else {})
            except Exception as e:
                ok, m = False, f"設定格式不對：{e}"
            return self.send(200 if ok else 400, json.dumps({"ok": ok, "msg": m, "settings": settings()}, ensure_ascii=False), TYPES[".json"])
        q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        name = (q.get("job") or [""])[0]
        if path != "/__local/run" or name not in MANUAL_JOBS:
            return self.send(404, "not found")
        ok = start(name, "手動按")
        self.send(200, json.dumps({"started": ok, "busy": status["job"]}, ensure_ascii=False), TYPES[".json"])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("--save", action="store_true", help="只存到 GitHub 然後結束")
    a = ap.parse_args()
    os.chdir(REPO)
    if a.save:
        do("save", "存到GitHub.bat")
        r = status["last"].get("save") or {}
        print("\n" + r.get("msg", ""))
        sys.exit(0 if r.get("ok") else 1)
    try:
        srv = http.server.ThreadingHTTPServer(("127.0.0.1", a.port), Handler)
    except OSError:
        print(f"連接埠 {a.port} 已經有人在用，可能交易猩猩已經開著了，直接開網頁。")
        webbrowser.open(f"http://127.0.0.1:{a.port}/live.html")
        time.sleep(3)
        return
    url = f"http://127.0.0.1:{a.port}/live.html"
    log(f"交易猩猩本機版啟動：{url}（關掉這個視窗就停止）")
    threading.Thread(target=scheduler, daemon=True).start()
    if not os.path.exists(os.path.join(DOCS, "live.html")) or in_session(now()):
        start("live", "剛啟動")
    if not a.no_browser:
        threading.Timer(1.5, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
