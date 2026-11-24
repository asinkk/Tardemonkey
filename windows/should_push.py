"""判斷這次算出來的即時訊號值不值得推上 GitHub（減少 commit 數）。

只比「交易上有意義」的部分，現價、更新時間、條件明細、移動停損每根 K 的小變動都不算：
  - v3 部位：方向、型態、進場時間與價格、起始停損、是否已保本、是否待出場
  - v3 訊號：方向、型態、時間
  - 週選紀錄 history/option_trades.json：每筆的成交、出場（不含即時價格欄位）
跟 GitHub 上 origin/main 的版本比，有不同就 exit 0（要推），一樣就 exit 1（跳過）。
用法：python windows/should_push.py   （在 repo 根目錄、live.py 跑完之後）
"""
import json, subprocess, sys

VOLATILE = {"last", "last_t", "env", "peak", "peak_t", "lag"}


def load(path, ref):
    try:
        if ref:
            txt = subprocess.run(["git", "show", f"{ref}:{path}"], capture_output=True, check=True).stdout.decode("utf-8")
        else:
            txt = open(path, encoding="utf-8").read()
        return json.loads(txt)
    except Exception:
        return None


def key(ref):
    live = load("docs/live.json", ref) or {}
    v3 = live.get("v3") or {}
    pos = v3.get("pos")
    if pos:
        pos = {k: pos.get(k) for k in ("side", "kind", "entry_t", "entry", "stop0", "be", "pending_exit")}
    sig = v3.get("sig")
    if sig:
        sig = {k: sig.get(k) for k in ("side", "kind", "at")}
    opt = (load("history/option_trades.json", ref) or {}).get("trades") or []
    opt = [{k: v for k, v in t.items() if k not in VOLATILE} for t in opt]
    return json.dumps({"pos": pos, "sig": sig, "opt": opt}, ensure_ascii=False, sort_keys=True)


if __name__ == "__main__":
    changed = key(None) != key("origin/main")
    print("訊號／部位／週選紀錄有變動，要推" if changed else "訊號／部位／週選紀錄沒變，不推")
    sys.exit(0 if changed else 1)
