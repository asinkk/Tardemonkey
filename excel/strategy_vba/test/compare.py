"""比對 VBA 與 Python 的輸出（數字容許 1e-5 誤差）。用法：python compare.py py.txt vba.txt"""
import sys
a = open(sys.argv[1], encoding="utf-8").read().splitlines()
b = open(sys.argv[2], encoding="utf-8", errors="replace").read().replace("\r", "").splitlines()
def eq(x, y):
    if x == y: return True
    try: return abs(float(x) - float(y)) <= 1e-5 * max(1, abs(float(x)))
    except ValueError: return False
bad = 0
for i in range(max(len(a), len(b))):
    x = a[i].split("\t") if i < len(a) else ["<無>"]
    y = b[i].split("\t") if i < len(b) else ["<無>"]
    if len(x) != len(y) or not all(eq(p, q) for p, q in zip(x, y)):
        bad += 1
        if bad <= 15: print(f"第 {i+1} 行\n  PY : {a[i] if i < len(a) else ''}\n  VBA: {b[i] if i < len(b) else ''}")
print(f"{len(a)} / {len(b)} 行，不同 {bad} 行")
sys.exit(1 if bad else 0)
