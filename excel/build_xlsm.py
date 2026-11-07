#!/usr/bin/env python3
"""產生 台指週選盤面.xlsm：openpyxl 建活頁簿，再自己組 vbaProject.bin（只放原始碼，Excel 開啟時會重新編譯）。

用法：python3 build_xlsm.py [輸出檔]
需要：openpyxl
參考：[MS-OVBA] 與 [MS-CFB] 規格。
"""
import io, os, random, struct, sys, uuid, zipfile

import openpyxl
from openpyxl.styles import Font, Alignment

HERE = os.path.dirname(os.path.abspath(__file__))
CODEPAGE = 950          # 繁體中文 Windows；VBA 原始碼以 Big5 存
ENC = "cp950"


# ---------------------------------------------------------------- MS-OVBA 壓縮
def _compress_chunk(chunk: bytes) -> bytes:
    out = bytearray()
    n = len(chunk)
    i = 0
    table = {}  # 3-byte prefix -> positions
    def add(pos):
        if pos + 3 <= n:
            table.setdefault(chunk[pos:pos + 3], []).append(pos)
    while i < n:
        flag_pos = len(out)
        out.append(0)
        flag = 0
        for bit in range(8):
            if i >= n:
                break
            best_len = best_off = 0
            if i > 0 and i + 3 <= n:
                bc = max(4, (i - 1).bit_length())
                max_len = (0xFFFF >> bc) + 3
                max_off = 1 << bc
                for j in reversed(table.get(chunk[i:i + 3], [])):
                    off = i - j
                    if off > max_off:
                        break
                    l = 0
                    while l < max_len and i + l < n and chunk[j + l] == chunk[i + l]:
                        l += 1
                    if l > best_len:
                        best_len, best_off = l, off
                        if l == max_len:
                            break
            if best_len >= 3:
                bc = max(4, (i - 1).bit_length())
                token = ((best_off - 1) << (16 - bc)) | (best_len - 3)
                out += struct.pack("<H", token)
                flag |= 1 << bit
                for p in range(i, i + best_len):
                    add(p)
                i += best_len
            else:
                out.append(chunk[i])
                add(i)
                i += 1
        out[flag_pos] = flag
    return bytes(out)


def ovba_compress(data: bytes) -> bytes:
    out = bytearray(b"\x01")
    for pos in range(0, max(len(data), 1), 4096):
        chunk = data[pos:pos + 4096]
        comp = _compress_chunk(chunk)
        if len(comp) > 4096:
            out += struct.pack("<H", 0x3FFF) + chunk.ljust(4096, b"\x00")
        else:
            out += struct.pack("<H", 0xB000 | (len(comp) + 2 - 3)) + comp
    return bytes(out)


# ---------------------------------------------------------------- MS-OVBA 資料加密（PROJECT 的 CMG/DPB/GC）
def ovba_encrypt(data: bytes, project_id: str) -> str:
    seed = random.randint(0, 255)
    version = 2
    proj_key = sum(project_id.encode("ascii")) & 0xFF
    out = [seed, seed ^ version, seed ^ proj_key]
    unenc1, enc1, enc2 = proj_key, seed ^ proj_key, seed ^ version
    ignored = (seed & 6) // 2
    stream = [random.randint(0, 255) for _ in range(ignored)] + list(struct.pack("<I", len(data))) + list(data)
    for b in stream:
        be = ((enc2 + unenc1) & 0xFF) ^ b
        out.append(be)
        enc2, enc1, unenc1 = enc1, be, b
    return "".join("%02X" % b for b in out)


# ---------------------------------------------------------------- dir 串流
def rec(rid, payload: bytes) -> bytes:
    return struct.pack("<HI", rid, len(payload)) + payload


def build_dir(modules, references):
    mb = lambda s: s.encode(ENC)
    u16 = lambda s: s.encode("utf-16-le")
    d = bytearray()
    d += rec(0x0001, struct.pack("<I", 1))                     # SYSKIND Win32
    d += rec(0x0002, struct.pack("<I", 0x0409))                # LCID
    d += rec(0x0014, struct.pack("<I", 0x0409))                # LCIDINVOKE
    d += rec(0x0003, struct.pack("<H", CODEPAGE))              # CODEPAGE
    d += rec(0x0004, b"VBAProject")                            # NAME
    d += rec(0x0005, b"") + struct.pack("<HI", 0x0040, 0)      # DOCSTRING
    d += rec(0x0006, b"") + struct.pack("<HI", 0x003D, 0)      # HELPFILEPATH
    d += rec(0x0007, struct.pack("<I", 0))                     # HELPCONTEXT
    d += rec(0x0008, struct.pack("<I", 0))                     # LIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)                 # VERSION
    d += rec(0x000C, b"") + struct.pack("<HI", 0x003C, 0)      # CONSTANTS
    for name, libid in references:
        d += rec(0x0016, mb(name)) + struct.pack("<HI", 0x003E, len(u16(name))) + u16(name)
        lb = mb(libid)
        d += struct.pack("<HII", 0x000D, 4 + len(lb) + 4 + 2, len(lb)) + lb + struct.pack("<IH", 0, 0)
    d += rec(0x000F, struct.pack("<H", len(modules)))          # MODULES count
    d += rec(0x0013, struct.pack("<H", 0xFFFF))                # PROJECTCOOKIE
    for m in modules:
        nm = m["name"]
        d += rec(0x0019, mb(nm))
        d += rec(0x0047, u16(nm))
        d += rec(0x001A, mb(nm)) + struct.pack("<HI", 0x0032, len(u16(nm))) + u16(nm)
        d += rec(0x001C, b"") + struct.pack("<HI", 0x0048, 0)
        d += rec(0x0031, struct.pack("<I", 0))                 # 原始碼從 0 開始（沒有 p-code）
        d += rec(0x001E, struct.pack("<I", 0))
        d += rec(0x002C, struct.pack("<H", 0xFFFF))
        d += struct.pack("<HI", 0x0022 if m["doc"] else 0x0021, 0)
        d += struct.pack("<HI", 0x002B, 0)
    d += struct.pack("<HI", 0x0010, 0)
    return bytes(d)


def build_project_stream(modules, project_id):
    lines = [f'ID="{project_id}"']
    for m in modules:
        if m["doc"]:
            lines.append(f"Document={m['name']}/&H00000000")
    for m in modules:
        if not m["doc"]:
            lines.append(f"Module={m['name']}")
    lines += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
              'CMG="%s"' % ovba_encrypt(struct.pack("<I", 0), project_id),
              'DPB="%s"' % ovba_encrypt(b"\x00", project_id),
              'GC="%s"' % ovba_encrypt(b"\xFF", project_id),
              "", "[Host Extender Info]",
              "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
              "", "[Workspace]"]
    for m in modules:
        lines.append(f"{m['name']}=0, 0, 0, 0, C")
    return ("\r\n".join(lines) + "\r\n").encode(ENC)


def build_projectwm(modules):
    out = bytearray()
    for m in modules:
        out += m["name"].encode(ENC) + b"\x00" + m["name"].encode("utf-16-le") + b"\x00\x00"
    out += b"\x00\x00"
    return bytes(out)


# ---------------------------------------------------------------- MS-CFB 寫入（v3、512 位元組磁區）
ENDOFCHAIN, FREESECT, FATSECT, NOSTREAM = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF


class Node:
    def __init__(self, name, data=None, children=None):
        self.name, self.data, self.children = name, data, children
        self.left = self.right = self.child = NOSTREAM
        self.color = 1
        self.start = ENDOFCHAIN
        self.size = 0


def _cfb_key(n):
    return (len(n.name), n.name.upper())


def _build_tree(nodes, depth, out_depths):
    if not nodes:
        return NOSTREAM
    mid = len(nodes) // 2
    root = nodes[mid]
    out_depths.append((root, depth))
    root.left = _build_tree(nodes[:mid], depth + 1, out_depths)
    root.right = _build_tree(nodes[mid + 1:], depth + 1, out_depths)
    return root.sid


def write_cfb(root_children):
    root = Node("Root Entry", children=root_children)
    entries = []

    def walk(n):
        n.sid = len(entries)
        entries.append(n)
        if n.children:
            for c in n.children:
                walk(c)
    walk(root)
    for n in entries:
        if n.children:
            kids = sorted(n.children, key=_cfb_key)
            depths = []
            n.child = _build_tree(kids, 0, depths)
            maxd = max(d for _, d in depths)
            perfect = len(kids) == (1 << (maxd + 1)) - 1
            for k, dpt in depths:      # 最深一層塗紅，維持紅黑樹的黑高相同
                k.color = 0 if (dpt == maxd and maxd > 0 and not perfect) else 1

    streams = [n for n in entries if n.data is not None]
    mini = [n for n in streams if len(n.data) < 4096]
    big = [n for n in streams if len(n.data) >= 4096]

    # 迷你串流
    ministream = bytearray()
    minifat = []
    for n in mini:
        n.size = len(n.data)
        if n.size == 0:
            n.start = ENDOFCHAIN
            continue
        cnt = (n.size + 63) // 64
        first = len(ministream) // 64
        n.start = first
        ministream += n.data.ljust(cnt * 64, b"\x00")
        for i in range(cnt):
            minifat.append(first + i + 1 if i < cnt - 1 else ENDOFCHAIN)
    root.size = len(ministream)

    def nsec(b):
        return (b + 511) // 512

    dir_bytes = len(entries) * 128
    n_dir = nsec(dir_bytes)
    n_minifat = nsec(len(minifat) * 4)
    n_ministream = nsec(len(ministream))
    n_big = [nsec(len(n.data)) for n in big]
    n_data = n_dir + n_minifat + n_ministream + sum(n_big)
    n_fat = 1
    while (n_fat + n_data) > n_fat * 128:
        n_fat += 1
    assert n_fat <= 109

    fat = []
    sec = 0

    def chain(count, kind=None):
        nonlocal sec
        start = sec
        for i in range(count):
            if kind == FATSECT:
                fat.append(FATSECT)
            else:
                fat.append(sec + 1 if i < count - 1 else ENDOFCHAIN)
            sec += 1
        return start if count else ENDOFCHAIN

    fat_start = chain(n_fat, FATSECT)
    dir_start = chain(n_dir)
    minifat_start = chain(n_minifat)
    root.start = chain(n_ministream) if n_ministream else ENDOFCHAIN
    for n, c in zip(big, n_big):
        n.size = len(n.data)
        n.start = chain(c)
    fat += [FREESECT] * (n_fat * 128 - len(fat))

    body = bytearray()
    body += b"".join(struct.pack("<I", x) for x in fat)
    dirb = bytearray()
    for n in entries:
        nm = n.name.encode("utf-16-le") + b"\x00\x00"
        typ = 5 if n is root else (1 if n.children is not None else 2)
        e = nm.ljust(64, b"\x00") + struct.pack("<HBB", len(nm), typ, n.color)
        e += struct.pack("<III", n.left, n.right, n.child)
        e += b"\x00" * 16 + struct.pack("<I", 0) + b"\x00" * 16
        e += struct.pack("<IQ", n.start if (n.data is not None or n is root) else 0, n.size)
        dirb += e
    while len(dirb) % 512:
        dirb += b"\x00" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM) + b"\x00" * 48
    body += dirb
    mf = b"".join(struct.pack("<I", x) for x in minifat)
    body += mf.ljust(n_minifat * 512, b"\xFF")
    body += bytes(ministream).ljust(n_ministream * 512, b"\x00")
    for n, c in zip(big, n_big):
        body += n.data.ljust(c * 512, b"\x00")

    hdr = bytearray()
    hdr += bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 16
    hdr += struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6) + b"\x00" * 6
    hdr += struct.pack("<IIII", 0, n_fat, dir_start, 0)
    hdr += struct.pack("<IIIII", 4096, minifat_start if n_minifat else ENDOFCHAIN, n_minifat, ENDOFCHAIN, 0)
    difat = [fat_start + i for i in range(n_fat)] + [FREESECT] * (109 - n_fat)
    hdr += b"".join(struct.pack("<I", x) for x in difat)
    assert len(hdr) == 512
    return bytes(hdr) + bytes(body)


# ---------------------------------------------------------------- VBA 專案
DOC_ATTR = ('Attribute VB_Name = "{name}"\r\nAttribute VB_Base = "0{{{guid}}}"\r\nAttribute VB_GlobalNameSpace = False\r\n'
            'Attribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\n'
            'Attribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n')
WB_GUID = "00020819-0000-0000-C000-000000000046"
WS_GUID = "00020820-0000-0000-C000-000000000046"

THISWORKBOOK_CODE = '''Option Explicit

Private Sub Workbook_Open()
    EnsureButton
    UpdateAll
End Sub
'''

REFERENCES = [
    ("VBA", r"*\G{000204EF-0000-0000-C000-000000000046}#4.2#9#C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL#Visual Basic For Applications"),
    ("Excel", r"*\G{00020813-0000-0000-C000-000000000046}#1.9#0#C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE#Microsoft Excel 16.0 Object Library"),
    ("stdole", r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation"),
]


def crlf(s):
    return s.replace("\r\n", "\n").replace("\n", "\r\n")


def build_vba_project(sheet_codenames):
    txo = open(os.path.join(HERE, "TXO.bas"), encoding="utf-8").read()
    modules = [{"name": "ThisWorkbook", "doc": True,
                "src": DOC_ATTR.format(name="ThisWorkbook", guid=WB_GUID) + crlf(THISWORKBOOK_CODE)}]
    for cn in sheet_codenames:
        modules.append({"name": cn, "doc": True, "src": DOC_ATTR.format(name=cn, guid=WS_GUID)})
    modules.append({"name": "TXO", "doc": False, "src": crlf(txo)})
    project_id = "{" + str(uuid.uuid4()).upper() + "}"
    vba_children = [Node("_VBA_PROJECT", data=bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00])),
                    Node("dir", data=ovba_compress(build_dir(modules, REFERENCES)))]
    for m in modules:
        vba_children.append(Node(m["name"], data=ovba_compress(m["src"].encode(ENC))))
    root_children = [Node("VBA", children=vba_children),
                     Node("PROJECT", data=build_project_stream(modules, project_id)),
                     Node("PROJECTwm", data=build_projectwm(modules))]
    return write_cfb(root_children)


# ---------------------------------------------------------------- 活頁簿
HELP = [
    ("台指週選盤面（Excel 巨集版）", None),
    ("", None),
    ("第一次使用", "b"),
    ("1. 下載後先在檔案上按右鍵 →「內容」→ 最下面勾「解除封鎖」→ 確定。沒做這步，Excel 會擋掉巨集。", None),
    ("2. 打開檔案，上方黃色列按「啟用內容」。之後每次打開都會自動抓最新資料。", None),
    ("3. 要再更新就按「盤面」右上角的「更新資料」按鈕（約 20–40 秒）。", None),
    ("4. 想讓台指 VIX 歷史一直累積到一年，更新後記得存檔。", None),
    ("", None),
    ("限制", "b"),
    ("只能用 Windows 版 Excel（2010 以後，建議 Microsoft 365）。Mac 版與手機版 Excel 不能執行巨集。", None),
    ("時間以電腦時鐘為準，請設成台北時間。日盤 08:45–13:45 先用日盤報價，其他時間先用夜盤報價。", None),
    ("每次更新會重建「盤面」「說明」以外的工作表，自己加的內容請放在別的活頁簿。", None),
    ("", None),
    ("工作表", "b"),
    ("盤面：各到期比較、籌碼與波動重點、最近到期的未平倉牆。", None),
    ("週選／月選：各履約價的 OI、OI 增減、價格、IV、最大痛點，附四張圖。", None),
    ("籌碼：三大法人選擇權淨部位、大額交易人、台指期近月與正逆價差、期貨法人淨部位。", None),
    ("波動：台指 VIX 的 IV Rank 與百分位、加權指數 20 日歷史波動、ATM IV ÷ HV。", None),
    ("美股：美國 VIX、S&P 500／那斯達克／道瓊期貨、費城半導體、台積電 ADR。", None),
    ("PC歷史：全市場 P/C ratio 近一個月。", None),
    ("", None),
    ("計算方式", "b"),
    ("合成期貨：取買權與賣權中價最接近的履約價 K，F = K + C − P。", None),
    ("預期波動：價平買權加賣權（跨式）價格。區間 = F ± 跨式。", None),
    ("IV：Black-76 模型、到期日 13:30 為終點，用即時買賣中價反推；沒有即時報價時用結算價。", None),
    ("P/C ratio：賣權除以買權。未平倉量用最近日盤收盤；成交量 P/C 分前一日與即時盤。", None),
    ("最大痛點：以收盤未平倉量計算，所有買方部位價值總和最小的結算價。", None),
    ("未平倉牆：合成期貨 ±3,000 點內 OI 最大的買權與賣權履約價。", None),
    ("OI 增減：今天日盤收盤 OI 減前一交易日。", None),
    ("外資選擇權方向：外資買權未平倉淨額減賣權未平倉淨額（億元），正值偏多。", None),
    ("正逆價差：台指期近月減加權指數，只算日盤。", None),
    ("IV Rank：（今天 − 區間最低）÷（最高 − 最低）。期交所只提供近 3 個月 VIX，存檔後會累積。", None),
    ("HV：加權指數日報酬 20 日標準差年化。", None),
    ("", None),
    ("資料來源：台灣期貨交易所（www.taifex.com.tw、mis.taifex.com.tw）、Yahoo Finance。僅供參考，不構成投資建議。", None),
]


def build_workbook(path):
    wb = openpyxl.Workbook()
    wb.code_name = "ThisWorkbook"
    ws = wb.active
    ws.title = "盤面"
    ws.sheet_properties.codeName = "Sheet1"
    ws.sheet_view.showGridLines = False
    ws["A1"] = "台指週選盤面"
    ws["A1"].font = Font(size=20, bold=True)
    ws["A3"] = "請按上方黃色列的「啟用內容」，或按右上角「更新資料」，就會開始抓資料（約 20–40 秒）。"
    ws["A4"] = "如果沒有看到「啟用內容」而是紅色的「安全性風險」，請先關閉檔案，在檔案上按右鍵 →「內容」→ 勾「解除封鎖」。"
    ws["A5"] = "使用方式與計算說明見「說明」工作表。"
    ws.column_dimensions["A"].width = 24
    h = wb.create_sheet("說明")
    h.sheet_properties.codeName = "Sheet2"
    h.sheet_view.showGridLines = False
    h.column_dimensions["A"].width = 120
    for i, (t, st) in enumerate(HELP, start=1):
        c = h.cell(row=i, column=1, value=t)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if i == 1:
            c.font = Font(size=16, bold=True)
        elif st == "b":
            c.font = Font(size=12, bold=True)
    buf = io.BytesIO()
    wb.save(buf)
    vba = build_vba_project(["Sheet1", "Sheet2"])

    src = zipfile.ZipFile(io.BytesIO(buf.getvalue()))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                t = data.decode("utf-8")
                t = t.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                              "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
                t = t.replace("<Default ", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default ', 1)
                data = t.encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                t = data.decode("utf-8")
                t = t.replace("</Relationships>",
                              '<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
                              'Target="vbaProject.bin"/></Relationships>')
                data = t.encode("utf-8")
            z.writestr(item, data)
        z.writestr("xl/vbaProject.bin", vba)


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "台指週選盤面.xlsm")
    build_workbook(out)
    print("已輸出", out)
