Option Explicit
' ============================================================
' 台指策略訊號（Excel 版）：計算
' 完全照 strategy/live.py、build_bars.py、backtest.py、v3.py（V3）、daily.py（D2）的算法改寫成 VBA。
' 這個模組只做計算，不碰網路（網路在 SNet、工作表在 SView），所以也能在 LibreOffice 裡測試。
' 時間一律用「分鐘序號」(Long) = Excel 日期序號 × 1440 + 當天第幾分鐘，避免浮點誤差。
' K 棒陣列從 0 開始編號，跟 Python 的索引一樣。
' ============================================================

Public Const COST As Double = 3          ' 每次來回滑價成本（點），R 的計算用
Public Const CAPITAL As Double = 200000  ' 本金 20 萬
Public Const MULT As Double = 10         ' 1 口微台每點 10 元
Public Const FEE As Double = 25          ' 每次下單 25 元
Public Const T_DAYOPEN As Long = 525     ' 08:45
Public Const T_DAYCLOSE As Long = 825    ' 13:45
Public Const T_NIGHT As Long = 900       ' 15:00

' 交易紀錄欄位（二維陣列 tr(欄位, 第幾筆)）
Public Const F_T As Long = 0        ' 訊號 K 開始時間
Public Const F_SIDE As Long = 1     ' 1 多、-1 空
Public Const F_KIND As Long = 2     ' 突破 / 拉回
Public Const F_RISK As Long = 3     ' R（初始風險點數）
Public Const F_RMUL As Long = 4     ' r（扣成本後的 R 倍數）
Public Const F_BARS As Long = 5
Public Const F_ENTRY As Long = 6
Public Const F_ENTRYT As Long = 7
Public Const F_STOP0 As Long = 8
Public Const F_STOP As Long = 9
Public Const F_EXITT As Long = 10
Public Const F_EXIT As Long = 11
Public Const F_OPEN As Long = 12
Public Const F_BE As Long = 13
Public Const F_WHY As Long = 14
Public Const F_LAST As Long = 14

' ------------------------------------------------------------
' 1 分K 資料：每個「單位」對應 Python 的一個檔案（history 的 Daily_*.csv 或即時盤別）
' ------------------------------------------------------------
Public gUN As Long
Public gUName() As String
Public gUDay() As Long          ' 檔名日期（交易日）
Public gULive() As Boolean
Public gUA() As Long            ' 這個單位的第一筆、最後一筆列號
Public gUB() As Long
Public gRN As Long
Public gRT() As Long
Public gRC() As Long            ' 契約月份 YYYYMM
Public gRO() As Double
Public gRH() As Double
Public gRL() As Double
Public gRCl() As Double
Public gRV() As Double
Public gHistLast As Long        ' history 最後一分鐘
Public gHistFiles As Long

' 即時 1 分K（期交所 MIS）：槽 1 = 日盤、2 = 夜盤
Public gLiveHas(1 To 2) As Boolean
Public gLiveStart(1 To 2) As Long
Public gLiveSym(1 To 2) As String
Public gLiveCName(1 To 2) As String
Public gLiveLast(1 To 2) As Double
Public gLiveRef(1 To 2) As Double
Public gLiveDate(1 To 2) As String
Public gLiveLastMin(1 To 2) As Long
Public gLiveMinutes(1 To 2) As Long
Public gLiveErr(1 To 2) As String
Public gLN As Long
Public gLS() As Long            ' 槽
Public gLC() As Long            ' 契約
Public gLT() As Long
Public gLO() As Double
Public gLH() As Double
Public gLL() As Double
Public gLCl() As Double
Public gLV() As Double

' 資料缺口（應該開盤卻沒有資料的盤別開始時間）
Public gGapN As Long
Public gGap() As Long

' ---- 全域結果（VBA 規定模組層級宣告必須放在所有程序之前）----
' v3 台指期 60 分K
Public gVN As Long
Public gVT() As Long
Public gVO() As Double
Public gVH() As Double
Public gVL() As Double
Public gVC() As Double
Public gVE20() As Double
Public gVE60() As Double
Public gVHI() As Double
Public gVMID() As Double
Public gVSD() As Double
Public gVATR() As Double
Public gVTr() As Variant
Public gVTn As Long
Public gVPos As Long
Public gVSig() As Variant       ' 有訊號時是一筆交易（tr2 的那一筆，0 To F_LAST），沒有時 gVSigHas = False
Public gVSigHas As Boolean
Public gVPend As String
Public gVPendHas As Boolean
' D2 台指期日K（只用日盤）
Public gDN As Long
Public gDT() As Long
Public gDO() As Double
Public gDH() As Double
Public gDL() As Double
Public gDC() As Double
Public gDE20() As Double
Public gDE60() As Double
Public gDHI() As Double
Public gDMID() As Double
Public gDSD() As Double
Public gDATR() As Double
Public gDI0 As Long
Public gDTr() As Variant
Public gDTn As Long
Public gDPos As Long
Public gDSig() As Variant
Public gDSigHas As Boolean
Public gDPend As String
Public gDPendHas As Boolean
' 加權指數回測（近 2 年）
Public gWVTr() As Variant
Public gWVTn As Long
Public gWDTr() As Variant
Public gWDTn As Long
Public gTwN As Long
Public gTwLast As Long
' 其他
Public gLastC As Double
Public gM1N As Long
Public gNowTm As Long
Public gStep As String

' 條件明細（顯示用）：策略 1 = v3、2 = D2；方向 1 = 多、2 = 空
Public gCkN As Long
Public gCkStrat() As Long
Public gCkSide() As Long
Public gCkGroup() As String
Public gCkName() As String
Public gCkOk() As Boolean
Public gCkDetail() As String
Public gStopOk(1 To 2, 1 To 2) As Boolean
Public gStopVal(1 To 2, 1 To 2) As Double
Public gStopRaw(1 To 2, 1 To 2) As Double
' 測試用：有值時 MisGet 改讀測試資料
Public gFixtureDir As String

' ============================================================
' 時間與數字小工具
' ============================================================
Public Function DayNum(ByVal y As Long, ByVal m As Long, ByVal d As Long) As Long
    DayNum = CLng(DateSerial(y, m, d))
End Function

Public Function TmOf(ByVal dn As Long, ByVal mins As Long) As Long
    TmOf = dn * 1440 + mins
End Function

Public Function TmFromDate(ByVal v As Date) As Long
    Dim d As Long
    d = Int(CDbl(v))
    TmFromDate = d * 1440 + Hour(v) * 60 + Minute(v)
End Function

Public Function DayOfTm(ByVal tm As Long) As Long
    DayOfTm = tm \ 1440
End Function

Public Function MinOfTm(ByVal tm As Long) As Long
    MinOfTm = tm - (tm \ 1440) * 1440
End Function

' 星期：一 = 0 … 日 = 6（跟 Python weekday() 一樣）
Public Function Wd(ByVal dn As Long) As Long
    Wd = (dn + 5) Mod 7
End Function

Public Function NextWeekday(ByVal dn As Long) As Long
    dn = dn + 1
    Do While Wd(dn) >= 5
        dn = dn + 1
    Loop
    NextWeekday = dn
End Function

Private Function NthSunday(ByVal y As Long, ByVal m As Long, ByVal n As Long) As Long
    Dim f As Long
    f = DayNum(y, m, 1)
    NthSunday = f + ((6 - Wd(f)) Mod 7) + 7 * (n - 1)
End Function

' 美國夏令時間（看日期）
Public Function UsDst(ByVal tm As Long) As Boolean
    Dim d As Long, y As Long
    d = DayOfTm(tm)
    y = Year(d)
    UsDst = (NthSunday(y, 3, 2) <= d) And (d < NthSunday(y, 11, 1))
End Function

' 可開新倉的 60 分K（看訊號 K 的開始時間）。mode 1 = 台指期（tx_revised）、2 = 加權指數（twii_revised）
Public Function Allowed(ByVal tm As Long, ByVal mode As Long) As Boolean
    Dim hm As Long
    hm = MinOfTm(tm)
    If mode = 2 Then
        Allowed = (hm \ 60 >= 9) And (hm \ 60 <= 12)
        Exit Function
    End If
    If hm = 525 Or hm = 585 Or hm = 645 Or hm = 705 Then Allowed = True: Exit Function
    If UsDst(tm) Then
        Allowed = (hm = 1260 Or hm = 1320 Or hm = 1380)
    Else
        Allowed = (hm = 1320 Or hm = 1380 Or hm = 0)
    End If
End Function

Public Function P2(ByVal n As Long) As String
    P2 = Right$("0" & CStr(n), 2)
End Function

' style 0 = YYYY-MM-DD HH:MM、1 = MM-DD HH:MM、2 = YYYY-MM-DD、3 = HH:MM、4 = MM-DD
Public Function FmtTm(ByVal tm As Long, Optional ByVal style As Long = 0) As String
    Dim d As Long, mi As Long, ds As String, hm As String
    d = tm \ 1440
    mi = tm - d * 1440
    ds = CStr(Year(d)) & "-" & P2(Month(d)) & "-" & P2(Day(d))
    hm = P2(mi \ 60) & ":" & P2(mi Mod 60)
    Select Case style
        Case 0: FmtTm = ds & " " & hm
        Case 1: FmtTm = Mid$(ds, 6) & " " & hm
        Case 2: FmtTm = ds
        Case 3: FmtTm = hm
        Case 4: FmtTm = Mid$(ds, 6)
    End Select
End Function

Public Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function

Public Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

Public Function Sel(ByVal cond As Boolean, ByVal a As Double, ByVal b As Double) As Double
    If cond Then Sel = a Else Sel = b
End Function

' Python 的 f"{x:,.0f}"（四捨六入五成雙，跟 Python 一樣）
Public Function N0(ByVal x As Double) As String
    N0 = Format$(Round(x), "#,##0")
End Function

' f"{x:+,.0f}"
Public Function S0(ByVal x As Double) As String
    Dim r As Double
    r = Round(x)
    If r < 0 Then S0 = "-" & Format$(-r, "#,##0") Else S0 = "+" & Format$(r, "#,##0")
End Function

' f"{x:+.nf}"
' 跟 Python round(x, n) 一樣看二進位值（VBA 的 Round(x, n) 會把 6.450000000000001 當成 6.45 再取偶數）
Public Function PyRound(ByVal x As Double, ByVal digits As Long) As Double
    PyRound = Round(x * 10 ^ digits) / 10 ^ digits
End Function

Public Function SN(ByVal x As Double, ByVal digits As Long) As String
    Dim r As Double, f As String
    r = PyRound(x, digits)
    f = "#,##0." & String$(digits, "0")
    If digits = 0 Then f = "#,##0"
    If r < 0 Then SN = "-" & Format$(-r, f) Else SN = "+" & Format$(r, f)
End Function

' 跟 Python f"{x:g}" 一樣的價格文字（整數不帶小數點）
Public Function NumG(ByVal x As Double) As String
    NumG = Trim$(Str$(x))
    If Left$(NumG, 1) = "." Then NumG = "0" & NumG
    If Left$(NumG, 2) = "-." Then NumG = "-0" & Mid$(NumG, 2)
End Function

Public Function FmtLag(ByVal m As Double) As String
    Dim k As Long
    k = Round(m)
    If k < 0 Then k = 0
    If k < 120 Then
        FmtLag = k & " 分鐘"
    ElseIf k < 2880 Then
        FmtLag = (k \ 60) & " 小時 " & (k Mod 60) & " 分"
    Else
        FmtLag = (k \ 1440) & " 天"
    End If
End Function

Public Function PathSep(ByVal p As String) As String
    If InStr(p, "/") > 0 And InStr(p, "\") = 0 Then PathSep = "/" Else PathSep = "\"
End Function

Public Function JoinPath(ByVal a As String, ByVal b As String) As String
    Dim s As String
    s = PathSep(a)
    If Right$(a, 1) = s Then JoinPath = a & b Else JoinPath = a & s & b
End Function

Public Function FileExists(ByVal p As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir(p)) > 0)
    If Err.Number <> 0 Then FileExists = False
    On Error GoTo 0
End Function

' 字串陣列排序（小資料量用插入排序）
Public Sub SortStrings(a() As String, ByVal n As Long)
    Dim i As Long, j As Long, v As String
    For i = 1 To n - 1
        v = a(i)
        j = i - 1
        Do While j >= 0
            If StrComp(a(j), v, vbBinaryCompare) <= 0 Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = v
    Next
End Sub

' 依兩個鍵排序索引（快速排序，鍵相同時保持原順序：最後用原索引比較）
Public Sub SortIdx2(idx() As Long, k1() As Double, k2() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Long, t As Long
    Do While hi - lo > 12
        p = idx((lo + hi) \ 2)
        i = lo: j = hi
        Do While i <= j
            Do While Less2(idx(i), p, k1, k2)
                i = i + 1
            Loop
            Do While Less2(p, idx(j), k1, k2)
                j = j - 1
            Loop
            If i <= j Then
                t = idx(i): idx(i) = idx(j): idx(j) = t
                i = i + 1: j = j - 1
            End If
        Loop
        If j - lo < hi - i Then
            SortIdx2 idx, k1, k2, lo, j
            lo = i
        Else
            SortIdx2 idx, k1, k2, i, hi
            hi = j
        End If
    Loop
    For i = lo + 1 To hi
        t = idx(i)
        j = i - 1
        Do While j >= lo
            If Not Less2(t, idx(j), k1, k2) Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = t
    Next
End Sub

Private Function Less2(ByVal a As Long, ByVal b As Long, k1() As Double, k2() As Double) As Boolean
    If k1(a) <> k1(b) Then
        Less2 = k1(a) < k1(b)
    ElseIf k2(a) <> k2(b) Then
        Less2 = k2(a) < k2(b)
    Else
        Less2 = a < b
    End If
End Function

' ============================================================
' 1 分K 資料：讀 history、接即時
' ============================================================
Public Sub ResetUnits()
    gUN = 0: gRN = 0: gHistLast = 0: gHistFiles = 0
    ReDim gUName(0 To 63): ReDim gUDay(0 To 63): ReDim gULive(0 To 63): ReDim gUA(0 To 63): ReDim gUB(0 To 63)
    ReDim gRT(0 To 65535): ReDim gRC(0 To 65535)
    ReDim gRO(0 To 65535): ReDim gRH(0 To 65535): ReDim gRL(0 To 65535): ReDim gRCl(0 To 65535): ReDim gRV(0 To 65535)
End Sub

Public Sub AddUnit(ByVal nm As String, ByVal dn As Long, ByVal live As Boolean)
    If gUN > UBound(gUName) Then
        ReDim Preserve gUName(0 To 2 * gUN): ReDim Preserve gUDay(0 To 2 * gUN): ReDim Preserve gULive(0 To 2 * gUN)
        ReDim Preserve gUA(0 To 2 * gUN): ReDim Preserve gUB(0 To 2 * gUN)
    End If
    gUName(gUN) = nm: gUDay(gUN) = dn: gULive(gUN) = live
    gUA(gUN) = gRN: gUB(gUN) = gRN - 1
    gUN = gUN + 1
End Sub

Public Sub AddRow(ByVal tm As Long, ByVal c As Long, ByVal o As Double, ByVal h As Double, ByVal l As Double, ByVal cl As Double, ByVal v As Double)
    Dim n As Long
    If gRN > UBound(gRT) Then
        n = 2 * gRN
        ReDim Preserve gRT(0 To n): ReDim Preserve gRC(0 To n)
        ReDim Preserve gRO(0 To n): ReDim Preserve gRH(0 To n): ReDim Preserve gRL(0 To n)
        ReDim Preserve gRCl(0 To n): ReDim Preserve gRV(0 To n)
    End If
    gRT(gRN) = tm: gRC(gRN) = c: gRO(gRN) = o: gRH(gRN) = h: gRL(gRN) = l: gRCl(gRN) = cl: gRV(gRN) = v
    gRN = gRN + 1
    gUB(gUN - 1) = gRN - 1
End Sub

' "20261105" + "1500" → 分鐘序號
Public Function TmOfYmdHm(ByVal ymd As String, ByVal hm As String) As Long
    Dim dn As Long
    dn = DayNum(Val(Left$(ymd, 4)), Val(Mid$(ymd, 5, 2)), Val(Mid$(ymd, 7, 2)))
    hm = Right$("0000" & hm, 4)
    TmOfYmdHm = dn * 1440 + Val(Left$(hm, 2)) * 60 + Val(Right$(hm, 2))
End Function

' 列出資料夾裡的 Daily_YYYY_MM_DD.csv（依檔名排序）
Public Function ListDaily(ByVal folder As String, fl() As String) As Long
    Dim f As String, n As Long
    ReDim fl(0 To 255)
    On Error Resume Next
    f = Dir(JoinPath(folder, "Daily_*.csv"))
    If Err.Number <> 0 Then f = ""
    On Error GoTo 0
    Do While Len(f) > 0
        If Len(f) = 20 And InStr(f, "_live") = 0 Then
            If n > UBound(fl) Then
                ReDim Preserve fl(0 To 2 * n)
            End If
            fl(n) = f
            n = n + 1
        End If
        f = Dir()
    Loop
    If n > 0 Then SortStrings fl, n
    ListDaily = n
End Function

Public Function DayOfDailyName(ByVal nm As String) As Long
    DayOfDailyName = DayNum(Val(Mid$(nm, 7, 4)), Val(Mid$(nm, 12, 2)), Val(Mid$(nm, 15, 2)))
End Function

' 讀 history 的 1 分K（date,time,contract,open,high,low,close,volume）
Public Function LoadHistory(ByVal folder As String) As Long
    Dim fl() As String, n As Long, i As Long, f As Integer, ln As String, a() As String
    Dim tm As Long, mx(0 To 1) As Long
    n = ListDaily(folder, fl)
    For i = 0 To n - 1
        AddUnit Left$(fl(i), 16), DayOfDailyName(fl(i)), False
        f = FreeFile
        Open JoinPath(folder, fl(i)) For Input As #f
        If Not EOF(f) Then Line Input #f, ln
        Do While Not EOF(f)
            Line Input #f, ln
            If Len(ln) > 10 Then
                a = Split(ln, ",")
                If UBound(a) >= 7 Then
                    tm = TmOfYmdHm(a(0), a(1))
                    AddRow tm, CLng(Val(a(2))), Val(a(3)), Val(a(4)), Val(a(5)), Val(a(6)), Val(a(7))
                    If i >= n - 2 Then
                        If tm > mx(i - (n - 2)) Then mx(i - (n - 2)) = tm
                    End If
                End If
            End If
        Loop
        Close #f
    Next
    gHistFiles = n
    gHistLast = mx(0)
    If mx(1) > gHistLast Then gHistLast = mx(1)
    LoadHistory = n
End Function

' ------------------------------------------------------------
' 即時 1 分K（期交所 MIS getChartData1M 的 JSON），不靠 Scripting.Dictionary
' ------------------------------------------------------------
Public Sub ResetLive()
    Dim k As Long
    For k = 1 To 2
        gLiveHas(k) = False: gLiveStart(k) = 0: gLiveSym(k) = "": gLiveCName(k) = "": gLiveLast(k) = 0: gLiveRef(k) = 0
        gLiveDate(k) = "": gLiveLastMin(k) = 0: gLiveMinutes(k) = 0: gLiveErr(k) = ""
    Next
    gLN = 0
    ReDim gLS(0 To 4095): ReDim gLC(0 To 4095): ReDim gLT(0 To 4095)
    ReDim gLO(0 To 4095): ReDim gLH(0 To 4095): ReDim gLL(0 To 4095): ReDim gLCl(0 To 4095): ReDim gLV(0 To 4095)
End Sub

' JSON 物件文字裡 "key":"value" 的值
Public Function JField(ByVal s As String, ByVal key As String, Optional ByVal from As Long = 1) As String
    Dim p As Long, q As Long
    p = InStr(from, s, """" & key & """:")
    If p = 0 Then Exit Function
    p = p + Len(key) + 3
    Do While Mid$(s, p, 1) = " "
        p = p + 1
    Loop
    If Mid$(s, p, 1) = """" Then
        q = InStr(p + 1, s, """")
        If q > 0 Then JField = Mid$(s, p + 1, q - p - 1)
    Else
        q = p
        Do While q <= Len(s) And InStr(",}]", Mid$(s, q, 1)) = 0
            q = q + 1
        Loop
        JField = Trim$(Mid$(s, p, q - p))
        If JField = "null" Then JField = ""
    End If
End Function

' getQuoteList 的 QuoteList → 台指期各月契約，依成交量由大到小（同量保持原順序）
Public Function ParseQuoteList(ByVal s As String, syms() As String, cnames() As String) As Long
    Dim p As Long, q As Long, ob As String, sym As String, n As Long, i As Long, j As Long
    Dim vols() As Double, v As Double, tS As String, tC As String
    ReDim syms(0 To 31): ReDim cnames(0 To 31): ReDim vols(0 To 31)
    p = InStr(s, """QuoteList"":")
    If p = 0 Then Err.Raise vbObjectError + 101, , "報價清單格式不對"
    Do
        p = InStr(p, s, "{")
        If p = 0 Then Exit Do
        q = InStr(p, s, "}")
        If q = 0 Then Exit Do
        ob = Mid$(s, p, q - p + 1)
        sym = JField(ob, "SymbolID")
        If Left$(sym, 3) = "TXF" And Len(sym) > 4 Then
            If InStr("ABCDEFGHIJKL", Mid$(sym, 4, 1)) > 0 Then
                If n > UBound(syms) Then
                    ReDim Preserve syms(0 To 2 * n)
                    ReDim Preserve cnames(0 To 2 * n)
                    ReDim Preserve vols(0 To 2 * n)
                End If
                syms(n) = sym: cnames(n) = JField(ob, "DispCName"): vols(n) = Val(JField(ob, "CTotalVolume"))
                n = n + 1
            End If
        End If
        p = q + 1
    Loop
    For i = 1 To n - 1      ' 穩定排序
        v = vols(i): tS = syms(i): tC = cnames(i)
        j = i - 1
        Do While j >= 0
            If vols(j) >= v Then Exit Do
            vols(j + 1) = vols(j): syms(j + 1) = syms(j): cnames(j + 1) = cnames(j)
            j = j - 1
        Loop
        vols(j + 1) = v: syms(j + 1) = tS: cnames(j + 1) = tC
    Next
    ParseQuoteList = n
End Function

' TXFK6-F → 202611
Public Function ContractCode(ByVal sym As String, ByVal yearNow As Long) As Long
    Dim y As Long
    y = (yearNow \ 10) * 10 + Val(Mid$(sym, 5, 1))
    If y < yearNow - 1 Then y = y + 10
    ContractCode = y * 100 + InStr("ABCDEFGHIJKL", Mid$(sym, 4, 1))
End Function

' 把一個契約的即時 1 分K 加進槽 slot。回傳 (分鐘數)。first = 這個盤別成交量最大的契約
Public Function ParseChart(ByVal s As String, ByVal slot As Long, ByVal contract As Long, ByVal isFirst As Boolean, _
                           ByVal sym As String, ByVal cname As String) As Long
    Dim qp As Long, cdt As String, bday As Long, p As Long, q As Long, inner As String, a() As String
    Dim hh As Long, mm As Long, dn As Long, ts As Long, k As Long, i0 As Long, cnt As Long, found As Long
    Dim o As Double, h As Double, l As Double, c As Double, v As Double, mn As Long, mxT As Long
    qp = InStr(s, """Quote"":")
    If qp = 0 Then Err.Raise vbObjectError + 102, , "1 分K 格式不對"
    cdt = JField(s, "CDate", qp)
    bday = DayNum(Val(Left$(cdt, 4)), Val(Mid$(cdt, 5, 2)), Val(Mid$(cdt, 7, 2)))
    i0 = gLN
    p = InStr(s, """Ticks"":")
    If p > 0 Then
        p = InStr(p, s, "[")
        Do
            p = InStr(p + 1, s, "[")
            If p = 0 Then Exit Do
            q = InStr(p, s, "]")
            If q = 0 Then Exit Do
            inner = Replace(Mid$(s, p + 1, q - p - 1), """", "")
            a = Split(inner, ",")
            If UBound(a) >= 5 Then
                hh = Val(Left$(a(0), 2)): mm = Val(Mid$(a(0), 3, 2))
                dn = bday
                If slot = 2 And hh < 15 Then dn = bday + 1
                ts = dn * 1440 + hh * 60 + mm - 1           ' MIS 標的是這分鐘的結束時間
                o = Val(a(1)): h = Val(a(2)): l = Val(a(3)): c = Val(a(4)): v = Int(Val(a(5)))
                found = -1
                For k = gLN - 1 To i0 Step -1
                    If gLT(k) = ts Then found = k: Exit For
                    If gLT(k) < ts Then Exit For
                Next
                If found < 0 Then
                    For k = i0 To gLN - 1
                        If gLT(k) = ts Then found = k: Exit For
                    Next
                End If
                If found >= 0 Then
                    gLH(found) = MaxD(gLH(found), h): gLL(found) = MinD(gLL(found), l): gLCl(found) = c: gLV(found) = gLV(found) + v
                Else
                    If gLN > UBound(gLT) Then
                        k = 2 * gLN
                        ReDim Preserve gLS(0 To k): ReDim Preserve gLC(0 To k): ReDim Preserve gLT(0 To k)
                        ReDim Preserve gLO(0 To k): ReDim Preserve gLH(0 To k): ReDim Preserve gLL(0 To k)
                        ReDim Preserve gLCl(0 To k): ReDim Preserve gLV(0 To k)
                    End If
                    gLS(gLN) = slot: gLC(gLN) = contract: gLT(gLN) = ts
                    gLO(gLN) = o: gLH(gLN) = h: gLL(gLN) = l: gLCl(gLN) = c: gLV(gLN) = v
                    gLN = gLN + 1
                End If
            End If
            p = q
            If Mid$(s, q + 1, 1) <> "," Then Exit Do
        Loop
    End If
    cnt = gLN - i0
    If cnt > 0 Then
        mn = gLT(i0): mxT = gLT(i0)
        For k = i0 To gLN - 1
            If gLT(k) < mn Then mn = gLT(k)
            If gLT(k) > mxT Then mxT = gLT(k)
        Next
        If isFirst Then
            gLiveHas(slot) = True
            gLiveStart(slot) = mn
            gLiveSym(slot) = sym
            gLiveCName(slot) = cname
            gLiveLast(slot) = Val(JField(s, "CLastPrice", qp))
            gLiveRef(slot) = Val(JField(s, "CRefPrice", qp))
            gLiveDate(slot) = cdt
            gLiveLastMin(slot) = mxT
            gLiveMinutes(slot) = cnt
        End If
    End If
    ParseChart = cnt
End Function

' 把 history 還沒有的即時分K，依盤別各當成一個單位接在最後（同 live.py merged_dir）
Public Sub AddLiveUnits()
    Dim order(1 To 2) As Long, k As Long, slot As Long, i As Long, j As Long, m As Long
    Dim idx() As Long, k1() As Double, k2() As Double, stem As String
    order(1) = 1: order(2) = 2
    If gLiveHas(1) And gLiveHas(2) Then
        If gLiveStart(2) < gLiveStart(1) Then order(1) = 2: order(2) = 1
    End If
    If gUN > 0 Then stem = gUName(gUN - 1) Else stem = "Daily_live"
    m = 0
    For k = 1 To 2
        slot = order(k)
        If gLiveHas(slot) Then
            m = m + 1
            ReDim idx(0 To gLN): ReDim k1(0 To gLN): ReDim k2(0 To gLN)
            j = 0
            For i = 0 To gLN - 1
                If gLS(i) = slot And gLT(i) > gHistLast Then
                    idx(j) = i: k1(i) = gLC(i): k2(i) = gLT(i)
                    j = j + 1
                End If
            Next
            If j > 0 Then
                SortIdx2 idx, k1, k2, 0, j - 1
                AddUnit stem & "_live" & m, 0, True
                For i = 0 To j - 1
                    AddRow gLT(idx(i)), gLC(idx(i)), gLO(idx(i)), gLH(idx(i)), gLL(idx(i)), gLCl(idx(i)), gLV(idx(i))
                Next
            End If
        End If
    Next
End Sub

' 資料缺口：history 最後一分鐘到即時第一分鐘之間、即時最後一分鐘到現在之間，平日 08:45 / 15:00 開始的盤
Public Sub FindGap(ByVal nowTm As Long)
    Dim i As Long, mn As Long, mx As Long, has As Boolean
    gGapN = 0
    ReDim gGap(0 To 63)
    For i = 0 To gLN - 1
        If gLT(i) > gHistLast And gLiveHas(gLS(i)) Then
            If Not has Then mn = gLT(i): mx = gLT(i): has = True
            If gLT(i) < mn Then mn = gLT(i)
            If gLT(i) > mx Then mx = gLT(i)
        End If
    Next
    If Not has Then
        ExpectedGap gHistLast, nowTm
    Else
        ExpectedGap gHistLast, mn
        ExpectedGap mx, nowTm
    End If
End Sub

Private Sub ExpectedGap(ByVal a As Long, ByVal b As Long)
    Dim d As Long, st As Long, k As Long
    If a <= 0 Then Exit Sub
    For d = a \ 1440 To b \ 1440
        If Wd(d) < 5 Then
            For k = 0 To 1
                st = d * 1440 + IIf(k = 0, T_DAYOPEN, T_NIGHT)
                If a < st And st <= b - 5 Then
                    If gGapN > UBound(gGap) Then
                        ReDim Preserve gGap(0 To 2 * gGapN)
                    End If
                    gGap(gGapN) = st
                    gGapN = gGapN + 1
                End If
            Next
        End If
    Next
End Sub

' ============================================================
' K 棒合成（build_bars.build）：每個單位用成交量最大的契約當近月，換月回溯調整
' ============================================================
Public Sub UnitFronts(front() As Long, off() As Double)
    ' 每個單位：各契約的總量與最後收盤（檔案順序的最後一筆），成交量最大的是近月；換月時舊資料加上新舊月價差
    Const MAXC As Long = 40
    Dim u As Long, k As Long, i As Long, best As Long, nc() As Long
    Dim cs() As Long, vs() As Double, ls() As Double, a As Double, b As Double, fa As Boolean, fb As Boolean
    If gUN = 0 Then Exit Sub
    ReDim front(0 To gUN - 1): ReDim off(0 To gUN - 1): ReDim nc(0 To gUN - 1)
    ReDim cs(0 To gUN - 1, 0 To MAXC): ReDim vs(0 To gUN - 1, 0 To MAXC): ReDim ls(0 To gUN - 1, 0 To MAXC)
    For u = 0 To gUN - 1
        For i = gUA(u) To gUB(u)
            For k = 0 To nc(u) - 1
                If cs(u, k) = gRC(i) Then Exit For
            Next
            If k = nc(u) And k < MAXC Then
                cs(u, k) = gRC(i)
                nc(u) = nc(u) + 1
            End If
            vs(u, k) = vs(u, k) + gRV(i)
            ls(u, k) = gRCl(i)
        Next
        best = 0
        For k = 1 To nc(u) - 1
            If vs(u, k) > vs(u, best) Then best = k
        Next
        front(u) = cs(u, best)
    Next
    For u = gUN - 1 To 1 Step -1
        off(u - 1) = off(u)
        If front(u) <> front(u - 1) Then
            fa = False: fb = False
            For k = 0 To nc(u - 1) - 1
                If cs(u - 1, k) = front(u) Then
                    a = ls(u - 1, k)
                    fa = True
                End If
                If cs(u - 1, k) = front(u - 1) Then
                    b = ls(u - 1, k)
                    fb = True
                End If
            Next
            If fa And fb Then off(u - 1) = off(u - 1) + a - b
        End If
    Next
End Sub

Public Function BuildBars(ByVal minutes As Long, bT() As Long, bO() As Double, bH() As Double, bL() As Double, _
                          bC() As Double, bV() As Double) As Long
    Dim front() As Long, off() As Double, u As Long, i As Long, n As Long, key As Long
    Dim t As Long, d As Long, st As Long, lastK As Long, k As Long, pos As Long, lo As Long, hi As Long, md As Long
    Dim o As Double, h As Double, l As Double, c As Double, j As Long
    n = 0
    ReDim bT(0 To 1023): ReDim bO(0 To 1023): ReDim bH(0 To 1023): ReDim bL(0 To 1023): ReDim bC(0 To 1023): ReDim bV(0 To 1023)
    If gUN = 0 Then BuildBars = 0: Exit Function
    UnitFronts front, off
    pos = -1
    For u = 0 To gUN - 1
        For i = gUA(u) To gUB(u)
            If gRC(i) = front(u) Then
                t = MinOfTm(gRT(i)): d = DayOfTm(gRT(i))
                If t >= T_DAYOPEN And t <= T_DAYCLOSE Then
                    st = d * 1440 + T_DAYOPEN: lastK = 300 \ minutes - 1
                Else
                    If t >= T_NIGHT Then st = d * 1440 + T_NIGHT Else st = (d - 1) * 1440 + T_NIGHT
                    lastK = 840 \ minutes - 1
                End If
                k = (gRT(i) - st) \ minutes
                If k > lastK Then k = lastK
                key = st + k * minutes
                o = gRO(i) + off(u): h = gRH(i) + off(u): l = gRL(i) + off(u): c = gRCl(i) + off(u)
                ' 跟 Python dict 一樣：第一次出現的分鐘定開盤，之後更新高低收量
                If pos >= 0 Then
                    If bT(pos) <> key Then pos = -1
                End If
                If pos < 0 Then
                    If n = 0 Then
                        lo = 0
                    ElseIf key > bT(n - 1) Then
                        lo = n
                    Else
                        lo = 0: hi = n - 1
                        Do While lo <= hi
                            md = (lo + hi) \ 2
                            If bT(md) = key Then pos = md: Exit Do
                            If bT(md) < key Then lo = md + 1 Else hi = md - 1
                        Loop
                    End If
                    If pos < 0 Then
                        If n > UBound(bT) Then
                            GrowBars bT, bO, bH, bL, bC, bV, 2 * n
                        End If
                        For j = n To lo + 1 Step -1
                            bT(j) = bT(j - 1): bO(j) = bO(j - 1): bH(j) = bH(j - 1): bL(j) = bL(j - 1): bC(j) = bC(j - 1): bV(j) = bV(j - 1)
                        Next
                        bT(lo) = key: bO(lo) = o: bH(lo) = h: bL(lo) = l: bC(lo) = c: bV(lo) = gRV(i)
                        n = n + 1
                        pos = lo
                        GoTo nextRow
                    End If
                End If
                If h > bH(pos) Then bH(pos) = h
                If l < bL(pos) Then bL(pos) = l
                bC(pos) = c
                bV(pos) = bV(pos) + gRV(i)
            End If
nextRow:
        Next
    Next
    BuildBars = n
End Function

Private Sub GrowBars(bT() As Long, bO() As Double, bH() As Double, bL() As Double, bC() As Double, bV() As Double, ByVal m As Long)
    ReDim Preserve bT(0 To m): ReDim Preserve bO(0 To m): ReDim Preserve bH(0 To m)
    ReDim Preserve bL(0 To m): ReDim Preserve bC(0 To m): ReDim Preserve bV(0 To m)
End Sub

' ============================================================
' D2 日K（只用日盤）：期交所每日行情 + 1 分K 合成（同 live.py tx_daily_bars, night=False）
' ============================================================
Public Function Utf8Text(b() As Byte) As String
    Dim i As Long, k As Long, c As Long, hi As Long, buf As String, ub As Long
    ub = UBound(b)
    If ub < 0 Then Exit Function
    buf = Space$(ub + 1)
    i = 0
    If ub >= 2 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then i = 3
    End If
    Do While i <= ub
        c = b(i)
        If c < &H80 Then
            i = i + 1
        ElseIf c < &HE0 And i + 1 <= ub Then
            c = ((c And &H1F) * &H40&) Or (b(i + 1) And &H3F)
            i = i + 2
        ElseIf c < &HF0 And i + 2 <= ub Then
            c = ((c And &HF) * &H1000&) Or ((b(i + 1) And &H3F) * &H40&) Or (b(i + 2) And &H3F)
            i = i + 3
        ElseIf i + 3 <= ub Then
            c = ((c And &H7) * &H40000) Or ((b(i + 1) And &H3F) * &H1000&) Or ((b(i + 2) And &H3F) * &H40&) Or (b(i + 3) And &H3F)
            i = i + 4
        Else
            c = &HFFFD&
            i = i + 1
        End If
        If c >= &H10000 Then
            c = c - &H10000
            hi = &HD800& + (c \ &H400&)
            k = k + 1: Mid$(buf, k, 1) = ChrW(hi)
            k = k + 1: Mid$(buf, k, 1) = ChrW(&HDC00& + (c And &H3FF&))
        Else
            k = k + 1: Mid$(buf, k, 1) = ChrW(c)
        End If
    Loop
    Utf8Text = Left$(buf, k)
End Function

Public Function ReadUtf8File(ByVal p As String) As String
    Dim f As Integer, b() As Byte, n As Long
    If Not FileExists(p) Then Exit Function
    f = FreeFile
    Open p For Binary Access Read As #f
    n = LOF(f)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #f, , b
    End If
    Close #f
    If n > 0 Then ReadUtf8File = Utf8Text(b)
End Function

Public Function SesDay() As String
    SesDay = ChrW(&H65E5) & ChrW(&H76E4)    ' 日盤
End Function

Public Function SesNight() As String
    SesNight = ChrW(&H591C) & ChrW(&H76E4)  ' 夜盤
End Function

' 回傳根數。dT 是日期 00:00 的分鐘序號。i0 = 第一個有 1 分K 的日子。mT.. = 調整後的近月日盤 1 分K
Public Function TxDailyBars(ByVal dailyCsv As String, dT() As Long, dOpn() As Double, dH() As Double, dL() As Double, dC() As Double, _
                            i0 As Long, mT() As Long, mO() As Double, mH() As Double, mL() As Double, mC() As Double, mN As Long) As Long
    Dim i As Long, j As Long, k As Long, u As Long, n As Long, t As Long, td As Long
    Dim idx() As Long, k1() As Double, k2() As Double, rTd() As Long
    ' --- 1 分K 的日盤（依交易日、契約、時間排序）
    ReDim idx(0 To gRN): ReDim k1(0 To gRN): ReDim k2(0 To gRN): ReDim rTd(0 To gRN)
    n = 0
    For u = 0 To gUN - 1
        For i = gUA(u) To gUB(u)
            t = MinOfTm(gRT(i))
            If t >= T_DAYOPEN And t <= T_DAYCLOSE Then
                If gULive(u) Then td = DayOfTm(gRT(i)) Else td = gUDay(u)
                rTd(i) = td
                idx(n) = i: k1(i) = CDbl(td) * 1000000# + gRC(i): k2(i) = gRT(i)
                n = n + 1
            End If
        Next
    Next
    If n > 1 Then SortIdx2 idx, k1, k2, 0, n - 1
    ' 每日每契約一筆：pD 日期、pC 契約、OHLCV、pRaw 是否來自 1 分K
    Dim pN As Long, pD() As Long, pC() As Long, pO() As Double, pH() As Double, pL() As Double, pCl() As Double, pV() As Double
    Dim pA() As Long, pB() As Long, pRaw() As Boolean
    ReDim pD(0 To 1023): ReDim pC(0 To 1023): ReDim pO(0 To 1023): ReDim pH(0 To 1023): ReDim pL(0 To 1023)
    ReDim pCl(0 To 1023): ReDim pV(0 To 1023): ReDim pA(0 To 1023): ReDim pB(0 To 1023): ReDim pRaw(0 To 1023)
    For j = 0 To n - 1
        i = idx(j)
        If pN > 0 Then
            If pD(pN - 1) = rTd(i) And pC(pN - 1) = gRC(i) And pRaw(pN - 1) Then
                k = pN - 1
                If gRH(i) > pH(k) Then pH(k) = gRH(i)
                If gRL(i) < pL(k) Then pL(k) = gRL(i)
                pCl(k) = gRCl(i): pV(k) = pV(k) + gRV(i): pB(k) = j
                GoTo nxt
            End If
        End If
        If pN > UBound(pD) Then
            GrowP pD, pC, pO, pH, pL, pCl, pV, pA, pB, pRaw, 2 * pN
        End If
        pD(pN) = rTd(i): pC(pN) = gRC(i): pO(pN) = gRO(i): pH(pN) = gRH(i): pL(pN) = gRL(i): pCl(pN) = gRCl(i)
        pV(pN) = gRV(i): pA(pN) = j: pB(pN) = j: pRaw(pN) = True
        pN = pN + 1
nxt:
    Next
    Dim rawN As Long, rawDays() As Long
    rawN = 0
    ReDim rawDays(0 To pN + 1)
    For k = 0 To pN - 1
        If rawN = 0 Then
            rawDays(0) = pD(k): rawN = 1
        ElseIf rawDays(rawN - 1) <> pD(k) Then
            rawDays(rawN) = pD(k): rawN = rawN + 1
        End If
    Next
    ' --- 期交所每日行情（只取日盤；有 1 分K 的日子不用）
    Dim txt As String, lines() As String, a() As String, dn As Long, sd As String
    txt = ReadUtf8File(dailyCsv)
    If Len(txt) > 0 Then
        sd = SesDay()
        lines = Split(Replace(txt, vbCr, ""), vbLf)
        For j = 1 To UBound(lines)
            If Len(lines(j)) > 10 Then
                a = Split(lines(j), ",")
                If UBound(a) >= 7 Then
                    If a(1) = sd Then
                        dn = DayNum(Val(Left$(a(0), 4)), Val(Mid$(a(0), 6, 2)), Val(Mid$(a(0), 9, 2)))
                        If Not InSorted(rawDays, rawN, dn) Then
                            If pN > UBound(pD) Then
                                GrowP pD, pC, pO, pH, pL, pCl, pV, pA, pB, pRaw, 2 * pN
                            End If
                            pD(pN) = dn: pC(pN) = CLng(Val(a(2))): pO(pN) = Val(a(3)): pH(pN) = Val(a(4)): pL(pN) = Val(a(5))
                            pCl(pN) = Val(a(6)): pV(pN) = Val(a(7)): pRaw(pN) = False
                            pN = pN + 1
                        End If
                    End If
                End If
            End If
        Next
    End If
    ' --- 依日期、契約排序
    Dim pidx() As Long, q1() As Double, q2() As Double
    ReDim pidx(0 To pN): ReDim q1(0 To pN): ReDim q2(0 To pN)
    For k = 0 To pN - 1
        pidx(k) = k: q1(k) = pD(k): q2(k) = pC(k)
    Next
    If pN > 1 Then SortIdx2 pidx, q1, q2, 0, pN - 1
    ' --- 每天的近月（成交量最大）
    Dim dn2 As Long, fr() As Long, frIdx() As Long, off() As Double, dFirst() As Long, dLast() As Long, best As Long
    ReDim dT(0 To pN): ReDim fr(0 To pN): ReDim frIdx(0 To pN): ReDim dFirst(0 To pN): ReDim dLast(0 To pN)
    dn2 = 0
    For j = 0 To pN - 1
        k = pidx(j)
        If dn2 = 0 Then
            dT(0) = pD(k): dFirst(0) = j: dn2 = 1
        ElseIf dT(dn2 - 1) <> pD(k) Then
            dLast(dn2 - 1) = j - 1
            dT(dn2) = pD(k): dFirst(dn2) = j: dn2 = dn2 + 1
        End If
    Next
    If dn2 > 0 Then dLast(dn2 - 1) = pN - 1
    For i = 0 To dn2 - 1
        best = pidx(dFirst(i))
        For j = dFirst(i) + 1 To dLast(i)
            If pV(pidx(j)) > pV(best) Then best = pidx(j)
        Next
        fr(i) = pC(best): frIdx(i) = best
    Next
    ReDim off(0 To dn2)
    For i = dn2 - 1 To 1 Step -1
        off(i - 1) = off(i)
        If fr(i) <> fr(i - 1) Then
            For j = dFirst(i - 1) To dLast(i - 1)
                If pC(pidx(j)) = fr(i) Then
                    off(i - 1) = off(i - 1) + pCl(pidx(j)) - pCl(frIdx(i - 1))
                    Exit For
                End If
            Next
        End If
    Next
    ReDim dOpn(0 To dn2): ReDim dH(0 To dn2): ReDim dL(0 To dn2): ReDim dC(0 To dn2)
    i0 = -1
    mN = 0
    ReDim mT(0 To n): ReDim mO(0 To n): ReDim mH(0 To n): ReDim mL(0 To n): ReDim mC(0 To n)
    For i = 0 To dn2 - 1
        k = frIdx(i)
        dOpn(i) = pO(k) + off(i): dH(i) = pH(k) + off(i): dL(i) = pL(k) + off(i): dC(i) = pCl(k) + off(i)
        If pRaw(k) Then
            If i0 < 0 Then i0 = i
            For j = pA(k) To pB(k)
                u = idx(j)
                mT(mN) = gRT(u): mO(mN) = gRO(u) + off(i): mH(mN) = gRH(u) + off(i): mL(mN) = gRL(u) + off(i): mC(mN) = gRCl(u) + off(i)
                mN = mN + 1
            Next
        End If
        dT(i) = dT(i) * 1440
    Next
    TxDailyBars = dn2
End Function

Private Function InSorted(a() As Long, ByVal n As Long, ByVal v As Long) As Boolean
    Dim lo As Long, hi As Long, m As Long
    lo = 0: hi = n - 1
    Do While lo <= hi
        m = (lo + hi) \ 2
        If a(m) = v Then InSorted = True: Exit Function
        If a(m) < v Then lo = m + 1 Else hi = m - 1
    Loop
End Function

Private Sub GrowP(pD() As Long, pC() As Long, pO() As Double, pH() As Double, pL() As Double, pCl() As Double, pV() As Double, _
                  pA() As Long, pB() As Long, pRaw() As Boolean, ByVal m As Long)
    ReDim Preserve pD(0 To m): ReDim Preserve pC(0 To m): ReDim Preserve pO(0 To m): ReDim Preserve pH(0 To m)
    ReDim Preserve pL(0 To m): ReDim Preserve pCl(0 To m): ReDim Preserve pV(0 To m): ReDim Preserve pA(0 To m)
    ReDim Preserve pB(0 To m): ReDim Preserve pRaw(0 To m)
End Sub

' ============================================================
' 加權指數 60 分K（回測用）：start,open,high,low,close,volume，跳過 13:30 那根
' ============================================================
Public Function LoadTwii(ByVal p As String, tT() As Long, tOpn() As Double, tH() As Double, tL() As Double, tC() As Double) As Long
    Dim f As Integer, ln As String, a() As String, n As Long, s As String
    ReDim tT(0 To 8191): ReDim tOpn(0 To 8191): ReDim tH(0 To 8191): ReDim tL(0 To 8191): ReDim tC(0 To 8191)
    If Not FileExists(p) Then LoadTwii = 0: Exit Function
    f = FreeFile
    Open p For Input As #f
    If Not EOF(f) Then Line Input #f, ln
    Do While Not EOF(f)
        Line Input #f, ln
        a = Split(ln, ",")
        If UBound(a) >= 4 Then
            s = a(0)
            If Right$(s, 5) <> "13:30" And Len(s) >= 16 Then
                If n > UBound(tT) Then
                    ReDim Preserve tT(0 To 2 * n): ReDim Preserve tOpn(0 To 2 * n): ReDim Preserve tH(0 To 2 * n)
                    ReDim Preserve tL(0 To 2 * n): ReDim Preserve tC(0 To 2 * n)
                End If
                tT(n) = DayNum(Val(Left$(s, 4)), Val(Mid$(s, 6, 2)), Val(Mid$(s, 9, 2))) * 1440& + Val(Mid$(s, 12, 2)) * 60 + Val(Mid$(s, 15, 2))
                tOpn(n) = Val(a(1)): tH(n) = Val(a(2)): tL(n) = Val(a(3)): tC(n) = Val(a(4))
                n = n + 1
            End If
        End If
    Loop
    Close #f
    LoadTwii = n
End Function

' 加權指數 60 分K → 日K
Public Function TwiiDaily(ByVal n As Long, tT() As Long, tOpn() As Double, tH() As Double, tL() As Double, tC() As Double, _
                          dT() As Long, dOpn() As Double, dH() As Double, dL() As Double, dC() As Double) As Long
    Dim i As Long, m As Long, d As Long
    ReDim dT(0 To n): ReDim dOpn(0 To n): ReDim dH(0 To n): ReDim dL(0 To n): ReDim dC(0 To n)
    m = 0
    For i = 0 To n - 1
        d = DayOfTm(tT(i)) * 1440
        If m > 0 Then
            If dT(m - 1) = d Then
                If tH(i) > dH(m - 1) Then dH(m - 1) = tH(i)
                If tL(i) < dL(m - 1) Then dL(m - 1) = tL(i)
                dC(m - 1) = tC(i)
                GoTo nx
            End If
        End If
        dT(m) = d: dOpn(m) = tOpn(i): dH(m) = tH(i): dL(m) = tL(i): dC(m) = tC(i)
        m = m + 1
nx:
    Next
    TwiiDaily = m
End Function

' ============================================================
' 指標（backtest.indicators）
' ============================================================
Private Sub EmaArr(ByVal n As Long, x() As Double, ByVal p As Long, out() As Double)
    Dim k As Double, i As Long
    ReDim out(0 To n)
    If n = 0 Then Exit Sub
    k = 2 / (p + 1)
    out(0) = x(0)
    For i = 1 To n - 1
        out(i) = out(i - 1) + k * (x(i) - out(i - 1))
    Next
End Sub

Public Sub CalcInd(ByVal n As Long, H() As Double, L() As Double, C() As Double, _
                   E20() As Double, E60() As Double, HI() As Double, BBM() As Double, BBS() As Double, ATR() As Double)
    Dim e12() As Double, e26() As Double, macd() As Double, sg() As Double, i As Long, j As Long
    Dim s As Double, m As Double, v As Double, tr As Double, a As Double
    EmaArr n, C, 20, E20
    EmaArr n, C, 60, E60
    EmaArr n, C, 12, e12
    EmaArr n, C, 26, e26
    ReDim macd(0 To n)
    For i = 0 To n - 1
        macd(i) = e12(i) - e26(i)
    Next
    EmaArr n, macd, 9, sg
    ReDim HI(0 To n)
    For i = 0 To n - 1
        HI(i) = macd(i) - sg(i)
    Next
    ReDim BBM(0 To n): ReDim BBS(0 To n)
    For i = 19 To n - 1
        s = 0
        For j = i - 19 To i
            s = s + C(j)
        Next
        m = s / 20
        v = 0
        For j = i - 19 To i
            v = v + (C(j) - m) * (C(j) - m)
        Next
        BBM(i) = m
        BBS(i) = Sqr(v / 20)
    Next
    ReDim ATR(0 To n)
    For i = 0 To n - 1
        If i = 0 Then
            tr = H(i) - L(i)
            a = tr
        Else
            tr = MaxD(H(i) - L(i), MaxD(Abs(H(i) - C(i - 1)), Abs(L(i) - C(i - 1))))
            a = a + (tr - a) / 14
        End If
        ATR(i) = a
    Next
End Sub

' ============================================================
' 交易紀錄容器
' ============================================================
Public Sub TrInit(tr() As Variant)
    ReDim tr(0 To F_LAST, 0 To 15)
End Sub

Public Sub TrAdd(tr() As Variant, n As Long, ByVal t As Long, ByVal side As Long, ByVal kind As String, ByVal R As Double, _
                 ByVal rr As Double, ByVal bars As Long, ByVal entry As Double, ByVal entryT As Long, ByVal stop0 As Double, _
                 ByVal stp As Double, ByVal exitT As Long, ByVal exitP As Double, ByVal isOpen As Boolean, ByVal be As Boolean, ByVal why As String)
    If n > UBound(tr, 2) Then
        ReDim Preserve tr(0 To F_LAST, 0 To 2 * n + 1)
    End If
    tr(F_T, n) = t: tr(F_SIDE, n) = side: tr(F_KIND, n) = kind: tr(F_RISK, n) = R: tr(F_RMUL, n) = rr
    tr(F_BARS, n) = bars: tr(F_ENTRY, n) = entry: tr(F_ENTRYT, n) = entryT: tr(F_STOP0, n) = stop0
    tr(F_STOP, n) = stp: tr(F_EXITT, n) = exitT: tr(F_EXIT, n) = exitP: tr(F_OPEN, n) = isOpen: tr(F_BE, n) = be
    tr(F_WHY, n) = why
    n = n + 1
End Sub

' ============================================================
' v3（60 分K）：v3.run 搭配 V3 參數
' ============================================================
Public Function RunV3(ByVal n As Long, T() As Long, O() As Double, H() As Double, L() As Double, C() As Double, _
                      E20() As Double, E60() As Double, HI() As Double, BBM() As Double, BBS() As Double, ATR() As Double, _
                      ByVal mode As Long, tr() As Variant) As Long
    Const BODY_ATR As Double = 0.8, STOP_ATR As Double = 1.5, BE_R As Double = 1#, TRAIL As Double = 5#
    Dim i As Long, j As Long, s As Long, sig As Long, kind As String, side As Long, cnt As Long
    Dim body As Double, rng As Double, e As Double, ext As Double, raw As Double, stp As Double, R As Double
    Dim pnl As Double, be As Boolean, best As Double, stop0 As Double, xt As Long, xp As Double, stillOpen As Boolean
    Dim worst As Double, px As Double, trail2 As Double, ok As Boolean
    TrInit tr
    cnt = 0
    i = 100
    Do While i < n - 1
        sig = 0
        If Allowed(T(i), mode) Then
            For s = 1 To -1 Step -2
                If (E20(i) - E60(i)) * s > 0 Then
                    body = (C(i) - O(i)) * s: rng = H(i) - L(i)
                    ok = False
                    If body > 0 And body >= 0.5 * rng Then
                        If body >= BODY_ATR * ATR(i - 1) Then
                            If (C(i) - (BBM(i) + 2 * BBS(i) * s)) * s > 0 Then
                                If HI(i) * s > 0 Then ok = True
                            End If
                        End If
                    End If
                    If ok Then sig = s: kind = "突破": Exit For
                    If body > 0 And body >= 0.5 * rng Then
                        If (Sel(s > 0, L(i), H(i)) - E20(i)) * s <= 0 Then
                            If (C(i) - E20(i)) * s > 0 Then
                                If (C(i) - Sel(s > 0, H(i - 1), L(i - 1))) * s > 0 Then
                                    If (E60(i) - E60(i - 5)) * s > 0 Then ok = True
                                End If
                            End If
                        End If
                    End If
                    If ok Then sig = s: kind = "拉回": Exit For
                End If
            Next
        End If
        If sig = 0 Then
            i = i + 1
        Else
            side = sig
            e = O(i + 1)
            ext = Sel(side > 0, L(i), H(i))
            raw = Abs(e - (ext - 0.1 * ATR(i) * side))
            If raw > (STOP_ATR + 0.5) * ATR(i) Then
                i = i + 1
            Else
                stp = e - MaxD(MinD(raw, STOP_ATR * ATR(i)), 0.5 * ATR(i)) * side
                R = Abs(e - stp)
                pnl = 0: be = False: j = i + 1: best = e
                stop0 = stp: xt = 0: xp = 0: stillOpen = False
                Do While j < n
                    worst = Sel(side > 0, L(j), H(j))
                    If (worst - stp) * side <= 0 Then
                        px = Sel((O(j) - stp) * side > 0, stp, O(j))
                        pnl = (px - e) * side: xt = T(j): xp = px
                        Exit Do
                    End If
                    If Not be Then
                        If (C(j) - e) * side >= BE_R * R Then be = True: stp = e + COST * side
                    End If
                    If j + 1 >= n Then
                        pnl = (C(j) - e) * side: xt = T(j): xp = C(j): stillOpen = True
                        Exit Do
                    End If
                    If be Then
                        best = MaxD(best * side, C(j) * side) * side
                        trail2 = best - TRAIL * ATR(j) * side
                        If (trail2 - stp) * side > 0 Then stp = trail2
                    End If
                    j = j + 1
                Loop
                TrAdd tr, cnt, T(i), side, kind, R, (pnl - COST) / R, j - i, e, T(i + 1), stop0, stp, xt, xp, stillOpen, be, ""
                i = j
            End If
        End If
    Loop
    RunV3 = cnt
End Function

' ============================================================
' D2（日K）：daily.run 搭配 D2 參數（收盤創 20 日新高/新低、實體 ≥ 高低差一半、EMA20/60 趨勢、收盤進場）
' sigFrom < 0 表示不限制
' ============================================================
Public Function RunD2(ByVal n As Long, T() As Long, O() As Double, H() As Double, L() As Double, C() As Double, _
                      E20() As Double, E60() As Double, ATR() As Double, ByVal sigFrom As Long, tr() As Variant) As Long
    Const NN As Long = 20, STOP_ATR As Double = 1#, BE_R As Double = 1#
    Dim i As Long, j As Long, s As Long, sig As Long, kind As String, side As Long, cnt As Long, k As Long
    Dim body As Double, rng As Double, e As Double, ext As Double, raw As Double, stp As Double, R As Double
    Dim pnl As Double, be As Boolean, stop0 As Double, xt As Long, xp As Double, stillOpen As Boolean, why As String
    Dim worst As Double, px As Double, gap As Boolean, mx As Double, mn As Double, brk As Boolean
    TrInit tr
    cnt = 0
    i = 60
    Do While i < n - 1
        sig = 0
        If sigFrom < 0 Or i >= sigFrom Then
            For s = 1 To -1 Step -2
                If (E20(i) - E60(i)) * s > 0 Then
                    body = (C(i) - O(i)) * s: rng = H(i) - L(i)
                    mx = C(i - NN): mn = C(i - NN)
                    For k = i - NN + 1 To i - 1
                        If C(k) > mx Then mx = C(k)
                        If C(k) < mn Then mn = C(k)
                    Next
                    If s > 0 Then brk = (C(i) > mx) Else brk = (C(i) < mn)
                    If body > 0 And body >= 0.5 * rng Then
                        If body >= 0# * ATR(i - 1) And brk Then sig = s: kind = "突破": Exit For
                    End If
                End If
            Next
        End If
        If sig = 0 Then
            i = i + 1
        Else
            side = sig
            e = C(i)
            e = e + 0 * ATR(i) * side
            ext = Sel(side > 0, L(i), H(i))
            raw = Abs(e - (ext - 0.1 * ATR(i) * side))
            If raw > (STOP_ATR + 0.5) * ATR(i) Then
                i = i + 1
            Else
                stp = e - MaxD(MinD(raw, STOP_ATR * ATR(i)), 0.5 * ATR(i)) * side
                R = Abs(e - stp)
                pnl = 0: be = False: j = i + 1
                stop0 = stp: xt = 0: xp = 0: stillOpen = False: why = ""
                Do While j < n
                    worst = Sel(side > 0, L(j), H(j))
                    If (worst - stp) * side <= 0 Then
                        gap = ((O(j) - stp) * side <= 0)
                        px = Sel(gap, O(j), stp)
                        pnl = (px - e) * side: xt = T(j): xp = px
                        If (stp - e) * side > COST Then
                            why = "移動停利"
                        ElseIf be Then
                            why = "保本停損"
                        Else
                            why = "初始停損"
                        End If
                        If gap Then why = why & "（跳空成交）"
                        Exit Do
                    End If
                    If Not be Then
                        If (C(j) - e) * side >= BE_R * R Then be = True: stp = e + COST * side
                    End If
                    If j + 1 >= n Then
                        pnl = (C(j) - e) * side: xt = T(j): xp = C(j): stillOpen = True: why = "持倉中"
                        Exit Do
                    End If
                    If be Then
                        If (C(j) - E20(j)) * side < 0 Then
                            pnl = (O(j + 1) - e) * side: xt = T(j + 1): xp = O(j + 1): why = "跌破EMA20"
                            j = j + 1
                            Exit Do
                        End If
                    End If
                    j = j + 1
                Loop
                TrAdd tr, cnt, T(i), side, kind, R, (pnl - COST) / R, j - i, e, T(i), stop0, stp, xt, xp, stillOpen, be, why
                i = j
            End If
        End If
    Loop
    RunD2 = cnt
End Function

' ============================================================
' 進出場時間精確到分鐘（live.refine_times）
' ============================================================
Private Function FirstTouch(mT() As Long, mH() As Double, mL() As Double, ByVal mN As Long, ByVal st As Long, ByVal en As Long, _
                            ByVal stp As Double, ByVal side As Long) As Long
    Dim lo As Long, hi As Long, m As Long, i As Long
    lo = 0: hi = mN
    Do While lo < hi
        m = (lo + hi) \ 2
        If mT(m) < st Then lo = m + 1 Else hi = m
    Loop
    i = lo
    Do While i < mN
        If mT(i) >= en Then Exit Do
        If (Sel(side > 0, mL(i), mH(i)) - stp) * side <= 0 Then FirstTouch = mT(i): Exit Function
        i = i + 1
    Loop
    FirstTouch = 0
End Function

Public Function ExitReasonV3(tr() As Variant, ByVal k As Long) As String
    Dim s As String
    If (tr(F_STOP, k) - tr(F_ENTRY, k)) * tr(F_SIDE, k) > COST Then
        s = "移動停利"
    ElseIf tr(F_BE, k) Then
        s = "保本停損"
    Else
        s = "初始停損"
    End If
    If Round(tr(F_EXIT, k)) <> Round(tr(F_STOP, k)) Then s = s & "（跳空成交）"
    ExitReasonV3 = s
End Function

Public Sub RefineV3(tr() As Variant, ByVal tn As Long, mT() As Long, mH() As Double, mL() As Double, ByVal mN As Long, ByVal barMin As Long)
    Dim k As Long, t As Long
    For k = 0 To tn - 1
        If tr(F_OPEN, k) Then
            tr(F_WHY, k) = ""
        Else
            tr(F_WHY, k) = ExitReasonV3(tr, k)
            t = FirstTouch(mT, mH, mL, mN, tr(F_EXITT, k), tr(F_EXITT, k) + barMin, tr(F_STOP, k), tr(F_SIDE, k))
            If t > 0 Then tr(F_EXITT, k) = t
        End If
    Next
End Sub

Public Sub RefineD2(tr() As Variant, ByVal tn As Long, mT() As Long, mH() As Double, mL() As Double, ByVal mN As Long, _
                    ByVal dayOpen As Long, ByVal dayClose As Long)
    Dim k As Long, t As Long, d As Long
    For k = 0 To tn - 1
        tr(F_ENTRYT, k) = (tr(F_ENTRYT, k) \ 1440) * 1440 + dayClose
        If Not tr(F_OPEN, k) Then
            d = tr(F_EXITT, k) \ 1440
            If Left$(tr(F_WHY, k), 2) = "跌破" Then
                tr(F_EXITT, k) = d * 1440 + dayOpen
            Else
                t = FirstTouch(mT, mH, mL, mN, d * 1440 + dayOpen, d * 1440 + dayClose + 1, tr(F_STOP, k), tr(F_SIDE, k))
                If t > 0 Then tr(F_EXITT, k) = t
            End If
        End If
    Next
End Sub

' 空單的「跌破」其實是收盤站上 EMA
Public Function WhyText(tr() As Variant, ByVal k As Long) As String
    WhyText = tr(F_WHY, k)
    If tr(F_SIDE, k) < 0 Then WhyText = Replace(WhyText, "跌破", "站上")
End Function

' 在最後多接一根開高低收都等於最後收盤的假 K（讓 run 檢查最後一根 K 的訊號與出場）
Public Sub WithDummy(ByVal n As Long, ByVal stp As Long, T() As Long, O() As Double, H() As Double, L() As Double, C() As Double, _
                     T2() As Long, O2() As Double, H2() As Double, L2() As Double, C2() As Double)
    Dim i As Long
    ReDim T2(0 To n): ReDim O2(0 To n): ReDim H2(0 To n): ReDim L2(0 To n): ReDim C2(0 To n)
    For i = 0 To n - 1
        T2(i) = T(i): O2(i) = O(i): H2(i) = H(i): L2(i) = L(i): C2(i) = C(i)
    Next
    T2(n) = T(n - 1) + stp: O2(n) = C(n - 1): H2(n) = C(n - 1): L2(n) = C(n - 1): C2(n) = C(n - 1)
End Sub

' ============================================================
' 全部計算（live.py main 的計算部分）。呼叫前先 ResetUnits/LoadHistory/ResetLive/(ParseChart)/AddLiveUnits
' ============================================================

Public Function ForwardStart() As Long
    ForwardStart = DayNum(2026, 11, 9) * 1440     ' 2026-11-09 00:00（兩套規則 11/7 定案，11/9 起算前測）
End Function

Public Function TwoYearsFrom() As Long
    TwoYearsFrom = DayNum(2024, 11, 7) * 1440
End Function

Private Sub CopyTrade(tr() As Variant, ByVal k As Long, out() As Variant)
    Dim f As Long
    ReDim out(0 To F_LAST)
    For f = 0 To F_LAST
        out(f) = tr(f, k)
    Next
End Sub

Private Sub FilterFrom(tr() As Variant, tn As Long, ByVal fromTm As Long)
    Dim k As Long, m As Long, f As Long
    m = 0
    For k = 0 To tn - 1
        If tr(F_T, k) >= fromTm Then
            For f = 0 To F_LAST
                tr(f, m) = tr(f, k)
            Next
            m = m + 1
        End If
    Next
    tn = m
End Sub

Public Sub ComputeAll(ByVal dailyCsv As String, ByVal twiiCsv As String, ByVal nowTm As Long)
    Dim T2() As Long, O2() As Double, H2() As Double, L2() As Double, C2() As Double, V() As Double
    Dim e20() As Double, e60() As Double, hi() As Double, md() As Double, sd() As Double, atr() As Double
    Dim tr2() As Variant, tn2 As Long, k As Long
    Dim mT() As Long, mO() As Double, mH() As Double, mL() As Double, mC() As Double, mV() As Double, mN As Long
    gNowTm = nowTm
    ' ---- v3：60 分K（日盤＋夜盤）
    gStep = "gVN = BuildBars"
    gVN = BuildBars(60, gVT, gVO, gVH, gVL, gVC, V)
    If gVN < 2 Then Err.Raise vbObjectError + 201, , "台指期 1 分K 不夠，算不出 60 分K"
    gLastC = gVC(gVN - 1)
    gStep = "CalcInd gVN, gVH"
    CalcInd gVN, gVH, gVL, gVC, gVE20, gVE60, gVHI, gVMID, gVSD, gVATR
    gStep = "gVTn = RunV3"
    gVTn = RunV3(gVN, gVT, gVO, gVH, gVL, gVC, gVE20, gVE60, gVHI, gVMID, gVSD, gVATR, 1, gVTr)
    gStep = "WithDummy gVN"
    WithDummy gVN, 60, gVT, gVO, gVH, gVL, gVC, T2, O2, H2, L2, C2
    CalcInd gVN + 1, H2, L2, C2, e20, e60, hi, md, sd, atr
    tn2 = RunV3(gVN + 1, T2, O2, H2, L2, C2, e20, e60, hi, md, sd, atr, 1, tr2)
    gStep = "PickState gVTr"
    PickState gVTr, gVTn, tr2, tn2, gVT(gVN - 1), gVPos, gVSig, gVSigHas, gVPend, gVPendHas
    gStep = "gM1N = BuildBars"
    gM1N = BuildBars(1, mT, mO, mH, mL, mC, mV)
    gStep = "RefineV3 gVTr"
    RefineV3 gVTr, gVTn, mT, mH, mL, gM1N, 60
    ' ---- D2：日K（只用日盤）
    gStep = "gDN = TxDailyBars"
    gDN = TxDailyBars(dailyCsv, gDT, gDO, gDH, gDL, gDC, gDI0, mT, mO, mH, mL, mC, mN)
    If gDN < 62 Then Err.Raise vbObjectError + 202, , "台指期日K 不到 62 根，算不出 D2（缺 tx_daily.csv？）"
    gStep = "CalcInd gDN, gDH"
    CalcInd gDN, gDH, gDL, gDC, gDE20, gDE60, gDHI, gDMID, gDSD, gDATR
    gStep = "gDTn = RunD2"
    gDTn = RunD2(gDN, gDT, gDO, gDH, gDL, gDC, gDE20, gDE60, gDATR, gDI0, gDTr)
    WithDummy gDN, 1440, gDT, gDO, gDH, gDL, gDC, T2, O2, H2, L2, C2
    CalcInd gDN + 1, H2, L2, C2, e20, e60, hi, md, sd, atr
    tn2 = RunD2(gDN + 1, T2, O2, H2, L2, C2, e20, e60, atr, gDI0, tr2)
    gStep = "PickState gDTr"
    PickState gDTr, gDTn, tr2, tn2, gDT(gDN - 1), gDPos, gDSig, gDSigHas, gDPend, gDPendHas
    gStep = "RefineD2 gDTr"
    RefineD2 gDTr, gDTn, mT, mH, mL, mN, T_DAYOPEN, T_DAYCLOSE
    ' ---- 條件明細
    gCkN = 0
    ReDim gCkStrat(0 To 63): ReDim gCkSide(0 To 63): ReDim gCkGroup(0 To 63): ReDim gCkName(0 To 63)
    ReDim gCkOk(0 To 63): ReDim gCkDetail(0 To 63)
    gStep = "ChecksV3 gVN"
    ChecksV3 gVN - 1
    gStep = "ChecksD2 gDN"
    ChecksD2 gDN - 1
    ' ---- 加權指數回測
    gStep = "ComputeTwii twiiCsv"
    ComputeTwii twiiCsv
End Sub

Public Sub ComputeTwii(ByVal twiiCsv As String)
    Dim tT() As Long, tOpn() As Double, tH() As Double, tL() As Double, tC() As Double, n As Long
    Dim dT() As Long, dOpn() As Double, dH() As Double, dL() As Double, dC() As Double, dn As Long
    Dim e20() As Double, e60() As Double, hi() As Double, md() As Double, sd() As Double, atr() As Double
    gWVTn = 0: gWDTn = 0
    TrInit gWVTr: TrInit gWDTr
    n = LoadTwii(twiiCsv, tT, tOpn, tH, tL, tC)
    gTwN = n
    If n < 200 Then Exit Sub
    gTwLast = tT(n - 1)
    CalcInd n, tH, tL, tC, e20, e60, hi, md, sd, atr
    gWVTn = RunV3(n, tT, tOpn, tH, tL, tC, e20, e60, hi, md, sd, atr, 2, gWVTr)
    FilterFrom gWVTr, gWVTn, TwoYearsFrom()
    RefineV3 gWVTr, gWVTn, tT, tH, tL, n, 60
    dn = TwiiDaily(n, tT, tOpn, tH, tL, tC, dT, dOpn, dH, dL, dC)
    If dn < 62 Then Exit Sub
    CalcInd dn, dH, dL, dC, e20, e60, hi, md, sd, atr
    gWDTn = RunD2(dn, dT, dOpn, dH, dL, dC, e20, e60, atr, -1, gWDTr)
    FilterFrom gWDTr, gWDTn, TwoYearsFrom()
    RefineD2 gWDTr, gWDTn, tT, tH, tL, n, 540, 810
End Sub

Private Sub PickState(tr() As Variant, ByVal tn As Long, tr2() As Variant, ByVal tn2 As Long, ByVal lastT As Long, _
                      pos As Long, sig() As Variant, sigHas As Boolean, pend As String, pendHas As Boolean)
    Dim k As Long
    pos = -1: sigHas = False: pend = "": pendHas = False
    For k = 0 To tn - 1
        If tr(F_OPEN, k) Then pos = k: Exit For
    Next
    For k = 0 To tn2 - 1
        If tr2(F_T, k) = lastT Then CopyTrade tr2, k, sig: sigHas = True: Exit For
    Next
    If pos >= 0 Then
        For k = 0 To tn2 - 1
            If tr2(F_T, k) = tr(F_T, pos) Then
                If Not tr2(F_OPEN, k) Then
                    pendHas = True
                    pend = WhyText(tr2, k)
                    If pend = "" Then pend = "出場條件成立"
                End If
                Exit For
            End If
        Next
    End If
End Sub

Private Sub Ck(ByVal strat As Long, ByVal side As Long, ByVal grp As String, ByVal nm As String, ByVal ok As Boolean, ByVal detail As String)
    If gCkN > UBound(gCkStrat) Then
        ReDim Preserve gCkStrat(0 To 2 * gCkN): ReDim Preserve gCkSide(0 To 2 * gCkN): ReDim Preserve gCkGroup(0 To 2 * gCkN)
        ReDim Preserve gCkName(0 To 2 * gCkN): ReDim Preserve gCkOk(0 To 2 * gCkN): ReDim Preserve gCkDetail(0 To 2 * gCkN)
    End If
    gCkStrat(gCkN) = strat: gCkSide(gCkN) = side: gCkGroup(gCkN) = grp: gCkName(gCkN) = nm
    gCkOk(gCkN) = ok: gCkDetail(gCkN) = detail
    gCkN = gCkN + 1
End Sub

Private Sub ProjStop(ByVal e As Double, ByVal ext As Double, ByVal a As Double, ByVal side As Long, ByVal stopAtr As Double, _
                     ByVal maxAtr As Double, ok As Boolean, stp As Double, raw As Double)
    raw = Abs(e - (ext - 0.1 * a * side))
    stp = e - MaxD(MinD(raw, stopAtr * a), 0.5 * a) * side
    ok = (raw <= maxAtr * a)
End Sub

Private Sub ChecksV3(ByVal i As Long)
    Dim s As Long, sd As Long, body As Double, rng As Double, strong As Boolean, bb As Double, prv As Double, gt As String
    Dim ok As Boolean, stp As Double, raw As Double, k As String
    For sd = 1 To 2
        If sd = 1 Then s = 1 Else s = -1
        body = (gVC(i) - gVO(i)) * s: rng = gVH(i) - gVL(i)
        strong = (body > 0 And body >= 0.5 * rng)
        bb = gVMID(i) + 2 * gVSD(i) * s
        prv = Sel(s > 0, gVH(i - 1), gVL(i - 1))
        If s > 0 Then gt = ">" Else gt = "<"
        ProjStop gVC(i), Sel(s > 0, gVL(i), gVH(i)), gVATR(i), s, 1.5, 2#, ok, stp, raw
        gStopOk(1, sd) = ok: gStopVal(1, sd) = stp: gStopRaw(1, sd) = raw
        If s > 0 Then k = "紅" Else k = "黑"
        Ck 1, sd, "common", "可進場時段", Allowed(gVT(i), 1), FmtTm(gVT(i), 3) & " 這根"
        Ck 1, sd, "common", "趨勢 EMA20 " & gt & " EMA60", (gVE20(i) - gVE60(i)) * s > 0, N0(gVE20(i)) & " / " & N0(gVE60(i))
        Ck 1, sd, "突破", k & "K 實體 ≥ 高低差一半", strong, "實體 " & S0(gVC(i) - gVO(i)) & "、高低差 " & N0(rng)
        Ck 1, sd, "突破", "實體 ≥ 0.8 ATR", body >= 0.8 * gVATR(i - 1), "需 " & N0(0.8 * gVATR(i - 1)) & " 點"
        Ck 1, sd, "突破", "收盤" & IIf(s > 0, "站上布林上軌", "跌破布林下軌"), (gVC(i) - bb) * s > 0, IIf(s > 0, "上", "下") & "軌 " & N0(bb)
        Ck 1, sd, "突破", "MACD 柱" & IIf(s > 0, "翻正", "翻負"), gVHI(i) * s > 0, SN(gVHI(i), 1)
        Ck 1, sd, "拉回", k & "K 實體 ≥ 高低差一半", strong, ""
        Ck 1, sd, "拉回", IIf(s > 0, "低點碰到", "高點碰到") & " EMA20", (Sel(s > 0, gVL(i), gVH(i)) - gVE20(i)) * s <= 0, "EMA20 " & N0(gVE20(i))
        Ck 1, sd, "拉回", "收盤" & IIf(s > 0, "站回", "壓回") & " EMA20", (gVC(i) - gVE20(i)) * s > 0, ""
        Ck 1, sd, "拉回", "收盤" & IIf(s > 0, "過前一根高點", "破前一根低點"), (gVC(i) - prv) * s > 0, N0(prv)
        Ck 1, sd, "拉回", "EMA60 比 5 根前" & IIf(s > 0, "高", "低"), (gVE60(i) - gVE60(i - 5)) * s > 0, S0(gVE60(i) - gVE60(i - 5))
    Next
End Sub

Private Sub ChecksD2(ByVal i As Long)
    Dim s As Long, sd As Long, body As Double, rng As Double, lvl As Double, gt As String, j As Long
    Dim ok As Boolean, stp As Double, raw As Double, k As String
    For sd = 1 To 2
        If sd = 1 Then s = 1 Else s = -1
        body = (gDC(i) - gDO(i)) * s: rng = gDH(i) - gDL(i)
        lvl = gDC(i - 20)
        For j = i - 19 To i - 1
            If s > 0 Then
                If gDC(j) > lvl Then lvl = gDC(j)
            Else
                If gDC(j) < lvl Then lvl = gDC(j)
            End If
        Next
        If s > 0 Then gt = ">" Else gt = "<"
        If s > 0 Then k = "紅" Else k = "黑"
        ProjStop gDC(i), Sel(s > 0, gDL(i), gDH(i)), gDATR(i), s, 1#, 1.5, ok, stp, raw
        gStopOk(2, sd) = ok: gStopVal(2, sd) = stp: gStopRaw(2, sd) = raw
        Ck 2, sd, "common", "趨勢 日EMA20 " & gt & " EMA60", (gDE20(i) - gDE60(i)) * s > 0, N0(gDE20(i)) & " / " & N0(gDE60(i))
        Ck 2, sd, "突破", "收盤創 20 日新" & IIf(s > 0, "高", "低"), (gDC(i) - lvl) * s > 0, "需" & IIf(s > 0, "高於", "低於") & " " & N0(lvl)
        Ck 2, sd, "突破", k & "K 實體 ≥ 高低差一半", body > 0 And body >= 0.5 * rng, _
           "開 " & N0(gDO(i)) & "、實體 " & S0(gDC(i) - gDO(i)) & "、高低差 " & N0(rng)
        Ck 2, sd, "突破", "停損距離 ≤ 1.5 ATR（不追太遠）", ok, N0(raw) & " 點 / 上限 " & N0(1.5 * gDATR(i))
    Next
End Sub

' ============================================================
' 測試用：把計算結果寫成文字檔，跟 Python（excel/strategy_vba/ref_dump.py）逐項比對
' ============================================================
Private Function DN6(ByVal x As Double) As String
    DN6 = Trim$(Str$(Round(x, 6)))
End Function

Private Sub DumpTrades(ByVal f As Integer, ByVal nm As String, tr() As Variant, ByVal tn As Long)
    Dim k As Long
    For k = 0 To tn - 1
        Print #f, "TR" & vbTab & nm & vbTab & FmtTm(tr(F_T, k)) & vbTab & tr(F_SIDE, k) & vbTab & tr(F_KIND, k) & vbTab & _
            DN6(tr(F_RISK, k)) & vbTab & DN6(tr(F_RMUL, k)) & vbTab & DN6(tr(F_ENTRY, k)) & vbTab & FmtTm(tr(F_ENTRYT, k)) & vbTab & _
            DN6(tr(F_STOP0, k)) & vbTab & DN6(tr(F_STOP, k)) & vbTab & IIf(tr(F_OPEN, k), "", FmtTm(tr(F_EXITT, k))) & vbTab & _
            DN6(tr(F_EXIT, k)) & vbTab & IIf(tr(F_OPEN, k), 1, 0) & vbTab & IIf(tr(F_BE, k), 1, 0) & vbTab & WhyText(tr, k)
    Next
End Sub

Public Function DumpAll(ByVal path As String) As String
    Dim f As Integer, i As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, "N60" & vbTab & gVN & vbTab & FmtTm(gVT(0)) & vbTab & FmtTm(gVT(gVN - 1))
    For i = 0 To gVN - 1
        Print #f, "B60" & vbTab & FmtTm(gVT(i)) & vbTab & DN6(gVO(i)) & vbTab & DN6(gVH(i)) & vbTab & DN6(gVL(i)) & vbTab & DN6(gVC(i))
    Next
    Print #f, "M1" & vbTab & gM1N
    Print #f, "ND" & vbTab & gDN & vbTab & gDI0
    For i = 0 To gDN - 1
        Print #f, "BD" & vbTab & FmtTm(gDT(i), 2) & vbTab & DN6(gDO(i)) & vbTab & DN6(gDH(i)) & vbTab & DN6(gDL(i)) & vbTab & DN6(gDC(i))
    Next
    i = gVN - 1
    Print #f, "IND60" & vbTab & DN6(gVE20(i)) & vbTab & DN6(gVE60(i)) & vbTab & DN6(gVHI(i)) & vbTab & DN6(gVMID(i)) & vbTab & DN6(gVSD(i)) & vbTab & DN6(gVATR(i))
    i = gDN - 1
    Print #f, "INDD" & vbTab & DN6(gDE20(i)) & vbTab & DN6(gDE60(i)) & vbTab & DN6(gDATR(i))
    DumpTrades f, "v3", gVTr, gVTn
    DumpTrades f, "d2", gDTr, gDTn
    DumpTrades f, "tw_v3", gWVTr, gWVTn
    DumpTrades f, "tw_d2", gWDTr, gWDTn
    Print #f, "POS" & vbTab & "v3" & vbTab & gVPos
    Print #f, "POS" & vbTab & "d2" & vbTab & gDPos
    If gVSigHas Then Print #f, "SIG" & vbTab & "v3" & vbTab & gVSig(F_SIDE) & vbTab & gVSig(F_KIND) & vbTab & DN6(gVSig(F_ENTRY)) & vbTab & DN6(gVSig(F_STOP0))
    If gDSigHas Then Print #f, "SIG" & vbTab & "d2" & vbTab & gDSig(F_SIDE) & vbTab & gDSig(F_KIND) & vbTab & DN6(gDSig(F_ENTRY)) & vbTab & DN6(gDSig(F_STOP0))
    If gVPendHas Then Print #f, "PEND" & vbTab & "v3" & vbTab & gVPend
    If gDPendHas Then Print #f, "PEND" & vbTab & "d2" & vbTab & gDPend
    For i = 0 To gCkN - 1
        Print #f, "CK" & vbTab & gCkStrat(i) & vbTab & gCkSide(i) & vbTab & gCkGroup(i) & vbTab & IIf(gCkOk(i), 1, 0) & vbTab & gCkName(i) & vbTab & gCkDetail(i)
    Next
    For i = 1 To 2
        Print #f, "STOP" & vbTab & "1" & vbTab & i & vbTab & IIf(gStopOk(1, i), 1, 0) & vbTab & DN6(gStopVal(1, i)) & vbTab & DN6(gStopRaw(1, i))
        Print #f, "STOP" & vbTab & "2" & vbTab & i & vbTab & IIf(gStopOk(2, i), 1, 0) & vbTab & DN6(gStopVal(2, i)) & vbTab & DN6(gStopRaw(2, i))
    Next
    For i = 0 To gGapN - 1
        Print #f, "GAP" & vbTab & FmtTm(gGap(i))
    Next
    Close #f
    DumpAll = "ok"
End Function

' ============================================================
' 即時 1 分K：每個盤別抓成交量前兩大的契約（同 live.py fetch_live）
' gFixtureDir 有值時改讀測試資料（quotes_0.json、quotes_1.json、chart_<SymbolID>.json）
' ============================================================

Public Function MisGet(ByVal ep As String, ByVal body As String, ByVal key As String) As String
    If Len(gFixtureDir) > 0 Then
        MisGet = ReadUtf8File(JoinPath(gFixtureDir, key & ".json"))
        If Len(MisGet) = 0 Then Err.Raise vbObjectError + 103, , "沒有測試資料 " & key
    Else
        MisGet = MisPost(ep, body)
    End If
End Function

Public Sub FetchLive(ByVal yearNow As Long)
    Dim mt As Long, slot As Long, q As String, syms() As String, cn() As String, n As Long, k As Long, c As String, ok As Boolean
    ResetLive
    For slot = 1 To 2
        mt = slot - 1
        ok = True
        On Error Resume Next
        q = MisGet("getQuoteList", "{""MarketType"":""" & mt & """,""SymbolType"":""F"",""KindID"":""1"",""CID"":""TXF"",""ExpireMonth"":""""," & _
                   """RowSize"":""全部"",""PageNo"":"""",""SortColumn"":"""",""AscDesc"":""A""}", "quotes_" & mt)
        If Err.Number = 0 Then n = ParseQuoteList(q, syms, cn)
        If Err.Number <> 0 Then
            gLiveErr(slot) = "報價清單抓取失敗：" & Err.Description
            ok = False
        End If
        Err.Clear
        On Error GoTo 0
        If ok Then
            If n > 2 Then n = 2
            For k = 0 To n - 1
                On Error Resume Next
                c = MisGet("getChartData1M", "{""SymbolID"":""" & syms(k) & """}", "chart_" & syms(k))
                If Err.Number = 0 Then ParseChart c, slot, ContractCode(syms(k), yearNow), (k = 0), syms(k), cn(k)
                If Err.Number <> 0 Then
                    If Len(gLiveErr(slot)) = 0 And Not gLiveHas(slot) Then gLiveErr(slot) = "1 分K 抓取失敗：" & Err.Description
                End If
                Err.Clear
                On Error GoTo 0
            Next
        End If
    Next
End Sub
