Attribute VB_Name = "TXO"
Option Explicit
' ============================================================
' 台指週選盤面 Excel 巨集版
' 按「更新資料」直接向期交所與 Yahoo Finance 抓資料，在 Excel 裡計算並畫圖。
' 不需要 Python、不靠 GitHub。只能在 Windows 版 Excel 執行。
' 資料來源：www.taifex.com.tw、mis.taifex.com.tw、query1.finance.yahoo.com
' 時間以這台電腦的時鐘為準（請設成台北時間）。
' ============================================================

Private Const UA As String = "Mozilla/5.0 (txo-excel)"

Private gErr As String
Private gNow As Date
Private gTradeDate As String
Private gPrevDate As String
Private gSum As Collection      ' 每個到期一個 Dictionary（摘要）
Private gTab As Collection      ' 每個到期一個 2D 陣列（履約價表）
Private gPC As Variant          ' 全市場 P/C 歷史 2D 陣列
Private gInst As Object         ' 三大法人選擇權：日期 -> Dictionary
Private gLarge As Object        ' 大額交易人：日期 -> Dictionary
Private gFut As Object          ' 期貨三大法人：日期 -> Dictionary
Private gFutLive As Object      ' 期貨即時：日盤/夜盤 -> Dictionary
Private gVol As Object          ' 波動率統計
Private gVolTab As Variant      ' 日期、VIX、HV20
Private gGlobal As Collection   ' 美股與 VIX

' JSON 解析用
Private jS As String
Private jP As Long
Private gChartTitle As String   ' 圖表標題：加第一條數列後才設定（空白圖表設標題有時會失敗）

' ============================================================
' 進入點
' ============================================================
Public Sub UpdateAll()
    Dim t0 As Double
    t0 = Timer
    gErr = ""
    gNow = Now
    Set gInst = Nothing: Set gLarge = Nothing: Set gFut = Nothing: Set gFutLive = Nothing
    Set gVol = Nothing: Set gGlobal = Nothing
    gPC = Empty: gVolTab = Empty
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "抓期交所選擇權資料..."
    On Error GoTo fail
    ComputeOptions
    On Error GoTo 0

    Application.StatusBar = "抓 P/C、法人、大額交易人資料..."
    SafeRun "PC"
    SafeRun "INST"
    SafeRun "LARGE"
    Application.StatusBar = "抓台指期資料..."
    SafeRun "FUT"
    SafeRun "FUTLIVE"
    Application.StatusBar = "抓 VIX 與歷史波動..."
    SafeRun "VOL"
    Application.StatusBar = "抓美股資料..."
    SafeRun "GLOBAL"

    Application.StatusBar = "產生工作表與圖表..."
    On Error GoTo fail
    RebuildSheets
    On Error GoTo 0

    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets("盤面").Activate
    ThisWorkbook.Worksheets("盤面").Range("A1").Select
    If gErr <> "" Then
        MsgBox "已更新，但以下資料抓取失敗，表上會留空：" & vbLf & gErr, vbExclamation, "台指週選盤面"
    End If
    Exit Sub
fail:
    Dim msg As String
    msg = Err.Description
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "更新失敗：" & msg & vbLf & vbLf & "請確認網路連線，稍後再按一次「更新資料」。", vbCritical, "台指週選盤面"
End Sub

Private Sub SafeRun(ByVal what As String)
    On Error GoTo bad
    Select Case what
        Case "PC": FetchPC
        Case "INST": FetchInst
        Case "LARGE": FetchLarge
        Case "FUT": FetchFut
        Case "FUTLIVE": FetchFutLive
        Case "VOL": FetchVol
        Case "GLOBAL": FetchGlobal
    End Select
    Exit Sub
bad:
    gErr = gErr & what & "：" & Err.Description & vbLf
End Sub

' 建立「更新資料」按鈕（活頁簿開啟時呼叫）
Public Sub EnsureButton()
    Dim ws As Worksheet, shp As Object
    Set ws = ThisWorkbook.Worksheets("盤面")
    On Error Resume Next
    Set shp = ws.Shapes("btnUpdate")
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(5, ws.Range("H1").Left, 4, 120, 30)
        shp.Name = "btnUpdate"
        shp.Fill.ForeColor.RGB = RGB(58, 91, 191)
        shp.Line.Visible = False
        With shp.TextFrame2
            .TextRange.Text = "更新資料"
            .TextRange.Font.Size = 12
            .TextRange.Font.Bold = True
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = 2
            .VerticalAnchor = 3
        End With
        shp.Placement = 3
    End If
    shp.OnAction = "UpdateAll"
End Sub

' ============================================================
' 網路
' ============================================================
' Big5 解碼：中文 Windows 直接用系統字碼頁（StrConv），其他語系才用 ADODB.Stream
Private Function Big5Text(ByVal b As Variant) As String
    Dim st As Object
    If IsNull(b) Or IsEmpty(b) Then Exit Function
    If UBound(b) < 0 Then Exit Function
    If StrConv(ChrB(&HA4) & ChrB(&HA4), vbUnicode) = ChrW(&H4E2D) Then
        Big5Text = StrConv(b, vbUnicode)
    Else
        Set st = CreateObject("ADODB.Stream")
        st.Type = 1
        st.Open
        st.Write b
        st.Position = 0
        st.Type = 2
        st.Charset = "big5"
        Big5Text = st.ReadText
        st.Close
    End If
End Function

' UTF-8 解碼（不靠 ADODB）
Private Function Utf8Text(ByVal b As Variant) As String
    Dim i As Long, k As Long, c As Long, hi As Long, buf As String
    If IsNull(b) Or IsEmpty(b) Then Exit Function
    If UBound(b) < LBound(b) Then Exit Function
    buf = Space$(UBound(b) - LBound(b) + 1)
    i = LBound(b)
    Do While i <= UBound(b)
        c = b(i)
        If c < &H80 Then
            i = i + 1
        ElseIf c < &HE0 And i + 1 <= UBound(b) Then
            c = ((c And &H1F) * &H40&) Or (b(i + 1) And &H3F)
            i = i + 2
        ElseIf c < &HF0 And i + 2 <= UBound(b) Then
            c = ((c And &HF) * &H1000&) Or ((b(i + 1) And &H3F) * &H40&) Or (b(i + 2) And &H3F)
            i = i + 3
        ElseIf i + 3 <= UBound(b) Then
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

Private Function Http(ByVal url As String, Optional ByVal body As String = "", _
                      Optional ByVal isJson As Boolean = False, Optional ByVal cs As String = "big5") As String
    Dim h As Object, tries As Long
    For tries = 1 To 2
        On Error Resume Next
        Err.Clear
        If tries = 1 Then
            Set h = CreateObject("MSXML2.ServerXMLHTTP.6.0")
            h.setTimeouts 15000, 15000, 30000, 90000
        Else
            Set h = CreateObject("MSXML2.XMLHTTP.6.0")
        End If
        If Len(body) = 0 Then
            h.Open "GET", url, False
        Else
            h.Open "POST", url, False
        End If
        h.setRequestHeader "User-Agent", UA
        h.setRequestHeader "If-Modified-Since", "Sat, 01 Jan 2000 00:00:00 GMT"
        If Len(body) > 0 Then
            If isJson Then
                h.setRequestHeader "Content-Type", "application/json"
            Else
                h.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            End If
            h.send body
        Else
            h.send
        End If
        If Err.Number = 0 Then
            If h.Status = 200 Then
                On Error GoTo 0
                If cs = "utf-8" Then
                    Http = Utf8Text(h.responseBody)
                Else
                    Http = Big5Text(h.responseBody)
                End If
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next
    Err.Raise vbObjectError + 1, , "連不上 " & Split(url, "?")(0)
End Function

Private Function D8(ByVal d As Date) As String
    D8 = Year(d) & "%2F" & Right$("0" & Month(d), 2) & "%2F" & Right$("0" & Day(d), 2)
End Function

Private Function CsvRows(ByVal txt As String) As Collection
    Dim c As New Collection, ln As Variant
    txt = Replace(txt, vbCr, "")
    For Each ln In Split(txt, vbLf)
        If Len(ln) > 10 Then
            If Left$(ln, 2) = "20" Then c.Add Split(ln, ",")
        End If
    Next
    Set CsvRows = c
End Function

' 期交所下載：結束日還沒有資料時會回錯誤頁，結束日往前退
Private Function TaifexCsv(ByVal url As String, ByVal params As String, ByVal d0 As Date, ByVal d1 As Date) As Collection
    Dim back As Long, c As Collection
    For back = 0 To 5
        Set c = CsvRows(Http(url, params & "&queryStartDate=" & D8(d0) & "&queryEndDate=" & D8(d1 - back)))
        If c.Count > 0 Then Exit For
    Next
    Set TaifexCsv = c
End Function

' ============================================================
' 小工具
' ============================================================
Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

Private Function DGet(ByVal d As Object, ByVal k As String, Optional ByVal dflt As Variant) As Variant
    If d.Exists(k) Then
        DGet = d(k)
    ElseIf IsMissing(dflt) Then
        DGet = Empty
    Else
        DGet = dflt
    End If
End Function

Private Function TxtNum(ByVal s As Variant) As Variant
    If IsNull(s) Or IsEmpty(s) Then TxtNum = Empty: Exit Function
    s = Trim$(CStr(s))
    If s = "" Or s = "-" Then TxtNum = Empty Else TxtNum = Val(s)
End Function

Private Function LngOf(ByVal s As Variant) As Long
    If IsNull(s) Or IsEmpty(s) Then Exit Function
    LngOf = CLng(Val(Trim$(CStr(s))))
End Function

Private Function DateFromSlash(ByVal s As String) As Date
    s = Trim$(s)
    DateFromSlash = DateSerial(Val(Left$(s, 4)), Val(Mid$(s, 6, 2)), Val(Mid$(s, 9, 2)))
End Function

Private Function DateFromYmd(ByVal s As String) As Date
    s = Trim$(s)
    DateFromYmd = DateSerial(Val(Left$(s, 4)), Val(Mid$(s, 5, 2)), Val(Mid$(s, 7, 2)))
End Function

Private Function Mx(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then Mx = a Else Mx = b
End Function

Private Sub SortArr(a As Variant)
    Dim i As Long, j As Long, v As Variant
    For i = LBound(a) + 1 To UBound(a)
        v = a(i)
        j = i - 1
        Do While j >= LBound(a)
            If a(j) <= v Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = v
    Next
End Sub

Private Function SortedKeys(ByVal d As Object) As Variant
    Dim a As Variant
    If d.Count = 0 Then SortedKeys = Array(): Exit Function
    a = d.Keys
    SortArr a
    SortedKeys = a
End Function

Private Function NoNull(ByVal v As Variant) As Variant
    If IsNull(v) Then NoNull = Empty Else NoNull = v
End Function

' ============================================================
' Black-76
' ============================================================
Private Function NCdf(ByVal x As Double) As Double
    NCdf = Application.WorksheetFunction.Norm_S_Dist(x, True)
End Function

Private Function Bs(ByVal F As Double, ByVal K As Double, ByVal T As Double, ByVal s As Double, ByVal cp As String) As Double
    Dim d1 As Double, d2 As Double
    d1 = (Log(F / K) + 0.5 * s * s * T) / (s * Sqr(T))
    d2 = d1 - s * Sqr(T)
    If cp = "C" Then
        Bs = F * NCdf(d1) - K * NCdf(d2)
    Else
        Bs = K * NCdf(-d2) - F * NCdf(-d1)
    End If
End Function

Private Function ImpliedVol(ByVal px As Variant, ByVal F As Double, ByVal K As Double, ByVal T As Double, ByVal cp As String) As Variant
    Dim intr As Double, lo As Double, hi As Double, m As Double, i As Long
    ImpliedVol = Empty
    If IsEmpty(px) Then Exit Function
    If cp = "C" Then intr = Mx(F - K, 0) Else intr = Mx(K - F, 0)
    If T <= 0 Or px <= intr + 0.05 Then Exit Function
    lo = 0.01
    hi = 5
    For i = 1 To 60
        m = (lo + hi) / 2
        If Bs(F, K, T, m, cp) > px Then
            hi = m
        Else
            lo = m
        End If
    Next
    If m > 0.011 And m < 4.99 Then ImpliedVol = m
End Function

' ============================================================
' 1. 選擇權：OI、結算價、即時報價 -> 各到期指標
' ============================================================
Private Sub ComputeOptions()
    Dim all As Collection, rows As New Collection, r As Variant
    Set all = TaifexCsv("https://www.taifex.com.tw/cht/3/optDataDown", _
                        "down_type=1&commodity_id=TXO&commodity_id2=", gNow - 10, gNow)
    gTradeDate = ""
    gPrevDate = ""
    For Each r In all
        If UBound(r) >= 20 Then
            If Trim$(r(1)) = "TXO" Then
                rows.Add r
                If Trim$(r(0)) > gTradeDate Then gTradeDate = Trim$(r(0))
            End If
        End If
    Next
    If rows.Count = 0 Then Err.Raise vbObjectError + 2, , "期交所選擇權每日行情沒有資料"
    For Each r In rows
        If Trim$(r(0)) < gTradeDate And Trim$(r(0)) > gPrevDate Then gPrevDate = Trim$(r(0))
    Next

    Dim td As Date, exps As Object, oi As Object, vol As Object, settle As Object, prevOI As Object
    Dim code As String, k As String, cp As String, sess As String, ed As Date
    td = DateFromSlash(gTradeDate)
    Set exps = NewDict(): Set oi = NewDict(): Set vol = NewDict()
    Set settle = NewDict(): Set prevOI = NewDict()
    For Each r In rows
        code = Trim$(r(2))
        If Trim$(r(4)) = "買權" Then cp = "C" Else cp = "P"
        k = code & "|" & CLng(Val(r(3))) & "|" & cp
        sess = Trim$(r(17))
        If Trim$(r(0)) = gTradeDate Then
            ed = DateFromYmd(r(20))
            If ed > td Then exps(code) = ed
            vol(k) = DGet(vol, k, 0) + LngOf(r(9))
            If sess = "一般" Then
                oi(k) = LngOf(r(11))
                settle(k) = TxtNum(r(10))
            End If
        ElseIf Trim$(r(0)) = gPrevDate And sess = "一般" Then
            prevOI(k) = LngOf(r(11))
        End If
    Next

    ' 依到期日排序，取最近兩個週選 + 最近的月選
    Dim tags As Object, kv As Variant, srt As Variant, i As Long, nW As Long, haveM As Boolean
    Dim chosen As New Collection
    Set tags = NewDict()
    For Each kv In exps.Keys
        tags(Format$(exps(kv), "yyyymmdd") & "|" & kv) = 1
    Next
    srt = SortedKeys(tags)
    For i = 0 To UBound(srt)
        code = Split(srt(i), "|")(1)
        If IsNumeric(code) Then
            If Not haveM Then chosen.Add code: haveM = True
        ElseIf nW < 2 Then
            chosen.Add code
            nW = nW + 1
        End If
    Next

    Set gSum = New Collection
    Set gTab = New Collection
    Dim c As Variant
    For Each c In chosen
        ComputeExpiry CStr(c), exps(c), oi, vol, settle, prevOI
    Next
    If gSum.Count = 0 Then Err.Raise vbObjectError + 3, , "沒有可計算的到期"
End Sub

Private Function LiveQuotes(ByVal code As String, ByVal mt As String, ByRef qt As Date) As Object
    Dim res As Object, body As String, root As Object, q As Object, x As Variant
    Dim good As Long, b As Variant, a As Variant, md As Variant, maxT As String, today As String
    Set res = NewDict()
    Set LiveQuotes = res
    On Error GoTo bad
    body = "{""MarketType"":""" & mt & """,""SymbolType"":""O"",""KindID"":""1"",""CID"":""TXO"",""ExpireMonth"":""" & code & _
           """,""RowSize"":""全部"",""PageNo"":"""",""SortColumn"":"""",""AscDesc"":""A""}"
    Set root = JsonObj(Http("https://mis.taifex.com.tw/futures/api/getQuoteListOption", body, True, "utf-8"))
    Set q = root("RtData")("QuoteList")
    today = Format$(gNow, "yyyymmdd")
    For Each x In q
        b = TxtNum(x("CBidPrice1"))
        a = TxtNum(x("CAskPrice1"))
        md = Empty
        If Not IsEmpty(b) And Not IsEmpty(a) Then
            If b > 0 And a > 0 And a >= b Then
                md = (a + b) / 2
                good = good + 1
            End If
        End If
        res(CLng(Val(x("StrikePrice"))) & "|" & x("CP")) = Array(md, LngOf(x("CTotalVolume")))
        If CStr(NoNull(x("CTime"))) <> "" And CStr(NoNull(x("CDate"))) = today Then
            If x("CDate") & x("CTime") > maxT Then maxT = x("CDate") & x("CTime")
        End If
    Next
    If good < 20 Then res.RemoveAll: Exit Function
    If maxT <> "" Then
        qt = DateFromYmd(Left$(maxT, 8)) + TimeSerial(Val(Mid$(maxT, 9, 2)), Val(Mid$(maxT, 11, 2)), Val(Mid$(maxT, 13, 2)))
    Else
        qt = gNow
    End If
    Exit Function
bad:
    res.RemoveAll
End Function

Private Sub ComputeExpiry(ByVal code As String, ByVal ed As Date, ByVal oi As Object, ByVal vol As Object, _
                          ByVal settle As Object, ByVal prevOI As Object)
    Dim pre As String, myOI As Object, kk As Variant, live As Object, qt As Date, sessName As String
    Dim order As Variant, j As Long, hm As Long
    pre = code & "|"
    Set myOI = NewDict()
    For Each kk In oi.Keys
        If Left$(kk, Len(pre)) = pre Then myOI(Mid$(kk, Len(pre) + 1)) = oi(kk)
    Next

    hm = Hour(gNow) * 100 + Minute(gNow)
    If hm >= 845 And hm < 1400 Then order = Array("0", "1") Else order = Array("1", "0")
    For j = 0 To 1
        qt = gNow
        Set live = LiveQuotes(code, CStr(order(j)), qt)
        If live.Count > 0 Then
            If order(j) = "0" Then sessName = "日盤" Else sessName = "夜盤"
            Exit For
        End If
    Next
    If sessName = "" Then
        sessName = "結算價"
        Set live = NewDict()
        qt = gNow
    End If

    ' 價格：即時中價優先，沒有就用結算價
    Dim keysU As Object, price As Object, strikesD As Object, p As Variant, arr As Variant
    Set keysU = NewDict(): Set price = NewDict(): Set strikesD = NewDict()
    For Each kk In myOI.Keys: keysU(kk) = 1: Next
    For Each kk In live.Keys: keysU(kk) = 1: Next
    For Each kk In keysU.Keys
        p = Empty
        If live.Exists(kk) Then
            arr = live(kk)
            p = arr(0)
        End If
        If IsEmpty(p) Then p = DGet(settle, pre & kk)
        If Not IsEmpty(p) Then
            If p > 0 Then price(kk) = p
        End If
        strikesD(CStr(Split(kk, "|")(0))) = CLng(Split(kk, "|")(0))
    Next
    Dim Ks As Variant, i As Long
    Ks = strikesD.Items
    If UBound(Ks) < 0 Then Exit Sub
    SortArr Ks

    ' 價平：買賣權價差最小的履約價
    Dim bestD As Double, K0 As Double, c0 As Double, p0 As Double, cpx As Variant, ppx As Variant
    bestD = -1
    For i = 0 To UBound(Ks)
        cpx = DGet(price, Ks(i) & "|C")
        ppx = DGet(price, Ks(i) & "|P")
        If Not IsEmpty(cpx) And Not IsEmpty(ppx) Then
            If bestD < 0 Or Abs(cpx - ppx) < bestD Then
                bestD = Abs(cpx - ppx): K0 = Ks(i): c0 = cpx: p0 = ppx
            End If
        End If
    Next
    If bestD < 0 Then Exit Sub

    Dim F As Double, T As Double, atmIV As Variant, straddle As Double
    F = K0 + c0 - p0
    T = (ed + TimeSerial(13, 30, 0)) - qt
    If T < 1 / 24 Then T = 1 / 24
    T = T / 365
    atmIV = ImpliedVol(c0, F, K0, T, "C")
    If IsEmpty(atmIV) Then atmIV = ImpliedVol(p0, F, K0, T, "P")
    straddle = c0 + p0

    ' OI、成交量、P/C
    Dim tc As Double, tp As Double, nc As Double, np As Double, kStr As Variant, Kn As Double, cpS As String
    Dim lvC As Double, lvP As Double, dvC As Double, dvP As Double
    For Each kk In myOI.Keys
        Kn = CDbl(Split(kk, "|")(0)): cpS = Split(kk, "|")(1)
        If cpS = "C" Then tc = tc + myOI(kk) Else tp = tp + myOI(kk)
        If Abs(Kn - F) <= 2500 Then
            If cpS = "C" Then nc = nc + myOI(kk) Else np = np + myOI(kk)
        End If
    Next
    For Each kk In live.Keys
        arr = live(kk)
        If Right$(kk, 1) = "C" Then lvC = lvC + arr(1) Else lvP = lvP + arr(1)
    Next
    For Each kk In vol.Keys
        If Left$(kk, Len(pre)) = pre Then
            If Right$(kk, 1) = "C" Then dvC = dvC + vol(kk) Else dvP = dvP + vol(kk)
        End If
    Next

    ' 最大痛點與未平倉牆
    Dim allKd As Object, allK As Variant, pay As Object, S As Double, tot As Double, mpK As Variant, mpV As Double
    Set allKd = NewDict()
    For Each kk In myOI.Keys
        allKd(CStr(Split(kk, "|")(0))) = CLng(Split(kk, "|")(0))
    Next
    Set pay = NewDict()
    mpK = Empty
    If allKd.Count > 0 Then
        Dim nO As Long, q As Long, oK() As Double, oC() As Boolean, oV() As Double
        nO = myOI.Count
        ReDim oK(1 To nO): ReDim oC(1 To nO): ReDim oV(1 To nO)
        q = 0
        For Each kk In myOI.Keys
            q = q + 1
            oK(q) = CDbl(Split(kk, "|")(0))
            oC(q) = (Right$(kk, 1) = "C")
            oV(q) = myOI(kk)
        Next
        allK = allKd.Items
        SortArr allK
        For i = 0 To UBound(allK)
            S = allK(i): tot = 0
            For q = 1 To nO
                If oC(q) Then
                    If S > oK(q) Then tot = tot + oV(q) * (S - oK(q))
                Else
                    If oK(q) > S Then tot = tot + oV(q) * (oK(q) - S)
                End If
            Next
            pay(CStr(allK(i))) = tot
            If IsEmpty(mpK) Or tot < mpV Then mpK = allK(i): mpV = tot
        Next
    End If
    Dim cwK As Variant, cwV As Double, pwK As Variant, pwV As Double
    cwK = Empty: pwK = Empty
    For i = 0 To UBound(Ks)
        If Abs(Ks(i) - F) <= 3000 Then
            If DGet(myOI, Ks(i) & "|C", 0) >= cwV And DGet(myOI, Ks(i) & "|C", 0) > 0 Then cwV = myOI(Ks(i) & "|C"): cwK = Ks(i)
            If DGet(myOI, Ks(i) & "|P", 0) >= pwV And DGet(myOI, Ks(i) & "|P", 0) > 0 Then pwV = myOI(Ks(i) & "|P"): pwK = Ks(i)
        End If
    Next

    Dim d As Object, lbl As String, shName As String
    Set d = NewDict()
    If IsNumeric(code) Then
        lbl = "月選（" & Month(ed) & "/" & Day(ed) & "）"
        shName = "月選 " & Month(ed) & "-" & Day(ed)
    Else
        lbl = "週選 " & Right$(code, 2) & "（" & Month(ed) & "/" & Day(ed) & "）"
        shName = "週選" & Right$(code, 2) & " " & Month(ed) & "-" & Day(ed)
    End If
    d("expiry") = code: d("label") = lbl: d("sheet") = shName
    d("expiry_date") = ed: d("days_left") = T * 365
    d("trade_date") = DateFromSlash(gTradeDate): d("quote_time") = qt: d("price_basis") = sessName
    d("forward") = F: d("atm_strike") = K0: d("atm_call") = c0: d("atm_put") = p0
    d("straddle") = straddle: d("move_pct") = straddle / F
    d("range_low") = F - straddle: d("range_high") = F + straddle
    d("atm_iv") = atmIV
    If IsEmpty(atmIV) Then d("sigma_pts") = Empty Else d("sigma_pts") = F * atmIV * Sqr(T)
    d("call_oi") = tc: d("put_oi") = tp
    If tc > 0 Then d("pc_oi") = tp / tc Else d("pc_oi") = Empty
    If nc > 0 Then d("pc_oi_near") = np / nc Else d("pc_oi_near") = Empty
    If dvC > 0 Then d("pc_vol_day") = dvP / dvC Else d("pc_vol_day") = Empty
    If lvC > 0 Then d("pc_vol_live") = lvP / lvC Else d("pc_vol_live") = Empty
    d("max_pain") = mpK
    d("call_wall") = cwK: d("call_wall_oi") = cwV
    d("put_wall") = pwK: d("put_wall_oi") = pwV
    gSum.Add d

    ' 履約價表（合成期貨 ±3000 點）
    Dim n As Long, tb() As Variant, row As Long, civ As Variant, piv As Variant
    For i = 0 To UBound(Ks)
        If Abs(Ks(i) - F) <= 3000 Then n = n + 1
    Next
    ReDim tb(1 To Mx(n, 1), 1 To 14)
    For i = 0 To UBound(Ks)
        If Abs(Ks(i) - F) <= 3000 Then
            row = row + 1
            cpx = DGet(price, Ks(i) & "|C"): ppx = DGet(price, Ks(i) & "|P")
            civ = ImpliedVol(cpx, F, CDbl(Ks(i)), T, "C")
            piv = ImpliedVol(ppx, F, CDbl(Ks(i)), T, "P")
            tb(row, 1) = Ks(i)
            tb(row, 2) = DGet(myOI, Ks(i) & "|C", 0)
            tb(row, 3) = DGet(myOI, Ks(i) & "|P", 0)
            tb(row, 4) = -tb(row, 3)
            If gPrevDate <> "" Then
                tb(row, 5) = tb(row, 2) - DGet(prevOI, pre & Ks(i) & "|C", 0)
                tb(row, 6) = tb(row, 3) - DGet(prevOI, pre & Ks(i) & "|P", 0)
            End If
            tb(row, 7) = cpx: tb(row, 8) = ppx
            tb(row, 9) = civ: tb(row, 10) = piv
            If Ks(i) >= F Then tb(row, 11) = civ Else tb(row, 11) = piv
            tb(row, 12) = DGet(vol, pre & Ks(i) & "|C", 0)
            tb(row, 13) = DGet(vol, pre & Ks(i) & "|P", 0)
            If pay.Exists(CStr(Ks(i))) Then tb(row, 14) = pay(CStr(Ks(i))) * 50 / 1000000#
        End If
    Next
    gTab.Add tb
End Sub

' ============================================================
' 2. 全市場 P/C 歷史
' ============================================================
Private Sub FetchPC()
    Dim c As Collection, r As Variant, n As Long, out() As Variant, i As Long
    Set c = TaifexCsv("https://www.taifex.com.tw/cht/3/pcRatioDown", "", gNow - 30, gNow)
    For Each r In c
        If UBound(r) >= 6 Then n = n + 1
    Next
    If n = 0 Then Err.Raise vbObjectError + 10, , "沒有資料"
    ReDim out(1 To n, 1 To 5)
    i = 0
    For Each r In c
        If UBound(r) >= 6 Then
            i = i + 1
            out(i, 1) = DateFromSlash(r(0))
            out(i, 2) = Val(r(6)) / 100
            out(i, 3) = Val(r(3)) / 100
            out(i, 4) = Val(r(4))
            out(i, 5) = Val(r(5))
        End If
    Next
    SortRows out, 1
    gPC = out
End Sub

Private Sub SortRows(a As Variant, ByVal col As Long)
    Dim i As Long, j As Long, c As Long, tmp As Variant
    For i = LBound(a, 1) + 1 To UBound(a, 1)
        j = i
        Do While j > LBound(a, 1)
            If a(j - 1, col) <= a(j, col) Then Exit Do
            For c = LBound(a, 2) To UBound(a, 2)
                tmp = a(j, c): a(j, c) = a(j - 1, c): a(j - 1, c) = tmp
            Next
            j = j - 1
        Loop
    Next
End Sub

' ============================================================
' 3. 三大法人選擇權（買賣權分計）
' ============================================================
Private Function WhoKey(ByVal s As String) As String
    s = Trim$(s)
    If s = "外資及陸資" Or s = "外資" Then
        WhoKey = "foreign"
    ElseIf s = "自營商" Then
        WhoKey = "dealer"
    ElseIf s = "投信" Then
        WhoKey = "trust"
    End If
End Function

Private Sub FetchInst()
    Dim c As Collection, r As Variant, who As String, cp As String, d As Object, dt As String
    Set gInst = NewDict()
    Set c = TaifexCsv("https://www.taifex.com.tw/cht/3/callsAndPutsDateDown", "commodityId=TXO", gNow - 30, gNow)
    For Each r In c
        If UBound(r) >= 15 Then
            who = WhoKey(r(3))
            If who <> "" Then
                dt = Trim$(r(0))
                If Not gInst.Exists(dt) Then gInst.Add dt, NewDict()
                Set d = gInst(dt)
                If UCase$(Trim$(r(2))) = "CALL" Then cp = "call" Else cp = "put"
                d(who & "_" & cp & "_oi") = Val(r(14))
                d(who & "_" & cp & "_amt") = Val(r(15)) / 100000#
                d(who & "_" & cp & "_trade") = Val(r(8))
            End If
        End If
    Next
    If gInst.Count = 0 Then Err.Raise vbObjectError + 11, , "沒有資料"
    Dim kv As Variant
    For Each kv In gInst.Keys
        Set d = gInst(kv)
        If d.Exists("foreign_call_amt") And d.Exists("foreign_put_amt") Then d("foreign_bias_amt") = d("foreign_call_amt") - d("foreign_put_amt")
    Next
End Sub

' ============================================================
' 4. 大額交易人（全部月份合計）
' ============================================================
Private Sub FetchLarge()
    Dim c As Collection, r As Variant, cp As String, kind As String, d As Object, dt As String
    Set gLarge = NewDict()
    Set c = TaifexCsv("https://www.taifex.com.tw/cht/3/largeTraderOptDown", "", gNow - 10, gNow)
    For Each r In c
        If UBound(r) >= 10 Then
            If Trim$(r(1)) = "TXO" And Trim$(r(4)) = "999999" Then
                dt = Trim$(r(0))
                If Not gLarge.Exists(dt) Then gLarge.Add dt, NewDict()
                Set d = gLarge(dt)
                If Trim$(r(3)) = "買權" Then cp = "call" Else cp = "put"
                If Trim$(r(5)) = "0" Then kind = "all" Else kind = "inst"
                d(kind & "_" & cp & "_top5_net") = Val(r(6)) - Val(r(7))
                d(kind & "_" & cp & "_top10_net") = Val(r(8)) - Val(r(9))
                d(cp & "_market_oi") = Val(r(10))
            End If
        End If
    Next
    If gLarge.Count = 0 Then Err.Raise vbObjectError + 12, , "沒有資料"
End Sub

' ============================================================
' 5. 台指期：三大法人淨部位、即時價格與正逆價差
' ============================================================
Private Sub FetchFut()
    Dim c As Collection, r As Variant, who As String, d As Object, dt As String
    Set gFut = NewDict()
    Set c = TaifexCsv("https://www.taifex.com.tw/cht/3/futContractsDateDown", "commodityId=TXF", gNow - 30, gNow)
    For Each r In c
        If UBound(r) >= 14 Then
            who = WhoKey(r(2))
            If who <> "" Then
                dt = Trim$(r(0))
                If Not gFut.Exists(dt) Then gFut.Add dt, NewDict()
                Set d = gFut(dt)
                d(who & "_net_oi") = Val(r(13))
                d(who & "_net_trade") = Val(r(7))
            End If
        End If
    Next
    If gFut.Count = 0 Then Err.Raise vbObjectError + 13, , "沒有資料"
End Sub

Private Sub FetchFutLive()
    Dim mt As Variant, body As String, root As Object, q As Object, x As Variant, nm As String
    Dim spot As Variant, f As Object, d As Object, lastP As Variant, refP As Variant
    Set gFutLive = NewDict()
    For Each mt In Array("0", "1")
        If mt = "0" Then nm = "日盤" Else nm = "夜盤"
        body = "{""MarketType"":""" & mt & """,""SymbolType"":""F"",""KindID"":""1"",""CID"":""TXF"",""ExpireMonth"":""""," & _
               """RowSize"":""全部"",""PageNo"":"""",""SortColumn"":"""",""AscDesc"":""A""}"
        Set root = JsonObj(Http("https://mis.taifex.com.tw/futures/api/getQuoteList", body, True, "utf-8"))
        Set q = root("RtData")("QuoteList")
        spot = Empty
        Set f = Nothing
        For Each x In q
            If InStr(CStr(NoNull(x("DispCName"))), "現貨") > 0 Then
                spot = TxtNum(x("CLastPrice"))
                If IsEmpty(spot) Then spot = TxtNum(x("CRefPrice"))
            ElseIf f Is Nothing And Left$(CStr(NoNull(x("SymbolID"))), 3) = "TXF" Then
                Set f = x
            End If
        Next
        If Not f Is Nothing Then
            Set d = NewDict()
            lastP = TxtNum(f("CLastPrice")): refP = TxtNum(f("CRefPrice"))
            d("name") = f("DispCName"): d("last") = lastP: d("ref") = refP
            If Not IsEmpty(lastP) And Not IsEmpty(refP) Then d("chg") = lastP - refP
            d("time") = CStr(NoNull(f("CDate"))) & " " & CStr(NoNull(f("CTime")))
            d("spot") = spot
            If nm = "日盤" And Not IsEmpty(lastP) And Not IsEmpty(spot) Then d("basis") = lastP - spot
            gFutLive.Add nm, d
        End If
    Next
End Sub

' ============================================================
' 6. 台指 VIX（期交所只留 3 個月，存在隱藏工作表累積）與 20 日歷史波動
' ============================================================
Private Function YahooChart(ByVal sym As String, ByVal rng As String) As Object
    Dim u As String, host As Variant, root As Object
    u = Replace(Replace(sym, "^", "%5E"), "=", "%3D")
    For Each host In Array("query1", "query2")
        On Error Resume Next
        Err.Clear
        Set root = JsonObj(Http("https://" & host & ".finance.yahoo.com/v8/finance/chart/" & u & "?range=" & rng & "&interval=1d", , , "utf-8"))
        If Err.Number = 0 And Not root Is Nothing Then
            On Error GoTo 0
            Exit For
        End If
        Set root = Nothing
        On Error GoTo 0
    Next
    If root Is Nothing Then Err.Raise vbObjectError + 20, , "Yahoo " & sym
    Set YahooChart = root("chart")("result")(1)
End Function

Private Function UnixToTW(ByVal ts As Double) As Date
    UnixToTW = DateSerial(1970, 1, 1) + (ts + 8 * 3600#) / 86400#
End Function

Private Function VixSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_vix")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "_vix"
        ws.Range("A1:B1").Value = Array("date", "vix")
        ws.Visible = 0
    End If
    Set VixSheet = ws
End Function

Private Sub FetchVol()
    Dim vix As Object, ws As Worksheet, i As Long, last As Long, m As Long, ym As String, txt As String
    Dim ln As Variant, tok As Variant, cutoff As Date, kv As Variant
    Set vix = NewDict()
    Set ws = VixSheet()
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    For i = 2 To last
        If IsDate(ws.Cells(i, 1).Value) Then vix(Format$(ws.Cells(i, 1).Value, "yyyymmdd")) = ws.Cells(i, 2).Value
    Next
    For m = 0 To 2
        ym = Format$(DateSerial(Year(gNow), Month(gNow) - m, 1), "yyyymm")
        txt = ""
        On Error Resume Next
        txt = Http("https://www.taifex.com.tw/file/taifex/Dailydownload/vix/log2data/" & ym & "new.txt")
        On Error GoTo 0
        For Each ln In Split(Replace(txt, vbCr, ""), vbLf)
            tok = Split(Application.WorksheetFunction.Trim(Replace(ln, vbTab, " ")), " ")
            If UBound(tok) >= 2 Then
                If Len(tok(0)) = 8 And Left$(tok(0), 2) = "20" Then
                    If IsNumeric(tok(0)) And IsNumeric(tok(2)) Then vix(tok(0)) = Val(tok(2))
                End If
            End If
        Next
    Next
    cutoff = Date - 366

    ' 加權指數日收盤 -> 20 日歷史波動
    Dim yc As Object, ts As Object, cl As Object, nP As Long, pxD() As Date, pxV() As Double, hv As Object
    Set hv = NewDict()
    On Error Resume Next
    Set yc = YahooChart("^TWII", "2y")
    On Error GoTo 0
    If Not yc Is Nothing Then
        Set ts = yc("timestamp")
        Set cl = yc("indicators")("quote")(1)("close")
        ReDim pxD(1 To ts.Count): ReDim pxV(1 To ts.Count)
        For i = 1 To ts.Count
            If Not IsNull(cl(i)) Then
                If cl(i) > 0 Then
                    nP = nP + 1
                    pxD(nP) = Int(UnixToTW(ts(i)))
                    pxV(nP) = cl(i)
                End If
            End If
        Next
        Dim rets() As Double, j As Long, mu As Double, sd As Double
        If nP > 21 Then
            ReDim rets(1 To nP - 1)
            For i = 2 To nP
                rets(i - 1) = Log(pxV(i) / pxV(i - 1))
            Next
            For i = 21 To nP
                ' rets(i-20 .. i-1) 對應到 pxD(i)
                mu = 0
                For j = i - 20 To i - 1: mu = mu + rets(j): Next
                mu = mu / 20
                sd = 0
                For j = i - 20 To i - 1: sd = sd + (rets(j) - mu) ^ 2: Next
                sd = Sqr(sd / 19)
                If pxD(i) >= cutoff Then hv(Format$(pxD(i), "yyyymmdd")) = sd * Sqr(252)
            Next
        End If
    Else
        gErr = gErr & "HV：Yahoo 加權指數抓不到" & vbLf
    End If

    ' 寫回 VIX 累積表（只留一年）
    Dim keep As Object, ks As Variant
    Set keep = NewDict()
    For Each kv In vix.Keys
        If DateFromYmd(kv) >= cutoff Then keep(kv) = vix(kv)
    Next
    ks = SortedKeys(keep)
    ws.Range("A2:B5000").ClearContents
    If keep.Count > 0 Then
        Dim vo() As Variant
        ReDim vo(1 To keep.Count, 1 To 2)
        For i = 0 To UBound(ks)
            vo(i + 1, 1) = DateFromYmd(ks(i)): vo(i + 1, 2) = keep(ks(i))
        Next
        ws.Range("A2").Resize(keep.Count, 2).Value = vo
        ws.Range("A2").Resize(keep.Count, 1).NumberFormat = "yyyy-mm-dd"
    End If

    ' 合併成圖表用的表：日期、VIX、HV20
    Dim allD As Object, ad As Variant
    Set allD = NewDict()
    For Each kv In keep.Keys: allD(kv) = 1: Next
    For Each kv In hv.Keys: allD(kv) = 1: Next
    ad = SortedKeys(allD)
    If allD.Count > 0 Then
        Dim tb() As Variant
        ReDim tb(1 To allD.Count, 1 To 3)
        For i = 0 To UBound(ad)
            tb(i + 1, 1) = DateFromYmd(ad(i))
            If keep.Exists(ad(i)) Then tb(i + 1, 2) = keep(ad(i)) / 100
            If hv.Exists(ad(i)) Then tb(i + 1, 3) = hv(ad(i))
        Next
        gVolTab = tb
    Else
        gVolTab = Empty
    End If

    ' 統計
    Set gVol = NewDict()
    Dim vs As Variant, hs As Variant, cur As Double
    If keep.Count > 0 Then
        vs = keep.Items
        cur = keep(ks(UBound(ks)))
        gVol("vix") = cur
        gVol("vix_date") = DateFromYmd(ks(UBound(ks)))
        gVol("vix_from") = DateFromYmd(ks(0))
        gVol("vix_days") = keep.Count
        gVol("vix_min") = Application.WorksheetFunction.Min(vs)
        gVol("vix_max") = Application.WorksheetFunction.Max(vs)
        RankInto vs, cur, "vix"
    End If
    If hv.Count > 0 Then
        Dim hk As Variant
        hk = SortedKeys(hv)
        hs = hv.Items
        cur = hv(hk(UBound(hk)))
        gVol("hv20") = cur
        RankInto hs, cur, "hv"
    End If
    If gSum.Count > 0 Then
        gVol("atm_iv") = gSum(1)("atm_iv")
        If Not IsEmpty(gVol("atm_iv")) And gVol.Exists("hv20") Then gVol("iv_hv_ratio") = gVol("atm_iv") / gVol("hv20")
    End If
End Sub

Private Sub RankInto(ByVal vals As Variant, ByVal cur As Double, ByVal pfx As String)
    Dim lo As Double, hi As Double, i As Long, n As Long
    lo = Application.WorksheetFunction.Min(vals)
    hi = Application.WorksheetFunction.Max(vals)
    If hi > lo Then gVol(pfx & "_rank") = (cur - lo) / (hi - lo)
    For i = LBound(vals) To UBound(vals)
        If vals(i) <= cur Then n = n + 1
    Next
    gVol(pfx & "_pct") = n / (UBound(vals) - LBound(vals) + 1)
End Sub

' ============================================================
' 7. 美股期貨與 VIX
' ============================================================
Private Sub FetchGlobal()
    Dim syms As Variant, names As Variant, i As Long, yc As Object, d As Object, cl As Object
    Dim closes() As Double, n As Long, j As Long, lastP As Variant, prevP As Variant, failed As Long
    syms = Array("^VIX", "ES=F", "NQ=F", "YM=F", "^SOX", "TSM")
    names = Array("美國 VIX", "S&P 500 期貨", "那斯達克 100 期貨", "道瓊期貨", "費城半導體", "台積電 ADR")
    Set gGlobal = New Collection
    For i = 0 To UBound(syms)
        Set d = NewDict()
        d("name") = names(i)
        d("symbol") = syms(i)
        Set yc = Nothing
        On Error Resume Next
        Set yc = YahooChart(CStr(syms(i)), "1mo")
        lastP = Empty
        lastP = yc("meta")("regularMarketPrice")
        d("time") = UnixToTW(yc("meta")("regularMarketTime"))
        On Error GoTo 0
        If Not yc Is Nothing And Not IsEmpty(lastP) Then
            Set cl = yc("indicators")("quote")(1)("close")
            ReDim closes(1 To cl.Count + 1)
            n = 0
            For j = 1 To cl.Count
                If Not IsNull(cl(j)) Then n = n + 1: closes(n) = cl(j)
            Next
            prevP = Empty
            If n >= 2 Then prevP = closes(n - 1)
            d("last") = lastP
            d("prev") = prevP
            If Not IsEmpty(prevP) Then
                If prevP > 0 Then d("chg_pct") = lastP / prevP - 1
            End If
            Dim sp() As Variant, s0 As Long
            s0 = n - 19
            If s0 < 1 Then s0 = 1
            If n > 0 Then
                ReDim sp(1 To n - s0 + 1)
                For j = s0 To n: sp(j - s0 + 1) = closes(j): Next
                d("spark") = sp
            End If
        Else
            failed = failed + 1
        End If
        gGlobal.Add d
    Next
    If failed = UBound(syms) + 1 Then Err.Raise vbObjectError + 30, , "Yahoo Finance 抓不到"
End Sub

' ============================================================
' JSON 解析（只處理一般格式）
' ============================================================
Private Function JsonObj(ByVal s As String) As Object
    Dim v As Variant
    jS = s
    jP = 1
    JValue v
    Set JsonObj = v
End Function

Private Sub JSkip()
    Dim ch As String
    Do While jP <= Len(jS)
        ch = Mid$(jS, jP, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then jP = jP + 1 Else Exit Do
    Loop
End Sub

Private Sub JValue(ByRef v As Variant)
    Dim ch As String, d As Object, c As Collection, k As String, x As Variant, st As Long
    JSkip
    If jP > Len(jS) Then Err.Raise vbObjectError + 40, , "JSON 格式錯誤"
    ch = Mid$(jS, jP, 1)
    Select Case ch
    Case "{"
        Set d = NewDict()
        jP = jP + 1
        JSkip
        If Mid$(jS, jP, 1) = "}" Then
            jP = jP + 1
        Else
            Do
                JSkip
                k = JString()
                JSkip
                jP = jP + 1
                x = Empty
                JValue x
                If IsObject(x) Then Set d(k) = x Else d(k) = x
                JSkip
                ch = Mid$(jS, jP, 1)
                jP = jP + 1
                If ch <> "," Then Exit Do
            Loop
        End If
        Set v = d
    Case "["
        Set c = New Collection
        jP = jP + 1
        JSkip
        If Mid$(jS, jP, 1) = "]" Then
            jP = jP + 1
        Else
            Do
                x = Empty
                JValue x
                c.Add x
                JSkip
                ch = Mid$(jS, jP, 1)
                jP = jP + 1
                If ch <> "," Then Exit Do
            Loop
        End If
        Set v = c
    Case """"
        v = JString()
    Case "t"
        v = True: jP = jP + 4
    Case "f"
        v = False: jP = jP + 5
    Case "n"
        v = Null: jP = jP + 4
    Case Else
        st = jP
        Do While jP <= Len(jS)
            If InStr("+-0123456789.eE", Mid$(jS, jP, 1)) = 0 Then Exit Do
            jP = jP + 1
        Loop
        v = Val(Mid$(jS, st, jP - st))
    End Select
End Sub

Private Function JString() As String
    Dim out As String, ch As String, q As Long, b As Long, e As Long
    jP = jP + 1
    Do While jP <= Len(jS)
        ch = Mid$(jS, jP, 1)
        If ch = """" Then
            jP = jP + 1
            Exit Do
        ElseIf ch = "\" Then
            jP = jP + 1
            ch = Mid$(jS, jP, 1)
            Select Case ch
                Case "n": out = out & vbLf
                Case "t": out = out & vbTab
                Case "r": out = out & vbCr
                Case "b", "f"
                Case "u"
                    out = out & ChrW(CLng("&H" & Mid$(jS, jP + 1, 4) & "&"))
                    jP = jP + 4
                Case Else: out = out & ch
            End Select
            jP = jP + 1
        Else
            q = InStr(jP, jS, """")
            b = InStr(jP, jS, "\")
            e = q
            If b > 0 And (b < q Or q = 0) Then e = b
            If e = 0 Then e = Len(jS) + 1
            out = out & Mid$(jS, jP, e - jP)
            jP = e
        End If
    Loop
    JString = out
End Function

' ============================================================
' 輸出：工作表與圖表
' ============================================================
Private Function CallClr() As Long
    CallClr = RGB(194, 65, 45)
End Function
Private Function PutClr() As Long
    PutClr = RGB(47, 125, 91)
End Function
Private Function BlueClr() As Long
    BlueClr = RGB(58, 91, 191)
End Function
Private Function AmberClr() As Long
    AmberClr = RGB(183, 121, 31)
End Function
Private Function GreyClr() As Long
    GreyClr = RGB(150, 158, 168)
End Function

Private Sub RebuildSheets()
    Dim ws As Worksheet, i As Long, keep As Variant, nm As String, del As Collection, x As Variant
    Application.DisplayAlerts = False
    Set del = New Collection
    For Each ws In ThisWorkbook.Worksheets
        nm = ws.Name
        If nm <> "盤面" And nm <> "說明" And nm <> "_vix" Then del.Add ws
    Next
    For Each x In del
        x.Delete
    Next
    Application.DisplayAlerts = True

    Dim pos As Worksheet
    Set pos = ThisWorkbook.Worksheets("盤面")
    For i = 1 To gSum.Count
        Set pos = WriteExpiry(gSum(i), gTab(i), pos)
    Next
    Set pos = WriteChips(pos)
    Set pos = WriteVol(pos)
    Set pos = WriteGlobal(pos)
    Set pos = WritePC(pos)
    WriteDashboard
End Sub

Private Function AddSheet(ByVal nm As String, ByVal after As Worksheet) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=after)
    ws.Name = nm
    ActiveWindow.DisplayGridlines = False
    ws.Cells.Font.Size = 10
    Set AddSheet = ws
End Function

Private Function NewChart(ByVal ws As Worksheet, ByVal at As Range, ByVal w As Double, ByVal h As Double, ByVal title As String) As Object
    Dim co As Object
    Set co = ws.ChartObjects.Add(at.Left, at.Top, w, h)
    With co.Chart
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
    End With
    gChartTitle = title
    Set NewChart = co.Chart
End Function

Private Function AddSeries(ByVal ch As Object, ByVal nm As String, ByVal vals As Range, ByVal cats As Range, _
                           ByVal clr As Long, ByVal ctype As Long) As Object
    Dim s As Object
    Set s = ch.SeriesCollection.NewSeries
    s.Name = nm
    s.Values = vals
    s.XValues = cats
    s.ChartType = ctype
    If ch.SeriesCollection.Count = 1 Then
        ch.HasTitle = True
        ch.ChartTitle.Text = gChartTitle
        ch.ChartTitle.Font.Size = 12
        ch.HasLegend = True
        ch.Legend.Position = -4107
    End If
    If ctype = 4 Then
        s.Format.Line.Visible = True
        s.Format.Line.ForeColor.RGB = clr
        s.Format.Line.Weight = 2
        s.MarkerStyle = -4142
    Else
        s.Format.Fill.ForeColor.RGB = clr
        s.Format.Line.Visible = False
    End If
    Set AddSeries = s
End Function

Private Sub HeaderRow(ByVal rg As Range, ByVal heads As Variant)
    rg.Resize(1, UBound(heads) - LBound(heads) + 1).Value = heads
    With rg.Resize(1, UBound(heads) - LBound(heads) + 1)
        .Font.Bold = True
        .Interior.Color = RGB(236, 239, 243)
        .HorizontalAlignment = -4108
    End With
End Sub

Private Function Fx(ByVal v As Variant, ByVal fmt As String) As String
    If IsEmpty(v) Or IsNull(v) Then Fx = "—" Else Fx = Format$(v, fmt)
End Function

Private Function WriteExpiry(ByVal d As Object, ByVal tb As Variant, ByVal after As Worksheet) As Worksheet
    Dim ws As Worksheet, n As Long, r0 As Long, ch As Object
    Set ws = AddSheet(d("sheet"), after)
    n = UBound(tb, 1)
    ws.Range("A1").Value = d("label") & "　到期 " & Format$(d("expiry_date"), "yyyy-mm-dd") & "，剩 " & Format$(d("days_left"), "0.0") & " 天"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "預期波動 ±" & Fx(d("straddle"), "#,##0") & " 點（" & Fx(d("move_pct"), "0.0%") & "），區間 " & _
        Fx(d("range_low"), "#,##0") & "–" & Fx(d("range_high"), "#,##0") & "；ATM IV " & Fx(d("atm_iv"), "0.0%") & _
        "；買權牆 " & Fx(d("call_wall"), "#,##0") & "、賣權牆 " & Fx(d("put_wall"), "#,##0") & "；最大痛點 " & Fx(d("max_pain"), "#,##0") & "。"
    ws.Range("A3").Value = "報價 " & Format$(d("quote_time"), "yyyy-mm-dd hh:nn") & "（" & d("price_basis") & "中價）；未平倉量截至 " & _
        Format$(d("trade_date"), "yyyy-mm-dd") & " 日盤收盤；OI 增減與 " & gPrevDate & " 比較。"
    ws.Range("A3").Font.Color = RGB(93, 102, 115)
    r0 = 5
    HeaderRow ws.Cells(r0, 1), Array("履約價", "買權OI", "賣權OI", "賣權OI(負)", "買權OI增減", "賣權OI增減", "買權價", "賣權價", _
                                     "買權IV", "賣權IV", "價外IV", "買權量", "賣權量", "結算買方價值(百萬)")
    ws.Cells(r0 + 1, 1).Resize(n, 14).Value = tb
    ws.Range(ws.Cells(r0 + 1, 1), ws.Cells(r0 + n, 6)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(r0 + 1, 4), ws.Cells(r0 + n, 4)).NumberFormat = "#,##0;#,##0"
    ws.Range(ws.Cells(r0 + 1, 5), ws.Cells(r0 + n, 6)).NumberFormat = "+#,##0;-#,##0;0"
    ws.Range(ws.Cells(r0 + 1, 7), ws.Cells(r0 + n, 8)).NumberFormat = "#,##0.0"
    ws.Range(ws.Cells(r0 + 1, 9), ws.Cells(r0 + n, 11)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(r0 + 1, 12), ws.Cells(r0 + n, 13)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(r0 + 1, 14), ws.Cells(r0 + n, 14)).NumberFormat = "#,##0"
    ws.Columns("A:N").ColumnWidth = 10
    ws.Columns("N").ColumnWidth = 16
    ws.Columns("D").Hidden = True
    ws.Range("A1:A3").WrapText = False
    ' 標出價平列
    Dim i As Long
    For i = 1 To n
        If tb(i, 1) = d("atm_strike") Then ws.Range(ws.Cells(r0 + i, 1), ws.Cells(r0 + i, 14)).Interior.Color = RGB(255, 242, 204)
    Next
    ws.Activate
    ws.Cells(r0 + 1, 2).Select
    ActiveWindow.FreezePanes = True

    Dim cats As Range, at As Range, w As Double, tick As Long
    Set cats = ws.Range(ws.Cells(r0 + 1, 1), ws.Cells(r0 + n, 1))
    tick = Application.WorksheetFunction.Max(1, Int(n / 12))
    w = 640
    Set at = ws.Range("P5")
    Set ch = NewChart(ws, at, w, 300, "未平倉牆（買權向上、賣權向下）")
    AddSeries ch, "買權 OI", cats.Offset(0, 1), cats, CallClr(), 51
    AddSeries ch, "賣權 OI", cats.Offset(0, 3), cats, PutClr(), 51
    ch.ChartGroups(1).Overlap = 100
    ch.ChartGroups(1).GapWidth = 20
    ch.Axes(2).TickLabels.NumberFormat = "#,##0;#,##0"
    ch.Axes(1).TickLabelSpacing = tick
    ch.Axes(1).TickLabelPosition = -4134

    Set at = ws.Range("P5").Offset(22, 0)
    Set ch = NewChart(ws, at, w, 260, "未平倉增減（與前一交易日比較）")
    AddSeries ch, "買權 OI 增減", cats.Offset(0, 4), cats, CallClr(), 51
    AddSeries ch, "賣權 OI 增減", cats.Offset(0, 5), cats, PutClr(), 51
    ch.ChartGroups(1).GapWidth = 30
    ch.Axes(2).TickLabels.NumberFormat = "+#,##0;-#,##0;0"
    ch.Axes(1).TickLabelSpacing = tick
    ch.Axes(1).TickLabelPosition = -4134

    Set at = ws.Range("P5").Offset(41, 0)
    Set ch = NewChart(ws, at, w / 2 - 5, 260, "IV 微笑（價外）")
    AddSeries ch, "價外 IV", cats.Offset(0, 10), cats, BlueClr(), 4
    ch.HasLegend = False
    ch.Axes(2).TickLabels.NumberFormat = "0%"
    ch.Axes(1).TickLabelSpacing = tick * 2
    ch.DisplayBlanksAs = 3

    Set ch = NewChart(ws, at.Offset(0, 0), w / 2 - 5, 260, "最大痛點曲線（百萬元）")
    ch.Parent.Left = at.Left + w / 2 + 5
    AddSeries ch, "結算時買方總價值", cats.Offset(0, 13), cats, BlueClr(), 4
    ch.HasLegend = False
    ch.Axes(2).TickLabels.NumberFormat = "#,##0"
    ch.Axes(1).TickLabelSpacing = tick * 2
    ch.DisplayBlanksAs = 3

    ws.Range("A1").Select
    Set WriteExpiry = ws
End Function

Private Function Latest2(ByVal d As Object, ByRef a As Object, ByRef b As Object) As Boolean
    Dim ks As Variant
    Set a = Nothing: Set b = Nothing
    If d Is Nothing Then Exit Function
    If d.Count = 0 Then Exit Function
    ks = SortedKeys(d)
    Set a = d(ks(UBound(ks)))
    If UBound(ks) >= 1 Then Set b = d(ks(UBound(ks) - 1))
    Latest2 = True
End Function

Private Function LatestKey(ByVal d As Object) As String
    Dim ks As Variant
    If d Is Nothing Then Exit Function
    If d.Count = 0 Then Exit Function
    ks = SortedKeys(d)
    LatestKey = ks(UBound(ks))
End Function

Private Function DDiff(ByVal a As Object, ByVal b As Object, ByVal k As String) As Variant
    DDiff = Empty
    If a Is Nothing Or b Is Nothing Then Exit Function
    If a.Exists(k) And b.Exists(k) Then DDiff = a(k) - b(k)
End Function

Private Function WriteChips(ByVal after As Worksheet) As Worksheet
    Dim ws As Worksheet, ks As Variant, i As Long, n As Long, d As Object, r As Long, ch As Object, out() As Variant
    Set ws = AddSheet("籌碼", after)
    ws.Range("A1").Value = "籌碼：三大法人、大額交易人、台指期"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B:K").ColumnWidth = 12

    ' 三大法人選擇權
    r = 3
    ws.Cells(r, 1).Value = "三大法人選擇權未平倉淨額（買方 - 賣方；金額單位：億元）"
    ws.Cells(r, 1).Font.Bold = True
    HeaderRow ws.Cells(r + 1, 1), Array("日期", "外資買權淨口數", "外資買權淨額", "外資賣權淨口數", "外資賣權淨額", "外資方向", _
                                        "自營商買權淨額", "自營商賣權淨額", "投信買權淨額", "投信賣權淨額")
    If Not gInst Is Nothing Then
        If gInst.Count > 0 Then
            ks = SortedKeys(gInst)
            n = UBound(ks) + 1
            ReDim out(1 To n, 1 To 10)
            For i = 0 To UBound(ks)
                Set d = gInst(ks(i))
                out(i + 1, 1) = DateFromSlash(ks(i))
                out(i + 1, 2) = DGet(d, "foreign_call_oi"): out(i + 1, 3) = DGet(d, "foreign_call_amt")
                out(i + 1, 4) = DGet(d, "foreign_put_oi"): out(i + 1, 5) = DGet(d, "foreign_put_amt")
                out(i + 1, 6) = DGet(d, "foreign_bias_amt")
                out(i + 1, 7) = DGet(d, "dealer_call_amt"): out(i + 1, 8) = DGet(d, "dealer_put_amt")
                out(i + 1, 9) = DGet(d, "trust_call_amt"): out(i + 1, 10) = DGet(d, "trust_put_amt")
            Next
            ws.Cells(r + 2, 1).Resize(n, 10).Value = out
            ws.Cells(r + 2, 1).Resize(n, 1).NumberFormat = "yyyy-mm-dd"
            ws.Cells(r + 2, 2).Resize(n, 9).NumberFormat = "+#,##0.00;-#,##0.00;0"
            ws.Cells(r + 2, 2).Resize(n, 1).NumberFormat = "+#,##0;-#,##0;0"
            ws.Cells(r + 2, 4).Resize(n, 1).NumberFormat = "+#,##0;-#,##0;0"
            Set ch = NewChart(ws, ws.Range("M3"), 520, 260, "外資選擇權淨部位（億元）")
            AddSeries ch, "外資買權", ws.Cells(r + 2, 3).Resize(n, 1), ws.Cells(r + 2, 1).Resize(n, 1), CallClr(), 4
            AddSeries ch, "外資賣權", ws.Cells(r + 2, 5).Resize(n, 1), ws.Cells(r + 2, 1).Resize(n, 1), PutClr(), 4
            AddSeries ch, "外資方向（買權 - 賣權）", ws.Cells(r + 2, 6).Resize(n, 1), ws.Cells(r + 2, 1).Resize(n, 1), BlueClr(), 4
            ch.Axes(1).CategoryType = 2
            ch.Axes(1).TickLabels.NumberFormat = "m/d"
            r = r + 2 + n + 2
        Else
            r = r + 4
        End If
    Else
        ws.Cells(r + 2, 1).Value = "（抓取失敗）"
        r = r + 4
    End If
    If r < 22 Then r = 22

    ' 大額交易人
    Dim a As Object, b As Object, rowsL As Variant, j As Long, k As String
    ws.Cells(r, 1).Value = "大額交易人選擇權淨部位（全部到期合計，買方 - 賣方，口）"
    ws.Cells(r, 1).Font.Bold = True
    HeaderRow ws.Cells(r + 1, 1), Array("交易人", "買權淨部位", "較前日", "賣權淨部位", "較前日")
    If Latest2(gLarge, a, b) Then
        rowsL = Array(Array("前五大", "all", "top5"), Array("前十大", "all", "top10"), _
                      Array("前五大特定法人", "inst", "top5"), Array("前十大特定法人", "inst", "top10"))
        For i = 0 To 3
            ws.Cells(r + 2 + i, 1).Value = rowsL(i)(0)
            k = rowsL(i)(1) & "_call_" & rowsL(i)(2) & "_net"
            ws.Cells(r + 2 + i, 2).Value = DGet(a, k): ws.Cells(r + 2 + i, 3).Value = DDiff(a, b, k)
            k = rowsL(i)(1) & "_put_" & rowsL(i)(2) & "_net"
            ws.Cells(r + 2 + i, 4).Value = DGet(a, k): ws.Cells(r + 2 + i, 5).Value = DDiff(a, b, k)
        Next
        ws.Cells(r + 2, 2).Resize(4, 4).NumberFormat = "+#,##0;-#,##0;0"
        ws.Cells(r + 6, 1).Value = "資料日期 " & LatestKey(gLarge) & "；全市場未沖銷：買權 " & Fx(DGet(a, "call_market_oi"), "#,##0") & _
            " 口、賣權 " & Fx(DGet(a, "put_market_oi"), "#,##0") & " 口。買權淨部位為負代表大戶整體偏賣買權（看不漲）。"
        ws.Cells(r + 6, 1).Font.Color = RGB(93, 102, 115)
    Else
        ws.Cells(r + 2, 1).Value = "（抓取失敗）"
    End If
    r = r + 9

    ' 台指期即時
    ws.Cells(r, 1).Value = "台指期近月"
    ws.Cells(r, 1).Font.Bold = True
    HeaderRow ws.Cells(r + 1, 1), Array("盤別", "合約", "最新", "漲跌", "現貨", "正逆價差", "時間")
    If Not gFutLive Is Nothing Then
        Dim kv As Variant
        i = 0
        For Each kv In gFutLive.Keys
            Set d = gFutLive(kv)
            ws.Cells(r + 2 + i, 1).Resize(1, 7).Value = Array(kv, DGet(d, "name"), DGet(d, "last"), DGet(d, "chg"), DGet(d, "spot"), DGet(d, "basis"), DGet(d, "time"))
            i = i + 1
        Next
        ws.Cells(r + 2, 3).Resize(2, 1).NumberFormat = "#,##0"
        ws.Cells(r + 2, 4).Resize(2, 1).NumberFormat = "+#,##0;-#,##0;0"
        ws.Cells(r + 2, 5).Resize(2, 1).NumberFormat = "#,##0.00"
        ws.Cells(r + 2, 6).Resize(2, 1).NumberFormat = "+#,##0.00;-#,##0.00;0"
        ws.Cells(r + 4, 1).Value = "正逆價差只算日盤（近月減加權指數）；夜盤時現貨已收盤。"
        ws.Cells(r + 4, 1).Font.Color = RGB(93, 102, 115)
    End If
    r = r + 6

    ' 期貨三大法人
    ws.Cells(r, 1).Value = "臺股期貨三大法人未平倉淨口數"
    ws.Cells(r, 1).Font.Bold = True
    HeaderRow ws.Cells(r + 1, 1), Array("日期", "外資淨未平倉", "投信淨未平倉", "自營商淨未平倉", "外資當日淨交易")
    If Not gFut Is Nothing Then
        If gFut.Count > 0 Then
            ks = SortedKeys(gFut)
            n = UBound(ks) + 1
            ReDim out(1 To n, 1 To 5)
            For i = 0 To UBound(ks)
                Set d = gFut(ks(i))
                out(i + 1, 1) = DateFromSlash(ks(i))
                out(i + 1, 2) = DGet(d, "foreign_net_oi"): out(i + 1, 3) = DGet(d, "trust_net_oi")
                out(i + 1, 4) = DGet(d, "dealer_net_oi"): out(i + 1, 5) = DGet(d, "foreign_net_trade")
            Next
            ws.Cells(r + 2, 1).Resize(n, 5).Value = out
            ws.Cells(r + 2, 1).Resize(n, 1).NumberFormat = "yyyy-mm-dd"
            ws.Cells(r + 2, 2).Resize(n, 4).NumberFormat = "+#,##0;-#,##0;0"
            Set ch = NewChart(ws, ws.Cells(r, 13), 520, 260, "臺股期貨法人淨未平倉（口）")
            AddSeries ch, "外資", ws.Cells(r + 2, 2).Resize(n, 1), ws.Cells(r + 2, 1).Resize(n, 1), BlueClr(), 4
            AddSeries ch, "投信", ws.Cells(r + 2, 3).Resize(n, 1), ws.Cells(r + 2, 1).Resize(n, 1), AmberClr(), 4
            ch.Axes(1).CategoryType = 2
            ch.Axes(1).TickLabels.NumberFormat = "m/d"
            ch.Axes(2).TickLabels.NumberFormat = "#,##0"
        End If
    End If
    ws.Range("A1").Select
    Set WriteChips = ws
End Function

Private Function WriteVol(ByVal after As Worksheet) As Worksheet
    Dim ws As Worksheet, n As Long, ch As Object
    Set ws = AddSheet("波動", after)
    ws.Range("A1").Value = "波動率位置：台指 VIX 與加權指數 20 日歷史波動"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("D:F").ColumnWidth = 12
    If Not gVol Is Nothing Then
        Dim lab As Variant, vals As Variant, q As Long
        lab = Array("台指 VIX", "VIX 日期", "IV Rank（區間位置）", "百分位", "區間", "區間起始日（交易日數）", _
                    "20 日歷史波動 HV", "HV 一年百分位", "最近週選 ATM IV", "ATM IV ÷ HV")
        vals = Array(DGet(gVol, "vix"), DGet(gVol, "vix_date"), DGet(gVol, "vix_rank"), DGet(gVol, "vix_pct"), _
                     Fx(DGet(gVol, "vix_min"), "0.00") & "–" & Fx(DGet(gVol, "vix_max"), "0.00"), _
                     Fx(DGet(gVol, "vix_from"), "yyyy-mm-dd") & "（" & DGet(gVol, "vix_days") & "）", _
                     DGet(gVol, "hv20"), DGet(gVol, "hv_pct"), DGet(gVol, "atm_iv"), DGet(gVol, "iv_hv_ratio"))
        For q = 0 To 9
            ws.Cells(3 + q, 1).Value = lab(q)
            ws.Cells(3 + q, 2).Value = vals(q)
        Next
        ws.Range("B3").NumberFormat = "0.00"
        ws.Range("B4").NumberFormat = "yyyy-mm-dd"
        ws.Range("B5:B6").NumberFormat = "0.0%"
        ws.Range("B9:B11").NumberFormat = "0.0%"
        ws.Range("B12").NumberFormat = "0.00"
        ws.Range("B3:B12").HorizontalAlignment = -4152
        ws.Range("A14").Value = "IV Rank =（今天 - 區間最低）÷（最高 - 最低）。期交所只提供近 3 個月 VIX，每次更新會存進隱藏的 _vix 表，存檔後會一直累積到一年。"
        ws.Range("A15").Value = "ATM IV ÷ HV 大於 1.2 代表選擇權比實際波動貴、對賣方有利；小於 0.9 代表便宜、對買方有利。"
        ws.Range("A14:A15").Font.Color = RGB(93, 102, 115)
    End If
    If Not IsEmpty(gVolTab) Then
        n = UBound(gVolTab, 1)
        HeaderRow ws.Range("D2"), Array("日期", "台指 VIX", "20 日 HV")
        ws.Range("D3").Resize(n, 3).Value = gVolTab
        ws.Range("D3").Resize(n, 1).NumberFormat = "yyyy-mm-dd"
        ws.Range("E3").Resize(n, 2).NumberFormat = "0.0%"
        Set ch = NewChart(ws, ws.Range("H2"), 560, 300, "台指 VIX 與 20 日歷史波動")
        AddSeries ch, "台指 VIX", ws.Range("E3").Resize(n, 1), ws.Range("D3").Resize(n, 1), BlueClr(), 4
        AddSeries ch, "20 日 HV", ws.Range("F3").Resize(n, 1), ws.Range("D3").Resize(n, 1), AmberClr(), 4
        ch.Axes(1).CategoryType = 3
        ch.Axes(1).TickLabels.NumberFormat = "yy/m"
        ch.Axes(2).TickLabels.NumberFormat = "0%"
        ch.DisplayBlanksAs = 3
    End If
    ws.Range("A1").Select
    Set WriteVol = ws
End Function

Private Function WriteGlobal(ByVal after As Worksheet) As Worksheet
    Dim ws As Worksheet, i As Long, d As Object, sp As Variant, r As Long, sg As Object
    Set ws = AddSheet("美股", after)
    ws.Range("A1").Value = "美股期貨與 VIX（Yahoo Finance，可能延遲）"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    HeaderRow ws.Range("A3"), Array("商品", "代號", "最新", "前一交易日", "漲跌", "近 20 日", "時間（台北）")
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B:E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 16
    If Not gGlobal Is Nothing Then
        For i = 1 To gGlobal.Count
            Set d = gGlobal(i)
            r = 3 + i
            ws.Cells(r, 1).Value = d("name")
            ws.Cells(r, 2).Value = d("symbol")
            ws.Cells(r, 3).Value = DGet(d, "last")
            ws.Cells(r, 4).Value = DGet(d, "prev")
            ws.Cells(r, 5).Value = DGet(d, "chg_pct")
            ws.Cells(r, 7).Value = DGet(d, "time")
            If d.Exists("spark") Then
                sp = d("spark")
                ws.Cells(r, 10).Resize(1, UBound(sp)).Value = sp
                Set sg = ws.Cells(r, 6).SparklineGroups.Add(1, "'" & ws.Name & "'!" & ws.Cells(r, 10).Resize(1, UBound(sp)).Address)
                sg.DisplayHidden = True
                sg.SeriesColor.Color = BlueClr()
            End If
        Next
        ws.Range("C4:D9").NumberFormat = "#,##0.00"
        ws.Range("E4:E9").NumberFormat = "+0.00%;-0.00%;0.00%"
        ws.Range("G4:G9").NumberFormat = "mm-dd hh:mm"
        ws.Range("E4:E9").FormatConditions.Delete
        With ws.Range("E4:E9").FormatConditions.Add(1, 5, "=0")
            .Font.Color = CallClr()
        End With
        With ws.Range("E4:E9").FormatConditions.Add(1, 6, "=0")
            .Font.Color = PutClr()
        End With
        ws.Range("J:AF").EntireColumn.Hidden = True
        ws.Range("A11").Value = "漲跌相對前一交易日收盤（紅漲綠跌）。台指夜盤常跟著美股期貨走。"
        ws.Range("A11").Font.Color = RGB(93, 102, 115)
    End If
    ws.Range("A1").Select
    Set WriteGlobal = ws
End Function

Private Function WritePC(ByVal after As Worksheet) As Worksheet
    Dim ws As Worksheet, n As Long, ch As Object
    Set ws = AddSheet("PC歷史", after)
    ws.Range("A1").Value = "全市場 P/C ratio（台指選擇權全部到期合計，近一個月）"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    HeaderRow ws.Range("A3"), Array("日期", "未平倉量 P/C", "成交量 P/C", "賣權未平倉", "買權未平倉")
    ws.Columns("A:E").ColumnWidth = 13
    If Not IsEmpty(gPC) Then
        n = UBound(gPC, 1)
        ws.Range("A4").Resize(n, 5).Value = gPC
        ws.Range("A4").Resize(n, 1).NumberFormat = "yyyy-mm-dd"
        ws.Range("B4").Resize(n, 2).NumberFormat = "0.00"
        ws.Range("D4").Resize(n, 2).NumberFormat = "#,##0"
        Set ch = NewChart(ws, ws.Range("G3"), 520, 260, "全市場 P/C ratio")
        AddSeries ch, "未平倉量 P/C", ws.Range("B4").Resize(n, 1), ws.Range("A4").Resize(n, 1), CallClr(), 4
        AddSeries ch, "成交量 P/C", ws.Range("C4").Resize(n, 1), ws.Range("A4").Resize(n, 1), PutClr(), 4
        ch.Axes(1).CategoryType = 2
        ch.Axes(1).TickLabels.NumberFormat = "m/d"
        ch.Axes(2).TickLabels.NumberFormat = "0.0"
    Else
        ws.Range("A4").Value = "（抓取失敗）"
    End If
    ws.Range("A1").Select
    Set WritePC = ws
End Function

Private Sub WriteDashboard()
    Dim ws As Worksheet, i As Long, d As Object, items As Variant, fmts As Variant, r As Long, co As Object
    Set ws = ThisWorkbook.Worksheets("盤面")
    ws.Activate
    ActiveWindow.FreezePanes = False
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop
    ws.Cells.Clear
    ws.Cells.Font.Size = 10
    ActiveWindow.DisplayGridlines = False
    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B:D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 3
    ws.Range("A1").Value = "台指週選盤面"
    ws.Range("A1").Font.Size = 20: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "更新於 " & Format$(gNow, "yyyy-mm-dd hh:nn") & "　未平倉量截至 " & gTradeDate & " 日盤收盤　來源：台灣期交所、Yahoo Finance"
    ws.Range("A2").Font.Color = RGB(93, 102, 115)
    If gErr <> "" Then
        ws.Range("A3").Value = "部分資料抓取失敗：" & Replace(gErr, vbLf, "；")
        ws.Range("A3").Font.Color = CallClr()
    End If

    ' 各到期比較
    items = Array("到期日", "剩餘天數", "報價時點", "價格來源", "合成期貨", "價平履約價", "價平買權", "價平賣權", _
                  "預期波動（跨式）", "預期波動 %", "區間下緣", "區間上緣", "ATM 隱含波動率", "1σ 點數", _
                  "P/C（未平倉）", "P/C（近價 ±2500）", "P/C（成交量，前一日）", "P/C（成交量，即時）", _
                  "最大痛點", "買權牆（壓力）", "買權牆 OI", "賣權牆（支撐）", "賣權牆 OI")
    Dim keys As Variant
    keys = Array("expiry_date", "days_left", "quote_time", "price_basis", "forward", "atm_strike", "atm_call", "atm_put", _
                 "straddle", "move_pct", "range_low", "range_high", "atm_iv", "sigma_pts", _
                 "pc_oi", "pc_oi_near", "pc_vol_day", "pc_vol_live", "max_pain", "call_wall", "call_wall_oi", "put_wall", "put_wall_oi")
    fmts = Array("yyyy-mm-dd", "0.0", "mm-dd hh:mm", "@", "#,##0", "#,##0", "#,##0.0", "#,##0.0", _
                 "#,##0", "0.0%", "#,##0", "#,##0", "0.0%", "#,##0", _
                 "0.00", "0.00", "0.00", "0.00", "#,##0", "#,##0", "#,##0", "#,##0", "#,##0")
    r = 5
    ws.Cells(r - 1, 1).Value = "各到期比較"
    ws.Cells(r - 1, 1).Font.Bold = True: ws.Cells(r - 1, 1).Font.Size = 12
    ws.Cells(r, 1).Value = "項目"
    For i = 1 To gSum.Count
        ws.Cells(r, 1 + i).Value = gSum(i)("label")
    Next
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 1 + gSum.Count))
        .Font.Bold = True
        .Interior.Color = RGB(236, 239, 243)
        .HorizontalAlignment = -4108
    End With
    Dim j As Long
    For j = 0 To UBound(items)
        ws.Cells(r + 1 + j, 1).Value = items(j)
        For i = 1 To gSum.Count
            Set d = gSum(i)
            ws.Cells(r + 1 + j, 1 + i).Value = DGet(d, CStr(keys(j)))
            ws.Cells(r + 1 + j, 1 + i).NumberFormat = fmts(j)
            ws.Cells(r + 1 + j, 1 + i).HorizontalAlignment = -4152
        Next
    Next
    ws.Range(ws.Cells(r + 9, 1), ws.Cells(r + 9, 1 + gSum.Count)).Font.Bold = True
    ws.Range(ws.Cells(r + 13, 1), ws.Cells(r + 13, 1 + gSum.Count)).Font.Bold = True

    ' 籌碼與波動
    Dim a As Object, b As Object, rr As Long, v As Variant
    rr = r + 2 + UBound(items) + 2
    ws.Cells(rr, 1).Value = "籌碼與波動"
    ws.Cells(rr, 1).Font.Bold = True: ws.Cells(rr, 1).Font.Size = 12
    HeaderRow ws.Cells(rr + 1, 1), Array("項目", "數值", "較前日", "說明")
    rr = rr + 2
    If Latest2(gFut, a, b) Then
        PutKV ws, rr, "外資台指期淨部位（口）", DGet(a, "foreign_net_oi"), DDiff(a, b, "foreign_net_oi"), "+#,##0;-#,##0;0", "資料日 " & LatestKey(gFut)
        rr = rr + 1
    End If
    If Not gFutLive Is Nothing Then
        If gFutLive.Exists("日盤") Then
            Set a = gFutLive("日盤")
            PutKV ws, rr, "期貨正逆價差（日盤）", DGet(a, "basis"), Empty, "+#,##0.0;-#,##0.0;0", _
                  "近月 " & Fx(DGet(a, "last"), "#,##0") & "，現貨 " & Fx(DGet(a, "spot"), "#,##0.00")
            rr = rr + 1
        End If
        If gFutLive.Exists("夜盤") Then
            Set a = gFutLive("夜盤")
            PutKV ws, rr, "台指期夜盤", DGet(a, "last"), DGet(a, "chg"), "#,##0", "漲跌相對日盤收盤，時間 " & DGet(a, "time")
            ws.Cells(rr, 3).NumberFormat = "+#,##0;-#,##0;0"
            rr = rr + 1
        End If
    End If
    If Latest2(gInst, a, b) Then
        PutKV ws, rr, "外資選擇權方向（億元）", DGet(a, "foreign_bias_amt"), DDiff(a, b, "foreign_bias_amt"), "+#,##0.00;-#,##0.00;0", _
              "買權淨額 " & Fx(DGet(a, "foreign_call_amt"), "0.00") & "、賣權淨額 " & Fx(DGet(a, "foreign_put_amt"), "0.00") & "；正值偏多"
        rr = rr + 1
    End If
    If Latest2(gLarge, a, b) Then
        PutKV ws, rr, "前十大交易人買權淨部位", DGet(a, "all_call_top10_net"), DDiff(a, b, "all_call_top10_net"), "+#,##0;-#,##0;0", _
              "賣權淨部位 " & Fx(DGet(a, "all_put_top10_net"), "#,##0")
        rr = rr + 1
    End If
    If Not gVol Is Nothing Then
        If gVol.Exists("vix") Then
            PutKV ws, rr, "台指 VIX", gVol("vix"), Empty, "0.00", _
                  "IV Rank " & Fx(DGet(gVol, "vix_rank"), "0%") & "、百分位 " & Fx(DGet(gVol, "vix_pct"), "0%") & "（" & DGet(gVol, "vix_days") & " 個交易日）"
            rr = rr + 1
        End If
        If gVol.Exists("hv20") Then
            PutKV ws, rr, "ATM IV ÷ 20 日 HV", DGet(gVol, "iv_hv_ratio"), Empty, "0.00", _
                  "IV " & Fx(DGet(gVol, "atm_iv"), "0.0%") & "、HV " & Fx(DGet(gVol, "hv20"), "0.0%") & "；大於 1.2 偏貴"
            rr = rr + 1
        End If
    End If
    If Not gGlobal Is Nothing Then
        For i = 1 To gGlobal.Count
            Set d = gGlobal(i)
            If d("symbol") = "^VIX" Or d("symbol") = "ES=F" Or d("symbol") = "NQ=F" Then
                PutKV ws, rr, d("name"), DGet(d, "last"), DGet(d, "chg_pct"), "#,##0.00", "時間 " & Fx(DGet(d, "time"), "mm-dd hh:nn")
                ws.Cells(rr, 3).NumberFormat = "+0.00%;-0.00%;0.00%"
                rr = rr + 1
            End If
        Next
    End If
    ws.Range(ws.Cells(rr, 1), ws.Cells(rr, 1)).Value = "各到期的履約價明細與圖表在後面的工作表；籌碼、波動、美股、P/C 歷史各有一張表。"
    ws.Cells(rr, 1).Font.Color = RGB(93, 102, 115)

    ' 最近到期的未平倉牆
    Dim src As Worksheet, n As Long, ch As Object, cats As Range
    Set src = ThisWorkbook.Worksheets(gSum(1)("sheet"))
    n = UBound(gTab(1), 1)
    Set cats = src.Range("A6").Resize(n, 1)
    Set ch = NewChart(ws, ws.Range("F5"), 560, 300, gSum(1)("label") & " 未平倉牆")
    AddSeries ch, "買權 OI", cats.Offset(0, 1), cats, CallClr(), 51
    AddSeries ch, "賣權 OI", cats.Offset(0, 3), cats, PutClr(), 51
    ch.ChartGroups(1).Overlap = 100
    ch.ChartGroups(1).GapWidth = 20
    ch.Axes(2).TickLabels.NumberFormat = "#,##0;#,##0"
    ch.Axes(1).TickLabelSpacing = Application.WorksheetFunction.Max(1, Int(n / 10))
    ch.Axes(1).TickLabelPosition = -4134
    EnsureButton
    ws.Range("A1").Select
End Sub

Private Sub PutKV(ByVal ws As Worksheet, ByVal r As Long, ByVal nm As String, ByVal v As Variant, ByVal dv As Variant, _
                  ByVal fmt As String, ByVal note As String)
    ws.Cells(r, 1).Value = nm
    ws.Cells(r, 2).Value = v
    ws.Cells(r, 2).NumberFormat = fmt
    ws.Cells(r, 3).Value = dv
    ws.Cells(r, 3).NumberFormat = fmt
    ws.Cells(r, 4).Value = note
    ws.Cells(r, 4).Font.Color = RGB(93, 102, 115)
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).HorizontalAlignment = -4152
End Sub
