Option Explicit
' ============================================================
' 台指策略訊號（Excel 版）：抓資料（只能在 Windows 版 Excel 執行）
'   - 期交所 MIS 即時 1 分K（getQuoteList、getChartData1M）
'   - 期交所「前 30 個交易日期貨每筆成交資料」→ 台指期 1 分K，存成 Daily_YYYY_MM_DD.csv（同 strategy/collect_ticks.py）
'   - 期交所台指期每日行情 futDataDown → tx_daily.csv（同 strategy/fetch_tx_daily.py）
'   - Yahoo 加權指數 60 分K → twii_60min.csv（同 strategy/fetch_twii.py，只給加權指數回測用）
' ============================================================

Private Const UA As String = "Mozilla/5.0 (txo-excel)"
Private Const MIS_URL As String = "https://mis.taifex.com.tw/futures/api/"
Private Const LIST_URL As String = "https://www.taifex.com.tw/cht/3/dlFutPrevious30DaysSalesData"
Private Const FILE_URL As String = "https://www.taifex.com.tw/file/taifex/Dailydownload/DailydownloadCSV/"
Private Const DAILY_URL As String = "https://www.taifex.com.tw/cht/3/futDataDown"

' ------------------------------------------------------------
' HTTP
' ------------------------------------------------------------
Public Function HttpBytes(ByVal url As String, Optional ByVal body As String = "", Optional ByVal ctype As String = "") As Variant
    Dim h As Object, tries As Long, st As Long, lastErr As String
    For tries = 1 To 3
        On Error Resume Next
        Err.Clear
        If tries < 3 Then
            Set h = CreateObject("MSXML2.ServerXMLHTTP.6.0")
            h.setTimeouts 15000, 15000, 60000, 180000
        Else
            Set h = CreateObject("MSXML2.XMLHTTP.6.0")
        End If
        If Len(body) = 0 Then
            h.Open "GET", url, False
        Else
            h.Open "POST", url, False
        End If
        h.setRequestHeader "User-Agent", UA
        h.setRequestHeader "Cache-Control", "no-cache"
        If Len(body) > 0 Then
            h.setRequestHeader "Content-Type", ctype
            h.send body      ' 直接送字串（跟盤面版 TXO.bas 一樣；送 Byte 陣列會被 MSXML 回「參數錯誤」）
        Else
            h.send
        End If
        If Err.Number = 0 Then
            st = h.Status
            If st = 200 Then
                HttpBytes = h.responseBody
                On Error GoTo 0
                Exit Function
            End If
            lastErr = "HTTP " & st
        Else
            lastErr = Err.Description
        End If
        On Error GoTo 0
    Next
    Err.Raise vbObjectError + 301, , "連不上 " & Split(url, "?")(0) & "（" & lastErr & "）"
End Function

Private Function BytesUtf8(ByVal v As Variant) As String
    Dim b() As Byte
    If IsEmpty(v) Then Exit Function
    b = v
    BytesUtf8 = Utf8Text(b)
End Function

' Big5：中文 Windows 用系統字碼頁，其他語系用 ADODB.Stream
Private Function BytesBig5(ByVal v As Variant) As String
    Dim st As Object, b() As Byte
    If IsEmpty(v) Then Exit Function
    b = v
    If UBound(b) < 0 Then Exit Function
    If StrConv(ChrB(&HA4) & ChrB(&HA4), vbUnicode) = ChrW(&H4E2D) Then
        BytesBig5 = StrConv(b, vbUnicode)
    Else
        Set st = CreateObject("ADODB.Stream")
        st.Type = 1
        st.Open
        st.Write b
        st.Position = 0
        st.Type = 2
        st.Charset = "big5"
        BytesBig5 = st.ReadText
        st.Close
    End If
End Function

Public Function MisPost(ByVal ep As String, ByVal body As String) As String
    MisPost = BytesUtf8(HttpBytes(MIS_URL & ep, body, "application/json"))
End Function

' ------------------------------------------------------------
' UTF-8 寫檔
' ------------------------------------------------------------
Public Sub WriteUtf8File(ByVal p As String, ByVal s As String)
    Dim b() As Byte, i As Long, k As Long, c As Long, n As Long, f As Integer, tmp As String
    n = Len(s)
    ReDim b(0 To 3 * n + 3)
    k = 0
    For i = 1 To n
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c < &H80 Then
            b(k) = c: k = k + 1
        ElseIf c < &H800 Then
            b(k) = &HC0 Or (c \ &H40): b(k + 1) = &H80 Or (c And &H3F): k = k + 2
        Else
            b(k) = &HE0 Or (c \ &H1000): b(k + 1) = &H80 Or ((c \ &H40) And &H3F): b(k + 2) = &H80 Or (c And &H3F): k = k + 3
        End If
    Next
    tmp = p & ".tmp"
    If FileExists(tmp) Then Kill tmp
    f = FreeFile
    Open tmp For Binary Access Write As #f
    If k > 0 Then
        ReDim Preserve b(0 To k - 1)
        Put #f, , b
    End If
    Close #f
    If FileExists(p) Then Kill p
    Name tmp As p
End Sub

Private Sub WriteBytes(ByVal p As String, ByVal v As Variant)
    Dim b() As Byte, f As Integer
    b = v
    If FileExists(p) Then Kill p
    f = FreeFile
    Open p For Binary Access Write As #f
    Put #f, , b
    Close #f
End Sub

Private Function TempDir() As String
    Dim t As String
    t = Environ$("TEMP")
    If Len(t) = 0 Then t = ThisWorkbook.Path
    TempDir = JoinPath(t, "txo_strategy")
    If Len(Dir(TempDir, vbDirectory)) = 0 Then MkDir TempDir
End Function

' ------------------------------------------------------------
' 期交所每筆成交 → 1 分K
' ------------------------------------------------------------
' 期交所目前提供的檔名（Daily_YYYY_MM_DD.zip，由舊到新）
Public Function TaifexDailyList(fl() As String) As Long
    Dim s As String, p As Long, nm As String, n As Long, k As Long, dup As Boolean
    s = BytesUtf8(HttpBytes(LIST_URL))
    ReDim fl(0 To 63)
    p = 1
    Do
        p = InStr(p, s, "Daily_")
        If p = 0 Then Exit Do
        nm = Mid$(s, p, 20)
        If nm Like "Daily_####_##_##.zip" Then
            dup = False
            For k = 0 To n - 1
                If fl(k) = nm Then dup = True: Exit For
            Next
            If Not dup Then
                If n > UBound(fl) Then
                    ReDim Preserve fl(0 To 2 * n)
                End If
                fl(n) = nm
                n = n + 1
            End If
        End If
        p = p + 6
    Loop
    If n = 0 Then Err.Raise vbObjectError + 302, , "期交所頁面沒有列出任何檔案，可能改版了"
    SortStrings fl, n
    TaifexDailyList = n
End Function

' 檔案有沒有那一天的日盤（沒有表示期交所只先放了前一晚的夜盤，日盤收盤後要重抓）
Public Function HasDaySession(ByVal p As String, ByVal dn As Long) As Boolean
    Dim f As Integer, ln As String, ymd As String, t As Long
    ymd = Year(dn) & P2(Month(dn)) & P2(Day(dn))
    f = FreeFile
    Open p For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        If Left$(ln, 8) = ymd Then
            t = Val(Mid$(ln, 10, 2)) * 60 + Val(Mid$(ln, 12, 2))
            If t >= T_DAYOPEN And t <= T_DAYCLOSE Then HasDaySession = True: Exit Do
        End If
    Loop
    Close #f
End Function

' 下載一個交易日的每筆成交 zip，轉成台指期（TX）1 分K 存到 dest
Public Function DownloadTicks(ByVal nm As String, ByVal dest As String) As Long
    Dim td As String, zp As String, csvp As String, shl As Object, ns As Object, it As Object, t0 As Double
    Dim f As Integer, ln As String, a() As String, prod As String, seen As Boolean, ct As String
    Dim keys As Object, key As String, k As Long, n As Long, px As Double, vol As Double
    Dim kO() As Double, kH() As Double, kL() As Double, kC() As Double, kV() As Double, kS() As String
    td = TempDir()
    zp = JoinPath(td, nm)
    csvp = JoinPath(td, Replace(nm, ".zip", ".csv"))
    WriteBytes zp, HttpBytes(FILE_URL & nm)
    If FileExists(csvp) Then Kill csvp
    ' 解壓縮：先用 Windows 內建的 tar，不行再用檔案總管的壓縮資料夾
    On Error Resume Next
    CreateObject("WScript.Shell").Run "tar -xf """ & zp & """ -C """ & td & """", 0, True
    On Error GoTo 0
    If Not FileExists(csvp) Then
        Set shl = CreateObject("Shell.Application")
        Set ns = shl.Namespace(CVar(zp))
        If ns Is Nothing Then Err.Raise vbObjectError + 303, , "解壓縮失敗 " & nm
        shl.Namespace(CVar(td)).CopyHere ns.Items, 4 + 16 + 1024
        t0 = Timer
        Do While Not FileExists(csvp)
            DoEvents
            If Timer - t0 > 120 Or Timer < t0 Then Err.Raise vbObjectError + 304, , "解壓縮逾時 " & nm
        Loop
        Application.Wait Now + TimeSerial(0, 0, 2)
    End If
    ' 讀每筆成交：只要台指期（商品代號 TX），不要價差單。檔案依商品排序，TX 讀完就停
    Set keys = CreateObject("Scripting.Dictionary")
    ReDim kO(0 To 4095): ReDim kH(0 To 4095): ReDim kL(0 To 4095): ReDim kC(0 To 4095): ReDim kV(0 To 4095): ReDim kS(0 To 4095)
    f = FreeFile
    Open csvp For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        If Mid$(ln, 10, 3) = "TX " Then
            a = Split(ln, ",")
            If UBound(a) >= 5 Then
                If Trim$(a(1)) = "TX" Then
                    seen = True
                    ct = Trim$(a(2))
                    If InStr(ct, "/") = 0 And IsNumeric(Trim$(a(4))) And IsNumeric(Trim$(a(5))) Then
                        px = Val(a(4)): vol = Val(a(5))
                        key = ct & "|" & Trim$(a(0)) & "|" & Left$(Right$("000000" & Trim$(a(3)), 6), 4)
                        If keys.Exists(key) Then
                            k = keys(key)
                            If px > kH(k) Then kH(k) = px
                            If px < kL(k) Then kL(k) = px
                            kC(k) = px: kV(k) = kV(k) + vol
                        Else
                            If n > UBound(kO) Then
                                ReDim Preserve kO(0 To 2 * n): ReDim Preserve kH(0 To 2 * n): ReDim Preserve kL(0 To 2 * n)
                                ReDim Preserve kC(0 To 2 * n): ReDim Preserve kV(0 To 2 * n): ReDim Preserve kS(0 To 2 * n)
                            End If
                            keys.Add key, n
                            kS(n) = key: kO(n) = px: kH(n) = px: kL(n) = px: kC(n) = px: kV(n) = vol
                            n = n + 1
                        End If
                    End If
                End If
            End If
        ElseIf seen Then
            Exit Do
        End If
    Loop
    Close #f
    On Error Resume Next
    Kill csvp
    Kill zp
    On Error GoTo 0
    If n = 0 Then Err.Raise vbObjectError + 305, , nm & " 沒有台指期資料"
    ' 依契約、日期、時間排序後寫檔（格式同 history/tx_1min）
    Dim idx() As Long, i As Long, j As Long, t As Long, out As String, parts() As String
    ReDim idx(0 To n - 1)
    For i = 0 To n - 1
        idx(i) = i
    Next
    SortKeyIdx kS, idx, 0, n - 1
    f = FreeFile
    Open dest & ".tmp" For Output As #f
    Print #f, "date,time,contract,open,high,low,close,volume"
    For i = 0 To n - 1
        j = idx(i)
        parts = Split(kS(j), "|")
        Print #f, parts(1) & "," & parts(2) & "," & parts(0) & "," & NumG(kO(j)) & "," & NumG(kH(j)) & "," & NumG(kL(j)) & "," & _
                  NumG(kC(j)) & "," & NumG(kV(j))
    Next
    Close #f
    If FileExists(dest) Then Kill dest
    Name dest & ".tmp" As dest
    DownloadTicks = n
End Function

Private Sub SortKeyIdx(ks() As String, idx() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As String, t As Long
    If hi <= lo Then Exit Sub
    p = ks(idx((lo + hi) \ 2))
    i = lo: j = hi
    Do While i <= j
        Do While StrComp(ks(idx(i)), p, vbBinaryCompare) < 0
            i = i + 1
        Loop
        Do While StrComp(ks(idx(j)), p, vbBinaryCompare) > 0
            j = j - 1
        Loop
        If i <= j Then
            t = idx(i): idx(i) = idx(j): idx(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    SortKeyIdx ks, idx, lo, j
    SortKeyIdx ks, idx, i, hi
End Sub

' 補 1 分K：下載資料夾裡還沒有的檔案；refreshNewest = True 時，最新 3 個檔案若缺日盤、而且那天日盤已收盤，也重抓
' 回傳更新的檔案數
Public Function CollectTicks(ByVal folder As String, ByVal nowTm As Long, ByVal refreshNewest As Boolean, msg As String) As Long
    Dim fl() As String, n As Long, i As Long, dest As String, need As Boolean, dn As Long, done As Long, k As Long
    n = TaifexDailyList(fl)
    For i = 0 To n - 1
        dest = JoinPath(folder, Replace(fl(i), ".zip", ".csv"))
        need = Not FileExists(dest)
        If Not need And refreshNewest And i >= n - 3 Then
            dn = DayOfDailyName(fl(i))
            If nowTm >= dn * 1440 + 16 * 60 + 30 Then        ' 期交所約 16:30 後放完整檔
                If Not HasDaySession(dest, dn) Then
                    k = Val(StateGet("retry_" & fl(i), "0"))
                    If nowTm - k >= 60 Then
                        need = True
                        StateSet "retry_" & fl(i), CStr(nowTm)
                    End If
                End If
            End If
        End If
        If need Then
            SetStatus "下載期交所每筆成交 " & fl(i) & "（" & (i + 1) & "/" & n & "）..."
            On Error Resume Next
            DownloadTicks fl(i), dest
            If Err.Number <> 0 Then
                msg = msg & fl(i) & "：" & Err.Description & vbLf
            Else
                done = done + 1
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next
    CollectTicks = done
End Function

' ------------------------------------------------------------
' 期交所台指期每日行情（日盤「一般」、夜盤「盤後」）→ tx_daily.csv
' ------------------------------------------------------------
Private Function Ymd8(ByVal dn As Long) As String
    Ymd8 = Year(dn) & "%2F" & P2(Month(dn)) & "%2F" & P2(Day(dn))
End Function

Public Function FetchTxDaily(ByVal p As String, ByVal todayDn As Long) As Long
    Dim txt As String, lines() As String, a() As String, j As Long, keys As Object, k As Variant
    Dim y As Long, m As Long, firstDn As Long, lastDn As Long, s As String, lastDate As String, got As Long
    Dim ses As String, ct As String, rws As Object, sd As String, sn As String
    Set rws = CreateObject("Scripting.Dictionary")
    sd = SesDay(): sn = SesNight()
    ' 舊資料
    txt = ReadUtf8File(p)
    If Len(txt) > 0 Then
        lines = Split(Replace(txt, vbCr, ""), vbLf)
        For j = 1 To UBound(lines)
            a = Split(lines(j), ",")
            If UBound(a) >= 7 Then
                rws(a(0) & "," & a(1) & "," & a(2)) = lines(j)
                If a(0) > lastDate Then lastDate = a(0)
            End If
        Next
    End If
    If Len(lastDate) > 0 Then
        firstDn = DayNum(Val(Left$(lastDate, 4)), Val(Mid$(lastDate, 6, 2)), Val(Mid$(lastDate, 9, 2))) - 10
    Else
        firstDn = DayNum(2024, 1, 1)
    End If
    y = Year(firstDn): m = Month(firstDn)
    Do While y * 12 + m <= Year(todayDn) * 12 + Month(todayDn)
        firstDn = DayNum(y, m, 1)
        lastDn = DayNum(y, m + 1, 1) - 1
        SetStatus "下載期交所台指期每日行情 " & y & "/" & P2(m) & "..."
        txt = BytesBig5(HttpBytes(DAILY_URL, "down_type=1&commodity_id=TX&queryStartDate=" & Ymd8(firstDn) & _
                                  "&queryEndDate=" & Ymd8(lastDn), "application/x-www-form-urlencoded"))
        If Left$(LTrim$(txt), 1) <> "<" Then
            lines = Split(Replace(txt, vbCr, ""), vbLf)
            For j = 1 To UBound(lines)
                a = Split(lines(j), ",")
                If UBound(a) >= 17 Then
                    If Trim$(a(1)) = "TX" Then
                        ses = Trim$(a(17))
                        If ses = ChrW(&H4E00) & ChrW(&H822C) Then           ' 一般
                            ses = sd
                        ElseIf ses = ChrW(&H76E4) & ChrW(&H5F8C) Then       ' 盤後
                            ses = sn
                        Else
                            ses = ""
                        End If
                        ct = Trim$(a(2))
                        If Len(ses) > 0 And InStr(ct, "/") = 0 And IsNumeric(ct) Then
                            If IsNumeric(Trim$(a(3))) And IsNumeric(Trim$(a(4))) And IsNumeric(Trim$(a(5))) And IsNumeric(Trim$(a(6))) _
                               And IsNumeric(Trim$(a(9))) Then
                                s = Replace(Trim$(a(0)), "/", "-")
                                rws(s & "," & ses & "," & ct) = s & "," & ses & "," & ct & "," & NumG(Val(a(3))) & "," & NumG(Val(a(4))) & "," & _
                                                                 NumG(Val(a(5))) & "," & NumG(Val(a(6))) & "," & CStr(CLng(Val(a(9))))
                                got = got + 1
                            End If
                        End If
                    End If
                End If
            Next
        End If
        m = m + 1
        If m > 12 Then m = 1: y = y + 1
    Loop
    If rws.Count = 0 Then Err.Raise vbObjectError + 306, , "沒有抓到台指期每日行情"
    ' 依（日期, 盤別, 契約）排序寫回（跟 Python sorted 一樣用字碼順序）
    Dim ks() As String, idx() As Long, i As Long, n As Long, out As String
    n = rws.Count
    ReDim ks(0 To n - 1): ReDim idx(0 To n - 1)
    i = 0
    For Each k In rws.Keys
        ks(i) = k: idx(i) = i: i = i + 1
    Next
    SortKeyIdx ks, idx, 0, n - 1
    Dim buf() As String
    ReDim buf(0 To n)
    buf(0) = "date,session,contract,open,high,low,close,volume"
    For i = 0 To n - 1
        buf(i + 1) = rws(ks(idx(i)))
    Next
    WriteUtf8File p, Join(buf, vbCrLf) & vbCrLf
    FetchTxDaily = got
End Function

' ------------------------------------------------------------
' Yahoo 加權指數 60 分K（只給加權指數回測用）→ twii_60min.csv
' ------------------------------------------------------------
Private Function JsonNums(ByVal s As String, ByVal key As String, ByVal from As Long) As Variant
    Dim p As Long, q As Long
    p = InStr(from, s, """" & key & """:[")
    If p = 0 Then JsonNums = Empty: Exit Function
    p = p + Len(key) + 4
    q = InStr(p, s, "]")
    JsonNums = Split(Mid$(s, p, q - p), ",")
End Function

Public Function FetchTwii(ByVal p As String, ByVal todayDn As Long) As Long
    Dim s As String, host As Variant, ts As Variant, qo As Variant, qh As Variant, ql As Variant, qc As Variant, qv As Variant
    Dim qp As Long, i As Long, rws As Object, txt As String, lines() As String, a() As String, k As Variant
    Dim t As Double, dn As Long, mi As Long, st As String, today As String, got As Long
    For Each host In Array("query1", "query2")
        On Error Resume Next
        Err.Clear
        s = BytesUtf8(HttpBytes("https://" & host & ".finance.yahoo.com/v8/finance/chart/%5ETWII?interval=60m&range=730d"))
        If Err.Number = 0 And InStr(s, """timestamp"":[") > 0 Then
            On Error GoTo 0
            Exit For
        End If
        s = ""
        On Error GoTo 0
    Next
    If Len(s) = 0 Then Err.Raise vbObjectError + 307, , "Yahoo 加權指數抓取失敗"
    ts = JsonNums(s, "timestamp", 1)
    qp = InStr(s, """quote"":[")
    qo = JsonNums(s, "open", qp): qh = JsonNums(s, "high", qp): ql = JsonNums(s, "low", qp)
    qc = JsonNums(s, "close", qp): qv = JsonNums(s, "volume", qp)
    Set rws = CreateObject("Scripting.Dictionary")
    txt = ReadUtf8File(p)
    If Len(txt) > 0 Then
        lines = Split(Replace(txt, vbCr, ""), vbLf)
        For i = 1 To UBound(lines)
            a = Split(lines(i), ",")
            If UBound(a) >= 5 Then rws(a(0)) = lines(i)
        Next
    End If
    today = FmtTm(todayDn * 1440, 2)
    For i = 0 To UBound(ts)
        If qo(i) <> "null" And qh(i) <> "null" And ql(i) <> "null" And qc(i) <> "null" Then
            t = Val(ts(i)) + 8 * 3600#
            dn = DayNum(1970, 1, 1) + Int(t / 86400#)
            mi = Int((t - Int(t / 86400#) * 86400#) / 60)
            st = FmtTm(dn * 1440& + mi, 0)
            If Right$(st, 5) <> "13:30" And Left$(st, 10) <> today Then
                rws(st) = st & "," & Format$(Val(qo(i)), "0.00") & "," & Format$(Val(qh(i)), "0.00") & "," & _
                           Format$(Val(ql(i)), "0.00") & "," & Format$(Val(qc(i)), "0.00") & "," & IIf(qv(i) = "null", "0", qv(i))
                got = got + 1
            End If
        End If
    Next
    Dim ks() As String, idx() As Long, n As Long, buf() As String
    n = rws.Count
    ReDim ks(0 To n - 1): ReDim idx(0 To n - 1): ReDim buf(0 To n)
    i = 0
    For Each k In rws.Keys
        ks(i) = k: idx(i) = i: i = i + 1
    Next
    SortKeyIdx ks, idx, 0, n - 1
    buf(0) = "start,open,high,low,close,volume"
    For i = 0 To n - 1
        buf(i + 1) = rws(ks(idx(i)))
    Next
    WriteUtf8File p, Join(buf, vbCrLf) & vbCrLf
    FetchTwii = got
End Function
