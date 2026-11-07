Option Explicit
' ============================================================
' 台指策略訊號（Excel 版）：更新流程與工作表
'   即時訊號：對應網頁 live.html（v3 與 D2 的部位、停損、訊號、進場條件、前測紀錄）
'   成交紀錄：對應網頁 trades.html（台指期每一筆、加權指數近 2 年回測每一筆）
' 資料存在活頁簿旁邊的 strategy_data 資料夾（tx_1min\Daily_*.csv、tx_daily.csv、twii_60min.csv、state.txt）
' ============================================================

Private Const SH_LIVE As String = "即時訊號"
Private Const SH_TRADES As String = "成交紀錄"
Private gBusy As Boolean
Private gNext As Date
Private gStK() As String
Private gStV() As String
Private gStN As Long
Private gStPath As String

' ------------------------------------------------------------
' 顏色
' ------------------------------------------------------------
Private Function cUp() As Long
    cUp = RGB(194, 65, 45)
End Function
Private Function cDn() As Long
    cDn = RGB(47, 125, 91)
End Function
Private Function cAccent() As Long
    cAccent = RGB(47, 95, 208)
End Function
Private Function cMuted() As Long
    cMuted = RGB(93, 102, 115)
End Function
Private Function cWarn() As Long
    cWarn = RGB(168, 101, 0)
End Function
Private Function cWarnBg() As Long
    cWarnBg = RGB(255, 244, 224)
End Function
Private Function cSigBg() As Long
    cSigBg = RGB(232, 239, 255)
End Function
Private Function cLine() As Long
    cLine = RGB(223, 227, 232)
End Function
Private Function cHead() As Long
    cHead = RGB(240, 242, 245)
End Function
Private Function cRowL() As Long
    cRowL = RGB(252, 238, 235)
End Function
Private Function cRowS() As Long
    cRowS = RGB(233, 245, 238)
End Function

' ------------------------------------------------------------
' 狀態檔（strategy_data\state.txt，key=value）
' ------------------------------------------------------------
Public Sub StateLoad(ByVal dd As String)
    Dim f As Integer, ln As String, p As Long
    gStPath = JoinPath(dd, "state.txt")
    gStN = 0
    ReDim gStK(0 To 63): ReDim gStV(0 To 63)
    If Not FileExists(gStPath) Then Exit Sub
    f = FreeFile
    Open gStPath For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        p = InStr(ln, "=")
        If p > 1 Then StateSet Left$(ln, p - 1), Mid$(ln, p + 1)
    Loop
    Close #f
End Sub

Public Sub StateSave()
    Dim f As Integer, i As Long
    If Len(gStPath) = 0 Then Exit Sub
    f = FreeFile
    Open gStPath For Output As #f
    For i = 0 To gStN - 1
        Print #f, gStK(i) & "=" & gStV(i)
    Next
    Close #f
End Sub

Public Function StateGet(ByVal k As String, ByVal dflt As String) As String
    Dim i As Long
    StateGet = dflt
    For i = 0 To gStN - 1
        If gStK(i) = k Then StateGet = gStV(i): Exit Function
    Next
End Function

Public Sub StateSet(ByVal k As String, ByVal v As String)
    Dim i As Long
    For i = 0 To gStN - 1
        If gStK(i) = k Then gStV(i) = v: Exit Sub
    Next
    If gStN > UBound(gStK) Then
        ReDim Preserve gStK(0 To 2 * gStN)
        ReDim Preserve gStV(0 To 2 * gStN)
    End If
    gStK(gStN) = k: gStV(gStN) = v
    gStN = gStN + 1
End Sub

Public Sub SetStatus(ByVal s As String)
    On Error Resume Next
    If Len(s) = 0 Then Application.StatusBar = False Else Application.StatusBar = s
    DoEvents
End Sub

Public Function DataDir() As String
    Dim d As String
    d = JoinPath(ThisWorkbook.Path, "strategy_data")
    If Len(Dir(d, vbDirectory)) = 0 Then MkDir d
    If Len(Dir(JoinPath(d, "tx_1min"), vbDirectory)) = 0 Then MkDir JoinPath(d, "tx_1min")
    DataDir = d
End Function

' ------------------------------------------------------------
' 盤別
' ------------------------------------------------------------
Public Function SessionState(ByVal tm As Long) As String
    Dim w As Long, t As Long
    w = Wd(tm \ 1440): t = MinOfTm(tm)
    If w < 5 And t >= T_DAYOPEN And t < T_DAYCLOSE Then
        SessionState = "日盤交易中"
    ElseIf (w < 5 And t >= T_NIGHT) Or (w >= 1 And w <= 5 And t < 300) Then
        SessionState = "夜盤交易中"
    Else
        SessionState = "休市"
    End If
End Function

' ============================================================
' 進入點
' ============================================================
Public Sub UpdateNow()
    RunUpdate True
    ScheduleNext
End Sub

Public Sub OnOpen()
    On Error Resume Next
    EnsureButtons
    On Error GoTo 0
    RunUpdate True
    ScheduleNext
End Sub

Public Sub OnClose()
    CancelNext
End Sub

Public Sub ToggleAuto()
    Dim dd As String
    dd = DataDir()
    StateLoad dd
    If StateGet("auto", "1") = "1" Then
        StateSet "auto", "0"
    Else
        StateSet "auto", "1"
    End If
    StateSave
    ScheduleNext
    ShowAutoStatus
End Sub

Public Sub AutoTick()
    Dim nowTm As Long, last As Long
    gNext = 0
    nowTm = TmFromDate(Now)
    StateLoad DataDir()
    last = Val(StateGet("prev_gen", "0"))
    ' 休市時每 60 分鐘才更新一次；交易時段每次都更新
    If SessionState(nowTm) <> "休市" Or nowTm - last >= 60 Then RunUpdate False
    ScheduleNext
End Sub

Public Sub ScheduleNext()
    Dim m As Long
    CancelNext
    If StateGet("auto", "1") <> "1" Then
        ShowAutoStatus
        Exit Sub
    End If
    m = Val(StateGet("auto_min", "5"))
    If m < 1 Then m = 5
    gNext = Now + TimeSerial(0, m, 0)
    On Error Resume Next
    Application.OnTime gNext, "'" & ThisWorkbook.Name & "'!AutoTick"
    On Error GoTo 0
    ShowAutoStatus
End Sub

Public Sub CancelNext()
    On Error Resume Next
    If gNext > 0 Then Application.OnTime gNext, "'" & ThisWorkbook.Name & "'!AutoTick", , False
    gNext = 0
End Sub

Private Sub ShowAutoStatus()
    Dim ws As Worksheet, s As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LIVE)
    If StateGet("auto", "1") = "1" Then
        s = "自動更新：開，每 " & Val(StateGet("auto_min", "5")) & " 分鐘（休市時每小時）"
        If gNext > 0 Then s = s & "・下次 " & Format$(gNext, "hh:nn")
    Else
        s = "自動更新：關（按「自動更新 開/關」打開）"
    End If
    ws.Range("B4").Value = s
    ws.Shapes("btnAuto").TextFrame2.TextRange.Text = IIf(StateGet("auto", "1") = "1", "自動更新：開", "自動更新：關")
End Sub

Public Sub EnsureButtons()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_LIVE)
    AddButton ws, "btnUpdate", "立即更新", "UpdateNow", ws.Range("E1").Left, 6, 110, 28, RGB(47, 95, 208)
    AddButton ws, "btnAuto", "自動更新：開", "ToggleAuto", ws.Range("E1").Left + 118, 6, 110, 28, RGB(93, 102, 115)
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal nm As String, ByVal txt As String, ByVal macro As String, _
                      ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal clr As Long)
    Dim shp As Object
    On Error Resume Next
    Set shp = ws.Shapes(nm)
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(5, x, y, w, h)
        shp.Name = nm
        shp.Fill.ForeColor.RGB = clr
        shp.Line.Visible = False
        With shp.TextFrame2
            .TextRange.Text = txt
            .TextRange.Font.Size = 11
            .TextRange.Font.Bold = True
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = 2
            .VerticalAnchor = 3
        End With
        shp.Placement = 3
    End If
    shp.OnAction = "'" & ThisWorkbook.Name & "'!" & macro
End Sub

' ============================================================
' 更新：抓資料 → 計算 → 畫工作表
' ============================================================
Public Sub RunUpdate(ByVal interactive As Boolean)
    Dim dd As String, d1 As String, nowTm As Long, todayDn As Long, errs As String, msg As String
    Dim fl() As String, nfiles As Long, collected As Boolean, got As Long, prevGen As Long, stage As String
    If gBusy Then Exit Sub
    gBusy = True
    On Error GoTo fail
    stage = "準備"
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    dd = DataDir()
    d1 = JoinPath(dd, "tx_1min")
    StateLoad dd
    nowTm = TmFromDate(Now)
    todayDn = nowTm \ 1440
    prevGen = Val(StateGet("prev_gen", "0"))

    ' 1) 期交所每筆成交 → 1 分K（第一次全部下載；之後每小時檢查一次有沒有新檔）
    stage = "期交所每筆成交"
    nfiles = ListDaily(d1, fl)
    If nfiles = 0 Then
        If interactive Then
            Application.ScreenUpdating = True
            MsgBox "第一次使用：資料夾裡沒有台指期 1 分K，要先從期交所下載最近 30 個交易日的每筆成交" & vbLf & _
                   "（每天一個約 5 MB 的壓縮檔），大約需要 3–10 分鐘。" & vbLf & vbLf & _
                   "下載後存在：" & d1 & vbLf & "之後只補新的日子。", vbInformation, "台指策略訊號"
            Application.ScreenUpdating = False
        End If
    End If
    If nfiles = 0 Or nowTm - Val(StateGet("last_collect", "0")) >= 60 Then
        On Error Resume Next
        got = CollectTicks(d1, nowTm, True, msg)
        If Err.Number <> 0 Then errs = errs & "期交所每筆成交：" & Err.Description & vbLf Else StateSet "last_collect", CStr(nowTm)
        Err.Clear
        On Error GoTo fail
        If Len(msg) > 0 Then errs = errs & msg
        collected = True
        StateSave
    End If

    ' 2) 台指期每日行情（D2 日K 暖機）與加權指數（回測），一天更新一次
    stage = "每日行情"
    If Not FileExists(JoinPath(dd, "tx_daily.csv")) Or Val(StateGet("last_daily", "0")) < todayDn Then
        On Error Resume Next
        FetchTxDaily JoinPath(dd, "tx_daily.csv"), todayDn
        If Err.Number <> 0 Then errs = errs & "期交所每日行情：" & Err.Description & vbLf Else StateSet "last_daily", CStr(todayDn)
        Err.Clear
        On Error GoTo fail
    End If
    If Not FileExists(JoinPath(dd, "twii_60min.csv")) Or Val(StateGet("last_twii", "0")) < todayDn Then
        On Error Resume Next
        FetchTwii JoinPath(dd, "twii_60min.csv"), todayDn
        If Err.Number <> 0 Then errs = errs & "Yahoo 加權指數（只影響回測）：" & Err.Description & vbLf Else StateSet "last_twii", CStr(todayDn)
        Err.Clear
        On Error GoTo fail
    End If

    ' 3) 即時 1 分K
    stage = "即時 1 分K"
    SetStatus "抓期交所即時 1 分K..."
    FetchLive Year(todayDn)

    ' 4) 接資料；有缺盤時去期交所補每日成交檔
    stage = "讀 1 分K"
    SetStatus "讀取 1 分K..."
    ResetUnits
    LoadHistory d1
    If gHistFiles = 0 Then Err.Raise vbObjectError + 401, , "沒有台指期 1 分K（" & d1 & "），請確認網路後再按一次「立即更新」"
    AddLiveUnits
    FindGap nowTm
    If gGapN > 0 And Not collected Then
        msg = ""
        On Error Resume Next
        got = CollectTicks(d1, nowTm, False, msg)
        Err.Clear
        On Error GoTo fail
        If got > 0 Then
            ResetUnits
            LoadHistory d1
            AddLiveUnits
            FindGap nowTm
        End If
    End If

    ' 5) 計算與輸出
    stage = "計算"
    SetStatus "計算 v3 與 D2..."
    ComputeAll JoinPath(dd, "tx_daily.csv"), JoinPath(dd, "twii_60min.csv"), nowTm
    stage = "輸出"
    SetStatus "更新工作表..."
    RenderLive nowTm, prevGen, errs
    RenderTrades nowTm
    StateSet "prev_gen", CStr(nowTm)
    StateSave
    SetStatus ""
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ShowAutoStatus
    gBusy = False
    Exit Sub
fail:
    Dim em As String
    em = "更新失敗（" & stage & "）：" & Err.Description
    SetStatus ""
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    gBusy = False
    On Error Resume Next
    With ThisWorkbook.Worksheets(SH_LIVE).Range("B5")
        .Value = "⚠ " & em & "（" & FmtTm(TmFromDate(Now), 1) & "）"
        .Font.Color = cWarn()
        .Interior.Color = cWarnBg()
    End With
    If interactive Then MsgBox em & vbLf & vbLf & "請確認網路連線，稍後再按一次「立即更新」。", vbExclamation, "台指策略訊號"
End Sub

' ============================================================
' 交易紀錄的顯示用欄位（live.trade_row）
' ============================================================
Private Function RowR(tr() As Variant, ByVal k As Long, ByVal useLast As Boolean) As Double
    Dim pts As Double
    If tr(F_OPEN, k) And useLast Then
        pts = (gLastC - tr(F_ENTRY, k)) * tr(F_SIDE, k)
        RowR = (pts - COST) / tr(F_RISK, k)
    Else
        RowR = tr(F_RMUL, k)
    End If
End Function

Private Function RowTwd(tr() As Variant, ByVal k As Long, ByVal useLast As Boolean) As Double
    Dim r As Double
    r = RowR(tr, k, useLast)
    RowTwd = Round((r * tr(F_RISK, k) + COST) * MULT - IIf(tr(F_OPEN, k), FEE, 2 * FEE))
End Function

Private Function RowPts(tr() As Variant, ByVal k As Long, ByVal useLast As Boolean) As Double
    RowPts = Round(RowR(tr, k, useLast) * tr(F_RISK, k) + COST)
End Function

Private Function RowPhase(tr() As Variant, ByVal k As Long) As String
    If tr(F_T, k) >= ForwardStart() Then RowPhase = "前測" Else RowPhase = "回測"
End Function

Private Function SideTxt(ByVal s As Long) As String
    If s > 0 Then SideTxt = "多" Else SideTxt = "空"
End Function

' R 合計：每筆先四捨五入到 0.01（同網頁），用「百分之一」整數加總避免浮點誤差，再取到 0.1（五入）
Private Function SumR1(ByVal hundredths As Double) As String
    Dim t As Double
    t = Int((Abs(hundredths) + 5) / 10)
    If hundredths < 0 And t > 0 Then SumR1 = "-" Else SumR1 = "+"
    SumR1 = SumR1 & Format$(t / 10, "0.0")
End Function

' ============================================================
' 即時訊號工作表
' ============================================================
Private Function Sh(ByVal nm As String) As Worksheet
    Set Sh = ThisWorkbook.Worksheets(nm)
End Function

Private Sub Txt(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal v As Variant, Optional ByVal sz As Double = 0, _
                Optional ByVal bold As Boolean = False, Optional ByVal clr As Long = -1, Optional ByVal fill As Long = -1)
    With ws.Cells(r, c)
        If VarType(v) = vbString Then .NumberFormat = "@"      ' 文字不要被 Excel 轉成數字（例如 +114.4）
        .Value = v
        If sz > 0 Then .Font.Size = sz
        If bold Then .Font.Bold = True
        If clr >= 0 Then .Font.Color = clr
        If fill >= 0 Then .Interior.Color = fill
    End With
End Sub

Private Sub Band(ByVal ws As Worksheet, ByVal r As Long, ByVal v As String, ByVal clr As Long, ByVal fill As Long, Optional ByVal bold As Boolean = False)
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 7)).Interior.Color = fill
    Txt ws, r, 2, v, 0, bold, clr
End Sub

Private Sub Money(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal v As Double)
    With ws.Cells(r, c)
        .Value = v
        .NumberFormat = "+#,##0;-#,##0;0"
        If v > 0 Then
            .Font.Color = cUp()
        ElseIf v < 0 Then
            .Font.Color = cDn()
        End If
    End With
End Sub

Public Sub RenderLive(ByVal nowTm As Long, ByVal prevGen As Long, ByVal errs As String)
    Dim ws As Worksheet, r As Long, s As String, k As Long, ln As Variant
    Set ws = Sh(SH_LIVE)
    ws.Cells.Clear
    On Error Resume Next
    ws.Hyperlinks.Delete
    On Error GoTo 0
    ws.Cells.Font.Size = 10
    ws.Columns(1).ColumnWidth = 2
    ws.Columns(2).ColumnWidth = 30
    ws.Columns(3).ColumnWidth = 17
    ws.Columns(4).ColumnWidth = 14
    ws.Columns(5).ColumnWidth = 30
    ws.Columns(6).ColumnWidth = 17
    ws.Columns(7).ColumnWidth = 14
    ws.Columns(8).ColumnWidth = 2
    Txt ws, 1, 2, "台指策略即時訊號", 18, True
    ws.Rows(1).RowHeight = 30
    Txt ws, 2, 2, SessionState(nowTm) & "・台北 " & FmtTm(nowTm) & " 更新・Excel 直接向期交所抓資料計算（不靠 GitHub）", 0, False, cMuted()
    With ws.Cells(3, 2)
        .Value = gLastC
        .NumberFormat = "#,##0"
        .Font.Size = 22
        .Font.Bold = True
        .HorizontalAlignment = -4131
    End With
    ws.Rows(3).RowHeight = 32
    s = ""
    For k = 1 To 2
        If gLiveHas(k) Then s = s & IIf(Len(s) > 0, "・", "") & IIf(k = 1, "日盤 ", "夜盤 ") & FmtTm(gLiveLastMin(k), 1)
    Next
    If Len(s) = 0 Then s = "沒有即時資料（history 到 " & FmtTm(gHistLast, 1) & "）"
    Txt ws, 3, 3, "台指期近月（日盤＋夜盤）・資料到 " & s, 0, False, cMuted()
    ws.Cells(3, 3).VerticalAlignment = -4108
    r = 5
    For k = 1 To 2
        If Len(gLiveErr(k)) > 0 Then
            Band ws, r, "⚠ " & IIf(k = 1, "日盤", "夜盤") & "：" & gLiveErr(k), cWarn(), cWarnBg()
            r = r + 1
        End If
    Next
    If Len(errs) > 0 Then
        For Each ln In Split(errs, vbLf)
            If Len(ln) > 0 Then
                Band ws, r, "⚠ " & ln, cWarn(), cWarnBg()
                r = r + 1
            End If
        Next
    End If
    If gGapN > 0 Then
        s = ""
        For k = 0 To gGapN - 1
            s = s & IIf(k > 0, "、", "") & FmtTm(gGap(k), 1)
        Next
        Band ws, r, "⚠ 資料缺口：" & s & " 開始的盤沒有資料（期交所即時資料沒抓到、每日檔案還沒出，或那天是假日）。指標與訊號可能不準。", cWarn(), cWarnBg()
        r = r + 1
    End If
    r = r + 1
    r = StratCard(ws, r, 1, nowTm, prevGen)
    r = r + 1
    r = StratCard(ws, r, 2, nowTm, prevGen)
    r = r + 1
    r = RecordBlock(ws, r, 1, "v3 模擬交易紀錄")
    r = r + 1
    r = RecordBlock(ws, r, 2, "D2 模擬交易紀錄")
    r = r + 1
    Txt ws, r, 2, "資料：兩套策略的訊號、部位、停損都只用台指期近月（TX）；v3 含夜盤，D2 日K 只算日盤。加權指數只用在「成交紀錄」的回測。", 9, False, cMuted()
    Txt ws, r + 1, 2, "「形成中」表示那根 K 還沒收完，條件用現價先算，收K前都可能變。試算：本金 20 萬、1 口微台（每點 10 元）、每次下單 25 元；R 是以初始停損點數為 1 的倍數。", 9, False, cMuted()
    Txt ws, r + 2, 2, "使用方式、資料來源與限制見「說明」工作表。僅供參考，不構成投資建議。", 9, False, cMuted()
    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
End Sub

Private Function StratCard(ByVal ws As Worksheet, ByVal r0 As Long, ByVal strat As Long, ByVal nowTm As Long, ByVal prevGen As Long) As Long
    Dim r As Long, forming As Boolean, lastT As Long, barEnd As Long, s As String, k As Long
    Dim posK As Long, sigHas As Boolean, pendHas As Boolean, pend As String, side As Long
    Dim tr() As Variant, tn As Long, sg() As Variant
    Dim pts As Double, risk As Double, entry As Double, stp As Double, st As Long, lag As Long
    r = r0
    If strat = 1 Then
        tr = gVTr: tn = gVTn: posK = gVPos: sigHas = gVSigHas: pendHas = gVPendHas: pend = gVPend
        If sigHas Then sg = gVSig
        lastT = gVT(gVN - 1): barEnd = lastT + 60
        forming = nowTm < barEnd
        Txt ws, r, 2, "v3・60 分K", 14, True
        Txt ws, r + 1, 2, "60 分K 收完才看訊號；日盤 08:45–11:45、夜盤 " & IIf(UsDst(nowTm), "21:00–23:00", "22:00–00:00") & " 開始的 K 才開新倉", 9, False, cMuted()
        s = "最後一根 60 分K：" & FmtTm(lastT, 1) & " 起" & IIf(forming, "，形成中（" & FmtTm(barEnd, 3) & " 收）", "，已收完")
    Else
        tr = gDTr: tn = gDTn: posK = gDPos: sigHas = gDSigHas: pendHas = gDPendHas: pend = gDPend
        If sigHas Then sg = gDSig
        lastT = gDT(gDN - 1)
        forming = nowTm < (lastT \ 1440) * 1440 + T_DAYCLOSE
        Txt ws, r, 2, "D2・日K", 14, True
        Txt ws, r + 1, 2, "收盤創 20 日新高/新低，13:45 收盤前進場；台指期日K 只算日盤（08:45–13:45），夜盤不算", 9, False, cMuted()
        s = "最後一根日K（只算日盤）：" & FmtTm(lastT, 4) & IIf(forming, "，形成中（13:45 收）", "，已收盤；夜盤不算")
    End If
    ws.Rows(r).RowHeight = 22
    ' 狀態
    With ws.Range(ws.Cells(r, 6), ws.Cells(r, 7))
        .Merge
        .HorizontalAlignment = -4108
        .Font.Bold = True
        .Font.Size = 12
        If posK >= 0 Then
            .Value = "持" & SideTxt(tr(F_SIDE, posK)) & "單"
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = IIf(tr(F_SIDE, posK) > 0, cUp(), cDn())
        Else
            .Value = "空手"
            .Font.Color = cMuted()
            .Interior.Color = cHead()
        End If
    End With
    Txt ws, r + 2, 2, s, 9, False, cMuted()
    r = r + 3
    ' 上次更新之後的進出場
    If prevGen > 0 Then
        Dim ev As String, n As Long
        n = 0
        For k = 0 To tn - 1
            side = tr(F_SIDE, k)
            If tr(F_ENTRYT, k) > prevGen Then
                If strat = 1 Then st = tr(F_T, k) + 60 Else st = (tr(F_T, k) \ 1440) * 1440 + T_DAYCLOSE
                ev = FmtTm(tr(F_ENTRYT, k), 1) & " " & SideTxt(side) & "單進場 " & N0(tr(F_ENTRY, k)) & "（訊號 " & FmtTm(st, 1) & _
                     " 成立，" & FmtTm(nowTm, 1) & " 更新才看到，晚了 " & FmtLag(nowTm - st) & "）"
                r = EventRow(ws, r, n, prevGen, ev)
            End If
            If Not tr(F_OPEN, k) And tr(F_EXITT, k) > prevGen Then
                ev = FmtTm(tr(F_EXITT, k), 1) & " " & SideTxt(side) & "單出場 " & N0(tr(F_EXIT, k)) & "（" & WhyText(tr, k) & "）"
                r = EventRow(ws, r, n, prevGen, ev)
            End If
        Next
    End If
    ' 訊號
    If sigHas Then
        Dim when As String, tag As String
        If strat = 1 Then
            when = "下一根 60 分K 開盤進場"
            st = sg(F_T) + 60
        Else
            when = IIf(forming, "13:45 收盤前進場", "應已在 13:45 收盤前進場")
            st = (sg(F_T) \ 1440) * 1440 + T_DAYCLOSE
        End If
        lag = nowTm - st
        If lag < 0 Then lag = 0
        If forming Then
            tag = "形成中：若以現價收K"
        Else
            tag = "訊號 " & FmtTm(st, 1) & " 成立" & IIf(lag > 10, "（晚了 " & FmtLag(lag) & "才看到）", "")
        End If
        Band ws, r, "● " & tag & "，做" & SideTxt(sg(F_SIDE)) & "（" & sg(F_KIND) & "）→ " & when & "，參考價 " & N0(sg(F_ENTRY)) & _
             "、停損 " & N0(sg(F_STOP0)) & "（" & N0(sg(F_RISK)) & " 點）", cAccent(), cSigBg(), True
        r = r + 1
    End If
    ' 部位
    If posK >= 0 Then
        side = tr(F_SIDE, posK)
        entry = tr(F_ENTRY, posK): stp = tr(F_STOP, posK): risk = tr(F_RISK, posK)
        pts = (gLastC - entry) * side
        If pendHas Then
            Band ws, r, "⚠ " & pend & "：下一根 K 開盤出場" & IIf(forming, "（以現價計，收K時再確認）", ""), cWarn(), cWarnBg(), True
            r = r + 1
        End If
        Txt ws, r, 2, "進場", 9, False, cMuted()
        Txt ws, r, 3, "目前停損", 9, False, cMuted()
        Txt ws, r, 5, "未實現", 9, False, cMuted()
        Txt ws, r, 6, "初始停損", 9, False, cMuted()
        Txt ws, r + 1, 2, Round(entry), 16, True
        Txt ws, r + 1, 3, Round(stp), 16, True
        Money ws, r + 1, 5, Round(pts)
        Txt ws, r + 1, 6, Round(tr(F_STOP0, posK)), 16, True
        ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1, 6)).NumberFormat = "#,##0"
        ws.Cells(r + 1, 5).NumberFormat = "+#,##0"" 點"";-#,##0"" 點"";0"" 點"""
        ws.Cells(r + 1, 5).Font.Size = 16: ws.Cells(r + 1, 5).Font.Bold = True
        ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1, 6)).HorizontalAlignment = -4131
        ws.Rows(r + 1).RowHeight = 24
        Txt ws, r + 2, 2, FmtTm(tr(F_ENTRYT, posK), 1) & "・" & tr(F_KIND, posK), 9, False, cMuted()
        Txt ws, r + 2, 3, "距現價 " & N0((gLastC - stp) * side) & " 點", 9, False, cMuted()
        Txt ws, r + 2, 5, S0(pts * MULT - FEE) & " 元（1 口微台）", 9, False, cMuted()
        Txt ws, r + 2, 6, "風險 " & N0(risk) & " 點", 9, False, cMuted()
        If tr(F_BE, posK) Then
            s = "已保本"
        Else
            s = "收盤到 " & N0(entry + 1# * risk * side) & " 就把停損移到進場價（保本）"
        End If
        If strat = 1 Then
            s = s & "。保本後停損改成「進場後" & IIf(side > 0, "最高收盤 − 5 ATR", "最低收盤 + 5 ATR") & "」，只往有利方向移動；目前 ATR " & N0(gVATR(gVN - 1)) & "。"
        Else
            s = s & "。保本後，收盤" & IIf(side > 0, "跌破", "站上") & "日 EMA20 就隔天開盤出場；目前 EMA20 " & N0(gDE20(gDN - 1)) & "。"
        End If
        Txt ws, r + 3, 2, s, 9, False, cMuted()
        r = r + 4
    End If
    ' 指標
    If strat = 1 Then
        k = gVN - 1
        s = "EMA20 " & N0(gVE20(k)) & "・EMA60 " & N0(gVE60(k)) & "・布林 " & N0(gVMID(k) - 2 * gVSD(k)) & "–" & N0(gVMID(k) + 2 * gVSD(k)) & "・ATR " & N0(gVATR(k))
    Else
        k = gDN - 1
        s = "日EMA20 " & N0(gDE20(k)) & "・EMA60 " & N0(gDE60(k)) & "・日ATR " & N0(gDATR(k)) & "・最後日K 開" & N0(gDO(k)) & " 高" & N0(gDH(k)) & _
            " 低" & N0(gDL(k)) & " 收" & N0(gDC(k))
    End If
    Txt ws, r, 2, s, 9, False, cMuted()
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 7)).Borders(8).Color = cLine()
    r = r + 1
    ' 條件明細
    If posK >= 0 Then
        Txt ws, r, 2, "進場條件明細（持倉中，不會再開新倉）", 10, True, cMuted()
    Else
        Txt ws, r, 2, "進場條件（看最後一根 K" & IIf(forming, "，以現價計）", "）"), 10, True
    End If
    r = r + 1
    r = ChecksBlock(ws, r, strat)
    ' 外框
    With ws.Range(ws.Cells(r0, 2), ws.Cells(r - 1, 7))
        .Borders(7).Color = cLine(): .Borders(8).Color = cLine(): .Borders(9).Color = cLine(): .Borders(10).Color = cLine()
    End With
    StratCard = r
End Function

Private Function EventRow(ByVal ws As Worksheet, ByVal r As Long, n As Long, ByVal prevGen As Long, ByVal ev As String) As Long
    If n = 0 Then
        Band ws, r, "上次更新（" & FmtTm(prevGen, 1) & "）之後：", cWarn(), cWarnBg()
        r = r + 1
    End If
    Band ws, r, "　" & ev, cWarn(), cWarnBg()
    n = n + 1
    EventRow = r + 1
End Function

Private Function ChecksBlock(ByVal ws As Worksheet, ByVal r0 As Long, ByVal strat As Long) As Long
    Dim sd As Long, col As Long, r As Long, i As Long, g As Variant, groups As Variant, allOk As Boolean, badge As String
    Dim gi As Long, title As String, nG As Long, rr As Long
    If strat = 1 Then groups = Array("突破", "拉回") Else groups = Array("突破")
    nG = UBound(groups) + 1
    For sd = 1 To 2
        col = IIf(sd = 1, 2, 5)
        r = r0
        badge = ""
        For Each g In groups
            allOk = True
            For i = 0 To gCkN - 1
                If gCkStrat(i) = strat And gCkSide(i) = sd And (gCkGroup(i) = "common" Or gCkGroup(i) = g) Then
                    If Not gCkOk(i) Then allOk = False
                End If
            Next
            If allOk Then badge = badge & IIf(Len(badge) > 0, "、", "") & g
        Next
        Txt ws, r, col, "做" & IIf(sd = 1, "多", "空"), 11, True, IIf(sd = 1, cUp(), cDn())
        With ws.Cells(r, col + 1)
            If Len(badge) > 0 Then
                .Value = badge & "條件全到"
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = cAccent()
                .Font.Bold = True
            Else
                .Value = "未成立"
                .Font.Color = cMuted()
                .Interior.Color = cHead()
            End If
            .HorizontalAlignment = -4108
        End With
        r = r + 1
        r = CkRows(ws, r, col, strat, sd, "common")
        For Each g In groups
            title = IIf(nG > 1, g & "進場", "進場條件")
            Txt ws, r, col, title, 9, True, cMuted()
            r = r + 1
            r = CkRows(ws, r, col, strat, sd, CStr(g))
        Next
        With ws.Range(ws.Cells(r, col), ws.Cells(r, col + 2))
            .Merge
            .Value = "若現在進場，停損約 " & N0(gStopVal(strat, sd)) & "（" & N0(Abs(gStopRaw(strat, sd))) & " 點前的端點，最多 1–1.5 ATR）"
            .Font.Size = 9
            .Font.Color = cMuted()
            .WrapText = True
            .VerticalAlignment = -4160
        End With
        ws.Rows(r).RowHeight = 26
        rr = r + 1
    Next
    ChecksBlock = rr
End Function

Private Function CkRows(ByVal ws As Worksheet, ByVal r As Long, ByVal col As Long, ByVal strat As Long, ByVal sd As Long, ByVal grp As String) As Long
    Dim i As Long
    For i = 0 To gCkN - 1
        If gCkStrat(i) = strat And gCkSide(i) = sd And gCkGroup(i) = grp Then
            Txt ws, r, col, IIf(gCkOk(i), "✓ ", "✗ ") & gCkName(i), 10, gCkOk(i), IIf(gCkOk(i), cAccent(), cMuted())
            Txt ws, r, col + 1, gCkDetail(i), 9, False, cMuted()
            ws.Cells(r, col + 1).HorizontalAlignment = -4152
            r = r + 1
        End If
    Next
    CkRows = r
End Function

Private Function RecordBlock(ByVal ws As Worksheet, ByVal r0 As Long, ByVal strat As Long, ByVal title As String) As Long
    Dim tr() As Variant, tn As Long, k As Long, r As Long, fs As Long, nF As Long, fw As Long, fR As Double, fT As Double
    Dim nB As Long, bR As Double, bT As Double, rr As Double, hasAny As Boolean
    If strat = 1 Then
        tr = gVTr: tn = gVTn
    Else
        tr = gDTr: tn = gDTn
    End If
    For k = 0 To tn - 1
        rr = PyRound(RowR(tr, k, True), 2)
        If RowPhase(tr, k) = "前測" Then
            If Not tr(F_OPEN, k) Then
                nF = nF + 1
                If rr > 0.1 Then fw = fw + 1
                fR = fR + Round(rr * 100)
            End If
            fT = fT + RowTwd(tr, k, True)
        Else
            If Not tr(F_OPEN, k) Then
                nB = nB + 1
                bR = bR + Round(rr * 100)
            End If
            bT = bT + RowTwd(tr, k, True)
        End If
    Next
    r = r0
    Txt ws, r, 2, title, 12, True
    r = r + 1
    Txt ws, r, 2, "前測（2026/11/09 起）已出場 " & nF & " 筆、賺 " & fw & " 筆、合計 " & SumR1(fR) & "R、" & S0(fT) & " 元（含未實現）", 9, False, cMuted()
    r = r + 1
    Dim heads As Variant
    heads = Array("進場", "進場價", "R", "出場", "出場價", "損益（元）")
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 7)).Value = heads
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 7))
        .Font.Color = cMuted()
        .Font.Size = 9
        .Interior.Color = cHead()
    End With
    r = r + 1
    For k = tn - 1 To 0 Step -1
        If RowPhase(tr, k) = "前測" Then
            hasAny = True
            RecordRow ws, r, tr, k
            r = r + 1
        End If
    Next
    If Not hasAny Then
        Txt ws, r, 2, "還沒有交易。", 9, False, cMuted()
        r = r + 1
    End If
    Txt ws, r, 2, "前測開始前（台指期累積資料回測）：" & nB & " 筆 " & SumR1(bR) & "R、" & S0(bT) & " 元", 9, False, cMuted()
    On Error Resume Next
    ws.Hyperlinks.Add ws.Cells(r, 5), "", "'" & SH_TRADES & "'!A1", "", "完整歷史成交紀錄 →"
    On Error GoTo 0
    r = r + 1
    RecordBlock = r
End Function

Private Sub RecordRow(ByVal ws As Worksheet, ByVal r As Long, tr() As Variant, ByVal k As Long)
    Dim side As Long
    side = tr(F_SIDE, k)
    ws.Cells(r, 2).Value = FmtTm(tr(F_ENTRYT, k), 1) & " " & SideTxt(side) & "・" & tr(F_KIND, k)
    ws.Cells(r, 3).Value = Round(tr(F_ENTRY, k))
    ws.Cells(r, 3).NumberFormat = "#,##0"
    ws.Cells(r, 4).Value = PyRound(RowR(tr, k, True), 2)
    ws.Cells(r, 4).NumberFormat = "+0.00;-0.00;0.00"
    If tr(F_OPEN, k) Then
        ws.Cells(r, 5).Value = "持倉中（停損 " & N0(tr(F_STOP, k)) & "）"
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 7)).Font.Bold = True
    Else
        ws.Cells(r, 5).Value = FmtTm(tr(F_EXITT, k), 1) & " " & WhyText(tr, k)
        ws.Cells(r, 6).Value = Round(tr(F_EXIT, k))
        ws.Cells(r, 6).NumberFormat = "#,##0"
    End If
    Money ws, r, 7, RowTwd(tr, k, True)
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 7)).Interior.Color = IIf(side > 0, cRowL(), cRowS())
End Sub

' ============================================================
' 成交紀錄工作表
' ============================================================
Public Sub RenderTrades(ByVal nowTm As Long)
    Dim ws As Worksheet, r As Long, w As Variant, i As Long
    Set ws = Sh(SH_TRADES)
    ws.Cells.Clear
    On Error Resume Next
    ws.Hyperlinks.Delete
    On Error GoTo 0
    ws.Cells.Font.Size = 10
    w = Array(2, 7, 17, 10, 10, 10, 17, 10, 24, 9, 8, 11, 12)
    For i = 0 To UBound(w)
        ws.Columns(i + 1).ColumnWidth = w(i)
    Next
    Txt ws, 1, 2, "台指策略成交紀錄", 18, True
    ws.Rows(1).RowHeight = 30
    Txt ws, 2, 2, "v3（60 分K）與 D2（日K）每一筆模擬交易・台北 " & FmtTm(nowTm) & " 產生・台指期持倉以現價 " & N0(gLastC) & " 計", 0, False, cMuted()
    On Error Resume Next
    ws.Hyperlinks.Add ws.Cells(2, 12), "", "'" & SH_LIVE & "'!A1", "", "← 即時訊號"
    On Error GoTo 0
    r = 4
    Txt ws, r, 2, "台指期・實際交易依據（" & FmtTm(gVT(0), 2) & " 起；v3 含夜盤，D2 只算日盤）", 12, True, cMuted()
    r = r + 1
    r = Ledger(ws, r, "v3・60 分K", "進場是訊號 K 收完後下一根 60 分K 開盤；停損出場時間是 1 分K 第一次碰到停損的那一分鐘。「前測」是 2026/11/9 規則定案後的交易。", gVTr, gVTn, True, True)
    r = Ledger(ws, r, "D2・日K（只算日盤）", "日K 只用日盤 08:45–13:45，夜盤不算；進場是訊號當天 13:45 收盤；停損出場時間是日盤 1 分K 第一次碰到停損的那一分鐘；保本後收盤破 EMA20 是隔天 08:45 開盤出場。", gDTr, gDTn, True, True)
    r = r + 1
    Txt ws, r, 2, "加權指數・回測用歷史資料（2024/11/07 起，近 2 年，只有日盤；不是實際交易依據）" & IIf(gTwN > 0, "・資料到 " & FmtTm(gTwLast, 2), "・沒有資料"), 12, True, cMuted()
    r = r + 1
    r = Ledger(ws, r, "v3・60 分K", "價格是加權指數點數，當成 1 口微台每點 10 元試算。資料只有 60 分K，所以時間是那根 60 分K 的開始時間（例如 10:00 代表 10:00–11:00 之間）。", gWVTr, gWVTn, False, False)
    r = Ledger(ws, r, "D2・日K", "進場是訊號當天 13:30 收盤；停損出場時間是第一根碰到停損的 60 分K；保本後收盤破 EMA20 是隔天 09:00 開盤出場。", gWDTr, gWDTn, False, False)
    r = r + 1
    Txt ws, r, 2, "試算：本金 200,000 元、固定 1 口微台（每點 10 元）、每次下單 25 元（來回 50 元）；點數是進出場價差，損益（元）已扣下單費用；R 是以初始停損點數為 1 的倍數（另扣 3 點滑價成本）。", 9, False, cMuted()
    Txt ws, r + 1, 2, "台指期價格已換算成目前近月合約（換月時回溯調整），所以早期的價格會跟當時實際報價差一個價差。", 9, False, cMuted()
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
    Sh(SH_LIVE).Activate
End Sub

Private Function Ledger(ByVal ws As Worksheet, ByVal r0 As Long, ByVal title As String, ByVal note As String, tr() As Variant, _
                        ByVal tn As Long, ByVal showPhase As Boolean, ByVal useLast As Boolean) As Long
    Dim r As Long, k As Long, eq As Double, pk As Double, dd As Double, nC As Long, win As Long, sumR As Double, unreal As Double
    Dim eqs() As Double, side As Long, c As Long, heads As Variant, rr As Double
    r = r0
    Txt ws, r, 2, title, 13, True
    r = r + 1
    Txt ws, r, 2, note, 9, False, cMuted()
    r = r + 1
    If tn = 0 Then
        Txt ws, r, 2, "還沒有交易。", 9, False, cMuted()
        Ledger = r + 2
        Exit Function
    End If
    ReDim eqs(0 To tn - 1)
    eq = CAPITAL: pk = CAPITAL: dd = 0
    For k = 0 To tn - 1
        eq = eq + RowTwd(tr, k, useLast)
        If eq > pk Then pk = eq
        If eq - pk < dd Then dd = eq - pk
        eqs(k) = eq
        rr = PyRound(RowR(tr, k, useLast), 2)
        If tr(F_OPEN, k) Then
            unreal = unreal + RowTwd(tr, k, useLast)
        Else
            nC = nC + 1
            If rr > 0.1 Then win = win + 1
            sumR = sumR + Round(rr * 100)
        End If
    Next
    ' 摘要
    Txt ws, r, 3, "帳戶", 9, False, cMuted()
    Txt ws, r, 5, "已出場", 9, False, cMuted()
    Txt ws, r, 7, "合計", 9, False, cMuted()
    Txt ws, r, 9, "最大回落", 9, False, cMuted()
    Txt ws, r + 1, 3, eq, 14, True: ws.Cells(r + 1, 3).NumberFormat = "#,##0"
    Txt ws, r + 1, 5, nC & " 筆", 14, True
    Txt ws, r + 1, 7, SumR1(sumR) & "R", 14, True
    Txt ws, r + 1, 9, dd, 14, True: ws.Cells(r + 1, 9).NumberFormat = "#,##0"
    Dim pct As Double
    If nC > 0 Then pct = win / nC Else pct = 0
    Txt ws, r + 2, 3, S0(eq - CAPITAL) & " 元（" & SN((eq - CAPITAL) / CAPITAL * 100, 1) & "%）", 9, False, cMuted()
    Txt ws, r + 2, 5, "賺 " & win & " 筆・勝率 " & Format$(pct, "0%"), 9, False, cMuted()
    Txt ws, r + 2, 7, IIf(unreal <> 0, "未實現 " & S0(unreal) & " 元", "沒有持倉"), 9, False, cMuted()
    Txt ws, r + 2, 9, "元", 9, False, cMuted()
    ws.Range(ws.Cells(r, 3), ws.Cells(r + 2, 10)).HorizontalAlignment = -4131
    r = r + 3
    heads = Array("階段", "進場時間", "方向", "進場價", "初始停損", "出場時間", "出場價", "出場原因", "點數", "R", "損益（元）", "帳戶")
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 13)).Value = heads
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 13))
        .Font.Color = cMuted()
        .Interior.Color = cHead()
        .Font.Bold = True
    End With
    If Not showPhase Then ws.Cells(r, 2).Value = ""
    r = r + 1
    For k = tn - 1 To 0 Step -1
        side = tr(F_SIDE, k)
        If showPhase Then
            ws.Cells(r, 2).Value = RowPhase(tr, k)
            If RowPhase(tr, k) = "前測" Then
                ws.Cells(r, 2).Font.Color = cAccent()
                ws.Cells(r, 2).Font.Bold = True
            End If
        End If
        ws.Cells(r, 3).Value = "'" & FmtTm(tr(F_ENTRYT, k))
        ws.Cells(r, 4).Value = SideTxt(side) & "・" & tr(F_KIND, k)
        ws.Cells(r, 5).Value = Round(tr(F_ENTRY, k))
        ws.Cells(r, 6).Value = Round(tr(F_STOP0, k))
        If tr(F_OPEN, k) Then
            ws.Cells(r, 7).Value = "持倉中"
            ws.Cells(r, 9).Value = "停損 " & N0(tr(F_STOP, k))
        Else
            ws.Cells(r, 7).Value = "'" & FmtTm(tr(F_EXITT, k))
            ws.Cells(r, 8).Value = Round(tr(F_EXIT, k))
            ws.Cells(r, 9).Value = WhyText(tr, k)
        End If
        ws.Cells(r, 10).Value = RowPts(tr, k, useLast)
        ws.Cells(r, 11).Value = PyRound(RowR(tr, k, useLast), 2)
        Money ws, r, 12, RowTwd(tr, k, useLast)
        ws.Cells(r, 13).Value = eqs(k)
        ws.Range(ws.Cells(r, 5), ws.Cells(r, 8)).NumberFormat = "#,##0"
        ws.Cells(r, 10).NumberFormat = "+#,##0;-#,##0;0"
        ws.Cells(r, 11).NumberFormat = "+0.00;-0.00;0.00"
        ws.Cells(r, 13).NumberFormat = "#,##0"
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 13)).Interior.Color = IIf(side > 0, cRowL(), cRowS())
        If tr(F_OPEN, k) Then ws.Range(ws.Cells(r, 3), ws.Cells(r, 13)).Font.Bold = True
        r = r + 1
    Next
    Ledger = r + 1
End Function

' 測試用：不連網，用資料夾裡的資料算一次並畫出兩張工作表
Public Function RenderOffline(ByVal dataDir As String, ByVal nowStr As String, ByVal fixtureDir As String) As String
    Dim nowTm As Long
    On Error GoTo bad
    nowTm = DayNum(Val(Left$(nowStr, 4)), Val(Mid$(nowStr, 6, 2)), Val(Mid$(nowStr, 9, 2))) * 1440& + Val(Mid$(nowStr, 12, 2)) * 60 + Val(Mid$(nowStr, 15, 2))
    ResetUnits
    LoadHistory JoinPath(dataDir, "tx_1min")
    ResetLive
    If Len(fixtureDir) > 0 Then
        gFixtureDir = fixtureDir
        FetchLive Year(nowTm \ 1440)
        gFixtureDir = ""
    End If
    AddLiveUnits
    FindGap nowTm
    ComputeAll JoinPath(dataDir, "tx_daily.csv"), JoinPath(dataDir, "twii_60min.csv"), nowTm
    RenderLive nowTm, nowTm - 600, ""
    RenderTrades nowTm
    RenderOffline = "ok"
    Exit Function
bad:
    RenderOffline = "ERR " & Err.Description
End Function
