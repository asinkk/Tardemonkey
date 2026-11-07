Option Explicit
' 測試入口（LibreOffice 或 Excel 都可以跑）：不連網，用 history 與測試資料算一次，結果寫成文字檔
Public Function RunOffline(ByVal dataDir As String, ByVal nowStr As String, ByVal fixtureDir As String, ByVal dumpPath As String) As String
    Dim nowTm As Long, t0 As Double, stp As String
    On Error GoTo bad
    t0 = Timer
    stp = "load"
    nowTm = DayNum(Val(Left$(nowStr, 4)), Val(Mid$(nowStr, 6, 2)), Val(Mid$(nowStr, 9, 2))) * 1440& + Val(Mid$(nowStr, 12, 2)) * 60 + Val(Mid$(nowStr, 15, 2))
    ResetUnits
    LoadHistory JoinPath(dataDir, "tx_1min")
    ResetLive
    If Len(fixtureDir) > 0 Then
        gFixtureDir = fixtureDir
        FetchLive Year(nowTm \ 1440)
        gFixtureDir = ""
    End If
    stp = "live"
    AddLiveUnits
    FindGap nowTm
    stp = "compute"
    ComputeAll JoinPath(dataDir, "tx_daily.csv"), JoinPath(dataDir, "twii_60min.csv"), nowTm
    stp = "dump"
    DumpAll dumpPath
    RunOffline = "units=" & gUN & " rows=" & gRN & " v3 trades=" & gVTn & " d2 trades=" & gDTn & " time=" & Format$(Timer - t0, "0.0")
    Exit Function
bad:
    RunOffline = "ERR " & stp & ": " & Err.Number & " " & Err.Description & " step=" & gStep
End Function
