'================================================================
'  模組名稱：modElecBill
'  說明：電費資料自動新增與圖表更新
'  使用方式：Alt+F11 → Insert Module → 貼上全部程式碼
'================================================================
Option Explicit

'--- 費率常數 -------------------------------------------------
Private Const RA_S As String = "1.68,2.45,3.70,5.04,6.24,8.46"
Private Const RA_N As String = "1.68,2.16,3.03,4.14,5.07,6.63"
Private Const RB_S As String = "1.78,2.55,3.80,5.14,6.44,8.86"
Private Const RB_N As String = "1.78,2.26,3.13,4.24,5.27,7.03"

'--- 累進電費計算 ---------------------------------------------
Public Function CalcFlow(ByVal deg As Long, _
                          ByVal season As String, _
                          ByVal ver As String) As Long
    Dim rStr As String
    If ver = "A" Then
        rStr = IIf(season = "夏季", RA_S, RA_N)
    Else
        rStr = IIf(season = "夏季", RB_S, RB_N)
    End If
    Dim r() As String: r = Split(rStr, ",")
    Dim lims(5) As Long
    lims(0)=120: lims(1)=330: lims(2)=500
    lims(3)=700: lims(4)=999: lims(5)=99999

    Dim total As Double, rem As Long, prev As Long
    rem = deg: prev = 0
    Dim i As Integer
    For i = 0 To 5
        Dim chunk As Long
        chunk = WorksheetFunction.Min(rem, lims(i) - prev)
        total = total + chunk * CDbl(r(i))
        rem = rem - chunk: prev = lims(i)
        If rem <= 0 Then Exit For
    Next i
    CalcFlow = CLng(WorksheetFunction.Round(total, 0))
End Function

'--- 自動判斷費率版本 -----------------------------------------
Public Function GetVersion(period As String) As String
    On Error Resume Next
    Dim yr As Long, mo As Long
    yr = CLng(Left(period, 3))
    mo = CLng(Mid(period, 5, 2))
    GetVersion = IIf(yr > 114 Or (yr = 114 And mo >= 10), "B", "A")
End Function

'--- 自動判斷半年 ---------------------------------------------
Public Function GetHalf(period As String) As String
    On Error Resume Next
    Dim mo As Long: mo = CLng(Mid(period, 5, 2))
    GetHalf = Left(period, 3) & IIf(mo <= 6, "上", "下")
End Function

'--- 主程序：從 Input Sheet 讀取並新增 -----------------------
Public Sub AddNewRecord()
    Dim wi As Worksheet
    Set wi = ThisWorkbook.Sheets("Input")

    '--- 讀取輸入值
    Dim period As String
    Dim degStr As String, totalStr As String
    Dim season As String

    period   = Trim(CStr(wi.Range("C4").Value))
    degStr   = Trim(CStr(wi.Range("C6").Value))
    totalStr = Trim(CStr(wi.Range("C8").Value))
    season   = Trim(CStr(wi.Range("C10").Value))

    '--- 驗證
    If period = "" Or degStr = "" Or totalStr = "" Then
        MsgBox "請填寫：帳單月份、用電度數、總電費", vbExclamation, "輸入不完整"
        Exit Sub
    End If
    If Len(period) <> 6 Or Mid(period, 4, 1) <> "/" Then
        MsgBox "帳單月份格式錯誤" & Chr(10) & "正確格式：115/07", vbExclamation, "格式錯誤"
        wi.Range("C4").Select: Exit Sub
    End If
    If Not IsNumeric(degStr) Or CLng(degStr) <= 0 Then
        MsgBox "用電度數請輸入正整數", vbExclamation, "格式錯誤"
        wi.Range("C6").Select: Exit Sub
    End If
    If Not IsNumeric(totalStr) Or CLng(totalStr) <= 0 Then
        MsgBox "總電費請輸入正整數", vbExclamation, "格式錯誤"
        wi.Range("C8").Select: Exit Sub
    End If
    If season <> "夏季" And season <> "非夏季" Then
        MsgBox "季節請選擇：夏季 或 非夏季", vbExclamation, "格式錯誤"
        wi.Range("C10").Select: Exit Sub
    End If

    Dim deg As Long:   deg   = CLng(degStr)
    Dim total As Long: total = CLng(totalStr)

    '--- 計算
    Dim ver As String:     ver     = GetVersion(period)
    Dim half As String:    half    = GetHalf(period)
    Dim yr As String:      yr      = Left(period, 3)
    Dim flowFee As Long:   flowFee = CalcFlow(deg, season, ver)
    Dim pubFee As Long:    pubFee  = total - flowFee
    Dim pct As Double:     pct     = CDbl(flowFee) / CDbl(total)

    '--- 確認對話框
    Dim msg As String
    msg = "確認新增以下資料？" & Chr(10) & Chr(10) & _
          "帳單月份：" & period & Chr(10) & _
          "用電度數：" & Format(deg, "#,##0") & " 度" & Chr(10) & _
          "總電費：NT$ " & Format(total, "#,##0") & Chr(10) & _
          "季節：" & season & "（費率" & ver & "）" & Chr(10) & Chr(10) & _
          "─────────────────" & Chr(10) & _
          "流動費用：NT$ " & Format(flowFee, "#,##0") & Chr(10) & _
          "公共費用：NT$ " & Format(pubFee, "#,##0") & Chr(10) & _
          "流動費用%：" & Format(pct, "0.0%")
    If MsgBox(msg, vbYesNo + vbQuestion, "確認新增") = vbNo Then Exit Sub

    '--- 寫入 Raw Data
    Application.ScreenUpdating = False
    Dim wrd As Worksheet: Set wrd = ThisWorkbook.Sheets("Raw Data")
    Dim nRow As Long
    nRow = wrd.Cells(wrd.Rows.Count, "A").End(xlUp).Row + 1

    With wrd
        .Cells(nRow, 1).Value  = period
        .Cells(nRow, 2).Value  = deg
        .Cells(nRow, 3).Value  = total
        .Cells(nRow, 4).Value  = season
        .Cells(nRow, 5).Value  = ver
        .Cells(nRow, 6).Value  = flowFee
        .Cells(nRow, 7).Value  = pubFee
        .Cells(nRow, 8).Value  = pct
        .Cells(nRow, 9).Value  = ""
        .Cells(nRow, 10).Value = half
        .Cells(nRow, 11).Value = yr

        '-- 格式套用
        Dim rng As Range
        Set rng = .Range(.Cells(nRow,1), .Cells(nRow,11))
        Dim bg As Long
        bg = IIf(nRow Mod 2 = 0, RGB(239,246,255), RGB(255,255,255))
        With rng
            .Interior.Color     = bg
            .Font.Name          = "Arial"
            .Font.Size          = 10
            .Borders.LineStyle  = xlContinuous
            .Borders.Color      = RGB(226,232,240)
            .Borders.Weight     = xlThin
        End With
        .Rows(nRow).RowHeight = 20
        .Cells(nRow,2).NumberFormat = "#,##0"
        .Cells(nRow,3).NumberFormat = "#,##0"
        .Cells(nRow,6).NumberFormat = "#,##0"
        .Cells(nRow,7).NumberFormat = "#,##0"
        .Cells(nRow,8).NumberFormat = "0.0%"
        .Cells(nRow,1).Font.Bold  = True
        .Cells(nRow,1).Font.Color = RGB(0,0,204)
        .Cells(nRow,2).Font.Color = RGB(0,0,204)
        .Cells(nRow,3).Font.Color = RGB(0,0,204)
        .Cells(nRow,5).Font.Color = RGB(4,120,87)
        .Cells(nRow,10).Font.Color= RGB(4,120,87)
        .Cells(nRow,11).Font.Color= RGB(4,120,87)
    End With

    '--- 更新 Dashboard 右側明細表
    Call RefreshDashboardTable

    '--- 強制重算（Data sheet 公式 + 圖表）
    Application.Calculate
    ThisWorkbook.RefreshAll

    '--- 清空輸入格
    wi.Range("C4").Value  = ""
    wi.Range("C6").Value  = ""
    wi.Range("C8").Value  = ""
    wi.Range("C10").Value = "非夏季"

    '--- 預覽顯示
    wi.Range("C17").Value = "版本 " & ver
    wi.Range("C18").Value = "NT$ " & Format(flowFee, "#,##0")
    wi.Range("C19").Value = "NT$ " & Format(pubFee,  "#,##0")
    wi.Range("C20").Value = half

    Application.ScreenUpdating = True

    MsgBox "✅ " & period & " 新增完成！" & Chr(10) & _
           "圖表與統計已自動更新。", vbInformation, "新增成功"
End Sub

'--- 更新 Dashboard 右側明細表 --------------------------------
Public Sub RefreshDashboardTable()
    Dim wrd As Worksheet, wdash As Worksheet
    Set wrd   = ThisWorkbook.Sheets("Raw Data")
    Set wdash = ThisWorkbook.Sheets("Dashboard")

    Dim TC As Long:  TC = 17   ' 從 Q 欄開始
    Dim SR As Long:  SR = 3    ' 從第 3 行開始

    '-- 清空舊內容
    Dim r As Long, c As Long
    For r = SR To SR + 55
        For c = TC To TC + 6
            wdash.Cells(r, c).Value = ""
        Next c
    Next r

    '-- 從 Raw Data 重填
    Dim lastRow As Long
    lastRow = wrd.Cells(wrd.Rows.Count,"A").End(xlUp).Row
    Dim ri As Long: ri = SR
    Dim i As Long

    For i = 4 To lastRow
        If CStr(wrd.Cells(i,1).Value) <> "" Then
            Dim period  As String: period  = CStr(wrd.Cells(i,1).Value)
            Dim season2 As String: season2 = CStr(wrd.Cells(i,4).Value)
            Dim deg2    As Long:   deg2    = CLng(wrd.Cells(i,2).Value)
            Dim fl      As Long:   fl      = CLng(wrd.Cells(i,6).Value)
            Dim tot     As Long:   tot     = CLng(wrd.Cells(i,3).Value)
            Dim pubV    As Variant: pubV   = wrd.Cells(i,7).Value

            Dim bg2 As Long
            bg2 = IIf(ri Mod 2 = 0, RGB(239,246,255), RGB(255,255,255))

            wdash.Cells(ri,TC).Value   = period
            wdash.Cells(ri,TC+1).Value = season2
            wdash.Cells(ri,TC+2).Value = Format(deg2,"#,##0") & " 度"
            wdash.Cells(ri,TC+3).Value = "$" & Format(fl,"#,##0")

            If Not IsNumeric(pubV) Or CDbl(pubV) < 0 Then
                wdash.Cells(ri,TC+4).Value = "溢繳抵扣"
                wdash.Cells(ri,TC+4).Font.Color = RGB(180,83,9)
                wdash.Cells(ri,TC+6).Value = "N/A"
            Else
                wdash.Cells(ri,TC+4).Value = "$" & Format(CLng(pubV),"#,##0")
                If tot > 0 Then
                    wdash.Cells(ri,TC+6).Value = Format(CDbl(fl)/CDbl(tot),"0.0%")
                End If
            End If
            wdash.Cells(ri,TC+5).Value = "$" & Format(tot,"#,##0")

            Dim rowRng As Range
            Set rowRng = wdash.Range(wdash.Cells(ri,TC), wdash.Cells(ri,TC+6))
            rowRng.Interior.Color = bg2
            rowRng.Font.Name = "Arial"
            rowRng.Font.Size = 9

            ri = ri + 1
        End If
    Next i

    '-- 更新彙總數字（SUMPRODUCT 公式已在 Data sheet，這裡做簡單加總）
    Dim sF As Long, sP As Long, sT As Long, cnt As Long
    For i = 4 To lastRow
        If CStr(wrd.Cells(i,1).Value) <> "" And IsNumeric(wrd.Cells(i,7).Value) _
           And CDbl(wrd.Cells(i,7).Value) >= 0 Then
            sF  = sF  + CLng(wrd.Cells(i,6).Value)
            sP  = sP  + CLng(wrd.Cells(i,7).Value)
            sT  = sT  + CLng(wrd.Cells(i,3).Value)
            cnt = cnt + 1
        End If
    Next i

    '-- 寫回彙總 block
    Dim sumStart As Long: sumStart = ri + 2
    Dim summaries(5,1) As Variant
    summaries(0,0) = "累計總電費":   summaries(0,1) = "NT$ " & Format(sT,"#,##0")
    summaries(1,0) = "流動費用":     summaries(1,1) = "NT$ " & Format(sF,"#,##0")
    summaries(2,0) = "公共費用":     summaries(2,1) = "NT$ " & Format(sP,"#,##0")
    If sT > 0 Then
        summaries(3,0) = "流動費用%": summaries(3,1) = Format(CDbl(sF)/CDbl(sT),"0.0%")
        summaries(4,0) = "公共費用%": summaries(4,1) = Format(CDbl(sP)/CDbl(sT),"0.0%")
    End If
    summaries(5,0) = "資料期數":     summaries(5,1) = cnt & " 期"

    Dim j2 As Integer
    For j2 = 0 To 5
        wdash.Cells(sumStart+j2, TC).Value   = summaries(j2,0)
        wdash.Cells(sumStart+j2, TC+4).Value = summaries(j2,1)
    Next j2
End Sub
