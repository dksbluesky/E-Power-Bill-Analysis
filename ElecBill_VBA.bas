'================================================================
'  模組名稱：modElecBill
'  說明：電費資料自動新增與圖表更新
'  版本：v2 (2025) - 更新雙月240度級距 + 節電獎勵欄位
'  使用方式：Alt+F11 → Insert Module → 貼上全部程式碼
'================================================================
Option Explicit

'--- 費率常數（雙月計費，第一段240度）--------------------------
' 版本A：113/04/01–114/09/30
Private Const RA_S As String = "1.68,2.45,3.70,5.04,6.24,8.46"
Private Const RA_N As String = "1.68,2.16,3.03,4.14,5.07,6.63"
' 版本B：114/10/01 起
Private Const RB_S As String = "1.78,2.55,3.80,5.14,6.44,8.86"
Private Const RB_N As String = "1.78,2.26,3.13,4.24,5.27,7.03"

'--- 雙月累進電費計算（第一段240度）----------------------------
Public Function CalcFlow(ByVal deg As Long, _
                          ByVal season As String, _
                          ByVal ver As String) As Double
    Dim rStr As String
    If ver = "A" Then
        rStr = IIf(season = "夏季", RA_S, RA_N)
    Else
        rStr = IIf(season = "夏季", RB_S, RB_N)
    End If
    Dim r() As String: r = Split(rStr, ",")

    ' 雙月級距：240, 660, 1000, 1400, 2000, 99999
    Dim lims(5) As Long
    lims(0)=240: lims(1)=660: lims(2)=1000
    lims(3)=1400: lims(4)=2000: lims(5)=99999

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
    CalcFlow = WorksheetFunction.Round(total, 1)
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

'--- 主程序：從 Input Sheet 新增 ------------------------------
Public Sub AddNewRecord()
    Dim wi As Worksheet
    Set wi = ThisWorkbook.Sheets("Input")

    '--- 讀取輸入值
    Dim period  As String: period   = Trim(CStr(wi.Range("C4").Value))
    Dim degStr  As String: degStr   = Trim(CStr(wi.Range("C6").Value))
    Dim totalStr As String: totalStr = Trim(CStr(wi.Range("C8").Value))
    Dim season  As String: season   = Trim(CStr(wi.Range("C10").Value))

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

    Dim deg   As Long: deg   = CLng(degStr)
    Dim total As Long: total = CLng(totalStr)

    '--- 計算
    Dim ver     As String: ver     = GetVersion(period)
    Dim half    As String: half    = GetHalf(period)
    Dim yr      As String: yr      = Left(period, 3)
    Dim flowFee As Double: flowFee = CalcFlow(deg, season, ver)
    Dim pubFee  As Double: pubFee  = total - flowFee  ' 節電獎勵新增後再調整
    Dim pct     As Double: pct     = CDbl(flowFee) / CDbl(total)

    '--- 確認對話框
    Dim msg As String
    msg = "確認新增以下資料？" & Chr(10) & Chr(10) & _
          "帳單月份：" & period & Chr(10) & _
          "用電度數：" & Format(deg, "#,##0") & " 度" & Chr(10) & _
          "總電費：NT$ " & Format(total, "#,##0") & Chr(10) & _
          "季節：" & season & "（費率版本" & ver & "）" & Chr(10) & Chr(10) & _
          "─────────────────" & Chr(10) & _
          "流動費用估算：NT$ " & Format(flowFee, "#,##0.0") & Chr(10) & _
          "公共費用估算：NT$ " & Format(pubFee, "#,##0.0") & Chr(10) & _
          "流動費用%：" & Format(pct, "0.0%") & Chr(10) & Chr(10) & _
          "⚠ 節電獎勵請新增後手動填入 L 欄"
    If MsgBox(msg, vbYesNo + vbQuestion, "確認新增") = vbNo Then Exit Sub

    '--- 寫入 Raw Data
    Application.ScreenUpdating = False
    Dim wrd As Worksheet: Set wrd = ThisWorkbook.Sheets("Raw Data")
    Dim nRow As Long
    nRow = wrd.Cells(wrd.Rows.Count, "A").End(xlUp).Row + 1

    With wrd
        ' A: 帳單月份
        .Cells(nRow, 1).Value  = period
        ' B: 度數
        .Cells(nRow, 2).Value  = deg
        ' C: 總電費
        .Cells(nRow, 3).Value  = total
        ' D: 季節
        .Cells(nRow, 4).Value  = season
        ' E: 費率版本（公式）
        .Cells(nRow, 5).Formula = _
            "=IF(A" & nRow & "="""","""",IF(OR(VALUE(LEFT(A" & nRow & ",3))>114," & _
            "AND(VALUE(LEFT(A" & nRow & ",3))=114,VALUE(MID(A" & nRow & ",5,2))>=10)),""B"",""A""))"
        ' F: 流動費用（公式，雙月240度級距）
        .Cells(nRow, 6).Formula = _
            "=IF(OR(B" & nRow & "="""",D" & nRow & "=""""),""""," & _
            "ROUND(MIN(B" & nRow & ",240)*INDEX($N$4:$S$7,MATCH(E" & nRow & "&D" & nRow & ",$M$4:$M$7,0),1)" & _
            "+MAX(MIN(B" & nRow & ",660)-240,0)*INDEX($N$4:$S$7,MATCH(E" & nRow & "&D" & nRow & ",$M$4:$M$7,0),2)" & _
            "+MAX(MIN(B" & nRow & ",1000)-660,0)*INDEX($N$4:$S$7,MATCH(E" & nRow & "&D" & nRow & ",$M$4:$M$7,0),3)" & _
            "+MAX(MIN(B" & nRow & ",1400)-1000,0)*INDEX($N$4:$S$7,MATCH(E" & nRow & "&D" & nRow & ",$M$4:$M$7,0),4)" & _
            "+MAX(MIN(B" & nRow & ",2000)-1400,0)*INDEX($N$4:$S$7,MATCH(E" & nRow & "&D" & nRow & ",$M$4:$M$7,0),5)" & _
            "+MAX(B" & nRow & "-2000,0)*INDEX($N$4:$S$7,MATCH(E" & nRow & "&D" & nRow & ",$M$4:$M$7,0),6),1))"
        ' G: 公共費用（公式，包含節電獎勵）
        .Cells(nRow, 7).Formula = _
            "=IF(OR(C" & nRow & "="""",F" & nRow & "=""""),"""",C" & nRow & "-F" & nRow & "+L" & nRow & ")"
        ' H: 流動%
        .Cells(nRow, 8).Formula = _
            "=IF(OR(C" & nRow & "=0,F" & nRow & "=""""),"""",F" & nRow & "/C" & nRow & ")"
        ' I: 備註
        .Cells(nRow, 9).Value  = ""
        ' J: 半年（公式）
        .Cells(nRow, 10).Formula = _
            "=IF(A" & nRow & "="""","""",LEFT(A" & nRow & ",3)&IF(VALUE(MID(A" & nRow & ",5,2))<=6,""上"",""下""))"
        ' K: 年度（公式）
        .Cells(nRow, 11).Formula = _
            "=IF(A" & nRow & "="""","""",LEFT(A" & nRow & ",3))"
        ' L: 節電獎勵（預設0，請手動更新為實際金額）
        .Cells(nRow, 12).Value = 0

        '-- 格式套用
        Dim rng As Range
        Set rng = .Range(.Cells(nRow, 1), .Cells(nRow, 12))
        Dim bg As Long
        bg = IIf(nRow Mod 2 = 0, RGB(239,246,255), RGB(255,255,255))
        With rng
            .Interior.Color    = bg
            .Font.Name         = "Arial"
            .Font.Size         = 10
            .Borders.LineStyle = xlContinuous
            .Borders.Color     = RGB(226,232,240)
            .Borders.Weight    = xlThin
        End With
        .Rows(nRow).RowHeight = 20

        ' 數字格式
        .Cells(nRow, 2).NumberFormat  = "#,##0"
        .Cells(nRow, 3).NumberFormat  = "#,##0"
        .Cells(nRow, 6).NumberFormat  = "#,##0.0"
        .Cells(nRow, 7).NumberFormat  = "#,##0.0"
        .Cells(nRow, 8).NumberFormat  = "0.0%"
        .Cells(nRow, 12).NumberFormat = "#,##0.0"

        ' 字色
        .Cells(nRow, 1).Font.Color  = RGB(0, 0, 204)
        .Cells(nRow, 1).Font.Bold   = True
        .Cells(nRow, 2).Font.Color  = RGB(0, 0, 204)
        .Cells(nRow, 3).Font.Color  = RGB(0, 0, 204)
        .Cells(nRow, 5).Font.Color  = RGB(4, 120, 87)
        .Cells(nRow, 10).Font.Color = RGB(4, 120, 87)
        .Cells(nRow, 11).Font.Color = RGB(4, 120, 87)
        .Cells(nRow, 12).Font.Color = RGB(148, 163, 184)  ' 灰色提示待填
    End With

    '--- 更新 Dashboard 右側明細表
    Call RefreshDashboardTable

    '--- 強制重算
    Application.Calculate
    ThisWorkbook.RefreshAll

    '--- 清空輸入格
    wi.Range("C4").Value  = ""
    wi.Range("C6").Value  = ""
    wi.Range("C8").Value  = ""
    wi.Range("C10").Value = "非夏季"

    '--- 預覽顯示
    wi.Range("C17").Value = "版本 " & ver
    wi.Range("C18").Value = "NT$ " & Format(flowFee, "#,##0.0") & "（估算）"
    wi.Range("C19").Value = "NT$ " & Format(pubFee, "#,##0.0") & "（節電獎勵待填L欄）"
    wi.Range("C20").Value = half

    Application.ScreenUpdating = True

    MsgBox "✅ " & period & " 新增完成！" & Chr(10) & Chr(10) & _
           "⚠ 請到 Raw Data 的 L 欄填入節電獎勵金額" & Chr(10) & _
           "  （115年以後帳單才有，無則保持 0）", vbInformation, "新增成功"
End Sub

'--- 更新 Dashboard 右側明細表 --------------------------------
Public Sub RefreshDashboardTable()
    Dim wrd   As Worksheet: Set wrd   = ThisWorkbook.Sheets("Raw Data")
    Dim wdash As Worksheet: Set wdash = ThisWorkbook.Sheets("Dashboard")

    Dim TC As Long: TC = 17   ' Q 欄
    Dim SR As Long: SR = 3    ' 第 3 行起

    ' 清空舊內容
    Dim r As Long, c As Long
    For r = SR To SR + 55
        For c = TC To TC + 6
            wdash.Cells(r, c).Value = ""
        Next c
    Next r

    ' 從 Raw Data 重填
    Dim lastRow As Long
    lastRow = wrd.Cells(wrd.Rows.Count, "A").End(xlUp).Row
    Dim ri As Long: ri = SR
    Dim i As Long

    For i = 4 To lastRow
        If CStr(wrd.Cells(i, 1).Value) <> "" Then
            Dim period2  As String: period2  = CStr(wrd.Cells(i, 1).Value)
            Dim season2  As String: season2  = CStr(wrd.Cells(i, 4).Value)
            Dim deg2     As Long:   deg2     = CLng(wrd.Cells(i, 2).Value)
            Dim fl       As Double
            Dim tot      As Long:   tot      = CLng(wrd.Cells(i, 3).Value)
            Dim pubV     As Variant: pubV    = wrd.Cells(i, 7).Value
            Dim nodeV    As Variant: nodeV   = wrd.Cells(i, 12).Value

            On Error Resume Next
            fl = CDbl(wrd.Cells(i, 6).Value)
            On Error GoTo 0

            Dim bg2 As Long
            bg2 = IIf(ri Mod 2 = 0, RGB(239,246,255), RGB(255,255,255))

            wdash.Cells(ri, TC).Value   = period2
            wdash.Cells(ri, TC+1).Value = season2
            wdash.Cells(ri, TC+2).Value = Format(deg2, "#,##0") & " 度"
            wdash.Cells(ri, TC+3).Value = "$" & Format(fl, "#,##0.0")

            If Not IsNumeric(pubV) Or CDbl(pubV) < 0 Then
                wdash.Cells(ri, TC+4).Value = "溢繳抵扣"
                wdash.Cells(ri, TC+4).Font.Color = RGB(180, 83, 9)
                wdash.Cells(ri, TC+6).Value = "N/A"
            Else
                wdash.Cells(ri, TC+4).Value = "$" & Format(CDbl(pubV), "#,##0.0")
                ' 節電獎勵
                If IsNumeric(nodeV) And CDbl(nodeV) > 0 Then
                    wdash.Cells(ri, TC+4).Value = "$" & Format(CDbl(pubV), "#,##0.0") & _
                        " (-$" & Format(CDbl(nodeV), "0.0") & "節電)"
                End If
                If tot > 0 Then
                    wdash.Cells(ri, TC+6).Value = Format(CDbl(fl)/CDbl(tot), "0.0%")
                End If
            End If
            wdash.Cells(ri, TC+5).Value = "$" & Format(tot, "#,##0")

            Dim rowRng As Range
            Set rowRng = wdash.Range(wdash.Cells(ri,TC), wdash.Cells(ri,TC+6))
            rowRng.Interior.Color = bg2
            rowRng.Font.Name = "Arial"
            rowRng.Font.Size = 9

            ri = ri + 1
        End If
    Next i

    ' 彙總
    Dim sF As Double, sP As Double, sT As Long, cnt As Long
    For i = 4 To lastRow
        If CStr(wrd.Cells(i,1).Value) <> "" And IsNumeric(wrd.Cells(i,7).Value) _
           And CDbl(wrd.Cells(i,7).Value) >= 0 Then
            sF  = sF  + CDbl(wrd.Cells(i,6).Value)
            sP  = sP  + CDbl(wrd.Cells(i,7).Value)
            sT  = sT  + CLng(wrd.Cells(i,3).Value)
            cnt = cnt + 1
        End If
    Next i

    Dim sumStart As Long: sumStart = ri + 2
    Dim summaries(5,1) As Variant
    summaries(0,0)="累計總電費":  summaries(0,1)="NT$ "&Format(sT,"#,##0")
    summaries(1,0)="流動費用":    summaries(1,1)="NT$ "&Format(sF,"#,##0.0")
    summaries(2,0)="公共費用":    summaries(2,1)="NT$ "&Format(sP,"#,##0.0")
    If sT > 0 Then
        summaries(3,0)="流動費用%": summaries(3,1)=Format(CDbl(sF)/CDbl(sT),"0.0%")
        summaries(4,0)="公共費用%": summaries(4,1)=Format(CDbl(sP)/CDbl(sT),"0.0%")
    End If
    summaries(5,0)="資料期數":    summaries(5,1)=cnt&" 期"

    Dim j2 As Integer
    For j2 = 0 To 5
        wdash.Cells(sumStart+j2, TC).Value   = summaries(j2,0)
        wdash.Cells(sumStart+j2, TC+4).Value = summaries(j2,1)
    Next j2
End Sub
