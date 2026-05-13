Option Explicit

' =====================================================================
'  使用流程
'  1. 將資料置於工作表（必含欄位 A=Case No, M=Product Name, AK=Month,
'     AL=Part No, AP=Cost）。在資料工作表上執行 SetupCompareMenu。
'  2. 切到 Menu 工作表，於 B3 (當月)、B4 (前月) 下拉選單選擇月份。
'  3. 執行 CreateWarrantyBubbleChartDiff 產出 BubbleData 與差異散佈圖。
'
'  資料過濾
'  - 僅納入「Warranty Coverage = Warranty」之資料列
'
'  圖表規格
'  - 散佈圖 (xlXYScatter)
'  - X 軸：Total Cost Diff   (BubbleData!J)
'  - Y 軸：Cost per Case Diff (BubbleData!H)
'  - 標記顏色：J > 0 紅、J < 0 綠、J = 0 灰
'  - 資料標籤：僅 Total Cost Diff 前 5 名（紅底紅邊）+ 後 5 名（綠底綠邊）
'  - BubbleData 依 J 欄 (Total Cost Diff) 由大到小排序，全部產品顯示
' =====================================================================

' ---------- Step A：建立月份比對選單 ----------
Sub SetupCompareMenu()
    Dim wsData As Worksheet, wsMenu As Worksheet
    Dim lastRow As Long, i As Long, k As Long
    Dim dictMonths As Object
    Dim monthVal As String

    Application.ScreenUpdating = False

    Set wsData = ActiveSheet
    If wsData.Name = "Menu" Or wsData.Name = "BubbleData" Then
        Application.ScreenUpdating = True
        MsgBox "請先切到原始資料工作表，再執行此巨集。", vbExclamation
        Exit Sub
    End If

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        Application.ScreenUpdating = True
        MsgBox "資料工作表沒有可讀資料。", vbExclamation
        Exit Sub
    End If

    ' 從 AK 欄收集唯一月份
    Set dictMonths = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        monthVal = Trim(CStr(wsData.Cells(i, "AK").Value))
        If monthVal <> "" And LCase(monthVal) <> "month" Then
            If Not dictMonths.Exists(monthVal) Then
                dictMonths.Add monthVal, True
            End If
        End If
    Next i

    If dictMonths.Count = 0 Then
        Application.ScreenUpdating = True
        MsgBox "在 AK 欄找不到任何月份資料。", vbExclamation
        Exit Sub
    End If

    ' 月份陣列並做字串遞增排序
    Dim months() As Variant
    ReDim months(1 To dictMonths.Count)
    k = 0
    Dim key As Variant
    For Each key In dictMonths.Keys
        k = k + 1
        months(k) = key
    Next key

    Dim a As Long, b As Long, tmp As Variant
    For a = 1 To k - 1
        For b = 1 To k - a
            If CStr(months(b)) > CStr(months(b + 1)) Then
                tmp = months(b)
                months(b) = months(b + 1)
                months(b + 1) = tmp
            End If
        Next b
    Next a

    ' 建立 / 清空 Menu 工作表
    On Error Resume Next
    Set wsMenu = ThisWorkbook.Sheets("Menu")
    On Error GoTo 0

    If wsMenu Is Nothing Then
        Set wsMenu = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsMenu.Name = "Menu"
    Else
        wsMenu.Cells.Clear
        On Error Resume Next
        wsMenu.Cells.Validation.Delete
        On Error GoTo 0
    End If

    ' 版面
    With wsMenu.Range("A1:B1")
        .Merge
        .Value = "月份比對選單"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .RowHeight = 28
    End With

    wsMenu.Range("A3").Value = "當月"
    wsMenu.Range("A4").Value = "前月"
    With wsMenu.Range("A3:A4")
        .Font.Bold = True
        .Interior.Color = RGB(217, 226, 243)
        .HorizontalAlignment = xlCenter
    End With
    With wsMenu.Range("B3:B4")
        .Interior.Color = RGB(255, 255, 204)
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 180, 180)
    End With

    ' 下拉清單來源放在隱藏的 E 欄
    For i = 1 To k
        wsMenu.Cells(i, "E").Value = months(i)
    Next i

    Dim monthRange As String
    monthRange = "=$E$1:$E$" & k

    With wsMenu.Range("B3").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=monthRange
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    With wsMenu.Range("B4").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=monthRange
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    If k >= 1 Then wsMenu.Range("B3").Value = months(k)
    If k >= 2 Then wsMenu.Range("B4").Value = months(k - 1)

    wsMenu.Range("G1").Value = "DataSheet"
    wsMenu.Range("G2").Value = wsData.Name
    wsMenu.Columns("E").Hidden = True
    wsMenu.Columns("G").Hidden = True

    wsMenu.Columns("A").ColumnWidth = 14
    wsMenu.Columns("B").ColumnWidth = 22

    wsMenu.Range("A6").Value = "操作說明"
    wsMenu.Range("A7").Value = "1. 在 B3、B4 下拉選單選擇要比對的月份"
    wsMenu.Range("A8").Value = "2. 執行巨集 CreateWarrantyBubbleChartDiff 產出散佈圖"
    wsMenu.Range("A6").Font.Bold = True
    wsMenu.Range("A6:A8").Font.Size = 10

    wsMenu.Activate
    wsMenu.Range("B3").Select

    Application.ScreenUpdating = True
    MsgBox "Menu 工作表已建立，請選擇當月與前月後執行 CreateWarrantyBubbleChartDiff。", vbInformation
End Sub

' ---------- Step B：依選單建立差異散佈圖 ----------
Sub CreateWarrantyBubbleChartDiff()
    Dim wsData As Worksheet, wsMenu As Worksheet, wsSummary As Worksheet
    Dim lastRow As Long, i As Long, r As Long
    Dim currMonth As String, prevMonth As String
    Dim dataSheetName As String

    Application.ScreenUpdating = False

    On Error Resume Next
    Set wsMenu = ThisWorkbook.Sheets("Menu")
    On Error GoTo 0
    If wsMenu Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "找不到 Menu 工作表，請先在資料表執行 SetupCompareMenu。", vbExclamation
        Exit Sub
    End If

    currMonth = Trim(CStr(wsMenu.Range("B3").Value))
    prevMonth = Trim(CStr(wsMenu.Range("B4").Value))
    dataSheetName = Trim(CStr(wsMenu.Range("G2").Value))

    If currMonth = "" Or prevMonth = "" Then
        Application.ScreenUpdating = True
        MsgBox "請先在 Menu 的 B3、B4 選擇當月與前月。", vbExclamation
        Exit Sub
    End If
    If currMonth = prevMonth Then
        Application.ScreenUpdating = True
        MsgBox "當月與前月不能相同，請重新選擇。", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets(dataSheetName)
    On Error GoTo 0
    If wsData Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "找不到資料工作表 [" & dataSheetName & "]，請重新執行 SetupCompareMenu。", vbExclamation
        Exit Sub
    End If

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        Application.ScreenUpdating = True
        MsgBox "資料工作表沒有可讀資料。", vbExclamation
        Exit Sub
    End If

    ' === 尋找 "Warranty Coverage" 欄位（標題列可能在第 1 或第 2 列，掃描前 5 列）===
    Dim warrCol As Long, lastCol As Long, hdr As Long
    Dim headerRow As Long, tmpCol As Long
    warrCol = 0
    lastCol = 1
    For headerRow = 1 To 5
        tmpCol = wsData.Cells(headerRow, wsData.Columns.Count).End(xlToLeft).Column
        If tmpCol > lastCol Then lastCol = tmpCol
    Next headerRow

    For headerRow = 1 To 5
        For hdr = 1 To lastCol
            If LCase(Trim(CStr(wsData.Cells(headerRow, hdr).Value))) = "warranty coverage" Then
                warrCol = hdr
                Exit For
            End If
        Next hdr
        If warrCol > 0 Then Exit For
    Next headerRow

    If warrCol = 0 Then
        Application.ScreenUpdating = True
        MsgBox "找不到 'Warranty Coverage' 欄位（請確認資料表前 5 列含此標題名稱）。", vbExclamation
        Exit Sub
    End If

    ' === Step 1：依「月份|產品」收集案件與成本（僅 Warranty Coverage = Warranty）===
    Dim dCases As Object, dTotal As Object, dProducts As Object
    Set dCases = CreateObject("Scripting.Dictionary")
    Set dTotal = CreateObject("Scripting.Dictionary")
    Set dProducts = CreateObject("Scripting.Dictionary")

    Dim caseNo As String, productName As String, partNo As String
    Dim monthVal As String, currentCost As Double
    Dim cKey As String, coverage As String

    For i = 2 To lastRow
        caseNo = Trim(CStr(wsData.Cells(i, "A").Value))
        productName = Trim(CStr(wsData.Cells(i, "M").Value))
        partNo = Trim(CStr(wsData.Cells(i, "AL").Value))
        monthVal = Trim(CStr(wsData.Cells(i, "AK").Value))
        coverage = Trim(CStr(wsData.Cells(i, warrCol).Value))
        currentCost = 0
        If IsNumeric(wsData.Cells(i, "AP").Value) Then
            currentCost = CDbl(wsData.Cells(i, "AP").Value)
        End If

        If caseNo <> "" And LCase(caseNo) <> "case no" Then
            If productName <> "" And LCase(productName) <> "product name" Then
                If monthVal = currMonth Or monthVal = prevMonth Then
                    If LCase(coverage) = "warranty" Then
                        cKey = monthVal & "|" & productName
                        If Not dCases.Exists(cKey) Then
                            dCases.Add cKey, CreateObject("Scripting.Dictionary")
                            dTotal.Add cKey, 0#
                        End If
                        If Not dCases(cKey).Exists(caseNo) Then
                            dCases(cKey).Add caseNo, True
                        End If
                        If Len(partNo) > 0 Then
                            If Left(partNo, 1) = "0" Or Left(partNo, 1) = "1" Then
                                dTotal(cKey) = dTotal(cKey) + currentCost
                            End If
                        End If
                        If Not dProducts.Exists(productName) Then
                            dProducts.Add productName, True
                        End If
                    End If
                End If
            End If
        End If
    Next i

    If dProducts.Count = 0 Then
        Application.ScreenUpdating = True
        MsgBox "選擇月份 (" & currMonth & " / " & prevMonth & ") 找不到資料。", vbExclamation
        Exit Sub
    End If

    ' === Step 2：彙整每個產品的當月 / 前月 / 差異 ===
    Dim n As Long
    n = dProducts.Count
    ' 欄位順序:
    '   1 Product, 2 Curr CostPerCase, 3 Curr QTY, 4 Curr Total
    '   5 Prev CostPerCase, 6 Prev QTY, 7 Prev Total
    '   8 Diff CostPerCase (Y), 9 Diff QTY, 10 Diff Total (X)
    Dim arr() As Variant
    ReDim arr(1 To n, 1 To 10)

    Dim key As Variant, k As Long
    Dim cur As String, prv As String
    Dim cQty As Long, cTotal As Double, cPer As Double
    Dim pQty As Long, pTotal As Double, pPer As Double

    k = 0
    For Each key In dProducts.Keys
        k = k + 1
        arr(k, 1) = key
        cur = currMonth & "|" & key
        prv = prevMonth & "|" & key

        cQty = 0: cTotal = 0
        If dCases.Exists(cur) Then
            cQty = dCases(cur).Count
            cTotal = dTotal(cur)
        End If
        If cQty > 0 Then cPer = cTotal / cQty Else cPer = 0

        pQty = 0: pTotal = 0
        If dCases.Exists(prv) Then
            pQty = dCases(prv).Count
            pTotal = dTotal(prv)
        End If
        If pQty > 0 Then pPer = pTotal / pQty Else pPer = 0

        arr(k, 2) = cPer:  arr(k, 3) = cQty:  arr(k, 4) = cTotal
        arr(k, 5) = pPer:  arr(k, 6) = pQty:  arr(k, 7) = pTotal
        arr(k, 8) = cPer - pPer
        arr(k, 9) = cQty - pQty
        arr(k, 10) = cTotal - pTotal
    Next key

    ' === Step 2.5：依 Total Cost Diff 由大到小排序（含正負） ===
    Dim a As Long, b As Long, c As Long, tmp As Variant
    For a = 1 To n - 1
        For b = 1 To n - a
            If arr(b, 10) < arr(b + 1, 10) Then
                For c = 1 To 10
                    tmp = arr(b, c)
                    arr(b, c) = arr(b + 1, c)
                    arr(b + 1, c) = tmp
                Next c
            End If
        Next b
    Next a

    ' === Step 3：寫入 BubbleData（全部產品） ===
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("BubbleData").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsSummary = ThisWorkbook.Sheets.Add(After:=wsMenu)
    wsSummary.Name = "BubbleData"

    With wsSummary.Range("A1:J1")
        .Value = Array("Product Name", _
                       currMonth & " Cost per Case", currMonth & " Repair QTY", currMonth & " Total Cost", _
                       prevMonth & " Cost per Case", prevMonth & " Repair QTY", prevMonth & " Total Cost", _
                       "Cost per Case Diff (Y)", "Repair QTY Diff", "Total Cost Diff (X)")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    wsSummary.Rows(1).RowHeight = 32

    For r = 1 To n
        wsSummary.Cells(r + 1, 1).Value = arr(r, 1)
        wsSummary.Cells(r + 1, 2).Value = arr(r, 2)
        wsSummary.Cells(r + 1, 3).Value = arr(r, 3)
        wsSummary.Cells(r + 1, 4).Value = arr(r, 4)
        wsSummary.Cells(r + 1, 5).Value = arr(r, 5)
        wsSummary.Cells(r + 1, 6).Value = arr(r, 6)
        wsSummary.Cells(r + 1, 7).Value = arr(r, 7)
        wsSummary.Cells(r + 1, 8).Value = arr(r, 8)
        wsSummary.Cells(r + 1, 9).Value = arr(r, 9)
        wsSummary.Cells(r + 1, 10).Value = arr(r, 10)
    Next r

    wsSummary.Range("B2:B" & (n + 1)).NumberFormat = "#,##0.00"
    wsSummary.Range("D2:D" & (n + 1)).NumberFormat = "#,##0.00"
    wsSummary.Range("E2:E" & (n + 1)).NumberFormat = "#,##0.00"
    wsSummary.Range("G2:G" & (n + 1)).NumberFormat = "#,##0.00"
    wsSummary.Range("H2:H" & (n + 1)).NumberFormat = "#,##0.00;[Red]-#,##0.00"
    wsSummary.Range("I2:I" & (n + 1)).NumberFormat = "0;[Red]-0"
    wsSummary.Range("J2:J" & (n + 1)).NumberFormat = "#,##0.00;[Red]-#,##0.00"

    ' J 欄條件式醒目：正值底紅、負值底綠
    With wsSummary.Range("J2:J" & (n + 1)).FormatConditions
        .Delete
        With .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            .Interior.Color = RGB(252, 228, 228)
            .Font.Color = RGB(155, 28, 42)
        End With
        With .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            .Interior.Color = RGB(226, 244, 230)
            .Font.Color = RGB(20, 110, 45)
        End With
    End With

    wsSummary.Columns("A:J").AutoFit

    ' === Step 4：X / Y 軸刻度（含負值區間，通過原點） ===
    ' X 軸 = J 欄 (Total Cost Diff)；Y 軸 = H 欄 (Cost per Case Diff)
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double
    Dim vCell As Range
    xMin = 0: xMax = 0: yMin = 0: yMax = 0
    For Each vCell In wsSummary.Range("J2:J" & (n + 1))
        If IsNumeric(vCell.Value) Then
            If vCell.Value > xMax Then xMax = vCell.Value
            If vCell.Value < xMin Then xMin = vCell.Value
        End If
    Next vCell
    For Each vCell In wsSummary.Range("H2:H" & (n + 1))
        If IsNumeric(vCell.Value) Then
            If vCell.Value > yMax Then yMax = vCell.Value
            If vCell.Value < yMin Then yMin = vCell.Value
        End If
    Next vCell

    Dim xLow As Double, xHigh As Double, xUnit As Double
    Dim yLow As Double, yHigh As Double, yUnit As Double

    xLow = NiceFloor(xMin * 1.15)
    xHigh = NiceCeiling(xMax * 1.15)
    If xLow = 0 And xHigh = 0 Then xHigh = 1
    If xLow = xHigh Then xHigh = xLow + 1
    xUnit = NiceStep(xHigh - xLow)
    If xUnit <= 0 Then xUnit = (xHigh - xLow) / 5

    yLow = NiceFloor(yMin * 1.15)
    yHigh = NiceCeiling(yMax * 1.15)
    If yLow = 0 And yHigh = 0 Then yHigh = 1
    If yLow = yHigh Then yHigh = yLow + 1
    yUnit = NiceStep(yHigh - yLow)
    If yUnit <= 0 Then yUnit = (yHigh - yLow) / 5

    ' === Step 5：繪製散佈圖 ===
    Dim cht As ChartObject
    Set cht = wsSummary.ChartObjects.Add(Left:=wsSummary.Range("M1").Left, Top:=10, Width:=1200, Height:=620)

    With cht.Chart
        .ChartType = xlXYScatter
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop

        Dim s As Series
        Set s = .SeriesCollection.NewSeries
        s.Name = "Diff (" & currMonth & " - " & prevMonth & ")"
        ' X 軸 = J 欄 (Total Cost Diff)；Y 軸 = H 欄 (Cost per Case Diff)
        s.XValues = wsSummary.Range("J2:J" & (n + 1))
        s.Values = wsSummary.Range("H2:H" & (n + 1))

        ' --- 關閉系列連接線（純散佈點）---
        On Error Resume Next
        s.Format.Line.Visible = msoFalse
        s.Border.LineStyle = xlNone
        On Error GoTo 0

        ' --- 美化：圖表外觀 ---
        On Error Resume Next
        With .ChartArea.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(248, 250, 253)
            .Solid
        End With
        .ChartArea.Format.Line.Visible = msoTrue
        .ChartArea.Format.Line.ForeColor.RGB = RGB(220, 220, 220)
        .ChartArea.Format.Line.Weight = 0.75
        .ChartArea.Font.Name = "Calibri"

        With .PlotArea.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
            .Solid
        End With
        .PlotArea.Format.Line.Visible = msoTrue
        .PlotArea.Format.Line.ForeColor.RGB = RGB(210, 215, 222)
        .PlotArea.Format.Line.Weight = 0.75
        On Error GoTo 0

        ' --- 標題 ---
        .HasTitle = True
        With .ChartTitle
            .Text = "Warranty Total Cost Diff: " & currMonth & " vs " & prevMonth & _
                    " (" & n & " products)"
            .Font.Size = 16
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Color = RGB(48, 56, 80)
        End With

        ' --- X 軸：Total Cost Diff ---
        .Axes(xlCategory).HasTitle = True
        With .Axes(xlCategory).AxisTitle
            .Text = "Total Cost Diff (" & currMonth & " - " & prevMonth & ")"
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = RGB(70, 80, 100)
            .Font.Name = "Calibri"
        End With
        On Error Resume Next
        With .Axes(xlCategory)
            .MinimumScale = xLow
            .MaximumScale = xHigh
            .MajorUnit = xUnit
            .TickLabels.NumberFormat = "#,##0;-#,##0"
            .TickLabels.Font.Size = 10
            .TickLabels.Font.Color = RGB(89, 89, 89)
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 232, 238)
            .MajorGridlines.Format.Line.DashStyle = msoLineDash
            .Format.Line.ForeColor.RGB = RGB(120, 130, 145)
            .Format.Line.Weight = 1.25
            .CrossesAt = 0
        End With
        On Error GoTo 0

        ' --- Y 軸：Cost per Case Diff ---
        .Axes(xlValue).HasTitle = True
        With .Axes(xlValue).AxisTitle
            .Text = "Cost per Case Diff (" & currMonth & " - " & prevMonth & ")"
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = RGB(70, 80, 100)
            .Font.Name = "Calibri"
            .Orientation = xlUpward
        End With
        On Error Resume Next
        With .Axes(xlValue)
            .MinimumScale = yLow
            .MaximumScale = yHigh
            .MajorUnit = yUnit
            .TickLabels.NumberFormat = "#,##0;-#,##0"
            .TickLabels.Font.Size = 10
            .TickLabels.Font.Color = RGB(89, 89, 89)
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 232, 238)
            .MajorGridlines.Format.Line.DashStyle = msoLineDash
            .Format.Line.ForeColor.RGB = RGB(120, 130, 145)
            .Format.Line.Weight = 1.25
            .CrossesAt = 0
        End With
        On Error GoTo 0

        .HasLegend = False

        ' --- 標記樣式 + 顏色（依 J 欄正負）---
        DoEvents
        s.MarkerStyle = xlMarkerStyleCircle
        s.MarkerSize = 12

        Dim p As Long, fillClr As Long, edgeClr As Long, diffVal As Double
        For p = 1 To s.Points.Count
            diffVal = 0
            If IsNumeric(wsSummary.Cells(p + 1, 10).Value) Then
                diffVal = CDbl(wsSummary.Cells(p + 1, 10).Value)
            End If
            If diffVal > 0 Then
                fillClr = RGB(220, 53, 69)    ' 紅：上升
                edgeClr = RGB(155, 28, 42)
            ElseIf diffVal < 0 Then
                fillClr = RGB(40, 167, 69)    ' 綠：下降
                edgeClr = RGB(20, 110, 45)
            Else
                fillClr = RGB(170, 170, 170)
                edgeClr = RGB(110, 110, 110)
            End If
            On Error Resume Next
            With s.Points(p)
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 12
                .MarkerForegroundColor = edgeClr
                .MarkerBackgroundColor = fillClr
                ' 不使用 Format.Line（避免畫出點與點連接線）
                .Format.Line.Visible = msoFalse
            End With
            On Error GoTo 0
        Next p

        ' --- 資料標籤：僅前 10 名 (最大正差) + 後 10 名 (最大負差)，靠右側垂直排列避免重疊 ---
        DoEvents

        ' 先用 ApplyDataLabels 建立 DataLabel 物件，再關閉所有點的顯示
        s.ApplyDataLabels Type:=xlDataLabelsShowValue, _
                          AutoText:=True, _
                          HasLeaderLines:=True
        With s.DataLabels
            .ShowSeriesName = False
            .ShowCategoryName = False
            .ShowValue = False
            .ShowBubbleSize = False
            .Font.Size = 9
            .Font.Bold = True
            .Font.Name = "Calibri"
        End With
        On Error Resume Next
        s.HasLeaderLines = True
        On Error GoTo 0

        Dim j As Long
        For j = 1 To s.Points.Count
            On Error Resume Next
            s.Points(j).HasDataLabel = False
            On Error GoTo 0
        Next j

        ' 計算前 5 名與後 5 名（n 不足 10 時自動退讓）
        Dim topCnt As Long, botCnt As Long
        If n >= 5 Then topCnt = 5 Else topCnt = n
        If (n - topCnt) >= 5 Then botCnt = 5 Else botCnt = n - topCnt

        ' 縮小 PlotArea，讓右側騰出標籤欄
        DoEvents
        Dim chartW As Double, chartH As Double
        chartW = .ChartArea.Width
        chartH = .ChartArea.Height
        Dim panelW As Double
        panelW = 220
        On Error Resume Next
        .PlotArea.Left = 50
        .PlotArea.Top = 55
        .PlotArea.Width = chartW - panelW - 60
        .PlotArea.Height = chartH - 120
        On Error GoTo 0
        DoEvents

        Dim plotL As Double, plotT As Double, plotW As Double, plotH As Double
        On Error Resume Next
        plotL = .PlotArea.InsideLeft
        plotT = .PlotArea.InsideTop
        plotW = .PlotArea.InsideWidth
        plotH = .PlotArea.InsideHeight
        On Error GoTo 0
        If plotW <= 0 Then
            plotL = 60: plotT = 60
            plotW = chartW - panelW - 80
            plotH = chartH - 130
        End If

        Dim labelX As Double
        labelX = plotL + plotW + 12

        ' 上半段放 top 10（正差，紅色），下半段放 bottom 10（負差，綠色）
        Dim halfH As Double
        halfH = plotH / 2#
        Dim topSpacing As Double, botSpacing As Double
        If topCnt > 0 Then topSpacing = halfH / topCnt Else topSpacing = 0
        If botCnt > 0 Then botSpacing = halfH / botCnt Else botSpacing = 0

        Dim lblText As String, kk As Long, bIdx As Long

        ' Top labels: 索引 1..topCnt（最大正差 = 成本上升 Top 5）
        For kk = 1 To topCnt
            lblText = " " & CStr(wsSummary.Cells(kk + 1, 1).Value) & " "
            On Error Resume Next
            With s.Points(kk)
                .HasDataLabel = True
                With .DataLabel
                    .ShowSeriesName = False
                    .ShowCategoryName = False
                    .ShowValue = False
                    .ShowBubbleSize = False
                    .Text = lblText
                    .Format.TextFrame2.TextRange.Text = lblText
                    .Font.Size = 11
                    .Font.Bold = True
                    .Font.Color = RGB(155, 28, 42)
                    .Font.Name = "Calibri"
                    ' --- 美化：淡紅底 + 紅邊框 ---
                    With .Format.Fill
                        .Visible = msoTrue
                        .ForeColor.RGB = RGB(253, 232, 232)
                        .Solid
                        .Transparency = 0.05
                    End With
                    With .Format.Line
                        .Visible = msoTrue
                        .ForeColor.RGB = RGB(220, 53, 69)
                        .Weight = 1
                        .DashStyle = msoLineSolid
                    End With
                    .Left = labelX
                    .Top = plotT + (kk - 1) * topSpacing + 4
                End With
            End With
            On Error GoTo 0
        Next kk

        ' Bottom labels: 索引 n - botCnt + 1 .. n（最大負差 = 成本下降 Top 5）
        For kk = 1 To botCnt
            bIdx = n - botCnt + kk
            lblText = " " & CStr(wsSummary.Cells(bIdx + 1, 1).Value) & " "
            On Error Resume Next
            With s.Points(bIdx)
                .HasDataLabel = True
                With .DataLabel
                    .ShowSeriesName = False
                    .ShowCategoryName = False
                    .ShowValue = False
                    .ShowBubbleSize = False
                    .Text = lblText
                    .Format.TextFrame2.TextRange.Text = lblText
                    .Font.Size = 11
                    .Font.Bold = True
                    .Font.Color = RGB(20, 110, 45)
                    .Font.Name = "Calibri"
                    ' --- 美化：淡綠底 + 綠邊框 ---
                    With .Format.Fill
                        .Visible = msoTrue
                        .ForeColor.RGB = RGB(232, 247, 232)
                        .Solid
                        .Transparency = 0.05
                    End With
                    With .Format.Line
                        .Visible = msoTrue
                        .ForeColor.RGB = RGB(40, 167, 69)
                        .Weight = 1
                        .DashStyle = msoLineSolid
                    End With
                    .Left = labelX
                    .Top = plotT + halfH + (kk - 1) * botSpacing + 4
                End With
            End With
            On Error GoTo 0
        Next kk
    End With

    wsSummary.Activate
    Application.ScreenUpdating = True
    MsgBox "已建立差異散佈圖：" & currMonth & " vs " & prevMonth & " (僅含 Warranty Coverage = Warranty，共 " & n & " 項產品)。" & vbCrLf & _
           "資料標籤：僅顯示 Total Cost Diff 前 5 名（紅）與後 5 名（綠）" & vbCrLf & _
           "Total Cost Diff > 0 → 紅色（成本上升）" & vbCrLf & _
           "Total Cost Diff < 0 → 綠色（成本下降）", vbInformation
End Sub

' ===== 取「漂亮」的上界（1 / 2 / 2.5 / 5 × 10^n）=====
Private Function NiceCeiling(ByVal v As Double) As Double
    If v <= 0 Then
        NiceCeiling = 0
        Exit Function
    End If
    Dim ex As Double, f As Double, nf As Double
    ex = Int(Log(v) / Log(10#))
    f = v / (10# ^ ex)
    If f <= 1 Then
        nf = 1
    ElseIf f <= 2 Then
        nf = 2
    ElseIf f <= 2.5 Then
        nf = 2.5
    ElseIf f <= 5 Then
        nf = 5
    Else
        nf = 10
    End If
    NiceCeiling = nf * (10# ^ ex)
End Function

' ===== 取「漂亮」的下界（負值對應的 floor）=====
Private Function NiceFloor(ByVal v As Double) As Double
    If v >= 0 Then
        NiceFloor = 0
        Exit Function
    End If
    NiceFloor = -NiceCeiling(-v)
End Function

' ===== 取「漂亮」的主要刻度間距 =====
Private Function NiceStep(ByVal rng As Double) As Double
    If rng <= 0 Then
        NiceStep = 1
        Exit Function
    End If
    Dim rough As Double, ex As Double, f As Double, nf As Double
    rough = rng / 6#
    ex = Int(Log(rough) / Log(10#))
    f = rough / (10# ^ ex)
    If f < 1.5 Then
        nf = 1
    ElseIf f < 3 Then
        nf = 2
    ElseIf f < 7 Then
        nf = 5
    Else
        nf = 10
    End If
    NiceStep = nf * (10# ^ ex)
End Function
