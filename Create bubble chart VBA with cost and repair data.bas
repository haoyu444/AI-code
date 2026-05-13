Sub CreateWarrantyBubbleChart()
    Dim wsData As Worksheet, wsSummary As Worksheet
    Dim lastRow As Long, i As Long, r As Long
    Dim caseNo As String, productName As String, partNo As String
    Dim currentCost As Double

    Application.ScreenUpdating = False

    Set wsData = ActiveSheet
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    ' === Step 1: 記錄 (產品→案件集合) 與 (產品→總成本) ===
    Dim dictProdCases As Object, dictProdTotal As Object
    Set dictProdCases = CreateObject("Scripting.Dictionary")
    Set dictProdTotal = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        caseNo = Trim(CStr(wsData.Cells(i, "A").Value))
        productName = Trim(CStr(wsData.Cells(i, "M").Value))
        partNo = Trim(CStr(wsData.Cells(i, "AL").Value))
        currentCost = 0
        If IsNumeric(wsData.Cells(i, "AP").Value) Then
            currentCost = CDbl(wsData.Cells(i, "AP").Value)
        End If

        If caseNo <> "" And LCase(caseNo) <> "case no" Then
            If productName <> "" And LCase(productName) <> "product name" Then
                If Not dictProdCases.Exists(productName) Then
                    dictProdCases.Add productName, CreateObject("Scripting.Dictionary")
                    dictProdTotal.Add productName, 0#
                End If
                If Not dictProdCases(productName).Exists(caseNo) Then
                    dictProdCases(productName).Add caseNo, True
                End If
                If Len(partNo) > 0 Then
                    If Left(partNo, 1) = "0" Or Left(partNo, 1) = "1" Then
                        dictProdTotal(productName) = dictProdTotal(productName) + currentCost
                    End If
                End If
            End If
        End If
    Next i

    ' === Step 2: 放入陣列並依 Total Cost 由大到小排序 ===
    Dim n As Long, k As Long
    n = dictProdCases.Count
    If n = 0 Then
        Application.ScreenUpdating = True
        MsgBox "沒有可用資料。", vbExclamation
        Exit Sub
    End If

    Dim arr() As Variant
    ReDim arr(1 To n, 1 To 4)
    k = 0
    Dim key As Variant
    For Each key In dictProdCases.Keys
        k = k + 1
        arr(k, 1) = key
        arr(k, 3) = dictProdCases(key).Count
        arr(k, 4) = dictProdTotal(key)
        If arr(k, 3) > 0 Then
            arr(k, 2) = arr(k, 4) / arr(k, 3)
        Else
            arr(k, 2) = 0
        End If
    Next key

    Dim a As Long, b As Long, c As Long, tmp As Variant
    For a = 1 To n - 1
        For b = 1 To n - a
            If arr(b, 4) < arr(b + 1, 4) Then
                For c = 1 To 4
                    tmp = arr(b, c)
                    arr(b, c) = arr(b + 1, c)
                    arr(b + 1, c) = tmp
                Next c
            End If
        Next b
    Next a

    Dim topN As Long
    If n > 20 Then
        topN = 20
    Else
        topN = n
    End If

    ' === Step 3: 寫入彙總工作表 ===
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("BubbleData").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsSummary = ThisWorkbook.Sheets.Add(After:=wsData)
    wsSummary.Name = "BubbleData"

    With wsSummary.Range("A1:D1")
        .Value = Array("Product Name", "Cost per Case (X)", "Repair QTY (Y)", "Total Cost (Size)")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
    End With

    For r = 1 To topN
        wsSummary.Cells(r + 1, 1).Value = arr(r, 1)
        wsSummary.Cells(r + 1, 2).Value = arr(r, 2)
        wsSummary.Cells(r + 1, 3).Value = arr(r, 3)
        wsSummary.Cells(r + 1, 4).Value = arr(r, 4)
    Next r
    wsSummary.Range("B2:B" & (topN + 1)).NumberFormat = "#,##0.00"
    wsSummary.Range("D2:D" & (topN + 1)).NumberFormat = "#,##0.00"
    wsSummary.Columns("A:D").AutoFit

    ' === Step 4: 計算 X / Y 軸最適刻度 ===
    Dim xMax As Double, yMax As Double, vCell As Range
    xMax = 0: yMax = 0
    For Each vCell In wsSummary.Range("B2:B" & (topN + 1))
        If IsNumeric(vCell.Value) Then
            If vCell.Value > xMax Then xMax = vCell.Value
        End If
    Next vCell
    For Each vCell In wsSummary.Range("C2:C" & (topN + 1))
        If IsNumeric(vCell.Value) Then
            If vCell.Value > yMax Then yMax = vCell.Value
        End If
    Next vCell
    If xMax <= 0 Then xMax = 1
    If yMax <= 0 Then yMax = 1

    Dim xLow As Double, xHigh As Double, xUnit As Double
    Dim yLow As Double, yHigh As Double, yUnit As Double
    xLow = 0
    xHigh = NiceCeiling(xMax * 1.15)
    xUnit = NiceStep(xHigh - xLow)
    If xUnit <= 0 Then xUnit = xHigh / 5
    yLow = 0
    yHigh = NiceCeiling(yMax * 1.15)
    If yHigh < 5 Then yHigh = 5
    yUnit = NiceStep(yHigh - yLow)
    If yUnit < 1 Then yUnit = 1

    ' === Step 5: 繪製泡泡圖 ===
    Dim cht As ChartObject
    Set cht = wsSummary.ChartObjects.Add(Left:=320, Top:=10, Width:=900, Height:=560)

    Dim palette(1 To 12) As Long
    palette(1) = RGB(91, 155, 213)
    palette(2) = RGB(237, 125, 49)
    palette(3) = RGB(165, 165, 165)
    palette(4) = RGB(255, 192, 0)
    palette(5) = RGB(68, 114, 196)
    palette(6) = RGB(112, 173, 71)
    palette(7) = RGB(158, 72, 14)
    palette(8) = RGB(99, 99, 99)
    palette(9) = RGB(153, 115, 0)
    palette(10) = RGB(38, 68, 120)
    palette(11) = RGB(67, 104, 43)
    palette(12) = RGB(217, 83, 113)

    With cht.Chart
        .ChartType = xlBubble
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop

        Dim s As Series
        Set s = .SeriesCollection.NewSeries
        s.Name = "Warranty Cost by Model"
        s.XValues = wsSummary.Range("B2:B" & (topN + 1))
        s.Values = wsSummary.Range("C2:C" & (topN + 1))
        s.BubbleSizes = "='" & wsSummary.Name & "'!" & wsSummary.Range("D2:D" & (topN + 1)).Address

        On Error Resume Next
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(250, 250, 252)
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
        On Error GoTo 0

        .HasTitle = True
        With .ChartTitle
            .Text = "Total Warranty Cost by Model (Top " & topN & ")"
            .Font.Size = 16
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Color = RGB(64, 64, 64)
        End With

        ' X 軸
        .Axes(xlCategory).HasTitle = True
        With .Axes(xlCategory).AxisTitle
            .Text = "Cost per Case"
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = RGB(89, 89, 89)
        End With
        On Error Resume Next
        With .Axes(xlCategory)
            .MinimumScale = xLow
            .MaximumScale = xHigh
            .MajorUnit = xUnit
            .TickLabels.NumberFormat = "#,##0"
            .TickLabels.Font.Size = 10
            .TickLabels.Font.Color = RGB(89, 89, 89)
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 230, 230)
            .MajorGridlines.Format.Line.DashStyle = msoLineDash
        End With
        On Error GoTo 0

        ' Y 軸
        .Axes(xlValue).HasTitle = True
        With .Axes(xlValue).AxisTitle
            .Text = "Repair QTY"
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = RGB(89, 89, 89)
            .Orientation = xlUpward
        End With
        On Error Resume Next
        With .Axes(xlValue)
            .MinimumScale = yLow
            .MaximumScale = yHigh
            .MajorUnit = yUnit
            .TickLabels.NumberFormat = "0"
            .TickLabels.Font.Size = 10
            .TickLabels.Font.Color = RGB(89, 89, 89)
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 230, 230)
            .MajorGridlines.Format.Line.DashStyle = msoLineDash
        End With
        On Error GoTo 0

        .HasLegend = False

        ' 泡泡上色 + 半透明 + 細邊框
        Dim p As Long, clr As Long
        For p = 1 To s.Points.Count
            clr = palette(((p - 1) Mod 12) + 1)
            On Error Resume Next
            With s.Points(p).Format.Fill
                .Visible = msoTrue
                .ForeColor.RGB = clr
                .Transparency = 0.35
            End With
            With s.Points(p).Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = clr
                .Weight = 1.25
            End With
            On Error GoTo 0
        Next p

        ' === 資料標籤：Product Name（用 TextFrame2.TextRange.Text 強制寫入）===
        s.HasDataLabels = True
        With s.DataLabels
            .ShowSeriesName = False
            .ShowCategoryName = False
            .ShowValue = False
            .ShowBubbleSize = False
        End With
        DoEvents

        Dim j As Long, lblText As String
        For j = 1 To s.Points.Count
            lblText = CStr(wsSummary.Cells(j + 1, 1).Value) & Chr(10) & _
                      "(" & wsSummary.Cells(j + 1, 3).Value & ")"
            With s.Points(j)
                .HasDataLabel = True
                With .DataLabel
                    .ShowSeriesName = False
                    .ShowCategoryName = False
                    .ShowValue = False
                    .ShowBubbleSize = False
                    .Position = xlLabelPositionCenter
                    ' --- 關鍵：直接寫入 TextFrame2 的文字內容 ---
                    On Error Resume Next
                    .Format.TextFrame2.TextRange.Text = lblText
                    On Error GoTo 0
                    ' --- 字型用 DataLabel 標準屬性（避免 Font.Fill 造成參數無效）---
                    .Font.Size = 8
                    .Font.Bold = True
                    .Font.Color = RGB(40, 40, 40)
                End With
            End With
        Next j
    End With

    wsSummary.Activate
    Application.ScreenUpdating = True
    MsgBox "已建立 Top " & topN & " 泡泡圖於 BubbleData 工作表！", vbInformation
End Sub

' === 取「漂亮」的上界（1 / 2 / 2.5 / 5 × 10^n）===
Private Function NiceCeiling(ByVal v As Double) As Double
    If v <= 0 Then
        NiceCeiling = 1
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

' === 取「漂亮」的主要刻度間距 ===
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
