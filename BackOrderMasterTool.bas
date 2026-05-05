Attribute VB_Name = "BackOrderMasterTool"
Option Explicit

' ============================================================
' Back Order Master Tool - Full Stable Version
' ============================================================

Private Const SRC_SHEET As String = "Raw Data"
Private Const GRP_SHEET As String = "CustomerGroups"
Private Const DSH_SHEET As String = "BackOrder_Dashboard"
Private Const GUIDE_SHEET As String = "操作說明"

Private Const C_CASENO   As Long = 1
Private Const C_CASEPROD As Long = 2
Private Const C_COUNTRY  As Long = 3
Private Const C_STATUS   As Long = 4
Private Const C_CUST     As Long = 5
Private Const C_WARRANTY As Long = 9
Private Const C_PARTNO   As Long = 15
Private Const C_PARTDESC As Long = 16
Private Const C_QTY      As Long = 17
Private Const C_PENDING  As Long = 20

Private Const MAX_COL As Long = 20
Private Const NM_GROUPS As String = "GroupList"

' ============================================================
' 初始化（第一次執行）
' ============================================================
Public Sub InitializeMasterWorkbook()
    Dim wsRaw As Worksheet
    Dim wsGrp As Worksheet
    Dim wsDash As Worksheet
    Dim wsGuide As Worksheet

    Application.ScreenUpdating = False

    Set wsRaw = GetOrCreateSheet(SRC_SHEET, ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    Set wsGrp = GetOrCreateSheet(GRP_SHEET, wsRaw)
    Set wsDash = GetOrCreateSheet(DSH_SHEET, wsGrp)
    Set wsGuide = GetOrCreateSheet(GUIDE_SHEET, wsDash)

    BuildDashboardLayout
    BuildUserGuideSheet

    Application.ScreenUpdating = True
    MsgBox "初始化完成。", vbInformation
End Sub

' ============================================================
' 一鍵全流程
' ============================================================
Public Sub Run_All_Update()
    Dim calcMode As XlCalculation

    On Error GoTo ErrHandler

    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ConsolidateRawDataIntoThisWorkbook
    BuildCustomerGroupsFast
    BuildDashboardLayout
    RefreshDashboard

    ThisWorkbook.Worksheets(DSH_SHEET).Activate
    MsgBox "更新完成。", vbInformation

SafeExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "執行時發生錯誤：" & vbCrLf & _
           "錯誤號碼: " & Err.Number & vbCrLf & _
           "錯誤訊息: " & Err.Description, vbExclamation
    Resume SafeExit
End Sub

Public Sub CreateBackOrderDashboard()
    BuildCustomerGroupsFast
    BuildDashboardLayout
    RefreshDashboard
End Sub

' ============================================================
' 彙整 Raw Data（固定輸出 A:T）
' + 自動辨識多種欄位名稱
' ============================================================
Public Sub ConsolidateRawDataIntoThisWorkbook()
    Dim fd As FileDialog
    Dim wsRaw As Worksheet
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim srcHeaderMap As Object
    Dim srcData As Variant
    Dim outArr() As Variant
    Dim srcLastRow As Long
    Dim rowCount As Long
    Dim outRow As Long
    Dim i As Long, r As Long, c As Long
    Dim srcCol As Long

    Set wsRaw = GetOrCreateSheet(SRC_SHEET, ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = True
        .Title = "請選擇要彙整的 Excel 檔案（可多選）"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xlsm;*.xls;*.xlsb", 1
        If .Show <> -1 Then Exit Sub
        If .SelectedItems.Count = 0 Then Exit Sub
    End With

    Application.ScreenUpdating = False

    wsRaw.Cells.Clear
    WriteFixedRawHeaders wsRaw
    outRow = 2

    For i = 1 To fd.SelectedItems.Count
        Set wbSrc = Nothing
        Set wsSrc = Nothing

        On Error Resume Next
        Set wbSrc = Workbooks.Open(fd.SelectedItems(i), ReadOnly:=True, UpdateLinks:=False)
        On Error GoTo 0
        If wbSrc Is Nothing Then GoTo NextFile

        On Error Resume Next
        Set wsSrc = wbSrc.Worksheets(SRC_SHEET)
        On Error GoTo 0
        If wsSrc Is Nothing Then
            wbSrc.Close SaveChanges:=False
            GoTo NextFile
        End If

        srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
        If srcLastRow < 2 Then
            wbSrc.Close SaveChanges:=False
            GoTo NextFile
        End If

        Set srcHeaderMap = BuildSourceHeaderMap(wsSrc)
        srcData = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(srcLastRow, MAX_COL)).Value2
        rowCount = UBound(srcData, 1)
        ReDim outArr(1 To rowCount, 1 To MAX_COL)

        For c = 1 To MAX_COL
            srcCol = GetMappedSourceColumn(srcHeaderMap, c)

            If srcCol > 0 And srcCol <= MAX_COL Then
                For r = 1 To rowCount
                    outArr(r, c) = CleanCellValueByColumn(srcData(r, srcCol), c)
                Next r
            Else
                For r = 1 To rowCount
                    outArr(r, c) = CleanCellValueByColumn(srcData(r, c), c)
                Next r
            End If
        Next c

        wsRaw.Cells(outRow, 1).Resize(rowCount, MAX_COL).Value = outArr
        outRow = outRow + rowCount

        wbSrc.Close SaveChanges:=False
NextFile:
    Next i

    NormalizeBackOrderStatus wsRaw

    With wsRaw
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(31, 73, 125)
        .Rows(1).Font.Color = RGB(255, 255, 255)
        .Columns("A:T").AutoFit
    End With

    Application.ScreenUpdating = True
End Sub

' ============================================================
' 建立 CustomerGroups
' A:B = 可分群資料
' F = 無法分群 customer
' G = Dropdown Source（排除 F）
' ============================================================
Private Sub BuildCustomerGroupsFast()
    Dim wsSrc As Worksheet
    Dim wsGrp As Worksheet
    Dim lastRow As Long
    Dim arrData As Variant
    Dim custToGrp As Object
    Dim grpToMembers As Object
    Dim ungroupable As Object
    Dim dictDrop As Object
    Dim custWithBO As Object
    Dim i As Long
    Dim rr As Long
    Dim m As Long
    Dim lastMapRow As Long
    Dim totalMembers As Long
    Dim outFRow As Long
    Dim outGRow As Long
    Dim custName As String
    Dim grpName As String
    Dim memberCust As String
    Dim gKey As Variant
    Dim k As Variant
    Dim outMap() As Variant
    Dim col As Collection

    Set wsSrc = GetOrCreateSheet(SRC_SHEET, ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    Set wsGrp = GetOrCreateSheet(GRP_SHEET, wsSrc)

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, C_CUST).End(xlUp).Row

    wsGrp.Cells.Clear
    With wsGrp
        .Range("A1").Value = "Customer Group"
        .Range("B1").Value = "Customer Name (Full)"
        .Range("D1").Value = "Group (Aligned)"
        .Range("F1").Value = "Ungroupable Customer"
        .Range("G1").Value = "Dropdown Source"
        StyleHeader .Range("A1:B1")
        StyleHeader .Range("D1")
        StyleHeader .Range("F1")
        StyleHeader .Range("G1")
    End With

    If lastRow < 2 Then
        wsGrp.Range("F2").Value = "(No Ungroupable Customer)"
        wsGrp.Range("G2").Value = "(No Group)"
        ResetGroupListName wsGrp, 2
        wsGrp.Columns("G").Hidden = True
        Exit Sub
    End If

    arrData = wsSrc.Range(wsSrc.Cells(2, C_STATUS), wsSrc.Cells(lastRow, C_CUST)).Value2

    Set custToGrp = CreateObject("Scripting.Dictionary")
    Set grpToMembers = CreateObject("Scripting.Dictionary")
    Set ungroupable = CreateObject("Scripting.Dictionary")
    Set dictDrop = CreateObject("Scripting.Dictionary")
    Set custWithBO = CreateObject("Scripting.Dictionary")

    custToGrp.CompareMode = vbTextCompare
    grpToMembers.CompareMode = vbTextCompare
    ungroupable.CompareMode = vbTextCompare
    dictDrop.CompareMode = vbTextCompare
    custWithBO.CompareMode = vbTextCompare

    For i = 1 To UBound(arrData, 1)
        If IsBackOrderStatus(SafeToString(arrData(i, 1))) Then
            custName = TrimTextKeepCase(SafeToString(arrData(i, 2)))
            If Len(custName) > 0 Then
                If Not custWithBO.Exists(custName) Then custWithBO.Add custName, 1
            End If
        End If
    Next i

    For i = 1 To UBound(arrData, 1)
        custName = TrimTextKeepCase(SafeToString(arrData(i, 2)))

        If Len(custName) > 0 Then
            grpName = DeriveGroupName(custName)

            If Len(grpName) = 0 Then
                If Not ungroupable.Exists(custName) Then ungroupable.Add custName, 1
            Else
                If Not custToGrp.Exists(custName) Then
                    custToGrp.Add custName, grpName

                    If grpToMembers.Exists(grpName) Then
                        grpToMembers(grpName).Add custName
                    Else
                        Set col = New Collection
                        col.Add custName
                        grpToMembers.Add grpName, col
                    End If
                End If
            End If
        End If
    Next i

    totalMembers = custToGrp.Count
    If totalMembers > 0 Then
        ReDim outMap(1 To totalMembers, 1 To 2)
        rr = 1

        For Each gKey In grpToMembers.Keys
            For m = 1 To grpToMembers(gKey).Count
                outMap(rr, 1) = CStr(gKey)
                outMap(rr, 2) = CStr(grpToMembers(gKey)(m))
                rr = rr + 1
            Next m
        Next gKey

        wsGrp.Range("A2").Resize(totalMembers, 2).Value = outMap
    End If

    lastMapRow = wsGrp.Cells(wsGrp.Rows.Count, 1).End(xlUp).Row
    If lastMapRow >= 2 Then
        wsGrp.Range("D2:D" & lastMapRow).Value = wsGrp.Range("A2:A" & lastMapRow).Value
    End If

    outFRow = 2
    If ungroupable.Count > 0 Then
        For Each k In ungroupable.Keys
            wsGrp.Cells(outFRow, 6).Value = CStr(k)
            outFRow = outFRow + 1
        Next k
    Else
        wsGrp.Range("F2").Value = "(No Ungroupable Customer)"
    End If

    If lastMapRow >= 2 Then
        For rr = 2 To lastMapRow
            grpName = Trim$(CStr(wsGrp.Cells(rr, 1).Value))
            memberCust = TrimTextKeepCase(SafeToString(wsGrp.Cells(rr, 2).Value))
            If IsValidDropdownGroup(grpName) Then
                If custWithBO.Exists(memberCust) Then
                    If Not dictDrop.Exists(grpName) Then dictDrop.Add grpName, 1
                End If
            End If
        Next rr
    End If

    outGRow = 2
    If dictDrop.Count > 0 Then
        For Each k In dictDrop.Keys
            wsGrp.Cells(outGRow, 7).Value = CStr(k)
            outGRow = outGRow + 1
        Next k
    Else
        wsGrp.Range("G2").Value = "(No Group)"
        outGRow = 3
    End If

    wsGrp.Range("G2:G" & outGRow - 1).Sort Key1:=wsGrp.Range("G2"), Order1:=xlAscending, Header:=xlNo
    ResetGroupListName wsGrp, outGRow - 1

    wsGrp.Columns("G").Hidden = True
    wsGrp.Columns("A:G").AutoFit
End Sub

Private Sub ResetGroupListName(ByVal wsGrp As Worksheet, ByVal lastRowG As Long)
    On Error Resume Next
    ThisWorkbook.Names(NM_GROUPS).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:=NM_GROUPS, _
        RefersTo:="=" & wsGrp.Range("G2:G" & lastRowG).Address(True, True, xlA1, True)
End Sub

Private Function IsValidDropdownGroup(ByVal s As String) As Boolean
    Dim t As String
    Dim key As String

    t = Trim$(CStr(s))
    If Len(t) = 0 Then Exit Function

    key = NormalizeKey(t)
    If key = "NOGROUP" Then Exit Function
    If key = "BACKORDERDASHBOARD" Then Exit Function
    If IsNumeric(t) Then Exit Function
    If Left$(t, 1) = "#" Then Exit Function
    If InStr(1, t, "/", vbTextCompare) > 0 Then Exit Function
    If InStr(1, t, "#", vbTextCompare) > 0 Then Exit Function

    IsValidDropdownGroup = True
End Function

' ============================================================
' Dashboard
' ============================================================
Private Sub BuildDashboardLayout()
    Dim wsSrc As Worksheet
    Dim wsGrp As Worksheet
    Dim wsDash As Worksheet
    Dim hdrs As Variant
    Dim c As Long

    Set wsSrc = GetOrCreateSheet(SRC_SHEET, ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    Set wsGrp = GetOrCreateSheet(GRP_SHEET, wsSrc)
    Set wsDash = GetOrCreateSheet(DSH_SHEET, wsGrp)

    wsDash.Cells.Clear

    With wsDash
        .Range("A1:I1").Merge
        .Range("A1").Value = "Back Order Parts Dashboard"
        .Range("A1").Font.Size = 15
        .Range("A1").Font.Bold = True

        .Range("A3").Value = "Select Customer Group:"
        .Range("A3").Font.Bold = True

        With .Range("B3").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & NM_GROUPS
            .IgnoreBlank = True
            .InCellDropdown = True
        End With

        .Range("B3").Interior.Color = RGB(255, 242, 204)
        .Range("B3").Font.Bold = True

        If Len(Trim$(CStr(wsGrp.Range("G2").Value))) > 0 Then
            .Range("B3").Value = wsGrp.Range("G2").Value
        End If

        .Range("G3").Value = "Total Back Order rows:"
        .Range("G3").Font.Bold = True
        .Range("H3").Value = 0

        hdrs = Array("Case No", "Case Product No", "country", "Customer", _
                     "Part Warranty Coverage", "Parts Number", "Parts Description", "Parts Qty", "Pending Days")

        For c = 0 To 8
            .Cells(5, c + 1).Value = hdrs(c)
        Next c

        StyleHeader .Range("A5:I5")
        .Columns("A:I").AutoFit
    End With
End Sub

Public Sub RefreshDashboard()
    Dim wsDash As Worksheet
    Dim wsGrp As Worksheet
    Dim wsSrc As Worksheet
    Dim selectedGroup As String
    Dim custInGroup As Object
    Dim lastGR As Long
    Dim lastSR As Long
    Dim arrGrp As Variant
    Dim arrSrc As Variant
    Dim outArr() As Variant
    Dim maxRows As Long
    Dim g As Long
    Dim i As Long
    Dim outN As Long
    Dim statusVal As String
    Dim custVal As String
    Dim cp As String

    Set wsDash = GetWorksheetSafe(DSH_SHEET)
    Set wsGrp = GetWorksheetSafe(GRP_SHEET)
    Set wsSrc = GetWorksheetSafe(SRC_SHEET)
    If wsDash Is Nothing Or wsGrp Is Nothing Or wsSrc Is Nothing Then Exit Sub

    selectedGroup = Trim$(CStr(wsDash.Range("B3").Value))
    If Len(selectedGroup) = 0 Then Exit Sub

    Set custInGroup = CreateObject("Scripting.Dictionary")
    custInGroup.CompareMode = vbTextCompare

    lastGR = wsGrp.Cells(wsGrp.Rows.Count, 1).End(xlUp).Row
    If lastGR >= 2 Then
        arrGrp = wsGrp.Range("A2:B" & lastGR).Value2
        For g = 1 To UBound(arrGrp, 1)
            If StrComp(Trim$(CStr(arrGrp(g, 1))), selectedGroup, vbTextCompare) = 0 Then
                custInGroup(TrimTextKeepCase(SafeToString(arrGrp(g, 2)))) = 1
            End If
        Next g
    End If

    wsDash.Range("A6:I1048576").ClearContents
    wsDash.Range("A6:I1048576").Interior.ColorIndex = xlNone

    lastSR = wsSrc.Cells(wsSrc.Rows.Count, C_CUST).End(xlUp).Row
    If lastSR < 2 Then Exit Sub

    arrSrc = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastSR, C_PENDING)).Value2
    maxRows = UBound(arrSrc, 1)
    ReDim outArr(1 To maxRows, 1 To 9)

    For i = 1 To UBound(arrSrc, 1)
        statusVal = SafeToString(arrSrc(i, C_STATUS))
        custVal = TrimTextKeepCase(SafeToString(arrSrc(i, C_CUST)))

        If IsBackOrderStatus(statusVal) Then
            If custInGroup.Exists(custVal) Then
                cp = SafeToString(arrSrc(i, C_CASEPROD))
                outN = outN + 1
                outArr(outN, 1) = arrSrc(i, C_CASENO)
                outArr(outN, 2) = cp
                outArr(outN, 3) = GetSecondThirdChars(cp)
                outArr(outN, 4) = arrSrc(i, C_CUST)
                outArr(outN, 5) = arrSrc(i, C_WARRANTY)
                outArr(outN, 6) = arrSrc(i, C_PARTNO)
                outArr(outN, 7) = arrSrc(i, C_PARTDESC)
                outArr(outN, 8) = arrSrc(i, C_QTY)
                outArr(outN, 9) = arrSrc(i, C_PENDING)
            End If
        End If
    Next i

    If outN = 0 Then
        wsDash.Range("A6").Value = "No Back Order records found for: " & selectedGroup
    Else
        wsDash.Range("A6").Resize(outN, 9).Value = outArr
    End If

    wsDash.Range("H3").Value = outN
    wsDash.Columns("A:I").AutoFit
End Sub

' ============================================================
' 操作說明
' ============================================================
Public Sub BuildUserGuideSheet()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim btn1 As Button
    Dim btn2 As Button
    Dim btn3 As Button

    Set ws = GetOrCreateSheet(GUIDE_SHEET, ThisWorkbook.Worksheets(1))
    ws.Cells.Clear

    On Error Resume Next
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    On Error GoTo 0

    ws.Range("A1").Value = "Back Order 工具使用流程"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True

    ws.Range("A3").Value = "Step 1：按【1. 選取多檔彙整Raw Data】"
    ws.Range("A4").Value = "Step 2：按【2. 產生Dashboard】"
    ws.Range("A5").Value = "Step 3：（可選）按【3. 一鍵全流程】"
    ws.Columns("A").ColumnWidth = 90

    Set btn1 = ws.Buttons.Add(ws.Range("A8").Left, ws.Range("A8").Top, 280, 34)
    btn1.Caption = "1. 選取多檔彙整Raw Data"
    btn1.OnAction = "ConsolidateRawDataIntoThisWorkbook"

    Set btn2 = ws.Buttons.Add(ws.Range("A10").Left, ws.Range("A10").Top, 280, 34)
    btn2.Caption = "2. 產生Dashboard"
    btn2.OnAction = "CreateBackOrderDashboard"

    Set btn3 = ws.Buttons.Add(ws.Range("A12").Left, ws.Range("A12").Top, 280, 34)
    btn3.Caption = "3. 一鍵全流程 Run_All_Update"
    btn3.OnAction = "Run_All_Update"
End Sub

' ============================================================
' Header Mapping
' ============================================================
Private Function BuildSourceHeaderMap(ByVal wsSrc As Worksheet) As Object
    Dim dict As Object
    Dim c As Long
    Dim h As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    For c = 1 To MAX_COL
        h = NormalizeHeader(wsSrc.Cells(1, c).Value)
        If Len(h) > 0 Then
            If Not dict.Exists(h) Then dict.Add h, c
        End If
    Next c

    Set BuildSourceHeaderMap = dict
End Function

Private Function GetMappedSourceColumn(ByVal srcHeaderMap As Object, ByVal outCol As Long) As Long
    Select Case outCol
        Case C_CASENO
            GetMappedSourceColumn = FindSourceColumnByAliases(srcHeaderMap, Array("CASE NO", "CASENO", "CASE NUMBER", "CASE ID", "SERVICE REQUEST"))
        Case C_CASEPROD
            GetMappedSourceColumn = FindSourceColumnByAliases(srcHeaderMap, Array("CASE PRODUCT NO", "CASE PRODUCT NUMBER", "CASE PROD NO", "PRODUCT NO", "PRODUCT NUMBER", "PRODUCT"))
        Case C_COUNTRY
            GetMappedSourceColumn = FindSourceColumnByAliases(srcHeaderMap, Array("COUNTRY", "REGION", "NATION"))
        Case C_STATUS
            GetMappedSourceColumn = FindSourceColumnByAliases(srcHeaderMap, Array("STATUS", "CASE STATUS", "ORDER STATUS", "BO STATUS", "BACK ORDER STATUS"))
        Case C_CUST
            GetMappedSourceColumn = FindSourceColumnByAliases(srcHeaderMap, Array("CUSTOMER", "CUSTOMER NAME", "END CUSTOMER", "CUSTOMER FULL NAME", "SOLD TO NAME", "ACCOUNT NAME", "CLIENT NAME"))
        Case C_WARRANTY
            GetMappedSourceColumn = FindSourceColumnByAliases(srcHeaderMap, Array("PART WARRANTY COVERAGE", "WARRANTY", "WARRANTY COVERAGE", "PARTS WARRANTY", "PART WARRANTY"))
        Case C_PARTNO
            GetMappedSourceColumn = FindSourceColumnByAliases(srcHeaderMap, Array("PARTS NUMBER", "PART NUMBER", "PART NO", "PN", "MATERIAL NUMBER"))
        Case C_PARTDESC
            GetMappedSourceColumn = FindSourceColumnByAliases(srcHeaderMap, Array("PARTS DESCRIPTION", "PART DESCRIPTION", "DESCRIPTION", "PART DESC", "MATERIAL DESCRIPTION"))
        Case C_QTY
            GetMappedSourceColumn = FindSourceColumnByAliases(srcHeaderMap, Array("PARTS QTY", "QTY", "QUANTITY", "PART QTY", "ORDER QTY"))
        Case C_PENDING
            GetMappedSourceColumn = FindSourceColumnByAliases(srcHeaderMap, Array("PENDING DAYS", "PENDING DAY", "AGING DAYS", "AGE", "DAYS PENDING", "OPEN DAYS"))
        Case Else
            GetMappedSourceColumn = 0
    End Select
End Function

Private Function FindSourceColumnByAliases(ByVal srcHeaderMap As Object, ByVal aliases As Variant) As Long
    Dim i As Long
    Dim key As String

    For i = LBound(aliases) To UBound(aliases)
        key = NormalizeHeader(CStr(aliases(i)))
        If srcHeaderMap.Exists(key) Then
            FindSourceColumnByAliases = CLng(srcHeaderMap(key))
            Exit Function
        End If
    Next i
End Function

' ============================================================
' Grouping - 品牌別名合併
' ============================================================
Private Function DeriveGroupName(ByVal customerName As String) As String
    Dim brandName As String
    Dim token As String

    brandName = MatchBrandAlias(customerName)
    If Len(brandName) > 0 Then
        DeriveGroupName = brandName
        Exit Function
    End If

    token = BrandToken(customerName)
    If Not IsGroupableToken(token) Then
        DeriveGroupName = ""
    Else
        DeriveGroupName = token
    End If
End Function

Private Function MatchBrandAlias(ByVal customerName As String) As String
    Dim norm As String

    norm = NormalizeKey(customerName)
    If Len(norm) = 0 Then Exit Function

    If HasAlias(norm, "PLANETFITNESS") Or HasAlias(norm, "PLANETFIT") Or HasAlias(norm, "PLANET") Then
        MatchBrandAlias = "PLANET FITNESS"
        Exit Function
    End If

    If HasAlias(norm, "SNAPFITNESS") Or HasAlias(norm, "SNAPFIT") Or HasAlias(norm, "SNAP") Then
        MatchBrandAlias = "SNAP FITNESS"
        Exit Function
    End If

    If HasAlias(norm, "BASICFIT") Or HasAlias(norm, "BASIC") Then
        MatchBrandAlias = "BASIC-FIT"
        Exit Function
    End If

    If HasAlias(norm, "ANYTIMEFITNESS") Or HasAlias(norm, "ANYTIMEFIT") Or HasAlias(norm, "ANYTIME") Then
        MatchBrandAlias = "ANYTIME FITNESS"
        Exit Function
    End If

    If HasAlias(norm, "WORLDGYM") Then
        MatchBrandAlias = "WORLD GYM"
        Exit Function
    End If

    If HasAlias(norm, "TRUEFITNESS") Or HasAlias(norm, "TRUEFIT") Then
        MatchBrandAlias = "TRUE FITNESS"
        Exit Function
    End If
End Function

Private Function HasAlias(ByVal normalizedCustomer As String, ByVal aliasKey As String) As Boolean
    Dim nCustomer As String
    Dim nAlias As String

    nCustomer = NormalizeKey(normalizedCustomer)
    nAlias = NormalizeKey(aliasKey)

    If Len(nCustomer) = 0 Then Exit Function
    If Len(nAlias) = 0 Then Exit Function

    If Len(nCustomer) > 255 Then nCustomer = Left$(nCustomer, 255)
    If Len(nAlias) > 80 Then nAlias = Left$(nAlias, 80)

    HasAlias = (nCustomer Like "*" & nAlias & "*")
End Function

Private Function BrandToken(ByVal nm As String) As String
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim token As String
    Dim codeVal As Long

    s = TrimTextKeepCase(nm)
    If Len(s) = 0 Then
        BrandToken = ""
        Exit Function
    End If

    If Len(s) > 255 Then s = Left$(s, 255)

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        codeVal = AscW(ch)

        If (codeVal >= 48 And codeVal <= 57) Or _
           (codeVal >= 65 And codeVal <= 90) Or _
           (codeVal >= 97 And codeVal <= 122) Then
            token = token & ch
        ElseIf Len(token) > 0 Then
            Exit For
        End If
    Next i

    BrandToken = UCase$(token)
End Function

Private Function IsGroupableToken(ByVal token As String) As Boolean
    Dim t As String

    t = Trim$(token)
    If Len(t) < 2 Then Exit Function
    If Len(t) > 40 Then Exit Function
    If IsNumeric(t) Then Exit Function
    If Left$(t, 1) = "#" Then Exit Function
    If InStr(1, t, "/", vbTextCompare) > 0 Then Exit Function
    If InStr(1, t, "#", vbTextCompare) > 0 Then Exit Function

    Select Case UCase$(t)
        Case "CO", "LTD", "INC", "LLC", "THE", "AND", "FOR"
            Exit Function
    End Select

    IsGroupableToken = True
End Function

' ============================================================
' Helpers
' ============================================================
Private Sub WriteFixedRawHeaders(ByVal ws As Worksheet)
    Dim hdrs As Variant
    Dim i As Long

    hdrs = Array( _
        "CASE NO", _
        "CASE PRODUCT NO", _
        "COUNTRY", _
        "STATUS", _
        "CUSTOMER", _
        "COL6", _
        "COL7", _
        "COL8", _
        "PART WARRANTY COVERAGE", _
        "COL10", _
        "COL11", _
        "COL12", _
        "COL13", _
        "COL14", _
        "PARTS NUMBER", _
        "PARTS DESCRIPTION", _
        "PARTS QTY", _
        "COL18", _
        "COL19", _
        "PENDING DAYS")

    For i = 0 To UBound(hdrs)
        ws.Cells(1, i + 1).Value = hdrs(i)
    Next i
End Sub

Private Sub NormalizeBackOrderStatus(ByVal wsRaw As Worksheet)
    Dim lastRawRow As Long
    Dim iRow As Long
    Dim v As String

    lastRawRow = wsRaw.Cells(wsRaw.Rows.Count, C_STATUS).End(xlUp).Row
    For iRow = 2 To lastRawRow
        v = Trim$(SafeToString(wsRaw.Cells(iRow, C_STATUS).Value))

        If v = "零件待料中" Then
            wsRaw.Cells(iRow, C_STATUS).Value = "BACK ORDER"
        ElseIf IsBackOrderStatus(v) Then
            wsRaw.Cells(iRow, C_STATUS).Value = "BACK ORDER"
        End If
    Next iRow
End Sub

Private Sub StyleHeader(ByVal rng As Range)
    With rng
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 73, 125)
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Function GetWorksheetSafe(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetSafe = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function GetOrCreateSheet(ByVal sheetName As String, ByVal insertAfter As Worksheet) As Worksheet
    Dim ws As Worksheet

    Set ws = GetWorksheetSafe(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=insertAfter)
        ws.Name = sheetName
    End If

    Set GetOrCreateSheet = ws
End Function

Private Function IsBackOrderStatus(ByVal s As String) As Boolean
    Dim v As String

    v = UCase$(Trim$(s))
    v = Replace(v, "-", "")
    v = Replace(v, " ", "")

    IsBackOrderStatus = (v = "BACKORDER" Or v = "BO")
End Function

Private Function GetSecondThirdChars(ByVal s As String) As String
    s = Trim$(CStr(s))

    If Len(s) >= 3 Then
        GetSecondThirdChars = Mid$(s, 2, 2)
    ElseIf Len(s) = 2 Then
        GetSecondThirdChars = Right$(s, 1)
    Else
        GetSecondThirdChars = ""
    End If
End Function

Private Function NormalizeHeader(ByVal v As Variant) As String
    NormalizeHeader = NormalizeKey(SafeToString(v))
End Function

Private Function CleanCellValueByColumn(ByVal v As Variant, ByVal colIndex As Long) As Variant
    If IsError(v) Then
        CleanCellValueByColumn = ""
        Exit Function
    End If

    Select Case colIndex
        Case C_CASENO, C_CASEPROD, C_COUNTRY, C_STATUS, C_CUST, C_WARRANTY, C_PARTNO, C_PARTDESC
            CleanCellValueByColumn = TrimTextKeepCase(SafeToString(v))
        Case Else
            CleanCellValueByColumn = PreserveValue(v)
    End Select
End Function

Private Function PreserveValue(ByVal v As Variant) As Variant
    If IsError(v) Then
        PreserveValue = ""
    ElseIf IsEmpty(v) Then
        PreserveValue = ""
    ElseIf IsNull(v) Then
        PreserveValue = ""
    ElseIf VarType(v) = vbString Then
        PreserveValue = TrimTextKeepCase(CStr(v))
    Else
        PreserveValue = v
    End If
End Function

Private Function SafeToString(ByVal v As Variant) As String
    On Error GoTo EH

    If IsError(v) Then
        SafeToString = ""
    ElseIf IsEmpty(v) Then
        SafeToString = ""
    ElseIf IsNull(v) Then
        SafeToString = ""
    Else
        SafeToString = CStr(v)
    End If
    Exit Function

EH:
    SafeToString = ""
End Function

Private Function TrimTextKeepCase(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim resultText As String
    Dim prevSpace As Boolean

    If Len(s) = 0 Then
        TrimTextKeepCase = ""
        Exit Function
    End If

    If Len(s) > 255 Then s = Left$(s, 255)

    resultText = ""
    prevSpace = True

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        Select Case AscW(ch)
            Case 9, 10, 13, 32, 160
                If Not prevSpace Then
                    resultText = resultText & " "
                    prevSpace = True
                End If
            Case Else
                resultText = resultText & ch
                prevSpace = False
        End Select
    Next i

    If Len(resultText) > 0 Then
        If Right$(resultText, 1) = " " Then
            resultText = Left$(resultText, Len(resultText) - 1)
        End If
    End If

    TrimTextKeepCase = resultText
End Function

Private Function NormalizeKey(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim codeVal As Long
    Dim resultText As String

    s = UCase$(TrimTextKeepCase(s))
    If Len(s) = 0 Then
        NormalizeKey = ""
        Exit Function
    End If

    If Len(s) > 255 Then s = Left$(s, 255)

    resultText = ""

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        codeVal = AscW(ch)

        If (codeVal >= 48 And codeVal <= 57) Or (codeVal >= 65 And codeVal <= 90) Then
            resultText = resultText & ch
        End If
    Next i

    NormalizeKey = resultText
End Function
