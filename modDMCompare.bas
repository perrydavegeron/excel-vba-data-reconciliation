Attribute VB_Name = "modDMCompare"
Option Explicit

' DM Compare Tool - Excel 2016 / Microsoft 365 (Windows)
' Entirely late-bound for dictionaries, so no references are necessary.

Private Const RESULTS_HEADER_ROW As Long = 7
Private Const RESULTS_FIRST_ROW As Long = 8
Private Const BLOCK_ROWS As Long = 500

Public Sub DM_CompareFiles()
    frmFileCompare.Show
End Sub

' Office Ribbon XML callback. Also usable when launched via Excel's Macro dialog.
Public Sub DM_RibbonCompare(ByVal control As IRibbonControl)
    DM_CompareFiles
End Sub

Public Function DM_GetWorkbook(ByVal fileName As String, ByRef openedHere As Boolean) As Workbook
    Dim b As Workbook
    openedHere = False
    For Each b In Application.Workbooks
        If StrComp(b.FullName, fileName, vbTextCompare) = 0 Then
            Set DM_GetWorkbook = b
            Exit Function
        End If
    Next b
    Set DM_GetWorkbook = Application.Workbooks.Open( _
        Filename:=fileName, UpdateLinks:=0, ReadOnly:=True, _
        AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    openedHere = True
End Function

Public Sub DM_LoadWorksheetNames(ByVal fileName As String, ByVal cbo As Object)
    Dim wb As Workbook, w As Worksheet, openedHere As Boolean, oldSecurity As Long
    Dim message As String
    On Error GoTo Fail
    oldSecurity = Application.AutomationSecurity
    Application.AutomationSecurity = 3
    Set wb = DM_GetWorkbook(fileName, openedHere)
    cbo.Clear
    For Each w In wb.Worksheets
        cbo.AddItem w.Name
    Next w
    If cbo.ListCount > 0 Then cbo.ListIndex = 0
CleanExit:
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.AutomationSecurity = oldSecurity
    Exit Sub
Fail:
    message = Err.Description
    On Error Resume Next
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.AutomationSecurity = oldSecurity
    On Error GoTo 0
    Err.Raise vbObjectError + 800, "DM_LoadWorksheetNames", message
End Sub

Public Function DM_HeaderKey(ByVal name As String) As String
    DM_HeaderKey = LCase$(Trim$(Replace(name, ChrW$(160), " ")))
End Function

Private Sub DM_ReadHeaders(ByVal ws As Worksheet, ByVal headerRow As Long, _
                           ByRef headerMap As Object, ByRef headerOrder As Collection, _
                           ByRef originalNames As Object, ByRef lastCol As Long)
    Dim c As Long, key As String, raw As Variant
    If headerRow < 1 Or headerRow > ws.Rows.Count Then
        Err.Raise vbObjectError + 801, , "Invalid header row for '" & ws.Name & "'."
    End If
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    Set headerMap = CreateObject("Scripting.Dictionary")
    Set originalNames = CreateObject("Scripting.Dictionary")
    Set headerOrder = New Collection
    headerMap.CompareMode = vbBinaryCompare
    originalNames.CompareMode = vbBinaryCompare
    For c = 1 To lastCol
        raw = ws.Cells(headerRow, c).Value2
        If IsError(raw) Then Err.Raise vbObjectError + 802, , _
            "Excel error found in the header row of '" & ws.Name & "'."
        key = DM_HeaderKey(CStr(raw))
        If Len(key) > 0 Then
            If headerMap.Exists(key) Then Err.Raise vbObjectError + 803, , _
                "Duplicate header '" & CStr(raw) & "' in '" & ws.Name & "'."
            headerMap.Add key, c
            originalNames.Add key, Trim$(CStr(raw))
            headerOrder.Add key
        End If
    Next c
    If headerMap.Count = 0 Then Err.Raise vbObjectError + 804, , _
        "No headers were found at row " & headerRow & " in '" & ws.Name & "'."
End Sub

Public Sub DM_LoadCommonIDs(ByVal pathA As String, ByVal sheetA As String, _
                            ByVal headerA As Long, ByVal pathB As String, _
                            ByVal sheetB As String, ByVal headerB As Long, _
                            ByVal cbo As Object)
    Dim wbA As Workbook, wbB As Workbook, wa As Worksheet, wb As Worksheet
    Dim openedA As Boolean, openedB As Boolean, oldSecurity As Long
    Dim mapA As Object, mapB As Object, namesA As Object, namesB As Object
    Dim orderA As Collection, orderB As Collection, key As Variant, lastA As Long, lastB As Long
    Dim errorMsg As String
    On Error GoTo Fail
    oldSecurity = Application.AutomationSecurity
    Application.AutomationSecurity = 3 ' msoAutomationSecurityForceDisable
    Set wbA = DM_GetWorkbook(pathA, openedA)
    Set wbB = DM_GetWorkbook(pathB, openedB)
    Set wa = wbA.Worksheets(sheetA)
    Set wb = wbB.Worksheets(sheetB)
    DM_ReadHeaders wa, headerA, mapA, orderA, namesA, lastA
    DM_ReadHeaders wb, headerB, mapB, orderB, namesB, lastB
    cbo.Clear
    For Each key In orderA
        If mapB.Exists(CStr(key)) Then cbo.AddItem namesA(key)
    Next key
CleanExit:
    On Error Resume Next
    If openedB Then wbB.Close SaveChanges:=False
    If openedA Then wbA.Close SaveChanges:=False
    Application.AutomationSecurity = oldSecurity
    On Error GoTo 0
    Exit Sub
Fail:
    errorMsg = Err.Description
    Resume CleanFail
CleanFail:
    On Error Resume Next
    If openedB Then wbB.Close SaveChanges:=False
    If openedA Then wbA.Close SaveChanges:=False
    Application.AutomationSecurity = oldSecurity
    On Error GoTo 0
    Err.Raise vbObjectError + 805, "DM_LoadCommonIDs", errorMsg
End Sub

Private Function DM_LastUsedRow(ByVal ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False)
    If f Is Nothing Then
        DM_LastUsedRow = 0
    Else
        DM_LastUsedRow = f.Row
    End If
End Function

Private Function DM_GetData(ByVal ws As Worksheet, ByVal firstRow As Long, _
                            ByVal lastCol As Long, ByRef nRows As Long) As Variant
    Dim lastRow As Long, v As Variant, temp() As Variant
    lastRow = DM_LastUsedRow(ws)
    If firstRow > lastRow Then
        nRows = 0
        DM_GetData = Empty
        Exit Function
    End If
    nRows = lastRow - firstRow + 1
    v = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, lastCol)).Value2
    If IsArray(v) Then
        DM_GetData = v
    Else
        ReDim temp(1 To 1, 1 To 1)
        temp(1, 1) = v
        DM_GetData = temp
    End If
End Function

Private Function DM_CellIsPresent(ByVal v As Variant) As Boolean
    If IsError(v) Then
        DM_CellIsPresent = True
    ElseIf IsNull(v) Or IsEmpty(v) Then
        DM_CellIsPresent = False
    Else
        DM_CellIsPresent = (Len(CStr(v)) > 0)
    End If
End Function

Private Function DM_RowHasData(ByRef data As Variant, ByVal rowIx As Long, _
                               ByVal headers As Object) As Boolean
    Dim col As Variant
    For Each col In headers.Items
        If DM_CellIsPresent(data(rowIx, CLng(col))) Then
            DM_RowHasData = True
            Exit Function
        End If
    Next col
End Function

Private Function DM_UniqueKey(ByVal v As Variant) As String
    If IsError(v) Then Err.Raise vbObjectError + 806, , "Unique IDs cannot contain Excel errors."
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    DM_UniqueKey = Trim$(CStr(v))
End Function

Private Sub DM_IndexRows(ByRef data As Variant, ByVal nRows As Long, _
                         ByVal headers As Object, ByVal idCol As Long, _
                         ByVal dataStart As Long, ByVal fileLabel As String, _
                         ByRef indexed As Object, ByRef order As Collection, _
                         ByRef dataCount As Long)
    Dim r As Long, key As String
    Set indexed = CreateObject("Scripting.Dictionary")
    indexed.CompareMode = vbBinaryCompare
    Set order = New Collection
    dataCount = 0
    For r = 1 To nRows
        If DM_RowHasData(data, r, headers) Then
            dataCount = dataCount + 1
            key = DM_UniqueKey(data(r, idCol))
            If Len(key) = 0 Then Err.Raise vbObjectError + 807, , _
                "Blank Unique ID in " & fileLabel & " at sheet row " & (dataStart + r - 1) & "."
            If indexed.Exists(key) Then Err.Raise vbObjectError + 808, , _
                "Duplicate Unique ID '" & key & "' in " & fileLabel & "."
            indexed.Add key, r
            order.Add r
        End If
    Next r
End Sub

Private Function DM_SameValue(ByVal a As Variant, ByVal b As Variant, _
                              ByVal wa As Worksheet, ByVal sourceRowA As Long, ByVal colA As Long, _
                              ByVal wb As Worksheet, ByVal sourceRowB As Long, ByVal colB As Long) As Boolean
    Dim blankA As Boolean, blankB As Boolean
    If IsError(a) Or IsError(b) Then
        If IsError(a) And IsError(b) Then
            DM_SameValue = (StrComp(wa.Cells(sourceRowA, colA).Text, _
                wb.Cells(sourceRowB, colB).Text, vbBinaryCompare) = 0)
        End If
        Exit Function
    End If
    blankA = (IsEmpty(a) Or IsNull(a))
    blankB = (IsEmpty(b) Or IsNull(b))
    If blankA Or blankB Then
        If blankA And blankB Then
            DM_SameValue = True
        ElseIf blankA Then
            DM_SameValue = (Len(CStr(b)) = 0)
        Else
            DM_SameValue = (Len(CStr(a)) = 0)
        End If
        Exit Function
    End If
    ' Value2 returns native numbers/dates as numeric variants.
    If DM_IsNumericVariant(a) And DM_IsNumericVariant(b) Then
        DM_SameValue = (CDbl(a) = CDbl(b))
    Else
        DM_SameValue = (StrComp(CStr(a), CStr(b), vbBinaryCompare) = 0)
    End If
End Function

Private Function DM_IsNumericVariant(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            DM_IsNumericVariant = True
    End Select
End Function

Private Function DM_ColumnLetter(ByVal index As Long) As String
    Dim letter As String
    Do While index > 0
        index = index - 1
        letter = Chr$(65 + (index Mod 26)) & letter
        index = index \ 26
    Loop
    DM_ColumnLetter = letter
End Function

Public Function DM_RunComparison(ByVal pathA As String, ByVal sheetA As String, _
                                 ByVal hdrA As Long, ByVal startA As Long, _
                                 ByVal pathB As String, ByVal sheetB As String, _
                                 ByVal hdrB As Long, ByVal startB As Long, _
                                 ByVal selectedID As String) As Boolean
    Dim beginTime As Double, elapsed As Double
    Dim wbA As Workbook, wbB As Workbook, wbOut As Workbook
    Dim wsA As Worksheet, wsB As Worksheet, wsOut As Worksheet
    Dim openedA As Boolean, openedB As Boolean
    Dim oldScreen As Boolean, oldEvents As Boolean, oldCalc As XlCalculation
    Dim oldSecurity As Long, settingsCaptured As Boolean
    Dim mapA As Object, mapB As Object, namesA As Object, namesB As Object
    Dim orderA As Collection, orderB As Collection, fields As Collection
    Dim keysA As Object, keysB As Object, rowsA As Collection, rowsB As Collection
    Dim rowsOnlyB As Collection
    Dim rawA As Variant, rawB As Variant
    Dim totalRowsA As Long, totalRowsB As Long, physicalRowsA As Long, physicalRowsB As Long
    Dim lastColA As Long, lastColB As Long, uniqueColA As Long, uniqueColB As Long
    Dim idKey As String, field As Variant
    Dim rA As Long, rB As Long, rowIx As Long, aCol As Long, bCol As Long
    Dim i As Long, j As Long, totalResult As Long, extraB As Long
    Dim compareFields As Long, resultCols As Long, statusCol As Long
    Dim totalMismatch As Long, thisSize As Long, blockBegin As Long
    Dim dataBlock() As Variant, matching As Boolean
    Dim recordStatus As String, valueA As Variant, valueB As Variant
    Dim headers() As Variant, reportLastRow As Long
    Dim cf As FormatCondition, fontRed As Long, fontGreen As Long
    Dim errorText As String, success As Boolean

    beginTime = Timer
    On Error GoTo Fail
    If startA <= hdrA Or startB <= hdrB Then Err.Raise vbObjectError + 810, , _
        "Data starting row must be below the header row in both files."
    If hdrA < 1 Or hdrB < 1 Or startA < 1 Or startB < 1 Then _
        Err.Raise vbObjectError + 811, , "Header and data row numbers must be positive."
    If Len(Trim$(selectedID)) = 0 Then Err.Raise vbObjectError + 812, , _
        "Please select a common Unique ID column."

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldCalc = Application.Calculation
    oldSecurity = Application.AutomationSecurity
    settingsCaptured = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = 3
    Application.StatusBar = "DM Compare: opening files..."

    Set wbA = DM_GetWorkbook(pathA, openedA)
    Set wbB = DM_GetWorkbook(pathB, openedB)
    Set wsA = wbA.Worksheets(sheetA)
    Set wsB = wbB.Worksheets(sheetB)
    DM_ReadHeaders wsA, hdrA, mapA, orderA, namesA, lastColA
    DM_ReadHeaders wsB, hdrB, mapB, orderB, namesB, lastColB
    idKey = DM_HeaderKey(selectedID)
    If Not mapA.Exists(idKey) Or Not mapB.Exists(idKey) Then _
        Err.Raise vbObjectError + 813, , "The selected Unique ID header was not found in both files."
    uniqueColA = CLng(mapA(idKey))
    uniqueColB = CLng(mapB(idKey))

    Set fields = New Collection
    For Each field In orderA
        If CStr(field) <> idKey Then
            If mapB.Exists(CStr(field)) Then fields.Add CStr(field)
        End If
    Next field
    If fields.Count = 0 Then Err.Raise vbObjectError + 814, , _
        "No common headers other than the Unique ID were found."
    compareFields = fields.Count
    ' Column 1 = Unique ID; each comparison uses (A, B, Status); last = Record Status.
    resultCols = 2 + (3 * compareFields)
    If resultCols > wsA.Columns.Count Then _
        Err.Raise vbObjectError + 815, , "Too many columns to fit in an Excel worksheet."

    Application.StatusBar = "DM Compare: loading source rows..."
    rawA = DM_GetData(wsA, startA, lastColA, physicalRowsA)
    rawB = DM_GetData(wsB, startB, lastColB, physicalRowsB)
    DM_IndexRows rawA, physicalRowsA, mapA, uniqueColA, startA, _
        "File A", keysA, rowsA, totalRowsA
    DM_IndexRows rawB, physicalRowsB, mapB, uniqueColB, startB, _
        "File B", keysB, rowsB, totalRowsB
    extraB = 0
    Set rowsOnlyB = New Collection
    For i = 1 To rowsB.Count
        rB = CLng(rowsB(i))
        idKey = DM_UniqueKey(rawB(rB, uniqueColB))
        If Not keysA.Exists(idKey) Then
            extraB = extraB + 1
            rowsOnlyB.Add rB
        End If
    Next i
    totalResult = totalRowsA + extraB
    If CDbl(RESULTS_HEADER_ROW) + CDbl(totalResult) > 1048576# Then _
        Err.Raise vbObjectError + 816, , "Results exceed Excel's 1,048,576-row limit."

    Application.StatusBar = "DM Compare: preparing results workbook..."
    Set wbOut = Application.Workbooks.Add(xlWBATWorksheet)
    Set wsOut = wbOut.Worksheets(1)
    wsOut.Name = "Comparison Results"
    reportLastRow = Application.Max(RESULTS_HEADER_ROW + 1, RESULTS_HEADER_ROW + totalResult)
    With wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(reportLastRow, resultCols)).Font
        .Name = "Aptos"
        .Size = 10
        .Color = RGB(35, 43, 55)
    End With
    wsOut.Range("A1").Value2 = "FILE COMPARISON RESULTS"
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.Size = 15
    wsOut.Range("A2").Value2 = "Total Data Rows - File A"
    wsOut.Range("A3").Value2 = "Total Data Rows - File B"
    wsOut.Range("A4").Value2 = "Total Mismatches"
    wsOut.Range("A5").Value2 = "Total Runtime"
    wsOut.Range("B2").Value2 = totalRowsA
    wsOut.Range("B3").Value2 = totalRowsB
    wsOut.Range("A2:A5").Font.Bold = True
    wsOut.Range("B2:B5").Font.Bold = True
    wsOut.Range("A2:A5").ColumnWidth = 27
    wsOut.Range("B2:B5").ColumnWidth = 20

    ReDim headers(1 To 1, 1 To resultCols)
    headers(1, 1) = namesA(DM_HeaderKey(selectedID))
    For j = 1 To compareFields
        field = fields(j)
        statusCol = 1 + j * 3
        headers(1, statusCol - 2) = namesA(field) & " (A)"
        headers(1, statusCol - 1) = namesB(field) & " (B)"
        headers(1, statusCol) = namesA(field) & " Status"
    Next j
    headers(1, resultCols) = "Record Status"
    wsOut.Cells(RESULTS_HEADER_ROW, 1).Resize(1, resultCols).Value2 = headers
    With wsOut.Cells(RESULTS_HEADER_ROW, 1).Resize(1, resultCols)
        .Font.Bold = True
        .Font.Color = RGB(28, 63, 99)
        .RowHeight = 28
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    wsOut.Columns(1).NumberFormat = wsA.Cells(startA, uniqueColA).NumberFormat
    wsOut.Columns(1).ColumnWidth = 20
    For j = 1 To compareFields
        statusCol = 1 + j * 3
        wsOut.Columns(statusCol - 2).ColumnWidth = 19
        wsOut.Columns(statusCol - 1).ColumnWidth = 19
        field = fields(j)
        wsOut.Columns(statusCol - 2).NumberFormat = wsA.Cells(startA, CLng(mapA(CStr(field)))).NumberFormat
        wsOut.Columns(statusCol - 1).NumberFormat = wsB.Cells(startB, CLng(mapB(CStr(field)))).NumberFormat
        wsOut.Columns(statusCol).ColumnWidth = 17
        If totalResult > 0 Then wsOut.Range(wsOut.Cells(RESULTS_FIRST_ROW, statusCol), _
                wsOut.Cells(RESULTS_HEADER_ROW + totalResult, statusCol)).Font.Bold = True
    Next j
    wsOut.Columns(resultCols).ColumnWidth = 20

    totalMismatch = 0
    If totalResult > 0 Then
        For blockBegin = 1 To totalResult Step BLOCK_ROWS
            thisSize = BLOCK_ROWS
            If blockBegin + thisSize - 1 > totalResult Then thisSize = totalResult - blockBegin + 1
            ReDim dataBlock(1 To thisSize, 1 To resultCols)
            Application.StatusBar = "DM Compare: comparing " & blockBegin & " of " & totalResult & " records..."
            For i = 1 To thisSize
                rowIx = blockBegin + i - 1
                rA = 0
                rB = 0
                If rowIx <= rowsA.Count Then
                    rA = CLng(rowsA(rowIx))
                    idKey = DM_UniqueKey(rawA(rA, uniqueColA))
                    If keysB.Exists(idKey) Then rB = CLng(keysB(idKey))
                    dataBlock(i, 1) = rawA(rA, uniqueColA)
                    If rB = 0 Then
                        recordStatus = "Only in File A"
                    Else
                        recordStatus = "Matched"
                    End If
                Else
                    ' Append B-only records in File B's original row order.
                    rB = CLng(rowsOnlyB(rowIx - rowsA.Count))
                    dataBlock(i, 1) = rawB(rB, uniqueColB)
                    recordStatus = "Only in File B"
                End If
                For j = 1 To compareFields
                    field = fields(j)
                    aCol = CLng(mapA(CStr(field)))
                    bCol = CLng(mapB(CStr(field)))
                    statusCol = 1 + j * 3
                    If rA > 0 Then dataBlock(i, statusCol - 2) = rawA(rA, aCol)
                    If rB > 0 Then dataBlock(i, statusCol - 1) = rawB(rB, bCol)
                    If rA > 0 And rB > 0 Then
                        valueA = rawA(rA, aCol)
                        valueB = rawB(rB, bCol)
                        matching = DM_SameValue(valueA, valueB, wsA, startA + rA - 1, aCol, _
                                                wsB, startB + rB - 1, bCol)
                        If matching Then
                            dataBlock(i, statusCol) = "Correct"
                        Else
                            dataBlock(i, statusCol) = "Mismatched"
                            totalMismatch = totalMismatch + 1
                        End If
                    End If
                Next j
                dataBlock(i, resultCols) = recordStatus
            Next i
            wsOut.Cells(RESULTS_FIRST_ROW + blockBegin - 1, 1).Resize(thisSize, resultCols).Value2 = dataBlock
        Next blockBegin
    End If
    wsOut.Range("B4").Value2 = totalMismatch
    fontRed = RGB(192, 31, 44)
    fontGreen = RGB(16, 125, 68)
    reportLastRow = RESULTS_HEADER_ROW + totalResult
    If totalResult > 0 Then
        For j = 1 To compareFields
            statusCol = 1 + j * 3
            Set cf = wsOut.Range(wsOut.Cells(RESULTS_FIRST_ROW, statusCol - 2), _
                    wsOut.Cells(reportLastRow, statusCol)).FormatConditions.Add( _
                    Type:=xlExpression, Formula1:="=$" & DM_ColumnLetter(statusCol) & _
                        RESULTS_FIRST_ROW & "=""Mismatched""")
            cf.Font.Color = fontRed
            Set cf = wsOut.Range(wsOut.Cells(RESULTS_FIRST_ROW, statusCol), _
                    wsOut.Cells(reportLastRow, statusCol)).FormatConditions.Add( _
                    Type:=xlExpression, Formula1:="=$" & DM_ColumnLetter(statusCol) & _
                        RESULTS_FIRST_ROW & "=""Correct""")
            cf.Font.Color = fontGreen
            cf.Font.Bold = True
        Next j
        Set cf = wsOut.Range(wsOut.Cells(RESULTS_FIRST_ROW, resultCols), _
            wsOut.Cells(reportLastRow, resultCols)).FormatConditions.Add( _
            Type:=xlExpression, Formula1:="=LEFT($" & DM_ColumnLetter(resultCols) & _
                RESULTS_FIRST_ROW & ",7)=""Only in""")
        cf.Font.Color = fontRed
    End If
    wsOut.Range(wsOut.Cells(RESULTS_HEADER_ROW, 1), _
        wsOut.Cells(Application.Max(RESULTS_HEADER_ROW + 1, reportLastRow), resultCols)).AutoFilter
    wsOut.Activate
    wsOut.Cells(RESULTS_FIRST_ROW, 2).Select
    ActiveWindow.FreezePanes = True
    elapsed = Timer - beginTime
    If elapsed < 0 Then elapsed = elapsed + 86400#
    wsOut.Range("B5").Value2 = Format$(elapsed, "0.00") & " seconds"
    wsOut.Range("B5").Font.Color = RGB(28, 63, 99)
    wbOut.Activate
    success = True
    DM_RunComparison = True
    GoTo CleanExit
Fail:
    errorText = Err.Description
    MsgBox "Comparison could not be completed:" & vbCrLf & errorText, vbExclamation, "DM Compare Tool"
CleanExit:
    On Error Resume Next
    Application.StatusBar = False
    If openedB Then wbB.Close SaveChanges:=False
    If openedA Then wbA.Close SaveChanges:=False
    If settingsCaptured Then
        Application.AutomationSecurity = oldSecurity
        Application.Calculation = oldCalc
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
    End If
    If Not success Then
        If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    Else
        wbOut.Activate
    End If
    On Error GoTo 0
End Function
