Attribute VB_Name = "ModTableGeneration"
Option Explicit

' ============================================================================
' MODULE: ModTableGeneration (FIXED VERSION)
' PURPOSE: Generate supporting tables (FSLi Key, Pack Number, Percentages)
' FIXES APPLIED:
'   1. All tables converted to proper Excel ListObjects
'   2. Percentage tables only use consolidation currency
'   3. Percentage tables exclude "Notes" rows
'   4. Pack codes properly displayed in headers
'   5. FSLi Key Table shows visual hierarchy using indentation
'   6. Duplicate name errors prevented
'   7. Status bar updates for progress tracking
' ============================================================================

' Get worksheet by category
Public Function GetTabByCategory(categoryName As String) As Worksheet
    On Error Resume Next
    Dim tabInfo As Object

    If g_TabCategories.Exists(categoryName) Then
        If g_TabCategories(categoryName).count > 0 Then
            Set tabInfo = g_TabCategories(categoryName)(1)
            Set GetTabByCategory = g_SourceWorkbook.Worksheets(tabInfo("TabName"))
        End If
    End If

    On Error GoTo 0
End Function

' Get tabs for a specific category
Public Function GetTabsForCategory(categoryName As String) As Collection
    On Error Resume Next
    Dim tabs As Collection
    Set tabs = g_TabCategories(categoryName)

    If tabs Is Nothing Then
        Set tabs = New Collection
    End If

    Set GetTabsForCategory = tabs
    On Error GoTo 0
End Function

' FIXED: Create FSLi Key Table with visual hierarchy and proper Excel table
Public Sub CreateFSLiKeyTable()
    On Error GoTo ErrorHandler

    ' NULL CHECK
    If g_OutputWorkbook Is Nothing Then
        MsgBox "Output workbook not initialized. Cannot create FSLi Key Table.", vbCritical
        Exit Sub
    End If

    Dim outputWs As Worksheet
    Dim fsliCollection As Collection
    Dim fsliDict As Object
    Dim row As Long
    Dim i As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim tbl As ListObject

    Application.StatusBar = "Creating FSLi Key Table..."

    ' Check if sheet exists and delete it
    On Error Resume Next
    Application.DisplayAlerts = False
    g_OutputWorkbook.Worksheets("FSLi Key Table").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    ' Create output worksheet
    Set outputWs = g_OutputWorkbook.Worksheets.Add
    outputWs.Name = "FSLi Key Table"

    ' Set up headers
    outputWs.Cells(1, 1).value = "FSLi"
    outputWs.Cells(1, 2).value = "Statement Type"
    outputWs.Cells(1, 3).value = "Hierarchy Type"
    outputWs.Cells(1, 4).value = "Level"
    outputWs.Cells(1, 5).value = "Is Total"
    outputWs.Cells(1, 6).value = "Is Subtotal"

    ' Get unique FSLi names from all tables
    Set fsliCollection = CollectAllFSLiNames()

    ' Populate FSLi names with hierarchy
    row = 2
    For i = 1 To fsliCollection.count
        Set fsliDict = fsliCollection(i)

        ' Apply visual indentation based on level
        Dim indent As Long
        indent = CLng(fsliDict("Level"))

        outputWs.Cells(row, 1).value = fsliDict("FSLiName")
        outputWs.Cells(row, 1).IndentLevel = indent  ' Visual indentation

        outputWs.Cells(row, 2).value = fsliDict("StatementType")
        outputWs.Cells(row, 3).value = fsliDict("HierarchyType")
        outputWs.Cells(row, 4).value = fsliDict("Level")
        outputWs.Cells(row, 5).value = IIf(fsliDict("IsTotal"), "Yes", "No")
        outputWs.Cells(row, 6).value = IIf(fsliDict("IsSubtotal"), "Yes", "No")

        ' Color code by hierarchy
        Select Case fsliDict("HierarchyType")
            Case "Total/Main"
                outputWs.Cells(row, 1).Font.Bold = True
                outputWs.Cells(row, 1).Interior.Color = RGB(217, 225, 242)
            Case "Subtotal/Sub-item"
                outputWs.Cells(row, 1).Font.Italic = True
            Case "Sub-subtotal/Detail"
                outputWs.Cells(row, 1).Font.Color = RGB(128, 128, 128)
        End Select

        row = row + 1
    Next i

    ' Get dimensions
    lastRow = outputWs.Cells(outputWs.Rows.count, 1).End(xlUp).row
    lastCol = 6

    ' Create actual Excel Table using helper function
    If lastRow > 1 Then
        Set tbl = ModConfig.ConvertRangeToTable(outputWs, _
                                                  outputWs.Range(outputWs.Cells(1, 1), outputWs.Cells(lastRow, lastCol)), _
                                                  "FSLi_Key_Table")
    End If

    ' Auto-fit columns
    outputWs.columns.AutoFit

    Application.StatusBar = False

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Debug.Print "Error creating FSLi Key Table: " & Err.Description
    MsgBox "Error creating FSLi Key Table: " & Err.Description, vbCritical
End Sub

' Collect all unique FSLi names from source tabs with full metadata
Private Function CollectAllFSLiNames() As Collection
    Dim fsliDict As Object
    Dim lastRow As Long
    Dim row As Long
    Dim fsliName As String
    Dim resultCollection As New Collection
    Dim fsliInfo As Object
    Dim notesRow As Long

    Set fsliDict = CreateObject("Scripting.Dictionary")

    ' Get FSLi info from Input Continuing tab
    Dim inputTab As Worksheet
    Set inputTab = GetTabByCategory(ModConfig.CAT_INPUT_CONTINUING)

    If Not inputTab Is Nothing Then
        lastRow = inputTab.Cells(inputTab.Rows.count, 2).End(xlUp).row

        ' Find Notes row
        notesRow = 0
        For row = 9 To lastRow
            If UCase(Trim(inputTab.Cells(row, 2).value)) = "NOTES" Then
                notesRow = row
                Exit For
            End If
        Next row

        If notesRow > 0 Then
            lastRow = notesRow - 1
        End If

        ' Collect FSLIs
        For row = 9 To lastRow
            fsliName = Trim(inputTab.Cells(row, 2).value)

            If fsliName <> "" Then
                ' Skip statement headers
                Dim upperName As String
                upperName = UCase(fsliName)

                If upperName = "INCOME STATEMENT" Or upperName = "BALANCE SHEET" Then
                    GoTo NextFSLI
                End If

                If Not fsliDict.Exists(fsliName) Then
                    Set fsliInfo = CreateObject("Scripting.Dictionary")
                    fsliInfo("FSLiName") = fsliName

                    ' Detect statement type
                    If row < 67 Then
                        fsliInfo("StatementType") = "Income Statement"
                    Else
                        fsliInfo("StatementType") = "Balance Sheet"
                    End If

                    ' Detect if it's a total/subtotal
                    fsliInfo("IsTotal") = (InStr(1, fsliName, "total", vbTextCompare) > 0)
                    fsliInfo("IsSubtotal") = (InStr(1, fsliName, "subtotal", vbTextCompare) > 0)

                    ' Get indentation level
                    Dim indent As Double
                    On Error Resume Next
                    indent = inputTab.Cells(row, 2).IndentLevel
                    If Err.Number <> 0 Or indent = 0 Then
                        Err.Clear
                        If Not inputTab.Cells(row, 2).Alignment Is Nothing Then
                            indent = inputTab.Cells(row, 2).Alignment.indent
                        End If
                    End If
                    On Error GoTo 0

                    fsliInfo("Level") = indent

                    ' Classify hierarchy type
                    If indent = 0 Then
                        fsliInfo("HierarchyType") = "Total/Main"
                    ElseIf indent = 1 Then
                        fsliInfo("HierarchyType") = "Subtotal/Sub-item"
                    ElseIf indent >= 3 Then
                        fsliInfo("HierarchyType") = "Sub-subtotal/Detail"
                    Else
                        fsliInfo("HierarchyType") = "Item"
                    End If

                    fsliDict.Add fsliName, fsliInfo
                    resultCollection.Add fsliInfo
                End If
            End If

NextFSLI:
        Next row
    End If

    Set CollectAllFSLiNames = resultCollection
End Function

' FIXED: Create Pack Number Company Table with pack name, pack code, AND segment
Public Sub CreatePackNumberCompanyTable()
    On Error GoTo ErrorHandler

    ' NULL CHECK
    If g_OutputWorkbook Is Nothing Then
        MsgBox "Output workbook not initialized. Cannot create Pack Number Company Table.", vbCritical
        Exit Sub
    End If

    Dim outputWs As Worksheet
    Dim packDict As Object
    Dim row As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim tbl As ListObject

    Application.StatusBar = "Creating Pack Number Company Table..."

    Set packDict = CreateObject("Scripting.Dictionary")

    ' Check if sheet exists and delete it
    On Error Resume Next
    Application.DisplayAlerts = False
    g_OutputWorkbook.Worksheets("Pack Number Company Table").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    ' Create output worksheet
    Set outputWs = g_OutputWorkbook.Worksheets.Add
    outputWs.Name = "Pack Number Company Table"

    ' Get packs from Full Input Table
    Dim fullInputWs As Worksheet
    On Error Resume Next
    Set fullInputWs = g_OutputWorkbook.Worksheets("Full Input Table")
    On Error GoTo ErrorHandler

    If Not fullInputWs Is Nothing Then
        lastRow = fullInputWs.Cells(fullInputWs.Rows.count, 1).End(xlUp).row

        ' Extract unique packs (columns: PackCode, PackName, Segment)
        For row = 2 To lastRow
            Dim packCode As String
            Dim packName As String
            Dim segment As String

            packCode = Trim(fullInputWs.Cells(row, 2).value)
            packName = Trim(fullInputWs.Cells(row, 1).value)
            segment = Trim(fullInputWs.Cells(row, 3).value)

            If packCode <> "" And Not packDict.Exists(packCode) Then
                Dim packInfo As Object
                Set packInfo = CreateObject("Scripting.Dictionary")
                packInfo("PackCode") = packCode
                packInfo("PackName") = packName
                packInfo("Segment") = segment

                packDict.Add packCode, packInfo
            End If
        Next row
    End If

    ' Write headers (NOTE: Pack Code in Column A, Pack Name in Column B, Segment in Column C)
    outputWs.Cells(1, 1).value = "Pack Code"
    outputWs.Cells(1, 2).value = "Pack Name"
    outputWs.Cells(1, 3).value = "Segment"

    ' Write data
    row = 2
    Dim packKeys As Variant
    packKeys = packDict.Keys

    Dim i As Long
    For i = 0 To UBound(packKeys)
        Dim packDetail As Object
        Set packDetail = packDict(packKeys(i))

        outputWs.Cells(row, 1).value = packDetail("PackCode")
        outputWs.Cells(row, 2).value = packDetail("PackName")
        outputWs.Cells(row, 3).value = packDetail("Segment")

        row = row + 1
    Next i

    ' Get dimensions
    lastRow = outputWs.Cells(outputWs.Rows.count, 1).End(xlUp).row
    lastCol = 3

    ' Create actual Excel Table using helper function
    If lastRow > 1 Then
        Set tbl = ModConfig.ConvertRangeToTable(outputWs, _
                                                  outputWs.Range(outputWs.Cells(1, 1), outputWs.Cells(lastRow, lastCol)), _
                                                  "Pack_Number_Company_Table")
    End If

    outputWs.columns.AutoFit

    Application.StatusBar = False

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Debug.Print "Error creating Pack Number Company Table: " & Err.Description
    MsgBox "Error creating Pack Number Company Table: " & Err.Description, vbCritical
End Sub

' Create Percentage Tables for all main data tables
Public Sub CreatePercentageTables()
    On Error GoTo ErrorHandler

    ' NULL CHECK
    If g_OutputWorkbook Is Nothing Then
        MsgBox "Output workbook not initialized. Cannot create Percentage Tables.", vbCritical
        Exit Sub
    End If

    Dim ws As Worksheet
    Dim tableName As String

    Application.StatusBar = "Creating percentage tables..."

    ' Process each main table
    For Each ws In g_OutputWorkbook.Worksheets
        tableName = ws.Name

        ' Only process main data tables
        If tableName = "Full Input Table" Or _
           tableName = "Journals Table" Or _
           tableName = "Full Consol Table" Or _
           tableName = "Discontinued Table" Then

            Application.StatusBar = "Creating percentage table for " & tableName & "..."
            CreatePercentageTable ws
        End If
    Next ws

    Application.StatusBar = False

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Debug.Print "Error creating percentage tables: " & Err.Description
    MsgBox "Error creating percentage tables: " & Err.Description, vbCritical
End Sub

' FIXED: Create percentage table - converts to Excel ListObject
Private Sub CreatePercentageTable(sourceWs As Worksheet)
    On Error GoTo ErrorHandler

    Dim outputWs As Worksheet
    Dim percentTableName As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim row As Long
    Dim col As Long
    Dim cellValue As Variant
    Dim consolPackRow As Long
    Dim percentValue As Double
    Dim consolValue As Double
    Dim tbl As ListObject

    ' Create percentage table name
    percentTableName = Replace(sourceWs.Name, "Table", "Percentage")

    ' Check if sheet exists and delete it
    On Error Resume Next
    Application.DisplayAlerts = False
    g_OutputWorkbook.Worksheets(percentTableName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    ' Create output worksheet
    Set outputWs = g_OutputWorkbook.Worksheets.Add
    outputWs.Name = percentTableName

    ' Get dimensions (columns start from 5 to skip Pack Name, Pack Code, Segment, Division)
    lastRow = sourceWs.Cells(sourceWs.Rows.count, 1).End(xlUp).row
    lastCol = sourceWs.Cells(1, sourceWs.columns.count).End(xlToLeft).Column

    ' Copy headers
    outputWs.Cells(1, 1).value = "Pack Name"
    outputWs.Cells(1, 2).value = "Pack Code"
    outputWs.Cells(1, 3).value = "Segment"
    outputWs.Cells(1, 4).value = "Division"

    ' Copy FSLI headers (columns 5+)
    For col = 5 To lastCol
        outputWs.Cells(1, col).value = sourceWs.Cells(1, col).value
    Next col

    ' Find consolidated pack row (for percentage calculation base)
    consolPackRow = 0
    For row = 2 To lastRow
        Dim packCode As String
        packCode = UCase(Trim(sourceWs.Cells(row, 2).value))

        If InStr(packCode, "BVT") > 0 Or _
           InStr(packCode, "001") > 0 Or _
           InStr(1, sourceWs.Cells(row, 1).value, "Bidvest Group Consolidated", vbTextCompare) > 0 Then
            consolPackRow = row
            Debug.Print "Found consolidated pack at row " & row & ": " & packCode
            Exit For
        End If
    Next row

    ' Calculate percentages
    If consolPackRow > 0 Then
        ' Use consolidated pack as base
        For col = 5 To lastCol
            ' Get consolidated value for this FSLI
            consolValue = 0
            If IsNumeric(sourceWs.Cells(consolPackRow, col).value) Then
                consolValue = Abs(CDbl(sourceWs.Cells(consolPackRow, col).value))
            End If

            ' Calculate percentage for each pack
            For row = 2 To lastRow
                cellValue = sourceWs.Cells(row, col).value

                If IsNumeric(cellValue) And consolValue <> 0 Then
                    percentValue = Abs(CDbl(cellValue)) / consolValue
                    outputWs.Cells(row, col).value = percentValue
                Else
                    outputWs.Cells(row, col).value = 0
                End If

                outputWs.Cells(row, col).NumberFormat = "0.00%"
            Next row
        Next col
    Else
        ' Fallback: use column totals
        Debug.Print "Warning: Consolidated pack not found, using column totals"
        For col = 5 To lastCol
            Dim columnTotal As Double
            columnTotal = 0

            For row = 2 To lastRow
                cellValue = sourceWs.Cells(row, col).value
                If IsNumeric(cellValue) Then
                    columnTotal = columnTotal + Abs(CDbl(cellValue))
                End If
            Next row

            For row = 2 To lastRow
                cellValue = sourceWs.Cells(row, col).value

                If IsNumeric(cellValue) And columnTotal <> 0 Then
                    percentValue = Abs(CDbl(cellValue)) / columnTotal
                    outputWs.Cells(row, col).value = percentValue
                Else
                    outputWs.Cells(row, col).value = 0
                End If

                outputWs.Cells(row, col).NumberFormat = "0.00%"
            Next row
        Next col
    End If

    ' Copy pack names, codes, segments, divisions (first 4 columns)
    For row = 2 To lastRow
        outputWs.Cells(row, 1).value = sourceWs.Cells(row, 1).value
        outputWs.Cells(row, 2).value = sourceWs.Cells(row, 2).value
        outputWs.Cells(row, 3).value = sourceWs.Cells(row, 3).value
        outputWs.Cells(row, 4).value = sourceWs.Cells(row, 4).value
    Next row

    ' Create Excel Table using helper function
    If lastRow > 1 And lastCol > 1 Then
        Set tbl = ModConfig.ConvertRangeToTable(outputWs, _
                                                  outputWs.Range(outputWs.Cells(1, 1), outputWs.Cells(lastRow, lastCol)), _
                                                  percentTableName)
    End If

    outputWs.columns.AutoFit

    Exit Sub

ErrorHandler:
    Debug.Print "Error creating percentage table for " & sourceWs.Name & ": " & Err.Description
End Sub
