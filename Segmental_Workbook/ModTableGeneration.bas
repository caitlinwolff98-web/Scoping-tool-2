Attribute VB_Name = "ModTableGeneration"
Option Explicit

'==============================================================================
' MODULE: ModTableGeneration
' PURPOSE: Generate FullTable and FullTablePercentage tables
' DESCRIPTION: Creates aggregated tables from segment data with FSLI columns
'              and pack rows, including percentage calculations
'==============================================================================

'==================== FULL TABLE GENERATION ====================

' Generate the main FullTable from all segment tabs
Public Function GenerateFullTable(ByVal sourceWb As Workbook, ByVal targetWb As Workbook) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "=== Starting FullTable Generation ==="

    ' Get all segment tabs
    Dim segmentTabs As Collection
    Set segmentTabs = GetSegmentTabs(sourceWb)

    If segmentTabs.count = 0 Then
        ShowError "No Segment Tabs", "No segment tabs found. Cannot generate FullTable."
        GenerateFullTable = False
        Exit Function
    End If

    ' Collect all unique FSLIs
    Dim allFSLIs As Collection
    Set allFSLIs = CollectAllFSLIsFromSegments(segmentTabs)

    LogMessage "Collected " & allFSLIs.count & " unique FSLIs"

    ' Collect all unique packs
    Dim allPacks As Collection
    Set allPacks = CollectAllPacks(segmentTabs)

    LogMessage "Collected " & allPacks.count & " unique packs"

    ' Create or clear the FullTable worksheet
    Dim wsFullTable As Worksheet
    Set wsFullTable = GetOrCreateWorksheet(targetWb, TABLE_FULL_TABLE)
    wsFullTable.Cells.Clear

    ' Build the table structure
    Call BuildFullTableStructure(wsFullTable, allFSLIs, allPacks, segmentTabs)

    LogMessage "=== FullTable Generation Complete ==="
    GenerateFullTable = True
    Exit Function

ErrorHandler:
    ShowError "FullTable Generation Error", "Error generating FullTable: " & Err.Description
    GenerateFullTable = False
End Function

' Build the structure and populate the FullTable
Private Sub BuildFullTableStructure(ByVal ws As Worksheet, ByVal fslis As Collection, ByVal packs As Collection, ByVal segmentTabs As Collection)
    On Error GoTo ErrorHandler

    LogMessage "Building FullTable structure..."

    Dim rowIdx As Long
    Dim colIdx As Long
    Dim fsli As Variant
    Dim pack As Variant
    Dim segmentWs As Worksheet

    ' Row 1: Header row with column names
    ws.Cells(1, 1).Value = "Pack Name"
    ws.Cells(1, 2).Value = "Pack Code"
    ws.Cells(1, 3).Value = "Segment"

    ' Row 1: FSLI columns (starting from column 4)
    colIdx = 4
    For Each fsli In fslis
        ws.Cells(1, colIdx).Value = CStr(fsli)
        colIdx = colIdx + 1
    Next fsli

    ' Format header row
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, colIdx - 1))
        .Font.Bold = True
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Data rows: One row per pack per segment
    rowIdx = 2

    For Each segmentWs In segmentTabs
        Dim segmentName As String
        segmentName = g_SegmentNames(segmentWs.Name)

        LogMessage "Processing segment: " & segmentName & " (Tab: " & segmentWs.Name & ")"

        ' Analyze Row 6 to get valid columns
        Dim lastCol As Long
        lastCol = GetDataColumnRange(segmentWs)

        Dim colAnalysis As Collection
        Set colAnalysis = AnalyzeRow6(segmentWs, COL_FSLI + 1, lastCol)

        ' Create a map of pack code -> column index for this segment
        Dim packColumnMap As Object
        Set packColumnMap = CreateDictionary()

        Dim analysis As Variant
        For Each analysis In colAnalysis
            If analysis.IsValid And analysis.PackCode <> "" Then
                If Not packColumnMap.Exists(analysis.PackCode) Then
                    packColumnMap.Add analysis.PackCode, analysis.ColumnIndex
                End If
            End If
        Next analysis

        ' For each pack, create a row
        For Each pack In packs
            Dim packName As String
            Dim packCode As String

            packName = pack(0)
            packCode = pack(1)

            ' Check if this pack exists in this segment
            If packColumnMap.Exists(packCode) Then
                Dim packColIdx As Long
                packColIdx = packColumnMap(packCode)

                ' Fill in pack info columns
                ws.Cells(rowIdx, 1).Value = packName
                ws.Cells(rowIdx, 2).Value = packCode
                ws.Cells(rowIdx, 3).Value = segmentName

                ' Fill in FSLI values
                colIdx = 4
                For Each fsli In fslis
                    Dim fsliValue As Variant
                    fsliValue = GetSegmentPackFSLIValue(segmentWs, CStr(fsli), packColIdx)

                    If IsValidNumber(fsliValue) Then
                        ws.Cells(rowIdx, colIdx).Value = CDbl(fsliValue)
                    Else
                        ws.Cells(rowIdx, colIdx).Value = 0
                    End If

                    colIdx = colIdx + 1
                Next fsli

                rowIdx = rowIdx + 1
            End If
        Next pack
    Next segmentWs

    ' Format as table
    If rowIdx > 2 Then
        Dim tableRange As Range
        Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowIdx - 1, 3 + fslis.count))

        ' Apply borders
        tableRange.Borders.LineStyle = xlContinuous
        tableRange.Borders.Weight = xlThin

        ' Auto-fit columns
        ws.Columns.AutoFit
    End If

    LogMessage "FullTable populated with " & (rowIdx - 2) & " data rows"

    Exit Sub

ErrorHandler:
    ShowError "Table Structure Error", "Error building FullTable structure: " & Err.Description
End Sub

'==================== FULL TABLE PERCENTAGE GENERATION ====================

' Generate FullTablePercentage with base amounts as 100%
Public Function GenerateFullTablePercentage(ByVal sourceWb As Workbook, ByVal targetWb As Workbook, ByVal summaryWs As Worksheet) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "=== Starting FullTablePercentage Generation ==="

    ' Find base amounts columns in summary tab
    Dim baseAmountsColumns As Collection
    Set baseAmountsColumns = FindBaseAmountsColumns(summaryWs, g_BaseAmountsLabel)

    If baseAmountsColumns.count = 0 Then
        ShowError "Base Amounts Not Found", ERR_BASE_AMOUNTS_NOT_FOUND
        GenerateFullTablePercentage = False
        Exit Function
    End If

    ' Get all FSLIs
    Dim allFSLIs As Collection
    Set allFSLIs = CollectAllFSLIsFromSegments(GetSegmentTabs(sourceWb))

    ' Calculate base amounts for each FSLI
    Dim baseAmounts As Object
    Set baseAmounts = CalculateBaseAmounts(summaryWs, baseAmountsColumns, allFSLIs)

    ' Create base amounts calculation table for transparency
    Call CreateBaseAmountsCalculationTable(targetWb, summaryWs, baseAmountsColumns, allFSLIs, baseAmounts)

    ' Get the FullTable worksheet
    Dim wsFullTable As Worksheet
    Set wsFullTable = GetWorksheetByName(targetWb, TABLE_FULL_TABLE)

    If wsFullTable Is Nothing Then
        ShowError "FullTable Not Found", "FullTable must be generated before FullTablePercentage."
        GenerateFullTablePercentage = False
        Exit Function
    End If

    ' Create or clear the FullTablePercentage worksheet
    Dim wsFullTablePct As Worksheet
    Set wsFullTablePct = GetOrCreateWorksheet(targetWb, TABLE_FULL_TABLE_PCT)
    wsFullTablePct.Cells.Clear

    ' Build the percentage table
    Call BuildFullTablePercentageStructure(wsFullTable, wsFullTablePct, baseAmounts)

    LogMessage "=== FullTablePercentage Generation Complete ==="
    GenerateFullTablePercentage = True
    Exit Function

ErrorHandler:
    ShowError "FullTablePercentage Error", "Error generating FullTablePercentage: " & Err.Description
    GenerateFullTablePercentage = False
End Function

' Build the FullTablePercentage structure
Private Sub BuildFullTablePercentageStructure(ByVal wsSource As Worksheet, ByVal wsTarget As Worksheet, ByVal baseAmounts As Object)
    On Error GoTo ErrorHandler

    LogMessage "Building FullTablePercentage structure..."

    ' Copy the entire FullTable structure
    wsSource.Cells.Copy wsTarget.Cells(1, 1)

    ' Get the dimensions
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = wsTarget.Cells(wsTarget.Rows.count, 1).End(xlUp).Row
    lastCol = wsTarget.Cells(1, wsTarget.Columns.count).End(xlToLeft).Column

    ' Convert FSLI values to percentages
    Dim rowIdx As Long
    Dim colIdx As Long
    Dim fsli As String
    Dim cellValue As Variant
    Dim baseAmount As Double
    Dim percentage As Double

    For rowIdx = 2 To lastRow ' Start from row 2 (skip header)
        For colIdx = 4 To lastCol ' Start from column 4 (first FSLI column)
            ' Get FSLI name from header
            fsli = SafeTrim(wsTarget.Cells(1, colIdx).Value)

            ' Get base amount for this FSLI
            If baseAmounts.Exists(fsli) Then
                baseAmount = baseAmounts(fsli)

                ' Get cell value
                cellValue = wsTarget.Cells(rowIdx, colIdx).Value

                If IsValidNumber(cellValue) And baseAmount <> 0 Then
                    percentage = (CDbl(cellValue) / baseAmount) * 100
                    wsTarget.Cells(rowIdx, colIdx).Value = percentage
                    wsTarget.Cells(rowIdx, colIdx).NumberFormat = "0.00%"
                Else
                    wsTarget.Cells(rowIdx, colIdx).Value = 0
                    wsTarget.Cells(rowIdx, colIdx).NumberFormat = "0.00%"
                End If
            End If
        Next colIdx
    Next rowIdx

    LogMessage "FullTablePercentage populated with percentage values"

    Exit Sub

ErrorHandler:
    ShowError "Percentage Table Error", "Error building FullTablePercentage: " & Err.Description
End Sub

'==================== BASE AMOUNTS CALCULATION TABLE ====================

' Create a table showing how base amounts were calculated
Private Sub CreateBaseAmountsCalculationTable(ByVal targetWb As Workbook, ByVal summaryWs As Worksheet, _
                                               ByVal baseAmountsColumns As Collection, ByVal fslis As Collection, _
                                               ByVal baseAmounts As Object)
    On Error GoTo ErrorHandler

    LogMessage "Creating Base Amounts Calculation table..."

    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet(targetWb, TABLE_BASE_CALC)
    ws.Cells.Clear

    ' Header row
    ws.Cells(1, 1).Value = "FSLI"
    Dim colIdx As Long
    Dim col As Variant
    colIdx = 2

    For Each col In baseAmountsColumns
        ws.Cells(1, colIdx).Value = "Column " & col
        colIdx = colIdx + 1
    Next col

    ws.Cells(1, colIdx).Value = "Total Base Amount"

    ' Format header
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, colIdx))
        .Font.Bold = True
        .Interior.Color = RGB(146, 208, 80)
        .Font.Color = RGB(0, 0, 0)
    End With

    ' Data rows
    Dim rowIdx As Long
    Dim fsli As Variant
    Dim fsliRow As Long
    Dim total As Double

    rowIdx = 2
    For Each fsli In fslis
        ws.Cells(rowIdx, 1).Value = CStr(fsli)

        ' Find FSLI row in summary
        fsliRow = FindFSLIRow(summaryWs, CStr(fsli))

        If fsliRow > 0 Then
            colIdx = 2
            total = 0

            For Each col In baseAmountsColumns
                Dim cellValue As Variant
                cellValue = summaryWs.Cells(fsliRow, CLng(col)).Value

                If IsValidNumber(cellValue) Then
                    ws.Cells(rowIdx, colIdx).Value = CDbl(cellValue)
                    total = total + CDbl(cellValue)
                Else
                    ws.Cells(rowIdx, colIdx).Value = 0
                End If

                colIdx = colIdx + 1
            Next col

            ws.Cells(rowIdx, colIdx).Value = total
        End If

        rowIdx = rowIdx + 1
    Next fsli

    ' Format as table
    Dim tableRange As Range
    Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowIdx - 1, colIdx))
    tableRange.Borders.LineStyle = xlContinuous
    tableRange.Borders.Weight = xlThin

    ' Auto-fit columns
    ws.Columns.AutoFit

    LogMessage "Base Amounts Calculation table created"

    Exit Sub

ErrorHandler:
    ShowError "Base Calculation Table Error", "Error creating base amounts calculation table: " & Err.Description
End Sub

'==================== HELPER FUNCTIONS ====================

' Collect all unique FSLIs from all segment tabs
Private Function CollectAllFSLIsFromSegments(ByVal segmentTabs As Collection) As Collection
    Dim allFSLIs As New Collection
    Dim fsliDict As Object
    Set fsliDict = CreateDictionary()

    Dim ws As Variant
    Dim fslis As Collection
    Dim fsli As Variant

    For Each ws In segmentTabs
        Set fslis = ExtractFSLIs(ws, 10) ' Start from row 10 to skip headers

        For Each fsli In fslis
            If Not fsliDict.Exists(CStr(fsli)) Then
                fsliDict.Add CStr(fsli), True
                allFSLIs.Add fsli
            End If
        Next fsli
    Next ws

    Set CollectAllFSLIsFromSegments = allFSLIs
End Function

' Get or create a worksheet
Private Function GetOrCreateWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = sheetName
    End If

    Set GetOrCreateWorksheet = ws
End Function

' Get worksheet by name
Private Function GetWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetByName = wb.Worksheets(sheetName)
End Function
