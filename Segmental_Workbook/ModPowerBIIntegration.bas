Attribute VB_Name = "ModPowerBIIntegration"
Option Explicit

'==============================================================================
' MODULE: ModPowerBIIntegration
' PURPOSE: Create all Power BI integration tables
' DESCRIPTION: Generates Scoping Control, Pack Number Company, FSLI Key,
'              Segment List, and other tables needed for Power BI dashboards
'==============================================================================

'==================== MAIN POWER BI TABLE GENERATION ====================

' Generate all Power BI integration tables
Public Function GenerateAllPowerBITables(ByVal sourceWb As Workbook, ByVal targetWb As Workbook) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "=== Starting Power BI Table Generation ==="

    ' Create Scoping Control table
    If Not CreateScopingControlTable(targetWb) Then
        LogMessage "WARNING: Failed to create Scoping Control table"
    End If

    ' Create Pack Number Company table
    If Not CreatePackNumberCompanyTable(targetWb) Then
        LogMessage "WARNING: Failed to create Pack Number Company table"
    End If

    ' Create FSLI Key table
    If Not CreateFSLIKeyTable(targetWb) Then
        LogMessage "WARNING: Failed to create FSLI Key table"
    End If

    ' Create Segment List table
    If Not CreateSegmentListTable(targetWb) Then
        LogMessage "WARNING: Failed to create Segment List table"
    End If

    ' Create Threshold Configuration table
    If Not CreateThresholdConfigurationTable(targetWb) Then
        LogMessage "WARNING: Failed to create Threshold Configuration table"
    End If

    ' Create Manual Scoping Adjustments table
    If Not CreateManualScopingTable(targetWb) Then
        LogMessage "WARNING: Failed to create Manual Scoping Adjustments table"
    End If

    ' Create Power BI Metadata sheet
    If Not CreatePowerBIMetadataSheet(targetWb) Then
        LogMessage "WARNING: Failed to create Power BI Metadata sheet"
    End If

    LogMessage "=== Power BI Table Generation Complete ==="
    GenerateAllPowerBITables = True
    Exit Function

ErrorHandler:
    ShowError "Power BI Generation Error", "Error generating Power BI tables: " & Err.Description
    GenerateAllPowerBITables = False
End Function

'==================== SCOPING CONTROL TABLE ====================

' Create the Scoping Control table for Power BI
Private Function CreateScopingControlTable(ByVal wb As Workbook) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "Creating Scoping Control table..."

    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet(wb, TABLE_SCOPING_CONTROL)
    ws.Cells.Clear

    ' Header row
    ws.Cells(1, 1).Value = "Pack Code"
    ws.Cells(1, 2).Value = "Pack Name"
    ws.Cells(1, 3).Value = "Segment"
    ws.Cells(1, 4).Value = "FSLI"
    ws.Cells(1, 5).Value = "Amount"
    ws.Cells(1, 6).Value = "Percentage"
    ws.Cells(1, 7).Value = "Is Scoped"
    ws.Cells(1, 8).Value = "Scoping Method"
    ws.Cells(1, 9).Value = "Threshold Level"
    ws.Cells(1, 10).Value = "Manual Override"
    ws.Cells(1, 11).Value = "Comments"

    ' Format header
    With ws.Range("A1:K1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Populate from FullTable
    Dim wsFullTable As Worksheet
    Set wsFullTable = GetWorksheetByName(wb, TABLE_FULL_TABLE)

    If Not wsFullTable Is Nothing Then
        Dim lastRow As Long
        Dim lastCol As Long
        Dim rowIdx As Long
        Dim colIdx As Long
        Dim outRow As Long

        lastRow = wsFullTable.Cells(wsFullTable.Rows.count, 1).End(xlUp).Row
        lastCol = wsFullTable.Cells(1, wsFullTable.Columns.count).End(xlToLeft).Column

        outRow = 2

        ' Loop through each row in FullTable
        For rowIdx = 2 To lastRow
            Dim packName As String
            Dim packCode As String
            Dim segment As String

            packName = wsFullTable.Cells(rowIdx, 1).Value
            packCode = wsFullTable.Cells(rowIdx, 2).Value
            segment = wsFullTable.Cells(rowIdx, 3).Value

            ' Loop through each FSLI column
            For colIdx = 4 To lastCol
                Dim fsli As String
                Dim amount As Variant

                fsli = wsFullTable.Cells(1, colIdx).Value
                amount = wsFullTable.Cells(rowIdx, colIdx).Value

                ' Write to Scoping Control
                ws.Cells(outRow, 1).Value = packCode
                ws.Cells(outRow, 2).Value = packName
                ws.Cells(outRow, 3).Value = segment
                ws.Cells(outRow, 4).Value = fsli
                ws.Cells(outRow, 5).Value = amount
                ws.Cells(outRow, 6).Value = 0 ' Will be calculated
                ws.Cells(outRow, 7).Value = "No" ' Default not scoped
                ws.Cells(outRow, 8).Value = "None"
                ws.Cells(outRow, 9).Value = ""
                ws.Cells(outRow, 10).Value = "No"
                ws.Cells(outRow, 11).Value = ""

                outRow = outRow + 1
            Next colIdx
        Next rowIdx

        LogMessage "Scoping Control table populated with " & (outRow - 2) & " rows"
    End If

    ' Format as table
    ws.Columns.AutoFit

    CreateScopingControlTable = True
    Exit Function

ErrorHandler:
    LogMessage "ERROR: Failed to create Scoping Control table: " & Err.Description
    CreateScopingControlTable = False
End Function

'==================== PACK NUMBER COMPANY TABLE ====================

' Create the Pack Number Company lookup table
Private Function CreatePackNumberCompanyTable(ByVal wb As Workbook) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "Creating Pack Number Company table..."

    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet(wb, TABLE_PACK_COMPANY)
    ws.Cells.Clear

    ' Header row
    ws.Cells(1, 1).Value = "Pack Code"
    ws.Cells(1, 2).Value = "Pack Name"
    ws.Cells(1, 3).Value = "Company Name"
    ws.Cells(1, 4).Value = "Company Code"
    ws.Cells(1, 5).Value = "Segment"
    ws.Cells(1, 6).Value = "Active"

    ' Format header
    With ws.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(112, 173, 71)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Populate with unique packs from g_PackList
    Dim rowIdx As Long
    Dim key As Variant
    Dim packInfo As Variant

    rowIdx = 2

    For Each key In g_PackList.Keys
        packInfo = g_PackList(key)

        ws.Cells(rowIdx, 1).Value = packInfo(1) ' Pack Code
        ws.Cells(rowIdx, 2).Value = packInfo(0) ' Pack Name
        ws.Cells(rowIdx, 3).Value = packInfo(0) ' Company Name (same as pack name initially)
        ws.Cells(rowIdx, 4).Value = packInfo(1) ' Company Code (same as pack code initially)
        ws.Cells(rowIdx, 5).Value = "" ' Segment - would need to determine
        ws.Cells(rowIdx, 6).Value = "Yes"

        rowIdx = rowIdx + 1
    Next key

    LogMessage "Pack Number Company table populated with " & (rowIdx - 2) & " packs"

    ws.Columns.AutoFit

    CreatePackNumberCompanyTable = True
    Exit Function

ErrorHandler:
    LogMessage "ERROR: Failed to create Pack Number Company table: " & Err.Description
    CreatePackNumberCompanyTable = False
End Function

'==================== FSLI KEY TABLE ====================

' Create the FSLI Key lookup table
Private Function CreateFSLIKeyTable(ByVal wb As Workbook) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "Creating FSLI Key table..."

    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet(wb, TABLE_FSLI_KEY)
    ws.Cells.Clear

    ' Header row
    ws.Cells(1, 1).Value = "FSLI"
    ws.Cells(1, 2).Value = "FSLI Description"
    ws.Cells(1, 3).Value = "Category"
    ws.Cells(1, 4).Value = "Financial Statement"
    ws.Cells(1, 5).Value = "Sort Order"

    ' Format header
    With ws.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(255, 192, 0)
        .Font.Color = RGB(0, 0, 0)
    End With

    ' Populate with unique FSLIs from g_FSLIList
    Dim rowIdx As Long
    Dim fsli As Variant

    rowIdx = 2

    For Each fsli In g_FSLIList.Keys
        ws.Cells(rowIdx, 1).Value = CStr(fsli)
        ws.Cells(rowIdx, 2).Value = CStr(fsli) ' Description same as FSLI initially
        ws.Cells(rowIdx, 3).Value = DetermineFSLICategory(CStr(fsli))
        ws.Cells(rowIdx, 4).Value = DetermineFinancialStatement(CStr(fsli))
        ws.Cells(rowIdx, 5).Value = rowIdx - 1

        rowIdx = rowIdx + 1
    Next fsli

    LogMessage "FSLI Key table populated with " & (rowIdx - 2) & " FSLIs"

    ws.Columns.AutoFit

    CreateFSLIKeyTable = True
    Exit Function

ErrorHandler:
    LogMessage "ERROR: Failed to create FSLI Key table: " & Err.Description
    CreateFSLIKeyTable = False
End Function

'==================== SEGMENT LIST TABLE ====================

' Create the Segment List table
Private Function CreateSegmentListTable(ByVal wb As Workbook) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "Creating Segment List table..."

    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet(wb, TABLE_SEGMENT_LIST)
    ws.Cells.Clear

    ' Header row
    ws.Cells(1, 1).Value = "Segment Name"
    ws.Cells(1, 2).Value = "Segment Code"
    ws.Cells(1, 3).Value = "Tab Name"
    ws.Cells(1, 4).Value = "Active"

    ' Format header
    With ws.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(155, 194, 230)
        .Font.Color = RGB(0, 0, 0)
    End With

    ' Populate with segments from g_SegmentNames
    Dim rowIdx As Long
    Dim key As Variant

    rowIdx = 2

    For Each key In g_SegmentNames.Keys
        ws.Cells(rowIdx, 1).Value = g_SegmentNames(key) ' Segment Name
        ws.Cells(rowIdx, 2).Value = g_SegmentNames(key) ' Segment Code (same initially)
        ws.Cells(rowIdx, 3).Value = CStr(key) ' Tab Name
        ws.Cells(rowIdx, 4).Value = "Yes"

        rowIdx = rowIdx + 1
    Next key

    LogMessage "Segment List table populated with " & (rowIdx - 2) & " segments"

    ws.Columns.AutoFit

    CreateSegmentListTable = True
    Exit Function

ErrorHandler:
    LogMessage "ERROR: Failed to create Segment List table: " & Err.Description
    CreateSegmentListTable = False
End Function

'==================== THRESHOLD CONFIGURATION TABLE ====================

' Create the Threshold Configuration table
Private Function CreateThresholdConfigurationTable(ByVal wb As Workbook) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "Creating Threshold Configuration table..."

    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet(wb, TABLE_THRESHOLD_CONFIG)
    ws.Cells.Clear

    ' Header row
    ws.Cells(1, 1).Value = "Threshold Name"
    ws.Cells(1, 2).Value = "Threshold Type"
    ws.Cells(1, 3).Value = "Percentage Value"
    ws.Cells(1, 4).Value = "Absolute Value"
    ws.Cells(1, 5).Value = "Apply To Segment"
    ws.Cells(1, 6).Value = "Apply To FSLI"
    ws.Cells(1, 7).Value = "Active"

    ' Format header
    With ws.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(244, 176, 132)
        .Font.Color = RGB(0, 0, 0)
    End With

    ' Add default threshold configurations
    Dim rowIdx As Long
    rowIdx = 2

    ' Materiality threshold
    ws.Cells(rowIdx, 1).Value = "Overall Materiality"
    ws.Cells(rowIdx, 2).Value = "Percentage"
    ws.Cells(rowIdx, 3).Value = 5 ' 5%
    ws.Cells(rowIdx, 4).Value = 0
    ws.Cells(rowIdx, 5).Value = "All"
    ws.Cells(rowIdx, 6).Value = "All"
    ws.Cells(rowIdx, 7).Value = "Yes"
    rowIdx = rowIdx + 1

    ' Performance materiality
    ws.Cells(rowIdx, 1).Value = "Performance Materiality"
    ws.Cells(rowIdx, 2).Value = "Percentage"
    ws.Cells(rowIdx, 3).Value = 2 ' 2%
    ws.Cells(rowIdx, 4).Value = 0
    ws.Cells(rowIdx, 5).Value = "All"
    ws.Cells(rowIdx, 6).Value = "All"
    ws.Cells(rowIdx, 7).Value = "Yes"
    rowIdx = rowIdx + 1

    ' Trivial threshold
    ws.Cells(rowIdx, 1).Value = "Trivial Threshold"
    ws.Cells(rowIdx, 2).Value = "Percentage"
    ws.Cells(rowIdx, 3).Value = 0.5 ' 0.5%
    ws.Cells(rowIdx, 4).Value = 0
    ws.Cells(rowIdx, 5).Value = "All"
    ws.Cells(rowIdx, 6).Value = "All"
    ws.Cells(rowIdx, 7).Value = "Yes"

    LogMessage "Threshold Configuration table created with default thresholds"

    ws.Columns.AutoFit

    CreateThresholdConfigurationTable = True
    Exit Function

ErrorHandler:
    LogMessage "ERROR: Failed to create Threshold Configuration table: " & Err.Description
    CreateThresholdConfigurationTable = False
End Function

'==================== MANUAL SCOPING TABLE ====================

' Create the Manual Scoping Adjustments table
Private Function CreateManualScopingTable(ByVal wb As Workbook) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "Creating Manual Scoping Adjustments table..."

    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet(wb, TABLE_MANUAL_SCOPING)
    ws.Cells.Clear

    ' Header row
    ws.Cells(1, 1).Value = "Pack Code"
    ws.Cells(1, 2).Value = "Segment"
    ws.Cells(1, 3).Value = "FSLI"
    ws.Cells(1, 4).Value = "Scope In/Out"
    ws.Cells(1, 5).Value = "Reason"
    ws.Cells(1, 6).Value = "Adjusted By"
    ws.Cells(1, 7).Value = "Adjustment Date"

    ' Format header
    With ws.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(191, 143, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    LogMessage "Manual Scoping Adjustments table created (empty template)"

    ws.Columns.AutoFit

    CreateManualScopingTable = True
    Exit Function

ErrorHandler:
    LogMessage "ERROR: Failed to create Manual Scoping table: " & Err.Description
    CreateManualScopingTable = False
End Function

'==================== POWER BI METADATA SHEET ====================

' Create the Power BI Metadata sheet with run information
Private Function CreatePowerBIMetadataSheet(ByVal wb As Workbook) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "Creating Power BI Metadata sheet..."

    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet(wb, PBI_METADATA_SHEET)
    ws.Cells.Clear

    ' Add metadata information
    ws.Cells(1, 1).Value = "Metadata Item"
    ws.Cells(1, 2).Value = "Value"

    With ws.Range("A1:B1")
        .Font.Bold = True
        .Interior.Color = RGB(0, 32, 96)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Cells(2, 1).Value = "Tool Version"
    ws.Cells(2, 2).Value = GetToolVersion()

    ws.Cells(3, 1).Value = "Generated Date"
    ws.Cells(3, 2).Value = Now

    ws.Cells(4, 1).Value = "Base Amounts Label"
    ws.Cells(4, 2).Value = g_BaseAmountsLabel

    ws.Cells(5, 1).Value = "Total Segments"
    ws.Cells(5, 2).Value = g_SegmentNames.count

    ws.Cells(6, 1).Value = "Total FSLIs"
    ws.Cells(6, 2).Value = g_FSLIList.count

    ws.Cells(7, 1).Value = "Total Packs"
    ws.Cells(7, 2).Value = g_PackList.count

    ws.Columns.AutoFit

    LogMessage "Power BI Metadata sheet created"

    CreatePowerBIMetadataSheet = True
    Exit Function

ErrorHandler:
    LogMessage "ERROR: Failed to create Power BI Metadata sheet: " & Err.Description
    CreatePowerBIMetadataSheet = False
End Function

'==================== HELPER FUNCTIONS ====================

' Determine FSLI category based on name patterns
Private Function DetermineFSLICategory(ByVal fsli As String) As String
    Dim fsliLower As String
    fsliLower = LCase(fsli)

    If InStr(fsliLower, "revenue") > 0 Or InStr(fsliLower, "income") > 0 Then
        DetermineFSLICategory = "Revenue"
    ElseIf InStr(fsliLower, "expense") > 0 Or InStr(fsliLower, "cost") > 0 Then
        DetermineFSLICategory = "Expense"
    ElseIf InStr(fsliLower, "asset") > 0 Then
        DetermineFSLICategory = "Asset"
    ElseIf InStr(fsliLower, "liability") > 0 Or InStr(fsliLower, "liabilities") > 0 Then
        DetermineFSLICategory = "Liability"
    ElseIf InStr(fsliLower, "equity") > 0 Then
        DetermineFSLICategory = "Equity"
    Else
        DetermineFSLICategory = "Other"
    End If
End Function

' Determine which financial statement the FSLI belongs to
Private Function DetermineFinancialStatement(ByVal fsli As String) As String
    Dim category As String
    category = DetermineFSLICategory(fsli)

    Select Case category
        Case "Revenue", "Expense"
            DetermineFinancialStatement = "Income Statement"
        Case "Asset", "Liability", "Equity"
            DetermineFinancialStatement = "Balance Sheet"
        Case Else
            DetermineFinancialStatement = "Other"
    End Select
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
