Attribute VB_Name = "ModDataProcessing"
Option Explicit

' ============================================================================
' MODULE: ModDataProcessing (FIXED VERSION)
' PURPOSE: Process consolidation data and analyze structure
' FIXES APPLIED:
'   1. Parse pack codes from END of pack names (format "PackName XX-####")
'   2. NO LONGER read from row 8 - parse from row 7 pack names
'   3. Properly exclude "NOTES" section
'   4. Exclude statement headers (INCOME STATEMENT, BALANCE SHEET)
'   5. Include ALL FSLIs including Net Revenue
'   6. Track segment/division for each pack
'   7. Capture indentation hierarchy correctly
'   8. Add status bar updates for user feedback
' ============================================================================

' Global variable for segment mapping
Private g_SegmentMapping As Object  ' Dictionary: PackCode -> Segment

' Load segment mapping from segment workbook
Public Function LoadSegmentMappingFromWorkbook() As Object
    On Error GoTo ErrorHandler

    Dim segmentWb As Workbook
    Dim segmentWs As Worksheet
    Dim segmentPath As String
    Dim lastRow As Long
    Dim row As Long
    Dim packCode As String
    Dim packName As String
    Dim segment As String
    Dim division As String
    Dim segmentDict As Object

    Set segmentDict = CreateObject("Scripting.Dictionary")

    ' Prompt user for segment workbook
    MsgBox "Please select the SEGMENT WORKBOOK that contains pack-to-segment mapping.", vbInformation, "Select Segment Workbook"

    segmentPath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx; *.xlsm; *.xls), *.xlsx; *.xlsm; *.xls", _
        title:="Select Segment Workbook")

    If segmentPath = "False" Then
        MsgBox "No segment workbook selected. Segments will be determined from tab names.", vbExclamation
        Set LoadSegmentMappingFromWorkbook = segmentDict
        Exit Function
    End If

    ' Open segment workbook
    Application.ScreenUpdating = False
    Application.StatusBar = "Opening segment workbook..."
    Set segmentWb = Workbooks.Open(segmentPath, ReadOnly:=True)

    ' Assume first sheet has the mapping
    Set segmentWs = segmentWb.Worksheets(1)

    ' Use default columns: A=PackCode, B=PackName, C=Segment, D=Division
    lastRow = segmentWs.Cells(segmentWs.Rows.count, 1).End(xlUp).row

    Application.StatusBar = "Loading segment mappings..."

    For row = 2 To lastRow  ' Assume row 1 is header
        packCode = Trim(segmentWs.Cells(row, 1).value)  ' Column A
        packName = Trim(segmentWs.Cells(row, 2).value)  ' Column B
        segment = Trim(segmentWs.Cells(row, 3).value)   ' Column C

        If segmentWs.Cells(row, 4).value <> "" Then
            division = Trim(segmentWs.Cells(row, 4).value)  ' Column D
        Else
            division = segment  ' Use segment as division if not specified
        End If

        If packCode <> "" And segment <> "" Then
            ' Store both segment and division
            Dim packInfo As Object
            Set packInfo = CreateObject("Scripting.Dictionary")
            packInfo("Segment") = segment
            packInfo("Division") = division
            packInfo("PackName") = packName

            segmentDict(packCode) = packInfo
            Debug.Print "Loaded mapping: " & packCode & " (" & packName & ") -> Segment: " & segment & ", Division: " & division
        End If
    Next row

    ' Close segment workbook
    segmentWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Loaded " & segmentDict.count & " pack-to-segment mappings from segment workbook.", vbInformation, "Segment Mapping Loaded"

    Set LoadSegmentMappingFromWorkbook = segmentDict
    Exit Function

ErrorHandler:
    If Not segmentWb Is Nothing Then
        segmentWb.Close SaveChanges:=False
    End If
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error loading segment workbook: " & Err.Description, vbCritical
    Set LoadSegmentMappingFromWorkbook = CreateObject("Scripting.Dictionary")
End Function

' Main processing orchestrator
Public Sub ProcessConsolidationData()
    On Error GoTo ErrorHandler

    Dim inputTab As Worksheet
    Dim discontinuedTab As Worksheet
    Dim journalsTab As Worksheet
    Dim consoleTab As Worksheet

    ' Get required tabs
    Application.StatusBar = "Finding Input Continuing tab..."
    Set inputTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_INPUT_CONTINUING)

    If inputTab Is Nothing Then
        MsgBox "Could not find Input Continuing tab. Process cannot continue.", vbCritical
        Application.StatusBar = False
        Exit Sub
    End If

    ' Process Input Continuing tab
    Application.StatusBar = "Processing Input Continuing tab..."
    ProcessInputTab inputTab

    ' Process other tabs if they exist
    Set discontinuedTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_DISCONTINUED)
    If Not discontinuedTab Is Nothing Then
        Application.StatusBar = "Processing Discontinued tab..."
        ProcessDiscontinuedTab discontinuedTab
    End If

    Set journalsTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_JOURNALS_CONTINUING)
    If Not journalsTab Is Nothing Then
        Application.StatusBar = "Processing Journals tab..."
        ProcessJournalsTab journalsTab
    End If

    Set consoleTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_CONSOLE_CONTINUING)
    If Not consoleTab Is Nothing Then
        Application.StatusBar = "Processing Consol tab..."
        ProcessConsolTab consoleTab
    End If

    ' Create supporting tables
    Application.StatusBar = "Creating FSLi Key Table..."
    CreateFSLiKeyTable

    Application.StatusBar = "Creating Pack Number Company Table..."
    CreatePackNumberCompanyTable

    Application.StatusBar = "Creating Percentage Tables..."
    CreatePercentageTables

    Application.StatusBar = False

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error in data processing: " & Err.Description, vbCritical
End Sub

' Process Input Continuing tab
Private Sub ProcessInputTab(ws As Worksheet)
    On Error GoTo ErrorHandler

    Dim lastCol As Long
    Dim lastRow As Long
    Dim columns As Collection
    Dim fsliList As Collection
    Dim selectedColumnType As String

    ' Step 1: Unmerge all cells
    Application.StatusBar = "Unmerging cells..."
    ws.Cells.UnMerge

    ' NEW: Load segment mapping from segment workbook FIRST
    If g_SegmentMapping Is Nothing Then
        Application.StatusBar = "Loading segment mappings..."
        Set g_SegmentMapping = LoadSegmentMappingFromWorkbook()
    End If

    ' Step 2: Detect columns and get user selection
    Application.StatusBar = "Detecting columns..."
    Set columns = DetectColumns(ws)
    selectedColumnType = PromptColumnSelection(columns)

    If selectedColumnType = "" Then
        MsgBox "No column type selected. Skipping Input tab.", vbExclamation
        Application.StatusBar = False
        Exit Sub
    End If

    ' Step 3: Analyze FSLi structure - FIXED to properly exclude Notes
    Application.StatusBar = "Analyzing FSLI structure..."
    Set fsliList = AnalyzeFSLiStructure(ws, selectedColumnType)

    ' Step 4: Create Full Input Table
    Application.StatusBar = "Creating Full Input Table..."
    CreateFullInputTable ws, columns, fsliList, selectedColumnType

    Application.StatusBar = False

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error processing Input tab: " & Err.Description, vbCritical
End Sub

' CRITICAL FIX: Detect columns - NOW PARSES pack codes from pack names
Private Function DetectColumns(ws As Worksheet) As Collection
    On Error GoTo ErrorHandler

    Dim columns As New Collection
    Dim col As Long
    Dim lastCol As Long
    Dim cellValue As String
    Dim colInfo As Object
    Dim packNameFull As String
    Dim packName As String
    Dim packCode As String

    ' Find last column with data in row 6
    lastCol = ws.Cells(6, ws.columns.count).End(xlToLeft).Column

    ' Analyze row 6 for column types
    For col = 1 To lastCol
        cellValue = Trim(ws.Cells(6, col).value)

        If cellValue <> "" Then
            Set colInfo = CreateObject("Scripting.Dictionary")
            colInfo("ColumnIndex") = col

            ' Determine column type
            If InStr(1, cellValue, "original", vbTextCompare) > 0 And _
               InStr(1, cellValue, "entity currency", vbTextCompare) > 0 Then
                colInfo("ColumnType") = "Original/Entity"
            ElseIf InStr(1, cellValue, "consolidation", vbTextCompare) > 0 And _
                   InStr(1, cellValue, "consolidation currency", vbTextCompare) > 0 Then
                colInfo("ColumnType") = "Consolidation/Consolidation"
            Else
                colInfo("ColumnType") = "Other"
            End If

            ' *** CRITICAL FIX: Get pack name from row 7 and PARSE pack code from it ***
            packNameFull = ""
            packName = ""
            packCode = ""

            If ws.Cells(7, col).value <> "" Then
                packNameFull = Trim(ws.Cells(7, col).value)

                ' Parse pack code from end of pack name using new function
                packCode = ModConfig.ParsePackCodeFromName(packNameFull)

                ' Get clean pack name without code
                If packCode <> "" Then
                    packName = ModConfig.ParsePackNameWithoutCode(packNameFull)
                Else
                    ' No code found, use full name as pack name
                    packName = packNameFull
                End If
            End If

            colInfo("PackNameFull") = packNameFull  ' Store original
            colInfo("PackName") = packName          ' Store without code
            colInfo("PackCode") = packCode          ' Store parsed code

            Debug.Print "Col " & col & ": PackName='" & packName & "', PackCode='" & packCode & "'"

            ' Get segment name from segment mapping
            If packCode <> "" Then
                colInfo("SegmentName") = GetSegmentForPackCode(packCode)
                colInfo("DivisionName") = GetDivisionForPackCode(packCode)
            Else
                colInfo("SegmentName") = "Unknown"
                colInfo("DivisionName") = "Unknown"
            End If

            columns.Add colInfo
        End If
    Next col

    Set DetectColumns = columns
    Exit Function

ErrorHandler:
    MsgBox "Error detecting columns: " & Err.Description, vbCritical
    Set DetectColumns = New Collection
End Function

' Public function to set segment mapping
Public Sub SetSegmentMapping(segmentDict As Object)
    Set g_SegmentMapping = segmentDict
    Debug.Print "Segment mapping set with " & segmentDict.count & " entries"
End Sub

' Get segment for a specific pack code from loaded mapping
Private Function GetSegmentForPackCode(packCode As String) As String
    On Error Resume Next

    If g_SegmentMapping Is Nothing Then
        GetSegmentForPackCode = "Unknown"
        Exit Function
    End If

    If g_SegmentMapping.Exists(packCode) Then
        Dim packInfo As Object
        Set packInfo = g_SegmentMapping(packCode)

        ' Check if it's a dictionary (new format) or string (old format)
        If TypeName(packInfo) = "Dictionary" Then
            GetSegmentForPackCode = packInfo("Segment")
        Else
            GetSegmentForPackCode = CStr(packInfo)
        End If
    Else
        GetSegmentForPackCode = "Unknown"
        Debug.Print "Warning: Pack code " & packCode & " not found in segment mapping"
    End If

    On Error GoTo 0
End Function

' Get division for a specific pack code from loaded mapping
Private Function GetDivisionForPackCode(packCode As String) As String
    On Error Resume Next

    If g_SegmentMapping Is Nothing Then
        GetDivisionForPackCode = "Unknown"
        Exit Function
    End If

    If g_SegmentMapping.Exists(packCode) Then
        Dim packInfo As Object
        Set packInfo = g_SegmentMapping(packCode)

        ' Check if it's a dictionary (new format) or string (old format)
        If TypeName(packInfo) = "Dictionary" Then
            If packInfo.Exists("Division") Then
                GetDivisionForPackCode = packInfo("Division")
            Else
                GetDivisionForPackCode = packInfo("Segment")
            End If
        Else
            GetDivisionForPackCode = CStr(packInfo)
        End If
    Else
        GetDivisionForPackCode = "Unknown"
        Debug.Print "Warning: Pack code " & packCode & " not found in segment mapping"
    End If

    On Error GoTo 0
End Function

' Prompt user to select column type
Private Function PromptColumnSelection(columns As Collection) As String
    Dim originalCount As Long
    Dim consolidationCount As Long
    Dim i As Long
    Dim colInfo As Object
    Dim msg As String
    Dim response As VbMsgBoxResult

    ' Count column types
    For i = 1 To columns.count
        Set colInfo = columns(i)
        If colInfo("ColumnType") = "Original/Entity" Then
            originalCount = originalCount + 1
        ElseIf colInfo("ColumnType") = "Consolidation/Consolidation" Then
            consolidationCount = consolidationCount + 1
        End If
    Next i

    ' Build message
    msg = "Column types detected in row 6:" & vbCrLf & vbCrLf

    If originalCount > 0 Then
        msg = msg & "- Original/Entity Currency: " & originalCount & " columns" & vbCrLf
    End If

    If consolidationCount > 0 Then
        msg = msg & "- Consolidation/Consolidation Currency: " & consolidationCount & " columns" & vbCrLf
    End If

    msg = msg & vbCrLf & "Which columns do you want to use?" & vbCrLf & vbCrLf
    msg = msg & "Click YES for Consolidation/Consolidation Currency (recommended)" & vbCrLf
    msg = msg & "Click NO for Original/Entity Currency"

    response = MsgBox(msg, vbYesNoCancel + vbQuestion, "Select Column Type")

    If response = vbYes Then
        PromptColumnSelection = "Consolidation/Consolidation"
    ElseIf response = vbNo Then
        PromptColumnSelection = "Original/Entity"
    Else
        PromptColumnSelection = ""
    End If
End Function

' FIXED: Analyze FSLi structure - properly exclude Notes and statement headers
Private Function AnalyzeFSLiStructure(ws As Worksheet, columnType As String) As Collection
    On Error GoTo ErrorHandler

    Dim fsliList As New Collection
    Dim row As Long
    Dim lastRow As Long
    Dim fsliName As String
    Dim fsliInfo As Object
    Dim currentStatement As String
    Dim notesStartRow As Long
    Dim indent As Double

    currentStatement = ""
    notesStartRow = 0

    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).row

    ' CRITICAL FIX: Find where "NOTES" starts
    For row = 9 To lastRow
        fsliName = Trim(ws.Cells(row, 2).value)
        If UCase(fsliName) = "NOTES" Then
            notesStartRow = row
            Debug.Print "Notes section found at row " & row
            Exit For
        End If
    Next row

    ' If Notes found, adjust lastRow
    If notesStartRow > 0 Then
        lastRow = notesStartRow - 1
        Debug.Print "Processing rows 9 to " & lastRow & " (excluding Notes)"
    End If

    ' Start from row 9 (after headers)
    For row = 9 To lastRow
        fsliName = Trim(ws.Cells(row, 2).value)

        ' Skip empty rows
        If fsliName = "" Then
            GoTo NextRow
        End If

        ' FIXED: Skip statement headers (but track statement type)
        If UCase(fsliName) = "INCOME STATEMENT" Then
            currentStatement = "Income Statement"
            Debug.Print "Found Income Statement header at row " & row & " - SKIPPING"
            GoTo NextRow
        ElseIf UCase(fsliName) = "BALANCE SHEET" Or _
               UCase(fsliName) = "STATEMENT OF FINANCIAL POSITION" Then
            currentStatement = "Balance Sheet"
            Debug.Print "Found Balance Sheet header at row " & row & " - SKIPPING"
            GoTo NextRow
        End If

        ' Skip other pure statement headers
        If IsStatementHeader(fsliName) Then
            Debug.Print "Skipping statement header at row " & row & ": " & fsliName
            GoTo NextRow
        End If

        ' Get indentation
        indent = DetectIndentationLevel(ws, row, 2)

        ' Create FSLi info dictionary
        Set fsliInfo = CreateObject("Scripting.Dictionary")
        fsliInfo("FSLiName") = fsliName
        fsliInfo("RowIndex") = row
        fsliInfo("StatementType") = currentStatement

        ' Detect if it's a total or subtotal based on name AND indentation
        fsliInfo("IsTotal") = (InStr(1, fsliName, "total", vbTextCompare) > 0)
        fsliInfo("IsSubtotal") = (InStr(1, fsliName, "subtotal", vbTextCompare) > 0) Or _
                                 (InStr(1, fsliName, "sub-total", vbTextCompare) > 0)

        ' FIXED: Use indentation to determine hierarchy
        fsliInfo("Level") = indent

        ' Classify by indentation
        If indent = 0 Then
            fsliInfo("HierarchyType") = "Total/Main"
        ElseIf indent = 1 Then
            fsliInfo("HierarchyType") = "Subtotal/Sub-item"
        ElseIf indent >= 3 Then
            fsliInfo("HierarchyType") = "Sub-subtotal/Detail"
        Else
            fsliInfo("HierarchyType") = "Item"
        End If

        ' Add to collection
        fsliList.Add fsliInfo

        Debug.Print "Added FSLI at row " & row & ": " & fsliName & " (Indent=" & indent & ", Type=" & fsliInfo("HierarchyType") & ")"

NextRow:
    Next row

    Debug.Print "Total FSLIs collected: " & fsliList.count

    Set AnalyzeFSLiStructure = fsliList
    Exit Function

ErrorHandler:
    MsgBox "Error analyzing FSLi structure: " & Err.Description, vbCritical
    Set AnalyzeFSLiStructure = New Collection
End Function

' Detect indentation level of a cell
Private Function DetectIndentationLevel(ws As Worksheet, row As Long, col As Long) As Double
    On Error Resume Next

    Dim cell As Range
    Set cell = ws.Cells(row, col)

    ' Try to get IndentLevel property first (this is the standard property)
    DetectIndentationLevel = cell.IndentLevel

    ' If that fails or returns 0, try Alignment.Indent (note: capital A)
    If Err.Number <> 0 Or DetectIndentationLevel = 0 Then
        Err.Clear
        DetectIndentationLevel = cell.Alignment.indent
    End If

    ' Default to 0 if still error
    If Err.Number <> 0 Then
        DetectIndentationLevel = 0
        Err.Clear
    End If

    On Error GoTo 0
End Function

' Check if a line is a statement header (not an actual FSLI)
Public Function IsStatementHeader(fsliName As String) As Boolean
    Dim upperName As String
    upperName = UCase(Trim(fsliName))

    ' Exact matches for statement headers
    IsStatementHeader = False

    If upperName = "INCOME STATEMENT" Or _
       upperName = "BALANCE SHEET" Or _
       upperName = "STATEMENT OF FINANCIAL POSITION" Or _
       upperName = "STATEMENT OF PROFIT OR LOSS" Or _
       upperName = "STATEMENT OF COMPREHENSIVE INCOME" Or _
       upperName = "CASH FLOW STATEMENT" Or _
       upperName = "STATEMENT OF CASH FLOWS" Or _
       upperName = "STATEMENT OF CHANGES IN EQUITY" Then
        IsStatementHeader = True
    End If
End Function

' Check if entire row is empty
Private Function IsRowEmpty(ws As Worksheet, row As Long) As Boolean
    Dim col As Long
    Dim lastCol As Long

    lastCol = ws.Cells(row, ws.columns.count).End(xlToLeft).Column

    For col = 1 To lastCol
        If ws.Cells(row, col).value <> "" Then
            IsRowEmpty = False
            Exit Function
        End If
    Next col

    IsRowEmpty = True
End Function

' Create Full Input Table
Private Sub CreateFullInputTable(sourceWs As Worksheet, columns As Collection, _
                                 fsliList As Collection, columnType As String)
    On Error GoTo ErrorHandler

    CreateGenericTable sourceWs, columns, fsliList, columnType, "Full Input Table"

    Exit Sub

ErrorHandler:
    MsgBox "Error creating Full Input Table: " & Err.Description, vbCritical
End Sub

' Process Discontinued tab
Private Sub ProcessDiscontinuedTab(ws As Worksheet)
    On Error GoTo ErrorHandler

    Dim columns As Collection
    Dim fsliList As Collection
    Dim selectedColumnType As String

    ws.Cells.UnMerge
    Set columns = DetectColumns(ws)
    selectedColumnType = "Consolidation/Consolidation"

    ' Check if we have columns of this type
    Dim hasColumns As Boolean
    hasColumns = False
    Dim i As Long
    Dim colInfo As Object
    For i = 1 To columns.count
        Set colInfo = columns(i)
        If colInfo("ColumnType") = selectedColumnType Then
            hasColumns = True
            Exit For
        End If
    Next i

    If Not hasColumns Then
        selectedColumnType = "Original/Entity"
    End If

    Set fsliList = AnalyzeFSLiStructure(ws, selectedColumnType)
    CreateDiscontinuedTable ws, columns, fsliList, selectedColumnType

    Exit Sub

ErrorHandler:
    MsgBox "Error processing Discontinued tab: " & Err.Description, vbCritical
End Sub

' Process Journals tab
Private Sub ProcessJournalsTab(ws As Worksheet)
    On Error GoTo ErrorHandler

    Dim columns As Collection
    Dim fsliList As Collection
    Dim selectedColumnType As String

    ws.Cells.UnMerge
    Set columns = DetectColumns(ws)
    selectedColumnType = "Consolidation/Consolidation"

    Dim hasColumns As Boolean
    hasColumns = False
    Dim i As Long
    Dim colInfo As Object
    For i = 1 To columns.count
        Set colInfo = columns(i)
        If colInfo("ColumnType") = selectedColumnType Then
            hasColumns = True
            Exit For
        End If
    Next i

    If Not hasColumns Then
        selectedColumnType = "Original/Entity"
    End If

    Set fsliList = AnalyzeFSLiStructure(ws, selectedColumnType)
    CreateJournalsTable ws, columns, fsliList, selectedColumnType

    Exit Sub

ErrorHandler:
    MsgBox "Error processing Journals tab: " & Err.Description, vbCritical
End Sub

' Process Consol tab
Private Sub ProcessConsolTab(ws As Worksheet)
    On Error GoTo ErrorHandler

    Dim columns As Collection
    Dim fsliList As Collection
    Dim selectedColumnType As String

    ws.Cells.UnMerge
    Set columns = DetectColumns(ws)
    selectedColumnType = "Consolidation/Consolidation"

    Dim hasColumns As Boolean
    hasColumns = False
    Dim i As Long
    Dim colInfo As Object
    For i = 1 To columns.count
        Set colInfo = columns(i)
        If colInfo("ColumnType") = selectedColumnType Then
            hasColumns = True
            Exit For
        End If
    Next i

    If Not hasColumns Then
        selectedColumnType = "Original/Entity"
    End If

    Set fsliList = AnalyzeFSLiStructure(ws, selectedColumnType)
    CreateConsolTable ws, columns, fsliList, selectedColumnType

    Exit Sub

ErrorHandler:
    MsgBox "Error processing Consol tab: " & Err.Description, vbCritical
End Sub

' Create Journals Table
Private Sub CreateJournalsTable(sourceWs As Worksheet, columns As Collection, _
                                fsliList As Collection, columnType As String)
    On Error GoTo ErrorHandler

    CreateGenericTable sourceWs, columns, fsliList, columnType, "Journals Table"

    Exit Sub

ErrorHandler:
    MsgBox "Error creating Journals Table: " & Err.Description, vbCritical
End Sub

' Create Consol Table
Private Sub CreateConsolTable(sourceWs As Worksheet, columns As Collection, _
                               fsliList As Collection, columnType As String)
    On Error GoTo ErrorHandler

    CreateGenericTable sourceWs, columns, fsliList, columnType, "Full Consol Table"

    Exit Sub

ErrorHandler:
    MsgBox "Error creating Consol Table: " & Err.Description, vbCritical
End Sub

' Create Discontinued Table
Private Sub CreateDiscontinuedTable(sourceWs As Worksheet, columns As Collection, _
                                    fsliList As Collection, columnType As String)
    On Error GoTo ErrorHandler

    CreateGenericTable sourceWs, columns, fsliList, columnType, "Discontinued Table"

    Exit Sub

ErrorHandler:
    MsgBox "Error creating Discontinued Table: " & Err.Description, vbCritical
End Sub

' FIXED: Create Generic Table - now includes segment name AND converts to Excel table
Private Sub CreateGenericTable(sourceWs As Worksheet, columns As Collection, _
                               fsliList As Collection, columnType As String, tableName As String)
    On Error GoTo ErrorHandler

    Dim outputWs As Worksheet
    Dim outRow As Long
    Dim outCol As Long
    Dim i As Long
    Dim j As Long
    Dim colInfo As Object
    Dim fsliInfo As Object
    Dim packDict As Object  ' To store pack details
    Dim packName As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim tbl As ListObject
    Dim packKey As String

    ' Check if g_OutputWorkbook is set (CRITICAL NULL CHECK)
    If g_OutputWorkbook Is Nothing Then
        MsgBox "Output workbook not initialized. Cannot create " & tableName, vbCritical
        Exit Sub
    End If

    ' Check if worksheet already exists and delete it
    On Error Resume Next
    Application.DisplayAlerts = False
    g_OutputWorkbook.Worksheets(tableName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    ' Create output worksheet
    Set outputWs = g_OutputWorkbook.Worksheets.Add
    outputWs.Name = tableName

    ' Use dictionary to store pack details including segment
    Set packDict = CreateObject("Scripting.Dictionary")

    For i = 1 To columns.count
        Set colInfo = columns(i)
        If colInfo("ColumnType") = columnType And colInfo("PackName") <> "" Then
            packName = colInfo("PackName")
            packKey = colInfo("PackCode") & "|" & packName

            If Not packDict.Exists(packKey) Then
                Dim packDetails As Object
                Set packDetails = CreateObject("Scripting.Dictionary")
                packDetails("PackName") = colInfo("PackName")
                packDetails("PackCode") = colInfo("PackCode")
                packDetails("SegmentName") = colInfo("SegmentName")
                packDetails("DivisionName") = colInfo("DivisionName")
                packDetails("ColumnIndex") = colInfo("ColumnIndex")

                packDict.Add packKey, packDetails
            End If
        End If
    Next i

    ' Write headers including segment and division
    outputWs.Cells(1, 1).value = "Pack Name"
    outputWs.Cells(1, 2).value = "Pack Code"
    outputWs.Cells(1, 3).value = "Segment"
    outputWs.Cells(1, 4).value = "Division"

    outCol = 5
    For i = 1 To fsliList.count
        Set fsliInfo = fsliList(i)
        outputWs.Cells(1, outCol).value = fsliInfo("FSLiName")
        outCol = outCol + 1
    Next i

    ' Write pack names and data
    outRow = 2
    Dim packKeys As Variant
    packKeys = packDict.Keys

    For i = 0 To UBound(packKeys)
        Dim packDetail As Object
        Set packDetail = packDict(packKeys(i))

        outputWs.Cells(outRow, 1).value = packDetail("PackName")
        outputWs.Cells(outRow, 2).value = packDetail("PackCode")
        outputWs.Cells(outRow, 3).value = packDetail("SegmentName")
        outputWs.Cells(outRow, 4).value = packDetail("DivisionName")

        Debug.Print "Writing pack: " & packDetail("PackCode") & " | Segment: " & packDetail("SegmentName")

        ' For each FSLi, find the value
        outCol = 5
        For j = 1 To fsliList.count
            Set fsliInfo = fsliList(j)

            ' Copy value from source
            Dim packCol As Long
            packCol = packDetail("ColumnIndex")

            If packCol > 0 Then
                outputWs.Cells(outRow, outCol).value = sourceWs.Cells(fsliInfo("RowIndex"), packCol).value
            End If

            outCol = outCol + 1
        Next j

        outRow = outRow + 1
    Next i

    ' Get dimensions for table
    lastRow = outputWs.Cells(outputWs.Rows.count, 1).End(xlUp).row
    lastCol = outputWs.Cells(1, outputWs.columns.count).End(xlToLeft).Column

    ' Create actual Excel Table using new helper function
    If lastRow > 1 And lastCol > 1 Then
        Set tbl = ModConfig.ConvertRangeToTable(outputWs, _
                                                 outputWs.Range(outputWs.Cells(1, 1), outputWs.Cells(lastRow, lastCol)), _
                                                 tableName)
    End If

    ' Auto-fit columns
    outputWs.columns.AutoFit

    Exit Sub

ErrorHandler:
    MsgBox "Error creating " & tableName & ": " & Err.Description, vbCritical
End Sub

' Create FSLi Key Table
Private Sub CreateFSLiKeyTable()
    ' Call the implementation in ModTableGeneration
    ModTableGeneration.CreateFSLiKeyTable
End Sub

' Create Pack Number Company Table
Private Sub CreatePackNumberCompanyTable()
    ' Call the implementation in ModTableGeneration
    ModTableGeneration.CreatePackNumberCompanyTable
End Sub

' Create Percentage Tables
Private Sub CreatePercentageTables()
    ' Call the implementation in ModTableGeneration
    ModTableGeneration.CreatePercentageTables
End Sub
