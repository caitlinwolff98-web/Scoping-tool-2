Attribute VB_Name = "ModSegmentAnalysis"
Option Explicit

' ============================================================================
' MODULE: ModSegmentAnalysis (FIXED VERSION)
' PURPOSE: Handle IAS 8 operating segment analysis and mapping
' DESCRIPTION: Processes segment reporting documents, extracts pack-to-segment
'              mappings, and creates analysis tables for Power BI integration
' FIXES APPLIED:
'   1. Added NULL CHECKS for g_OutputWorkbook (fixes Error 91)
'   2. Fixed encoding issues ("‰Û¢" -> "•")
'   3. Added status bar updates
'   4. Ensure all tables are converted to Excel ListObjects
' ============================================================================

' Global variables for segment analysis
Public g_SegmentWorkbook As Workbook
Public g_SegmentTabCategories As Object ' Dictionary for segment tab categorization

' ============================================================================
' MAIN ENTRY POINT
' ============================================================================

' ProcessSegmentDocument - Main orchestrator for segment reporting analysis
' Returns: True if successful, False if cancelled or error
Public Function ProcessSegmentDocument() As Boolean
    On Error GoTo ErrorHandler

    Dim segmentWorkbookName As String
    Dim result As VbMsgBoxResult

    ' CRITICAL NULL CHECK: Ensure g_OutputWorkbook is initialized
    If g_OutputWorkbook Is Nothing Then
        MsgBox "Error: Output workbook not initialized. Cannot process segment document.", vbCritical
        ProcessSegmentDocument = False
        Exit Function
    End If

    ' Ask user if they want to process segment reporting document
    result = MsgBox("Do you have an IAS 8 Operating Segment reporting document?" & vbCrLf & vbCrLf & _
                    "This optional document allows you to:" & vbCrLf & _
                    "• Map packs to their operating segments" & vbCrLf & _
                    "• Analyze scoping coverage by segment" & vbCrLf & _
                    "• Create segment-level reporting in Power BI" & vbCrLf & vbCrLf & _
                    "Click YES if you have a segment document to process." & vbCrLf & _
                    "Click NO to skip segment analysis.", _
                    vbYesNo + vbQuestion, "Segment Reporting Document")

    If result = vbNo Then
        ProcessSegmentDocument = True ' Skip is not an error
        Exit Function
    End If

    ' Step 1: Get segment workbook name
    Application.StatusBar = "Getting segment workbook..."
    segmentWorkbookName = GetSegmentWorkbookName()
    If segmentWorkbookName = "" Then
        MsgBox "No segment workbook name provided. Skipping segment analysis.", vbInformation
        ProcessSegmentDocument = True ' Skip is not an error
        Application.StatusBar = False
        Exit Function
    End If

    ' Step 2: Validate and set segment workbook reference
    Application.StatusBar = "Validating segment workbook..."
    If Not SetSegmentWorkbook(segmentWorkbookName) Then
        MsgBox "Could not find segment workbook '" & segmentWorkbookName & "'." & vbCrLf & _
               "Please ensure it is open in Excel." & vbCrLf & vbCrLf & _
               "Skipping segment analysis.", vbExclamation
        ProcessSegmentDocument = True ' Skip is not an error
        Application.StatusBar = False
        Exit Function
    End If

    ' Step 3: Discover segment tabs
    Application.StatusBar = "Discovering segment tabs..."
    Dim segmentTabList As Collection
    Set segmentTabList = DiscoverSegmentTabs()

    If segmentTabList.count = 0 Then
        MsgBox "No tabs found in segment workbook. Skipping segment analysis.", vbExclamation
        ProcessSegmentDocument = True
        Application.StatusBar = False
        Exit Function
    End If

    ' Step 4: Categorize segment tabs
    Application.StatusBar = "Categorizing segment tabs..."
    If Not CategorizeSegmentTabs(segmentTabList) Then
        MsgBox "Segment tab categorization was cancelled. Skipping segment analysis.", vbInformation
        ProcessSegmentDocument = True
        Application.StatusBar = False
        Exit Function
    End If

    ' Step 5: Extract segment pack mappings
    Application.StatusBar = "Extracting segment pack mappings..."
    Dim segmentMappings As Collection
    Set segmentMappings = ExtractSegmentPackMappings()

    If segmentMappings.count = 0 Then
        MsgBox "No segment pack mappings could be extracted." & vbCrLf & _
               "Please verify the segment document structure." & vbCrLf & vbCrLf & _
               "Skipping segment analysis.", vbExclamation
        ProcessSegmentDocument = True
        Application.StatusBar = False
        Exit Function
    End If

    ' Step 6: Match segment packs to consolidation packs
    Application.StatusBar = "Matching segment packs to consolidation document..."
    Dim matchedMappings As Collection
    Set matchedMappings = MatchSegmentToConsolidationPacks(segmentMappings)

    If matchedMappings.count = 0 Then
        MsgBox "Could not match any segment packs to consolidation packs." & vbCrLf & _
               "Please verify pack names and codes match between documents." & vbCrLf & vbCrLf & _
               "Skipping segment analysis.", vbExclamation
        ProcessSegmentDocument = True
        Application.StatusBar = False
        Exit Function
    End If

    ' Step 7: Create segment analysis tables in output workbook
    Application.StatusBar = "Creating segment analysis tables..."
    CreateSegmentPackMappingTable matchedMappings
    CreateSegmentSummaryTable matchedMappings

    ' Success
    Application.StatusBar = False
    MsgBox "Segment analysis completed successfully!" & vbCrLf & vbCrLf & _
           "Created tables:" & vbCrLf & _
           "• Segment_Pack_Mapping" & vbCrLf & _
           "• Segment_Summary" & vbCrLf & vbCrLf & _
           "These tables can be imported into Power BI for segment-level scoping analysis.", _
           vbInformation, "Segment Analysis Complete"

    ProcessSegmentDocument = True
    Exit Function

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error processing segment document: " & Err.Description & vbCrLf & vbCrLf & _
           "Segment analysis will be skipped.", vbCritical
    ProcessSegmentDocument = True ' Don't fail entire process
End Function

' ============================================================================
' WORKBOOK MANAGEMENT FUNCTIONS
' ============================================================================

' GetSegmentWorkbookName - Prompt user for segment workbook name
Private Function GetSegmentWorkbookName() As String
    Dim workbookName As String

    workbookName = InputBox("Enter the name of the SEGMENT REPORTING workbook:" & vbCrLf & vbCrLf & _
                           "Example: ""Bidvest_Segment_Reporting_2024.xlsx""" & vbCrLf & vbCrLf & _
                           "IMPORTANT:" & vbCrLf & _
                           "• The workbook must be OPEN in Excel" & vbCrLf & _
                           "• Enter the EXACT name including extension (.xlsx or .xlsm)" & vbCrLf & _
                           "• This is the document showing IAS 8 operating segments", _
                           "Segment Reporting Workbook Name", "")

    GetSegmentWorkbookName = Trim(workbookName)
End Function

' SetSegmentWorkbook - Validate and set reference to segment workbook
Private Function SetSegmentWorkbook(workbookName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim wb As Workbook

    ' Try to find the workbook
    For Each wb In Application.Workbooks
        If wb.Name = workbookName Then
            Set g_SegmentWorkbook = wb
            SetSegmentWorkbook = True
            Exit Function
        End If
    Next wb

    ' Not found
    SetSegmentWorkbook = False
    Exit Function

ErrorHandler:
    SetSegmentWorkbook = False
End Function

' ============================================================================
' TAB DISCOVERY AND CATEGORIZATION
' ============================================================================

' DiscoverSegmentTabs - Get list of all worksheets in segment workbook
Private Function DiscoverSegmentTabs() As Collection
    On Error GoTo ErrorHandler

    Dim tabList As New Collection
    Dim ws As Worksheet

    For Each ws In g_SegmentWorkbook.Worksheets
        tabList.Add ws.Name
    Next ws

    Set DiscoverSegmentTabs = tabList
    Exit Function

ErrorHandler:
    Set DiscoverSegmentTabs = New Collection
End Function

' CategorizeSegmentTabs - Prompt user to categorize each segment tab
Private Function CategorizeSegmentTabs(tabList As Collection) As Boolean
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim tabName As String
    Dim Category As String
    Dim segmentName As String
    Dim msg As String
    Dim categoryDict As Object

    Set g_SegmentTabCategories = CreateObject("Scripting.Dictionary")
    Set categoryDict = CreateObject("Scripting.Dictionary")

    ' Category mapping for segment document
    categoryDict.Add "1", "Segment Tab"
    categoryDict.Add "2", "Segment Summary Tab"
    categoryDict.Add "9", "Uncategorized (Skip)"

    ' Build instruction message (with fixed encoding)
    msg = "SEGMENT TAB CATEGORIZATION" & vbCrLf & vbCrLf
    msg = msg & "For each tab, enter the category number:" & vbCrLf & vbCrLf
    msg = msg & "1 = Segment Tab (contains pack data for a specific segment)" & vbCrLf
    msg = msg & "2 = Segment Summary Tab (summary of all segments)" & vbCrLf
    msg = msg & "9 = Uncategorized (skip this tab)" & vbCrLf & vbCrLf
    msg = msg & "NOTE: For Segment Tabs (1), you will also enter the segment name." & vbCrLf
    msg = msg & "Example segment names: ""Food Services"", ""Freight"", ""Office Products""" & vbCrLf & vbCrLf
    msg = msg & "Click OK to begin categorization."

    MsgBox msg, vbInformation, "Segment Tab Categorization"

    ' Categorize each tab
    For i = 1 To tabList.count
        tabName = tabList(i)

        ' Prompt for category
        Category = InputBox("Tab " & i & " of " & tabList.count & ": """ & tabName & """" & vbCrLf & vbCrLf & _
                           "Enter category number:" & vbCrLf & _
                           "1 = Segment Tab" & vbCrLf & _
                           "2 = Segment Summary" & vbCrLf & _
                           "9 = Uncategorized" & vbCrLf & vbCrLf & _
                           "Enter category:", "Categorize: " & tabName, "1")

        ' Validate category
        If Not categoryDict.Exists(Category) Then
            MsgBox "Invalid category '" & Category & "'. Please enter 1, 2, or 9.", vbExclamation
            i = i - 1 ' Retry this tab
            GoTo NextTab
        End If

        ' If Segment Tab, prompt for segment name
        segmentName = ""
        If Category = "1" Then
            segmentName = InputBox("Enter the SEGMENT NAME for this tab:" & vbCrLf & vbCrLf & _
                                  "Examples:" & vbCrLf & _
                                  "• Food Services" & vbCrLf & _
                                  "• Freight" & vbCrLf & _
                                  "• Office Products" & vbCrLf & _
                                  "• Automotive" & vbCrLf & vbCrLf & _
                                  "This name will be used in segment analysis.", _
                                  "Segment Name for: " & tabName, "")

            If Trim(segmentName) = "" Then
                MsgBox "Segment name is required for Segment Tabs. Please re-enter.", vbExclamation
                i = i - 1 ' Retry this tab
                GoTo NextTab
            End If
        End If

        ' Store categorization (format: "Category|SegmentName")
        g_SegmentTabCategories.Add tabName, Category & "|" & segmentName

NextTab:
    Next i

    CategorizeSegmentTabs = True
    Exit Function

ErrorHandler:
    MsgBox "Error during segment tab categorization: " & Err.Description, vbCritical
    CategorizeSegmentTabs = False
End Function

' ============================================================================
' SEGMENT PACK EXTRACTION
' ============================================================================

' ExtractSegmentPackMappings - Extract pack names/codes from segment tabs
Private Function ExtractSegmentPackMappings() As Collection
    On Error GoTo ErrorHandler

    Dim mappings As New Collection
    Dim tabName As Variant
    Dim categoryInfo As String
    Dim categoryParts() As String
    Dim Category As String
    Dim segmentName As String
    Dim ws As Worksheet

    ' Iterate through categorized tabs
    For Each tabName In g_SegmentTabCategories.Keys
        categoryInfo = g_SegmentTabCategories(tabName)
        categoryParts = Split(categoryInfo, "|")
        Category = categoryParts(0)

        ' Process only Segment Tabs (category 1)
        If Category = "1" Then
            segmentName = categoryParts(1)
            Set ws = g_SegmentWorkbook.Worksheets(CStr(tabName))

            ' Extract packs from this segment tab
            ExtractPacksFromSegmentTab ws, segmentName, mappings
        End If
    Next tabName

    Set ExtractSegmentPackMappings = mappings
    Exit Function

ErrorHandler:
    MsgBox "Error extracting segment pack mappings: " & Err.Description, vbCritical
    Set ExtractSegmentPackMappings = New Collection
End Function

' ExtractPacksFromSegmentTab - Extract pack info from a single segment tab
Private Sub ExtractPacksFromSegmentTab(ws As Worksheet, segmentName As String, mappings As Collection)
    On Error GoTo ErrorHandler

    Dim col As Long
    Dim lastCol As Long
    Dim cellValue As String
    Dim parsedPack As Object
    Dim mapping As Object

    ' Find last column with data in row 7 (pack names are in row 7)
    lastCol = ws.Cells(7, ws.columns.count).End(xlToLeft).Column

    ' Scan row 7 for pack entries
    For col = 1 To lastCol
        cellValue = Trim(ws.Cells(7, col).value)

        ' Skip empty cells
        If cellValue <> "" Then
            ' FIXED: Parse pack code from pack name using new function
            Dim packCode As String
            Dim packName As String

            packCode = ModConfig.ParsePackCodeFromName(cellValue)
            packName = ModConfig.ParsePackNameWithoutCode(cellValue)

            If packCode <> "" Then
                ' Create mapping entry
                Set mapping = CreateObject("Scripting.Dictionary")
                mapping("SegmentName") = segmentName
                mapping("PackNameCode") = cellValue ' Original combined format
                mapping("PackName") = packName
                mapping("PackCode") = packCode
                mapping("ColumnIndex") = col
                mapping("SourceTab") = ws.Name

                ' Add to mappings collection
                mappings.Add mapping
                Debug.Print "Extracted: Segment=" & segmentName & ", Pack=" & packName & ", Code=" & packCode
            End If
        End If
    Next col

    Exit Sub

ErrorHandler:
    ' Log error but continue processing other tabs
    Debug.Print "Error extracting packs from segment tab " & ws.Name & ": " & Err.Description
End Sub

' ============================================================================
' MATCHING LOGIC: Segment Document → Consolidation Document
' ============================================================================

' MatchSegmentToConsolidationPacks - Match segment packs to consolidation packs
Private Function MatchSegmentToConsolidationPacks(segmentMappings As Collection) As Collection
    On Error GoTo ErrorHandler

    Dim matchedMappings As New Collection
    Dim consolidationPacks As Object
    Dim i As Long
    Dim mapping As Object
    Dim matchedPack As Object
    Dim matchCount As Long
    Dim unmatchedCount As Long

    ' Build dictionary of consolidation packs for fast lookup
    Set consolidationPacks = BuildConsolidationPacksDictionary()

    If consolidationPacks.count = 0 Then
        MsgBox "No consolidation packs available for matching." & vbCrLf & _
               "Please ensure consolidation document was processed first.", vbExclamation
        Set MatchSegmentToConsolidationPacks = matchedMappings
        Exit Function
    End If

    ' Match each segment pack to consolidation pack
    matchCount = 0
    unmatchedCount = 0

    For i = 1 To segmentMappings.count
        Set mapping = segmentMappings(i)

        ' Try to find match in consolidation
        Set matchedPack = FindConsolidationPackMatch(mapping, consolidationPacks)

        If Not matchedPack Is Nothing Then
            ' Add matched consolidation pack info to mapping
            mapping("ConsolPackName") = matchedPack("PackName")
            mapping("ConsolPackCode") = matchedPack("PackCode")
            mapping("ConsolDivision") = matchedPack("Division")
            mapping("MatchMethod") = matchedPack("MatchMethod")
            mapping("IsMatched") = True
            matchCount = matchCount + 1
        Else
            ' No match found
            mapping("ConsolPackName") = "[Not Matched]"
            mapping("ConsolPackCode") = "[Not Matched]"
            mapping("ConsolDivision") = ""
            mapping("MatchMethod") = "No Match"
            mapping("IsMatched") = False
            unmatchedCount = unmatchedCount + 1
        End If

        matchedMappings.Add mapping
    Next i

    ' Report matching statistics
    MsgBox "Segment Pack Matching Results:" & vbCrLf & vbCrLf & _
           "Total segment packs: " & segmentMappings.count & vbCrLf & _
           "Successfully matched: " & matchCount & vbCrLf & _
           "Unmatched: " & unmatchedCount & vbCrLf & vbCrLf & _
           IIf(unmatchedCount > 0, "Review unmatched packs in Segment_Pack_Mapping table.", "All packs matched successfully!"), _
           IIf(unmatchedCount = 0, vbInformation, vbExclamation), _
           "Matching Results"

    Set MatchSegmentToConsolidationPacks = matchedMappings
    Exit Function

ErrorHandler:
    MsgBox "Error matching segment packs: " & Err.Description, vbCritical
    Set MatchSegmentToConsolidationPacks = New Collection
End Function

' BuildConsolidationPacksDictionary - Build dictionary of consolidation packs
Private Function BuildConsolidationPacksDictionary() As Object
    On Error GoTo ErrorHandler

    Dim packsDict As Object
    Set packsDict = CreateObject("Scripting.Dictionary")

    ' CRITICAL NULL CHECK: Ensure g_OutputWorkbook is not Nothing
    If g_OutputWorkbook Is Nothing Then
        Debug.Print "ERROR: g_OutputWorkbook is Nothing in BuildConsolidationPacksDictionary"
        Set BuildConsolidationPacksDictionary = packsDict
        Exit Function
    End If

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim packName As String
    Dim packCode As String
    Dim division As String
    Dim packInfo As Object

    ' Find Pack Number Company Table in output workbook
    On Error Resume Next
    Set ws = g_OutputWorkbook.Worksheets("Pack Number Company Table")
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        Debug.Print "Warning: Pack Number Company Table not found"
        Set BuildConsolidationPacksDictionary = packsDict
        Exit Function
    End If

    ' Find last row
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    ' Start from row 2 (skip header)
    For row = 2 To lastRow
        packCode = Trim(ws.Cells(row, 1).value) ' Column A: Pack Code
        packName = Trim(ws.Cells(row, 2).value) ' Column B: Pack Name
        division = Trim(ws.Cells(row, 3).value) ' Column C: Division

        If packCode <> "" Then
            ' Store pack info
            Set packInfo = CreateObject("Scripting.Dictionary")
            packInfo("PackName") = packName
            packInfo("PackCode") = packCode
            packInfo("Division") = division

            ' Add with pack code as key
            If Not packsDict.Exists(packCode) Then
                packsDict.Add packCode, packInfo
            End If
        End If
    Next row

    Set BuildConsolidationPacksDictionary = packsDict
    Exit Function

ErrorHandler:
    Debug.Print "Error in BuildConsolidationPacksDictionary: " & Err.Description
    Set BuildConsolidationPacksDictionary = CreateObject("Scripting.Dictionary")
End Function

' FindConsolidationPackMatch - Find matching consolidation pack
Private Function FindConsolidationPackMatch(segmentMapping As Object, consolidationPacks As Object) As Object
    On Error GoTo ErrorHandler

    Dim packCode As String
    Dim matchedPack As Object

    packCode = segmentMapping("PackCode")

    ' Try exact pack code match
    If consolidationPacks.Exists(packCode) Then
        Set matchedPack = consolidationPacks(packCode)
        matchedPack("MatchMethod") = "Exact Code"
        Set FindConsolidationPackMatch = matchedPack
        Exit Function
    End If

    ' No match found
    Set FindConsolidationPackMatch = Nothing
    Exit Function

ErrorHandler:
    Set FindConsolidationPackMatch = Nothing
End Function

' ============================================================================
' OUTPUT TABLE GENERATION
' ============================================================================

' FIXED: CreateSegmentPackMappingTable - With null check for g_OutputWorkbook
Private Sub CreateSegmentPackMappingTable(matchedMappings As Collection)
    On Error GoTo ErrorHandler

    ' CRITICAL NULL CHECK
    If g_OutputWorkbook Is Nothing Then
        MsgBox "Error: Output workbook not initialized. Cannot create Segment Pack Mapping table.", vbCritical
        Exit Sub
    End If

    Dim ws As Worksheet
    Dim row As Long
    Dim i As Long
    Dim mapping As Object
    Dim tbl As ListObject
    Dim lastRow As Long
    Dim lastCol As Long

    ' Delete existing worksheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    g_OutputWorkbook.Worksheets("Segment_Pack_Mapping").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    ' Create worksheet
    Set ws = g_OutputWorkbook.Worksheets.Add
    ws.Name = "Segment_Pack_Mapping"

    ' Write headers
    ws.Cells(1, 1).value = "Segment Name"
    ws.Cells(1, 2).value = "Pack Name (Segment Doc)"
    ws.Cells(1, 3).value = "Pack Code (Segment Doc)"
    ws.Cells(1, 4).value = "Pack Name (Consol Doc)"
    ws.Cells(1, 5).value = "Pack Code"
    ws.Cells(1, 6).value = "Division"
    ws.Cells(1, 7).value = "Match Status"
    ws.Cells(1, 8).value = "Match Method"
    ws.Cells(1, 9).value = "Source Tab"

    ' Write data
    row = 2
    For i = 1 To matchedMappings.count
        Set mapping = matchedMappings(i)

        ws.Cells(row, 1).value = mapping("SegmentName")
        ws.Cells(row, 2).value = mapping("PackName")
        ws.Cells(row, 3).value = mapping("PackCode")
        ws.Cells(row, 4).value = mapping("ConsolPackName")
        ws.Cells(row, 5).value = mapping("ConsolPackCode")
        ws.Cells(row, 6).value = mapping("ConsolDivision")
        ws.Cells(row, 7).value = IIf(mapping("IsMatched"), "Matched", "Unmatched")
        ws.Cells(row, 8).value = mapping("MatchMethod")
        ws.Cells(row, 9).value = mapping("SourceTab")

        row = row + 1
    Next i

    ' Get dimensions
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    lastCol = 9

    ' Create Excel Table using helper function
    If lastRow > 1 Then
        Set tbl = ModConfig.ConvertRangeToTable(ws, _
                                                  ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), _
                                                  "Segment_Pack_Mapping")
    End If

    ' Auto-fit columns
    ws.columns.AutoFit

    Exit Sub

ErrorHandler:
    MsgBox "Error creating Segment Pack Mapping table: " & Err.Description, vbCritical
End Sub

' FIXED: CreateSegmentSummaryTable - With null check for g_OutputWorkbook
Private Sub CreateSegmentSummaryTable(matchedMappings As Collection)
    On Error GoTo ErrorHandler

    ' CRITICAL NULL CHECK
    If g_OutputWorkbook Is Nothing Then
        MsgBox "Error: Output workbook not initialized. Cannot create Segment Summary table.", vbCritical
        Exit Sub
    End If

    Dim ws As Worksheet
    Dim segmentStats As Object
    Dim mapping As Object
    Dim i As Long
    Dim segmentName As String
    Dim stats As Object
    Dim row As Long
    Dim tbl As ListObject
    Dim lastRow As Long

    ' Build statistics dictionary
    Set segmentStats = CreateObject("Scripting.Dictionary")

    For i = 1 To matchedMappings.count
        Set mapping = matchedMappings(i)
        segmentName = mapping("SegmentName")

        ' Initialize stats for this segment if not exists
        If Not segmentStats.Exists(segmentName) Then
            Set stats = CreateObject("Scripting.Dictionary")
            stats("TotalPacks") = 0
            stats("MatchedPacks") = 0
            stats("UnmatchedPacks") = 0
            segmentStats.Add segmentName, stats
        End If

        ' Update stats
        Set stats = segmentStats(segmentName)
        stats("TotalPacks") = stats("TotalPacks") + 1
        If mapping("IsMatched") Then
            stats("MatchedPacks") = stats("MatchedPacks") + 1
        Else
            stats("UnmatchedPacks") = stats("UnmatchedPacks") + 1
        End If
    Next i

    ' Delete existing worksheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    g_OutputWorkbook.Worksheets("Segment_Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    ' Create worksheet
    Set ws = g_OutputWorkbook.Worksheets.Add
    ws.Name = "Segment_Summary"

    ' Write headers
    ws.Cells(1, 1).value = "Segment Name"
    ws.Cells(1, 2).value = "Total Packs"
    ws.Cells(1, 3).value = "Matched Packs"
    ws.Cells(1, 4).value = "Unmatched Packs"
    ws.Cells(1, 5).value = "Match Rate %"

    ' Write data
    row = 2
    Dim segment As Variant
    For Each segment In segmentStats.Keys
        Set stats = segmentStats(segment)

        ws.Cells(row, 1).value = segment
        ws.Cells(row, 2).value = stats("TotalPacks")
        ws.Cells(row, 3).value = stats("MatchedPacks")
        ws.Cells(row, 4).value = stats("UnmatchedPacks")

        ' Calculate match rate
        If stats("TotalPacks") > 0 Then
            ws.Cells(row, 5).value = stats("MatchedPacks") / stats("TotalPacks")
            ws.Cells(row, 5).NumberFormat = "0.0%"
        Else
            ws.Cells(row, 5).value = 0
        End If

        row = row + 1
    Next segment

    ' Get dimensions
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    ' Create Excel Table using helper function
    If lastRow > 1 Then
        Set tbl = ModConfig.ConvertRangeToTable(ws, _
                                                  ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 5)), _
                                                  "Segment_Summary")
    End If

    ' Auto-fit columns
    ws.columns.AutoFit

    ' Add totals row
    row = lastRow + 2
    ws.Cells(row, 1).value = "TOTAL"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 2).Formula = "=SUM(B2:B" & lastRow & ")"
    ws.Cells(row, 3).Formula = "=SUM(C2:C" & lastRow & ")"
    ws.Cells(row, 4).Formula = "=SUM(D2:D" & lastRow & ")"
    ws.Cells(row, 5).Formula = "=C" & row & "/B" & row
    ws.Cells(row, 5).NumberFormat = "0.0%"

    Exit Sub

ErrorHandler:
    MsgBox "Error creating Segment Summary table: " & Err.Description, vbCritical
End Sub
