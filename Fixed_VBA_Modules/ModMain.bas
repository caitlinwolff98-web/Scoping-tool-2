Attribute VB_Name = "ModMain"
Option Explicit

' ============================================================================
' MODULE: ModMain (FIXED VERSION)
' PURPOSE: Main entry point for the Bidvest Scoping Tool
' DESCRIPTION: Orchestrates the entire process of analyzing consolidation
'              workbooks, categorizing tabs, and creating structured tables
' FIXES APPLIED:
'   1. Added comprehensive status bar updates throughout
'   2. Ensure all tables are formatted as Excel ListObjects
'   3. Add null checks before accessing workbooks
'   4. Fixed encoding issues in messages
'   5. Call table formatting at the end to ensure consistency
' ============================================================================

' Global variables for workbook references
Public g_SourceWorkbook As Workbook
Public g_OutputWorkbook As Workbook
Public g_TabCategories As Object ' Dictionary for tab categorization
Public g_ConsolidatedPackCode As String ' Pack code for consolidated entity
Public g_ConsolidatedPackName As String ' Pack name for consolidated entity

' Main entry point - called when user clicks the button
Public Sub StartScopingTool()
    On Error GoTo ErrorHandler

    Dim workbookName As String
    Dim result As VbMsgBoxResult

    ' Display welcome message (with fixed encoding)
    result = MsgBox("Welcome to the Bidvest Consolidation Scoping Tool!" & vbCrLf & vbCrLf & _
                    "This tool will:" & vbCrLf & _
                    "• Analyze your consolidation workbook" & vbCrLf & _
                    "• Categorize tabs for processing" & vbCrLf & _
                    "• Create structured tables for Power BI" & vbCrLf & _
                    "• Process segment reporting (optional)" & vbCrLf & _
                    "• Generate scoping analysis" & vbCrLf & vbCrLf & _
                    "Click OK to continue or Cancel to exit.", _
                    vbOKCancel + vbInformation, "Bidvest Scoping Tool")

    If result = vbCancel Then
        Application.StatusBar = False
        Exit Sub
    End If

    ' Step 1: Get the workbook name from user
    Application.StatusBar = "Waiting for workbook name..."
    workbookName = GetWorkbookName()
    If workbookName = "" Then
        MsgBox "No workbook name provided. Process cancelled.", vbExclamation
        Application.StatusBar = False
        Exit Sub
    End If

    ' Step 2: Validate and set source workbook
    Application.StatusBar = "Opening source workbook..."
    If Not SetSourceWorkbook(workbookName) Then
        MsgBox "Could not find workbook '" & workbookName & "'. Please ensure it is open.", vbCritical
        Application.StatusBar = False
        Exit Sub
    End If

    ' Step 3: Discover and list all tabs
    Application.StatusBar = "Discovering worksheets..."
    Dim tabList As Collection
    Set tabList = DiscoverTabs()

    If tabList.count = 0 Then
        MsgBox "No tabs found in the workbook.", vbExclamation
        Application.StatusBar = False
        Exit Sub
    End If

    ' Step 4: Categorize tabs
    Application.StatusBar = "Waiting for tab categorization..."
    If Not CategorizeTabs(tabList) Then
        MsgBox "Tab categorization was cancelled. Process terminated.", vbInformation
        Application.StatusBar = False
        Exit Sub
    End If

    ' Step 5: Validate required categories
    Application.StatusBar = "Validating categories..."
    If Not ValidateCategories() Then
        MsgBox "Required tabs are missing. Please ensure all mandatory categories are assigned.", vbCritical
        Application.StatusBar = False
        Exit Sub
    End If

    ' Step 5a: Select consolidated entity (to exclude from scoping)
    Application.StatusBar = "Waiting for consolidated entity selection..."
    If Not SelectConsolidatedEntity() Then
        MsgBox "Consolidated entity selection was cancelled. Process terminated.", vbInformation
        Application.StatusBar = False
        Exit Sub
    End If

    ' Step 5b: Load segment mapping workbook
    Application.StatusBar = "Loading segment mapping workbook..."
    Dim segmentMapping As Object
    Set segmentMapping = ModDataProcessing.LoadSegmentMappingFromWorkbook()

    ' Store in global variable
    ModDataProcessing.SetSegmentMapping segmentMapping
    Application.StatusBar = False

    ' Step 6: Create output workbook for tables
    Application.StatusBar = "Creating output workbook..."
    CreateOutputWorkbook

    ' Step 7: Process data and create tables
    Application.StatusBar = "Processing consolidation data..."
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Call main data processing (status bar updates are inside)
    ModDataProcessing.ProcessConsolidationData

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Step 8: Create interactive dashboard
    Application.StatusBar = "Creating interactive dashboard..."
    ModInteractiveDashboard.CreateInteractiveDashboard
    Application.StatusBar = False

    ' Step 9: Process IAS 8 Segment Reporting Document (optional)
    Application.StatusBar = "Checking for segment reporting document..."
    Dim segmentProcessed As Boolean
    segmentProcessed = ModSegmentAnalysis.ProcessSegmentDocument()
    Application.StatusBar = False

    ' Step 10: CRITICAL - Ensure all tables are properly formatted as Excel ListObjects
    Application.StatusBar = "Finalizing table formatting..."
    ConvertAllTablesToListObjects
    Application.StatusBar = False

    ' Step 11: Save the output workbook
    Application.StatusBar = "Saving output workbook..."
    SaveOutputWorkbook
    Application.StatusBar = False

    ' Step 12: Display completion message
    Dim completionMsg As String
    completionMsg = "Scoping tool completed successfully!" & vbCrLf & vbCrLf & _
                   "Output saved as: " & g_OutputWorkbook.Name & vbCrLf & _
                   "Location: " & g_OutputWorkbook.Path & vbCrLf & vbCrLf & _
                   "Generated assets:" & vbCrLf & _
                   "• Full Input Table (consolidation currency)" & vbCrLf & _
                   "• Full Input Percentage" & vbCrLf & _
                   "• FSLi Key Table" & vbCrLf & _
                   "• Pack Number Company Table" & vbCrLf & _
                   "• Manual Scoping interface" & vbCrLf & _
                   "• Scoping Control Table" & vbCrLf & _
                   "• Analytics Dashboard" & vbCrLf

    ' Add segment tables message if processed
    If segmentProcessed Then
        completionMsg = completionMsg & "• Segment Pack Mapping" & vbCrLf & _
                       "• Segment Summary" & vbCrLf
    End If

    completionMsg = completionMsg & vbCrLf & _
                   "All tables have been converted to Excel ListObjects for easy filtering and analysis!"

    MsgBox completionMsg, vbInformation, "Process Complete"

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical, "Error"
End Sub

' NEW: Convert all data ranges to Excel ListObjects for better usability
Private Sub ConvertAllTablesToListObjects()
    On Error Resume Next

    ' NULL CHECK
    If g_OutputWorkbook Is Nothing Then
        Debug.Print "Error: g_OutputWorkbook is Nothing in ConvertAllTablesToListObjects"
        Exit Sub
    End If

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    Dim tbl As ListObject

    ' List of worksheet names that should have tables
    Dim tableSheets As Variant
    tableSheets = Array("Full Input Table", "Full Input Percentage", _
                        "FSLi Key Table", "Pack Number Company Table", _
                        "Journals Table", "Journals Percentage", _
                        "Full Consol Table", "Full Consol Percentage", _
                        "Discontinued Table", "Discontinued Percentage", _
                        "Scoping Control Table", "Manual Scoping", _
                        "Segment_Pack_Mapping", "Segment_Summary")

    Dim sheetName As Variant
    For Each sheetName In tableSheets
        ' Try to get the worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = g_OutputWorkbook.Worksheets(CStr(sheetName))
        On Error GoTo NextSheet

        If Not ws Is Nothing Then
            ' Check if there's already a ListObject
            If ws.ListObjects.count = 0 Then
                ' Find data range
                lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
                lastCol = ws.Cells(1, ws.columns.count).End(xlToLeft).Column

                If lastRow > 1 And lastCol >= 1 Then
                    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

                    ' Convert to table
                    Set tbl = ModConfig.ConvertRangeToTable(ws, dataRange, CStr(sheetName))

                    If Not tbl Is Nothing Then
                        Debug.Print "Converted " & sheetName & " to ListObject: " & tbl.Name
                    End If
                End If
            Else
                Debug.Print sheetName & " already has a ListObject"
            End If
        End If

NextSheet:
    Next sheetName

    On Error GoTo 0
End Sub

' Categorize tabs using ModTabCategorization
Private Function CategorizeTabs(tabList As Collection) As Boolean
    ' Call categorization module
    CategorizeTabs = ModTabCategorization.CategorizeTabs(tabList)
End Function

' Validate that required categories have been assigned
Private Function ValidateCategories() As Boolean
    On Error Resume Next

    ValidateCategories = False

    ' Check if Input Continuing category exists and has at least one tab
    If Not g_TabCategories.Exists(ModConfig.CAT_INPUT_CONTINUING) Then
        Exit Function
    End If

    If g_TabCategories(ModConfig.CAT_INPUT_CONTINUING).count = 0 Then
        Exit Function
    End If

    ValidateCategories = True
    On Error GoTo 0
End Function

' Select consolidated entity to exclude from scoping
Private Function SelectConsolidatedEntity() As Boolean
    On Error GoTo ErrorHandler

    Dim inputTab As Worksheet
    Dim lastCol As Long
    Dim col As Long
    Dim packNames As String
    Dim packCodes As String
    Dim selection As String
    Dim selectedIndex As Long

    ' Get Input Continuing tab
    Set inputTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_INPUT_CONTINUING)
    If inputTab Is Nothing Then
        SelectConsolidatedEntity = False
        Exit Function
    End If

    ' Build list of packs from row 7 (pack names)
    lastCol = inputTab.Cells(7, inputTab.columns.count).End(xlToLeft).Column

    packNames = "Select the consolidated entity pack:" & vbCrLf & vbCrLf
    Dim packCount As Long
    packCount = 0

    For col = 2 To lastCol
        Dim packNameFull As String
        Dim packCode As String

        packNameFull = Trim(inputTab.Cells(7, col).value)

        If packNameFull <> "" Then
            packCount = packCount + 1

            ' Parse pack code from pack name
            packCode = ModConfig.ParsePackCodeFromName(packNameFull)

            packNames = packNames & packCount & ". " & packNameFull
            If packCode <> "" Then
                packNames = packNames & " [" & packCode & "]"
            End If
            packNames = packNames & vbCrLf
        End If
    Next col

    ' Prompt user to select consolidated entity
    selection = InputBox(packNames & vbCrLf & "Enter the number of the consolidated entity:", _
                        "Select Consolidated Entity", "1")

    If selection = "" Then
        SelectConsolidatedEntity = False
        Exit Function
    End If

    ' Parse selection and extract pack code
    selectedIndex = CLng(selection)
    packCount = 0

    For col = 2 To lastCol
        packNameFull = Trim(inputTab.Cells(7, col).value)

        If packNameFull <> "" Then
            packCount = packCount + 1

            If packCount = selectedIndex Then
                g_ConsolidatedPackName = packNameFull
                g_ConsolidatedPackCode = ModConfig.ParsePackCodeFromName(packNameFull)

                MsgBox "Selected consolidated entity:" & vbCrLf & _
                       "Name: " & g_ConsolidatedPackName & vbCrLf & _
                       "Code: " & g_ConsolidatedPackCode, _
                       vbInformation, "Consolidated Entity Selected"

                SelectConsolidatedEntity = True
                Exit Function
            End If
        End If
    Next col

    SelectConsolidatedEntity = False
    Exit Function

ErrorHandler:
    MsgBox "Error selecting consolidated entity: " & Err.Description, vbCritical
    SelectConsolidatedEntity = False
End Function

' Save output workbook with standardized name
Private Sub SaveOutputWorkbook()
    On Error GoTo ErrorHandler

    ' NULL CHECK
    If g_OutputWorkbook Is Nothing Then
        Debug.Print "Error: g_OutputWorkbook is Nothing in SaveOutputWorkbook"
        Exit Sub
    End If

    Dim savePath As String
    Dim fileName As String
    Dim timestamp As String

    ' Create timestamp for unique filename
    timestamp = Format(Now, "yyyy-mm-dd_hhnnss")

    ' Standard output file name with timestamp
    fileName = "Bidvest_Scoping_Tool_Output_" & timestamp & ".xlsx"

    ' Use the same directory as the source workbook
    If Not g_SourceWorkbook Is Nothing Then
        savePath = g_SourceWorkbook.Path & Application.PathSeparator & fileName
    Else
        ' Fallback to user's documents folder
        savePath = Environ("USERPROFILE") & "\Documents\" & fileName
    End If

    ' Save the workbook
    Application.DisplayAlerts = False
    g_OutputWorkbook.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = True
    ' If save fails, just leave it unsaved for user to manually save
    Debug.Print "Could not auto-save output workbook: " & Err.Description
    MsgBox "Could not auto-save the output workbook." & vbCrLf & _
           "Please save it manually using File > Save As.", vbExclamation
End Sub

' Get workbook name from user
Private Function GetWorkbookName() As String
    Dim userInput As String

    userInput = InputBox( _
        "Please enter the exact name of the consolidation workbook." & vbCrLf & vbCrLf & _
        "Instructions:" & vbCrLf & _
        "1. Open the consolidation workbook" & vbCrLf & _
        "2. Copy the workbook name from the title bar" & vbCrLf & _
        "3. Paste it below (include .xlsx or .xlsm extension)", _
        "Enter Workbook Name", _
        "")

    GetWorkbookName = Trim(userInput)
End Function

' Set the source workbook reference
Private Function SetSourceWorkbook(workbookName As String) As Boolean
    On Error Resume Next

    ' Use centralized function from ModConfig
    Set g_SourceWorkbook = ModConfig.GetWorkbookByName(workbookName)

    SetSourceWorkbook = Not (g_SourceWorkbook Is Nothing)
    On Error GoTo 0
End Function

' Discover all tabs in the source workbook
Private Function DiscoverTabs() As Collection
    Dim tabs As New Collection
    Dim ws As Worksheet

    ' NULL CHECK
    If g_SourceWorkbook Is Nothing Then
        Set DiscoverTabs = tabs
        Exit Function
    End If

    For Each ws In g_SourceWorkbook.Worksheets
        tabs.Add ws.Name
    Next ws

    Set DiscoverTabs = tabs
End Function

' Create the output workbook for generated tables
Private Sub CreateOutputWorkbook()
    Application.StatusBar = "Creating new output workbook..."

    Set g_OutputWorkbook = Workbooks.Add
    g_OutputWorkbook.Worksheets(1).Name = "Control Panel"

    ' Add professional informational sheet
    With g_OutputWorkbook.Worksheets("Control Panel")
        ' Title
        .Range("A1").value = "BIDVEST SCOPING TOOL - OUTPUT WORKBOOK"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Color = RGB(0, 112, 192)

        ' Version and date
        .Range("A2").value = "Version: " & ModConfig.TOOL_VERSION
        .Range("A3").value = "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")

        ' Instructions
        .Range("A5").value = "INSTRUCTIONS"
        .Range("A5").Font.Bold = True
        .Range("A5").Font.Size = 12

        .Range("A6").value = "This workbook contains scoping analysis tables:"
        .Range("A7").value = "• Full Input Table - Consolidated data with all packs"
        .Range("A8").value = "• Full Input Percentage - Percentages relative to consolidated entity"
        .Range("A9").value = "• FSLi Key Table - List of all Financial Statement Line Items"
        .Range("A10").value = "• Pack Number Company Table - Pack codes, names, and segments"
        .Range("A11").value = "• Manual Scoping - Interactive scoping interface"
        .Range("A12").value = "• Scoping Control Table - Summary with amounts and percentages"
        .Range("A13").value = "• Analytics Dashboard - Scoping coverage analysis"

        ' Format
        .Range("A7:A13").IndentLevel = 1
        .columns("A:A").ColumnWidth = 80

        ' Note about tables
        .Range("A15").value = "NOTE: All data ranges have been converted to Excel Tables (ListObjects)"
        .Range("A16").value = "for easy filtering, sorting, and analysis. Use the filter dropdowns in headers."
        .Range("A15:A16").Font.Italic = True
        .Range("A15:A16").Font.Color = RGB(0, 128, 0)
    End With

    Application.StatusBar = False
End Sub
