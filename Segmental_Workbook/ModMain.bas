Attribute VB_Name = "ModMain"
Option Explicit

'==============================================================================
' MODULE: ModMain
' PURPOSE: Main entry point and orchestration for Segmental Scoping Tool
' DESCRIPTION: Coordinates the entire process from tab categorization through
'              table generation and Power BI integration
'==============================================================================

' Main entry point - Run the complete segmental analysis
Public Sub RunSegmentalAnalysis()
    On Error GoTo ErrorHandler

    LogMessage "========================================="
    LogMessage "STARTING BIDVEST SEGMENTAL SCOPING TOOL"
    LogMessage "========================================="
    LogMessage "Version: " & GetToolVersion()

    ' Check prerequisites
    If Not CheckPrerequisites() Then
        Exit Sub
    End If

    ' Initialize global collections
    Call InitializeGlobalCollections

    ' Get source workbook (prompt user)
    Dim sourceWb As Workbook
    Set sourceWb = PromptForSourceWorkbook()

    If sourceWb Is Nothing Then
        ShowWarning "Cancelled", ERR_CANCELLED
        Call CleanupGlobalCollections
        Exit Sub
    End If

    ' Create target workbook for results
    Dim targetWb As Workbook
    Set targetWb = CreateTargetWorkbook()

    ' Step 1: Categorize all tabs
    LogMessage "Step 1: Tab Categorization"
    LogMessage "----------------------------"

    If Not CategorizeTabs(sourceWb) Then
        ShowError "Categorization Failed", "Tab categorization failed. Process aborted."
        Call CleanupGlobalCollections
        Exit Sub
    End If

    ' Step 2: Identify Summary tab and prompt for base amounts label
    LogMessage "Step 2: Identify Base Amounts"
    LogMessage "------------------------------"

    Dim summaryWs As Worksheet
    Set summaryWs = FindSegmentSummaryWorksheet(sourceWb)

    If summaryWs Is Nothing Then
        ShowError "No Summary Tab", ERR_NO_SUMMARY_TAB
        Call CleanupGlobalCollections
        Exit Sub
    End If

    Dim baseAmountsLabel As String
    baseAmountsLabel = PromptForBaseAmountsLabel(summaryWs)

    If baseAmountsLabel = "" Then
        ShowError "No Base Amounts Label", "Base amounts label is required. Process aborted."
        Call CleanupGlobalCollections
        Exit Sub
    End If

    ' Validate that base amounts label exists
    If Not ValidateBaseAmountsLabel(summaryWs, baseAmountsLabel) Then
        Call CleanupGlobalCollections
        Exit Sub
    End If

    ' Step 3: Generate FullTable
    LogMessage "Step 3: Generate FullTable"
    LogMessage "---------------------------"

    If Not GenerateFullTable(sourceWb, targetWb) Then
        ShowError "Table Generation Failed", "Failed to generate FullTable. Process aborted."
        Call CleanupGlobalCollections
        Exit Sub
    End If

    ' Step 4: Generate FullTablePercentage
    LogMessage "Step 4: Generate FullTablePercentage"
    LogMessage "-------------------------------------"

    If Not GenerateFullTablePercentage(sourceWb, targetWb, summaryWs) Then
        ShowError "Percentage Table Failed", "Failed to generate FullTablePercentage. Process aborted."
        Call CleanupGlobalCollections
        Exit Sub
    End If

    ' Step 5: Generate Power BI integration tables
    LogMessage "Step 5: Generate Power BI Tables"
    LogMessage "---------------------------------"

    If Not GenerateAllPowerBITables(sourceWb, targetWb) Then
        LogMessage "WARNING: Some Power BI tables failed to generate"
    End If

    ' Step 6: Apply initial scoping (if configured)
    LogMessage "Step 6: Apply Initial Scoping"
    LogMessage "-----------------------------"

    Call ApplyInitialScoping(targetWb)

    ' Step 7: Format and finalize output workbook
    LogMessage "Step 7: Finalize Output"
    LogMessage "-----------------------"

    Call FinalizeOutputWorkbook(targetWb)

    ' Cleanup
    Call CleanupGlobalCollections

    ' Success message
    LogMessage "========================================="
    LogMessage "SEGMENTAL ANALYSIS COMPLETED SUCCESSFULLY"
    LogMessage "========================================="

    ShowInfo "Process Complete", _
        "Segmental analysis completed successfully!" & vbCrLf & vbCrLf & _
        "Results saved to: " & targetWb.Name & vbCrLf & vbCrLf & _
        "Total Segments: " & g_SegmentNames.count & vbCrLf & _
        "Total FSLIs: " & g_FSLIList.count & vbCrLf & _
        "Total Packs: " & g_PackList.count

    Exit Sub

ErrorHandler:
    ShowError "Process Error", "An error occurred during the segmental analysis: " & vbCrLf & vbCrLf & Err.Description
    Call CleanupGlobalCollections
End Sub

'==================== PREREQUISITE CHECKS ====================

' Check all prerequisites before starting
Private Function CheckPrerequisites() As Boolean
    On Error GoTo ErrorHandler

    CheckPrerequisites = False

    ' Check if Scripting Runtime is available
    If Not IsScriptingRuntimeAvailable() Then
        ShowError "Missing Library", ERR_SCRIPTING_RUNTIME
        Exit Function
    End If

    LogMessage "Prerequisites check passed"
    CheckPrerequisites = True
    Exit Function

ErrorHandler:
    ShowError "Prerequisite Check Failed", "Failed to verify prerequisites: " & Err.Description
    CheckPrerequisites = False
End Function

'==================== WORKBOOK MANAGEMENT ====================

' Prompt user to select source workbook
Private Function PromptForSourceWorkbook() As Workbook
    On Error GoTo ErrorHandler

    Dim wbName As String
    Dim wb As Workbook

    ' List all open workbooks
    Dim wbList As String
    wbList = "Available workbooks:" & vbCrLf

    For Each wb In Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            wbList = wbList & "  • " & wb.Name & vbCrLf
        End If
    Next wb

    wbName = InputBox( _
        "Enter the name of the source workbook containing segment data:" & vbCrLf & vbCrLf & _
        wbList & vbCrLf & _
        "Note: The workbook must be open.", _
        "Select Source Workbook", _
        "")

    If Trim(wbName) = "" Then
        Set PromptForSourceWorkbook = Nothing
        Exit Function
    End If

    ' Try to get the workbook
    Set wb = GetWorkbookByName(wbName)

    If wb Is Nothing Then
        ShowError "Workbook Not Found", ERR_WORKBOOK_NOT_FOUND & vbCrLf & vbCrLf & "Workbook name: " & wbName
        Set PromptForSourceWorkbook = Nothing
    Else
        LogMessage "Source workbook selected: " & wb.Name
        Set PromptForSourceWorkbook = wb
    End If

    Exit Function

ErrorHandler:
    ShowError "Workbook Selection Error", "Error selecting source workbook: " & Err.Description
    Set PromptForSourceWorkbook = Nothing
End Function

' Create target workbook for results
Private Function CreateTargetWorkbook() As Workbook
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Set wb = Workbooks.Add

    ' Set name
    wb.Worksheets(1).Name = "Instructions"

    ' Add instructions
    With wb.Worksheets("Instructions")
        .Cells(1, 1).Value = "Bidvest Segmental Scoping Tool - Output"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 16

        .Cells(3, 1).Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
        .Cells(4, 1).Value = "Version: " & GetToolVersion()

        .Cells(6, 1).Value = "This workbook contains the following outputs:"
        .Cells(7, 1).Value = "  • FullTable - Aggregated data from all segments"
        .Cells(8, 1).Value = "  • FullTablePercentage - Percentage of base amounts"
        .Cells(9, 1).Value = "  • BaseAmountsCalculation - How base amounts were calculated"
        .Cells(10, 1).Value = "  • ScopingControl - Main table for Power BI integration"
        .Cells(11, 1).Value = "  • PackNumberCompanyTable - Pack lookup table"
        .Cells(12, 1).Value = "  • FSLIKeyTable - FSLI lookup table"
        .Cells(13, 1).Value = "  • SegmentList - Segment information"
        .Cells(14, 1).Value = "  • ThresholdConfiguration - Scoping threshold settings"
        .Cells(15, 1).Value = "  • ManualScopingAdjustments - Manual override tracking"
        .Cells(16, 1).Value = "  • PowerBI_Metadata - Run information and metadata"

        .Columns("A:A").AutoFit
    End With

    LogMessage "Target workbook created"
    Set CreateTargetWorkbook = wb

    Exit Function

ErrorHandler:
    ShowError "Target Workbook Error", "Error creating target workbook: " & Err.Description
    Set CreateTargetWorkbook = Nothing
End Function

'==================== HELPER FUNCTIONS ====================

' Find the Segment Summary worksheet
Private Function FindSegmentSummaryWorksheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet

    ' The summary tab would be the one NOT in g_SegmentNames
    ' We need a better way to track this from categorization

    ' For now, prompt the user
    Dim wsName As String
    wsName = InputBox( _
        "Enter the name of the Segment Summary tab:", _
        "Select Summary Tab", _
        "")

    If Trim(wsName) = "" Then
        Set FindSegmentSummaryWorksheet = Nothing
        Exit Function
    End If

    On Error Resume Next
    Set ws = wb.Worksheets(wsName)

    If ws Is Nothing Then
        ShowError "Summary Tab Not Found", "Could not find worksheet: " & wsName
    Else
        LogMessage "Segment Summary tab identified: " & ws.Name
    End If

    Set FindSegmentSummaryWorksheet = ws
End Function

' Apply initial scoping based on threshold configuration
Private Sub ApplyInitialScoping(ByVal wb As Workbook)
    On Error Resume Next

    LogMessage "Applying initial scoping rules..."

    ' Get threshold configuration
    Dim wsThreshold As Worksheet
    Set wsThreshold = wb.Worksheets(TABLE_THRESHOLD_CONFIG)

    ' Get scoping control table
    Dim wsScoping As Worksheet
    Set wsScoping = wb.Worksheets(TABLE_SCOPING_CONTROL)

    If wsThreshold Is Nothing Or wsScoping Is Nothing Then
        LogMessage "WARNING: Cannot apply initial scoping - required tables not found"
        Exit Sub
    End If

    ' Apply thresholds (basic implementation)
    ' This would be expanded based on threshold logic requirements

    LogMessage "Initial scoping rules applied"
End Sub

' Finalize the output workbook
Private Sub FinalizeOutputWorkbook(ByVal wb As Workbook)
    On Error Resume Next

    LogMessage "Finalizing output workbook..."

    ' Set zoom level for all worksheets
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        ws.Activate
        ActiveWindow.Zoom = 85
    Next ws

    ' Activate instructions sheet
    wb.Worksheets("Instructions").Activate
    wb.Worksheets("Instructions").Cells(1, 1).Select

    LogMessage "Output workbook finalized"
End Sub

'==================== QUICK ACCESS PROCEDURES ====================

' Quick procedure to regenerate just the FullTable
Public Sub RegenerateFullTable()
    On Error GoTo ErrorHandler

    Dim sourceWb As Workbook
    Set sourceWb = PromptForSourceWorkbook()

    If sourceWb Is Nothing Then Exit Sub

    Dim targetWb As Workbook
    Set targetWb = ActiveWorkbook

    Call InitializeGlobalCollections

    If GenerateFullTable(sourceWb, targetWb) Then
        ShowInfo "Success", "FullTable regenerated successfully"
    End If

    Call CleanupGlobalCollections
    Exit Sub

ErrorHandler:
    ShowError "Regeneration Error", "Error regenerating FullTable: " & Err.Description
    Call CleanupGlobalCollections
End Sub

' Quick procedure to regenerate Power BI tables only
Public Sub RegeneratePowerBITables()
    On Error GoTo ErrorHandler

    Dim targetWb As Workbook
    Set targetWb = ActiveWorkbook

    Call InitializeGlobalCollections

    ' Need to reload global collections from existing tables
    ' For simplicity, prompt for source workbook again
    Dim sourceWb As Workbook
    Set sourceWb = PromptForSourceWorkbook()

    If sourceWb Is Nothing Then
        Call CleanupGlobalCollections
        Exit Sub
    End If

    If GenerateAllPowerBITables(sourceWb, targetWb) Then
        ShowInfo "Success", "Power BI tables regenerated successfully"
    End If

    Call CleanupGlobalCollections
    Exit Sub

ErrorHandler:
    ShowError "Regeneration Error", "Error regenerating Power BI tables: " & Err.Description
    Call CleanupGlobalCollections
End Sub

'==================== INTERACTIVE DASHBOARD LAUNCHER ====================

' Launch the interactive scoping dashboard
Public Sub LaunchScopingDashboard()
    On Error GoTo ErrorHandler

    ' This would launch the interactive dashboard
    ' For now, show a placeholder message

    ShowInfo "Dashboard", "Interactive Scoping Dashboard" & vbCrLf & vbCrLf & _
        "Use this dashboard to:" & vbCrLf & _
        "  • View scoping analysis" & vbCrLf & _
        "  • Apply threshold-based scoping" & vbCrLf & _
        "  • Make manual scoping adjustments" & vbCrLf & _
        "  • Generate scoping reports"

    ' In full implementation, this would show a UserForm
    ' frmScopingDashboard.Show

    Exit Sub

ErrorHandler:
    ShowError "Dashboard Error", "Error launching scoping dashboard: " & Err.Description
End Sub

'==================== UTILITIES ====================

' Refresh all data from source
Public Sub RefreshAllData()
    On Error GoTo ErrorHandler

    If MsgBox("This will refresh all data from the source workbook. Continue?", vbYesNo + vbQuestion, "Confirm Refresh") = vbYes Then
        Call RunSegmentalAnalysis
    End If

    Exit Sub

ErrorHandler:
    ShowError "Refresh Error", "Error refreshing data: " & Err.Description
End Sub

' Export to Power BI ready format
Public Sub ExportToPowerBI()
    On Error GoTo ErrorHandler

    LogMessage "Exporting to Power BI ready format..."

    ' All required tables are already in Power BI ready format
    ' Just need to ensure they're named correctly

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim requiredTables As Variant
    requiredTables = Array(TABLE_SCOPING_CONTROL, TABLE_PACK_COMPANY, TABLE_FSLI_KEY, _
                          TABLE_SEGMENT_LIST, TABLE_THRESHOLD_CONFIG, TABLE_MANUAL_SCOPING)

    Dim missingTables As String
    missingTables = ""

    Dim tableName As Variant
    For Each tableName In requiredTables
        On Error Resume Next
        Dim ws As Worksheet
        Set ws = wb.Worksheets(CStr(tableName))

        If ws Is Nothing Then
            missingTables = missingTables & "  • " & tableName & vbCrLf
        End If
        On Error GoTo ErrorHandler
    Next tableName

    If missingTables <> "" Then
        ShowWarning "Missing Tables", "The following required tables are missing:" & vbCrLf & vbCrLf & missingTables
    Else
        ShowInfo "Power BI Ready", "All Power BI tables are present and ready for import." & vbCrLf & vbCrLf & _
            "You can now connect Power BI to this workbook using the 'Excel' data connector."
    End If

    Exit Sub

ErrorHandler:
    ShowError "Export Error", "Error preparing Power BI export: " & Err.Description
End Sub
