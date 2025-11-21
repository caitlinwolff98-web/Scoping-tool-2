Attribute VB_Name = "ModInteractiveDashboard"
Option Explicit

'==============================================================================
' MODULE: ModInteractiveDashboard
' PURPOSE: Interactive dashboard for scoping analysis and adjustments
' DESCRIPTION: Provides dynamic dashboard functionality for threshold scoping,
'              manual adjustments, and real-time scoping analysis
'==============================================================================

'==================== DASHBOARD INITIALIZATION ====================

' Initialize and display the interactive scoping dashboard
Public Sub ShowInteractiveDashboard()
    On Error GoTo ErrorHandler

    LogMessage "Initializing Interactive Dashboard..."

    ' Verify required tables exist
    If Not ValidateDashboardPrerequisites() Then
        Exit Sub
    End If

    ' Create or activate dashboard worksheet
    Dim wsDashboard As Worksheet
    Set wsDashboard = CreateDashboardWorksheet()

    ' Build dashboard controls and displays
    Call BuildDashboardLayout(wsDashboard)

    ' Initialize dashboard data
    Call RefreshDashboardData(wsDashboard)

    LogMessage "Interactive Dashboard loaded successfully"
    wsDashboard.Activate

    Exit Sub

ErrorHandler:
    ShowError "Dashboard Error", "Error initializing dashboard: " & Err.Description
End Sub

'==================== DASHBOARD LAYOUT ====================

' Create the dashboard worksheet structure
Private Function CreateDashboardWorksheet() As Worksheet
    On Error Resume Next

    Dim ws As Worksheet
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    ' Try to get existing dashboard
    Set ws = wb.Worksheets("Scoping Dashboard")

    If ws Is Nothing Then
        ' Create new dashboard
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        ws.Name = "Scoping Dashboard"
    Else
        ' Clear existing dashboard
        ws.Cells.Clear
    End If

    Set CreateDashboardWorksheet = ws
End Function

' Build the dashboard layout with controls and displays
Private Sub BuildDashboardLayout(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler

    ' Title Section
    With ws.Range("A1:J1")
        .Merge
        .Value = "SEGMENTAL SCOPING DASHBOARD"
        .Font.Bold = True
        .Font.Size = 16
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With

    ' Summary Statistics Section
    ws.Cells(3, 1).Value = "SUMMARY STATISTICS"
    ws.Cells(3, 1).Font.Bold = True
    ws.Cells(3, 1).Font.Size = 12

    ws.Cells(4, 1).Value = "Total Items:"
    ws.Cells(4, 2).Value = "=COUNTA(ScopingControl!A:A)-1"

    ws.Cells(5, 1).Value = "Items Scoped In:"
    ws.Cells(5, 2).Value = "=COUNTIF(ScopingControl!G:G,""Yes"")"

    ws.Cells(6, 1).Value = "Items Scoped Out:"
    ws.Cells(6, 2).Value = "=COUNTIF(ScopingControl!G:G,""No"")"

    ws.Cells(7, 1).Value = "Scoping Coverage %:"
    ws.Cells(7, 2).Value = "=IF(B4=0,0,B5/B4)"
    ws.Cells(7, 2).NumberFormat = "0.0%"

    ws.Cells(9, 1).Value = "Total Segments:"
    ws.Cells(9, 2).Value = "=COUNTA(SegmentList!A:A)-1"

    ws.Cells(10, 1).Value = "Total FSLIs:"
    ws.Cells(10, 2).Value = "=COUNTA(FSLIKeyTable!A:A)-1"

    ws.Cells(11, 1).Value = "Total Packs:"
    ws.Cells(11, 2).Value = "=COUNTA(PackNumberCompanyTable!A:A)-1"

    ' Threshold Configuration Section
    ws.Cells(3, 5).Value = "THRESHOLD CONFIGURATION"
    ws.Cells(3, 5).Font.Bold = True
    ws.Cells(3, 5).Font.Size = 12

    ws.Cells(4, 5).Value = "Materiality %:"
    ws.Cells(4, 6).Value = 5
    ws.Cells(4, 6).NumberFormat = "0.0%"
    ws.Cells(4, 6).Interior.Color = RGB(255, 255, 200)

    ws.Cells(5, 5).Value = "Performance %:"
    ws.Cells(5, 6).Value = 2
    ws.Cells(5, 6).NumberFormat = "0.0%"
    ws.Cells(5, 6).Interior.Color = RGB(255, 255, 200)

    ws.Cells(6, 5).Value = "Trivial %:"
    ws.Cells(6, 6).Value = 0.5
    ws.Cells(6, 6).NumberFormat = "0.0%"
    ws.Cells(6, 6).Interior.Color = RGB(255, 255, 200)

    ' Action Buttons Section (will be created as shapes/buttons)
    ws.Cells(3, 8).Value = "ACTIONS"
    ws.Cells(3, 8).Font.Bold = True
    ws.Cells(3, 8).Font.Size = 12

    ws.Cells(4, 8).Value = "[Apply Thresholds]"
    ws.Cells(5, 8).Value = "[Clear Scoping]"
    ws.Cells(6, 8).Value = "[Export Results]"
    ws.Cells(7, 8).Value = "[Refresh Data]"

    ' Scoping by Segment Section
    ws.Cells(13, 1).Value = "SCOPING BY SEGMENT"
    ws.Cells(13, 1).Font.Bold = True
    ws.Cells(13, 1).Font.Size = 12

    ws.Cells(14, 1).Value = "Segment"
    ws.Cells(14, 2).Value = "Total Items"
    ws.Cells(14, 3).Value = "Scoped In"
    ws.Cells(14, 4).Value = "Scoped Out"
    ws.Cells(14, 5).Value = "Coverage %"

    With ws.Range("A14:E14")
        .Font.Bold = True
        .Interior.Color = RGB(146, 208, 80)
    End With

    ' Scoping by FSLI Section
    ws.Cells(13, 7).Value = "SCOPING BY FSLI"
    ws.Cells(13, 7).Font.Bold = True
    ws.Cells(13, 7).Font.Size = 12

    ws.Cells(14, 7).Value = "FSLI"
    ws.Cells(14, 8).Value = "Total Items"
    ws.Cells(14, 9).Value = "Scoped In"
    ws.Cells(14, 10).Value = "Coverage %"

    With ws.Range("G14:J14")
        .Font.Bold = True
        .Interior.Color = RGB(255, 192, 0)
    End With

    ' Format columns
    ws.Columns("A:J").AutoFit

    Exit Sub

ErrorHandler:
    LogMessage "ERROR: Failed to build dashboard layout: " & Err.Description
End Sub

'==================== DASHBOARD DATA REFRESH ====================

' Refresh dashboard data and calculations
Public Sub RefreshDashboardData(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler

    LogMessage "Refreshing dashboard data..."

    ' Populate scoping by segment
    Call PopulateScopingBySegment(ws)

    ' Populate scoping by FSLI
    Call PopulateScopingByFSLI(ws)

    ' Recalculate all formulas
    ws.Calculate

    LogMessage "Dashboard data refreshed"

    Exit Sub

ErrorHandler:
    LogMessage "ERROR: Failed to refresh dashboard data: " & Err.Description
End Sub

' Populate scoping statistics by segment
Private Sub PopulateScopingBySegment(ByVal ws As Worksheet)
    On Error Resume Next

    Dim wsScoping As Worksheet
    Set wsScoping = ActiveWorkbook.Worksheets("ScopingControl")

    If wsScoping Is Nothing Then Exit Sub

    ' Get unique segments
    Dim segments As Object
    Set segments = CreateDictionary()

    Dim lastRow As Long
    Dim rowIdx As Long

    lastRow = wsScoping.Cells(wsScoping.Rows.count, 3).End(xlUp).Row

    For rowIdx = 2 To lastRow
        Dim segment As String
        segment = SafeTrim(wsScoping.Cells(rowIdx, 3).Value)

        If segment <> "" And Not segments.Exists(segment) Then
            segments.Add segment, 0
        End If
    Next rowIdx

    ' Populate segment statistics
    Dim outRow As Long
    outRow = 15

    Dim seg As Variant
    For Each seg In segments.Keys
        ws.Cells(outRow, 1).Value = seg
        ws.Cells(outRow, 2).Value = "=COUNTIF(ScopingControl!C:C,""" & seg & """)"
        ws.Cells(outRow, 3).Value = "=COUNTIFS(ScopingControl!C:C,""" & seg & """,ScopingControl!G:G,""Yes"")"
        ws.Cells(outRow, 4).Value = "=COUNTIFS(ScopingControl!C:C,""" & seg & """,ScopingControl!G:G,""No"")"
        ws.Cells(outRow, 5).Value = "=IF(B" & outRow & "=0,0,C" & outRow & "/B" & outRow & ")"
        ws.Cells(outRow, 5).NumberFormat = "0.0%"

        outRow = outRow + 1
    Next seg
End Sub

' Populate scoping statistics by FSLI
Private Sub PopulateScopingByFSLI(ByVal ws As Worksheet)
    On Error Resume Next

    Dim wsScoping As Worksheet
    Set wsScoping = ActiveWorkbook.Worksheets("ScopingControl")

    If wsScoping Is Nothing Then Exit Sub

    ' Get unique FSLIs
    Dim fslis As Object
    Set fslis = CreateDictionary()

    Dim lastRow As Long
    Dim rowIdx As Long

    lastRow = wsScoping.Cells(wsScoping.Rows.count, 4).End(xlUp).Row

    For rowIdx = 2 To lastRow
        Dim fsli As String
        fsli = SafeTrim(wsScoping.Cells(rowIdx, 4).Value)

        If fsli <> "" And Not fslis.Exists(fsli) Then
            fslis.Add fsli, 0
        End If
    Next rowIdx

    ' Populate FSLI statistics (top 10 only)
    Dim outRow As Long
    outRow = 15

    Dim f As Variant
    Dim count As Long
    count = 0

    For Each f In fslis.Keys
        If count >= 10 Then Exit For ' Limit to top 10

        ws.Cells(outRow, 7).Value = f
        ws.Cells(outRow, 8).Value = "=COUNTIF(ScopingControl!D:D,""" & f & """)"
        ws.Cells(outRow, 9).Value = "=COUNTIFS(ScopingControl!D:D,""" & f & """,ScopingControl!G:G,""Yes"")"
        ws.Cells(outRow, 10).Value = "=IF(H" & outRow & "=0,0,I" & outRow & "/H" & outRow & ")"
        ws.Cells(outRow, 10).NumberFormat = "0.0%"

        outRow = outRow + 1
        count = count + 1
    Next f
End Sub

'==================== THRESHOLD SCOPING ACTIONS ====================

' Apply threshold-based scoping
Public Sub ApplyThresholdScoping()
    On Error GoTo ErrorHandler

    LogMessage "Applying threshold-based scoping..."

    Dim wsScoping As Worksheet
    Set wsScoping = ActiveWorkbook.Worksheets("ScopingControl")

    If wsScoping Is Nothing Then
        ShowError "Missing Table", "ScopingControl table not found"
        Exit Sub
    End If

    ' Get threshold values from dashboard
    Dim wsDashboard As Worksheet
    Set wsDashboard = ActiveWorkbook.Worksheets("Scoping Dashboard")

    Dim materialityThreshold As Double
    Dim performanceThreshold As Double

    If Not wsDashboard Is Nothing Then
        materialityThreshold = wsDashboard.Cells(4, 6).Value
        performanceThreshold = wsDashboard.Cells(5, 6).Value
    Else
        materialityThreshold = 0.05 ' 5%
        performanceThreshold = 0.02 ' 2%
    End If

    ' Apply thresholds
    Dim lastRow As Long
    Dim rowIdx As Long

    lastRow = wsScoping.Cells(wsScoping.Rows.count, 1).End(xlUp).Row

    Dim scopedCount As Long
    scopedCount = 0

    For rowIdx = 2 To lastRow
        Dim percentage As Variant
        percentage = wsScoping.Cells(rowIdx, 6).Value ' Percentage column

        If IsValidNumber(percentage) Then
            Dim pct As Double
            pct = CDbl(percentage)

            ' Scope in if >= performance threshold
            If pct >= performanceThreshold Then
                wsScoping.Cells(rowIdx, 7).Value = "Yes" ' Is Scoped
                wsScoping.Cells(rowIdx, 8).Value = "Threshold"
                wsScoping.Cells(rowIdx, 9).Value = performanceThreshold & "%"
                scopedCount = scopedCount + 1
            Else
                wsScoping.Cells(rowIdx, 7).Value = "No"
                wsScoping.Cells(rowIdx, 8).Value = "Below Threshold"
                wsScoping.Cells(rowIdx, 9).Value = ""
            End If
        End If
    Next rowIdx

    LogMessage "Threshold scoping applied. Scoped in " & scopedCount & " items"

    ShowInfo "Scoping Applied", "Threshold-based scoping applied successfully." & vbCrLf & vbCrLf & _
        "Items scoped in: " & scopedCount & vbCrLf & _
        "Performance threshold: " & Format(performanceThreshold, "0.0%")

    ' Refresh dashboard if it exists
    If Not wsDashboard Is Nothing Then
        Call RefreshDashboardData(wsDashboard)
    End If

    Exit Sub

ErrorHandler:
    ShowError "Scoping Error", "Error applying threshold scoping: " & Err.Description
End Sub

' Clear all scoping decisions
Public Sub ClearAllScoping()
    On Error GoTo ErrorHandler

    If MsgBox("This will clear all scoping decisions. Continue?", vbYesNo + vbQuestion, "Confirm Clear") <> vbYes Then
        Exit Sub
    End If

    Dim wsScoping As Worksheet
    Set wsScoping = ActiveWorkbook.Worksheets("ScopingControl")

    If wsScoping Is Nothing Then Exit Sub

    Dim lastRow As Long
    Dim rowIdx As Long

    lastRow = wsScoping.Cells(wsScoping.Rows.count, 1).End(xlUp).Row

    For rowIdx = 2 To lastRow
        wsScoping.Cells(rowIdx, 7).Value = "No" ' Is Scoped
        wsScoping.Cells(rowIdx, 8).Value = "None" ' Scoping Method
        wsScoping.Cells(rowIdx, 9).Value = "" ' Threshold Level
        wsScoping.Cells(rowIdx, 10).Value = "No" ' Manual Override
    Next rowIdx

    ShowInfo "Cleared", "All scoping decisions have been cleared"

    ' Refresh dashboard
    Dim wsDashboard As Worksheet
    Set wsDashboard = ActiveWorkbook.Worksheets("Scoping Dashboard")
    If Not wsDashboard Is Nothing Then
        Call RefreshDashboardData(wsDashboard)
    End If

    Exit Sub

ErrorHandler:
    ShowError "Clear Error", "Error clearing scoping: " & Err.Description
End Sub

'==================== VALIDATION ====================

' Validate that all required tables exist for dashboard
Private Function ValidateDashboardPrerequisites() As Boolean
    On Error Resume Next

    ValidateDashboardPrerequisites = False

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim requiredTables As Variant
    requiredTables = Array("ScopingControl", "SegmentList", "FSLIKeyTable", "PackNumberCompanyTable")

    Dim missingTables As String
    missingTables = ""

    Dim tableName As Variant
    For Each tableName In requiredTables
        Dim ws As Worksheet
        Set ws = Nothing
        Set ws = wb.Worksheets(CStr(tableName))

        If ws Is Nothing Then
            missingTables = missingTables & "  • " & tableName & vbCrLf
        End If
    Next tableName

    If missingTables <> "" Then
        ShowError "Missing Tables", "The following required tables are missing:" & vbCrLf & vbCrLf & missingTables & vbCrLf & _
            "Please run the complete analysis first."
        Exit Function
    End If

    ValidateDashboardPrerequisites = True
End Function
