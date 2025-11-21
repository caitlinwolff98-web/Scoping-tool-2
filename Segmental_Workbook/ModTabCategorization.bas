Attribute VB_Name = "ModTabCategorization"
Option Explicit

'==============================================================================
' MODULE: ModTabCategorization
' PURPOSE: Handle tab categorization and segment name prompting
' DESCRIPTION: Prompts user to categorize tabs as Segment Summary or Segment tabs
'              and suggests tab names as segment names
'==============================================================================

' Main function to categorize all tabs
Public Function CategorizeTabs(ByVal sourceWb As Workbook) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "Starting tab categorization..."

    Dim ws As Worksheet
    Dim tabCategories As Object
    Set tabCategories = CreateDictionary()

    Dim categorizedCount As Long
    categorizedCount = 0

    ' Loop through all worksheets and prompt for categorization
    For Each ws In sourceWb.Worksheets
        If Not IsHiddenSheet(ws) Then
            Dim category As String
            Dim segmentName As String

            ' Prompt user to categorize the tab
            category = PromptForCategory(ws, segmentName)

            If category = "" Then
                ' User cancelled
                ShowWarning "Categorization Cancelled", ERR_CANCELLED
                CategorizeTabs = False
                Exit Function
            End If

            ' Store categorization
            tabCategories.Add ws.Name, Array(category, segmentName)

            ' Track segment names for Segment Tab category
            If category = CAT_SEGMENT_TAB Then
                If Not g_SegmentNames.Exists(ws.Name) Then
                    g_SegmentNames.Add ws.Name, segmentName
                End If
                categorizedCount = categorizedCount + 1
            ElseIf category = CAT_SEGMENT_SUMMARY Then
                categorizedCount = categorizedCount + 1
            End If

            LogMessage "Tab '" & ws.Name & "' categorized as: " & category & IIf(segmentName <> "", " (Segment: " & segmentName & ")", "")
        End If
    Next ws

    ' Validate that we have at least one Segment Summary and one Segment Tab
    Dim hasSummary As Boolean
    Dim hasSegments As Boolean

    hasSummary = False
    hasSegments = False

    Dim key As Variant
    For Each key In tabCategories.Keys
        If tabCategories(key)(0) = CAT_SEGMENT_SUMMARY Then hasSummary = True
        If tabCategories(key)(0) = CAT_SEGMENT_TAB Then hasSegments = True
    Next key

    If Not hasSummary Then
        ShowError "Missing Required Tabs", ERR_NO_SUMMARY_TAB
        CategorizeTabs = False
        Exit Function
    End If

    If Not hasSegments Then
        ShowError "Missing Required Tabs", ERR_NO_SEGMENT_TABS
        CategorizeTabs = False
        Exit Function
    End If

    LogMessage "Tab categorization completed successfully. Categorized " & categorizedCount & " tabs."
    CategorizeTabs = True
    Exit Function

ErrorHandler:
    ShowError "Categorization Error", "An error occurred during tab categorization: " & Err.Description
    CategorizeTabs = False
End Function

' Prompt user to categorize a single tab
Private Function PromptForCategory(ByVal ws As Worksheet, ByRef segmentName As String) As String
    On Error GoTo ErrorHandler

    Dim response As VbMsgBoxResult
    Dim message As String
    Dim suggestedSegmentName As String

    ' Suggest tab name as segment name
    suggestedSegmentName = ws.Name

    message = "How should the tab '" & ws.Name & "' be categorized?" & vbCrLf & vbCrLf
    message = message & "Click:" & vbCrLf
    message = message & "• YES - This is the Segment Summary Tab" & vbCrLf
    message = message & "• NO - This is a Segment Tab" & vbCrLf
    message = message & "• CANCEL - Skip this tab (Uncategorized)"

    response = MsgBox(message, vbYesNoCancel + vbQuestion, "Categorize Tab: " & ws.Name)

    Select Case response
        Case vbYes
            ' Segment Summary Tab
            PromptForCategory = CAT_SEGMENT_SUMMARY
            segmentName = ""

        Case vbNo
            ' Segment Tab - prompt for segment name
            PromptForCategory = CAT_SEGMENT_TAB

            ' Prompt for segment name with suggestion
            segmentName = InputBox( _
                "Enter the Segment Name for tab '" & ws.Name & "':" & vbCrLf & vbCrLf & _
                "This name will be used to identify the segment in reports and analysis.", _
                "Segment Name", _
                suggestedSegmentName)

            ' If user cancels segment name input, use tab name as default
            If Trim(segmentName) = "" Then
                segmentName = suggestedSegmentName
            End If

        Case vbCancel
            ' Uncategorized
            PromptForCategory = CAT_UNCATEGORIZED
            segmentName = ""

    End Select

    Exit Function

ErrorHandler:
    ShowError "Prompt Error", "Error prompting for category: " & Err.Description
    PromptForCategory = ""
    segmentName = ""
End Function

' Check if worksheet is hidden
Private Function IsHiddenSheet(ByVal ws As Worksheet) As Boolean
    IsHiddenSheet = (ws.Visible <> xlSheetVisible)
End Function

' Get all tabs categorized as Segment Summary
Public Function GetSegmentSummaryTabs(ByVal sourceWb As Workbook) As Collection
    Dim tabs As New Collection
    Dim ws As Worksheet

    For Each ws In sourceWb.Worksheets
        ' Check if this tab was categorized as Segment Summary
        ' In practice, we'd store this categorization info
        ' For now, we'll use a naming convention or stored data
    Next ws

    Set GetSegmentSummaryTabs = tabs
End Function

' Get all tabs categorized as Segment Tabs
Public Function GetSegmentTabs(ByVal sourceWb As Workbook) As Collection
    Dim tabs As New Collection
    Dim ws As Worksheet

    ' Return all tabs that are in the g_SegmentNames dictionary
    Dim key As Variant
    For Each key In g_SegmentNames.Keys
        On Error Resume Next
        Set ws = sourceWb.Worksheets(CStr(key))
        If Not ws Is Nothing Then
            tabs.Add ws
        End If
        On Error GoTo 0
    Next key

    Set GetSegmentTabs = tabs
End Function

' Prompt user to identify the Base Amounts label in the Summary tab
Public Function PromptForBaseAmountsLabel(ByVal summaryWs As Worksheet) As String
    On Error GoTo ErrorHandler

    Dim message As String
    Dim response As String
    Dim defaultLabel As String

    defaultLabel = MARKER_BASE_AMOUNTS

    ' Scan Row 7 for potential labels
    Dim colIdx As Long
    Dim potentialLabels As String
    potentialLabels = ""

    For colIdx = 1 To 50 ' Scan first 50 columns
        Dim cellValue As String
        cellValue = SafeTrim(summaryWs.Cells(ROW_BASE_AMOUNTS, colIdx).Value)

        If cellValue <> "" Then
            If potentialLabels <> "" Then potentialLabels = potentialLabels & ", "
            potentialLabels = potentialLabels & cellValue
        End If
    Next colIdx

    message = "In the Segment Summary tab, Row " & ROW_BASE_AMOUNTS & " should contain labels identifying 'Base amounts' columns." & vbCrLf & vbCrLf
    message = message & "Potential labels found in Row " & ROW_BASE_AMOUNTS & ":" & vbCrLf
    message = message & potentialLabels & vbCrLf & vbCrLf
    message = message & "Please enter the exact text that identifies the 'Base amounts' columns:" & vbCrLf
    message = message & "(The tool will look for columns containing this text)"

    response = InputBox(message, "Identify Base Amounts Label", defaultLabel)

    If Trim(response) = "" Then
        ShowWarning "No Label Specified", "Using default label: " & defaultLabel
        response = defaultLabel
    End If

    PromptForBaseAmountsLabel = Trim(response)
    g_BaseAmountsLabel = Trim(response)

    LogMessage "Base amounts label set to: " & g_BaseAmountsLabel

    Exit Function

ErrorHandler:
    ShowError "Prompt Error", "Error prompting for base amounts label: " & Err.Description
    PromptForBaseAmountsLabel = defaultLabel
    g_BaseAmountsLabel = defaultLabel
End Function

' Validate that the base amounts label exists in the summary tab
Public Function ValidateBaseAmountsLabel(ByVal summaryWs As Worksheet, ByVal label As String) As Boolean
    On Error GoTo ErrorHandler

    Dim colIdx As Long
    Dim foundCount As Long
    foundCount = 0

    For colIdx = 1 To 100 ' Scan first 100 columns
        Dim cellValue As String
        cellValue = SafeTrim(summaryWs.Cells(ROW_BASE_AMOUNTS, colIdx).Value)

        If InStr(1, cellValue, label, vbTextCompare) > 0 Then
            foundCount = foundCount + 1
        End If
    Next colIdx

    If foundCount = 0 Then
        ShowError "Validation Error", ERR_BASE_AMOUNTS_NOT_FOUND & vbCrLf & vbCrLf & "Label searched: " & label
        ValidateBaseAmountsLabel = False
    Else
        LogMessage "Found " & foundCount & " columns with base amounts label: " & label
        ValidateBaseAmountsLabel = True
    End If

    Exit Function

ErrorHandler:
    ValidateBaseAmountsLabel = False
End Function

' Get the Segment Summary worksheet
Public Function GetSegmentSummaryWorksheet(ByVal sourceWb As Workbook) As Worksheet
    Dim ws As Worksheet

    ' In a real implementation, we'd store which tab was categorized as Summary
    ' For now, we'll search the g_SegmentNames to find tabs NOT in it
    ' and assume the first one found is the summary

    For Each ws In sourceWb.Worksheets
        If Not ws.Visible = xlSheetHidden Then
            ' Check if this worksheet is NOT in segment names
            If Not g_SegmentNames.Exists(ws.Name) Then
                ' Could be summary tab - would need additional logic
                ' For simplicity, we'll need to store this during categorization
            End If
        End If
    Next ws

    ' Better approach: Store the summary tab name during categorization
    ' For now, return Nothing and let calling code handle
    Set GetSegmentSummaryWorksheet = Nothing
End Function

' Store categorization information for later retrieval
Private Sub StoreCategorization(ByVal tabName As String, ByVal category As String, Optional ByVal segmentName As String = "")
    ' This would store the categorization in a hidden sheet or module-level collection
    ' Implementation depends on persistence requirements
End Sub

' Retrieve stored categorization
Private Function GetStoredCategory(ByVal tabName As String) As String
    ' Retrieve previously stored categorization
    ' Implementation depends on persistence requirements
    GetStoredCategory = ""
End Function
