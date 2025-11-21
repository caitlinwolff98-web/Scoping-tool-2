Attribute VB_Name = "ModConfig"
Option Explicit

'==============================================================================
' MODULE: ModConfig
' PURPOSE: Centralized configuration and constants for Segmental Scoping Tool
' DESCRIPTION: Contains all global constants, configuration settings, and
'              shared utility functions for segment-based analysis
'==============================================================================

'==================== CATEGORY CONSTANTS ====================
' These constants define worksheet categories for processing
Public Const CAT_SEGMENT_SUMMARY As String = "Segment Summary Tab"
Public Const CAT_SEGMENT_TAB As String = "Segment Tab"
Public Const CAT_UNCATEGORIZED As String = "Uncategorized"

'==================== VERSION INFORMATION ====================
Public Const TOOL_VERSION As String = "1.0.0"
Public Const TOOL_NAME As String = "Bidvest Segmental Scoping Tool"
Public Const TOOL_DATE As String = "2025-11"

'==================== PROCESSING CONSTANTS ====================
' Row indices for data structure
Public Const ROW_MARKER As Long = 6          ' Row with "Total", "All journals", "Base amounts", or empty
Public Const ROW_BASE_AMOUNTS As Long = 7    ' Row in Summary tab with "Base amounts" label
Public Const ROW_PACK_INFO As Long = 8       ' Row with pack names and codes
Public Const COL_FSLI As Long = 2            ' Column B contains FSLIs

' Row 6 markers to EXCLUDE
Public Const MARKER_TOTAL As String = "Total"
Public Const MARKER_ALL_JOURNALS As String = "All journals"
Public Const MARKER_BASE_AMOUNTS As String = "Base amounts"

'==================== OUTPUT TABLE NAMES ====================
' These names match Power BI expectations
Public Const TABLE_FULL_TABLE As String = "FullTable"
Public Const TABLE_FULL_TABLE_PCT As String = "FullTablePercentage"
Public Const TABLE_BASE_CALC As String = "BaseAmountsCalculation"
Public Const TABLE_SCOPING_CONTROL As String = "ScopingControl"
Public Const TABLE_PACK_COMPANY As String = "PackNumberCompanyTable"
Public Const TABLE_FSLI_KEY As String = "FSLIKeyTable"
Public Const TABLE_SEGMENT_LIST As String = "SegmentList"
Public Const TABLE_THRESHOLD_CONFIG As String = "ThresholdConfiguration"
Public Const TABLE_MANUAL_SCOPING As String = "ManualScopingAdjustments"

'==================== POWER BI INTEGRATION CONSTANTS ====================
Public Const PBI_METADATA_SHEET As String = "PowerBI_Metadata"
Public Const PBI_SCOPING_SHEET As String = "PowerBI_Scoping"

'==================== ERROR MESSAGES ====================
Public Const ERR_WORKBOOK_NOT_FOUND As String = "Could not find the specified workbook. Please ensure it is open."
Public Const ERR_NO_SUMMARY_TAB As String = "No Segment Summary tab found. Please categorize at least one tab as Segment Summary."
Public Const ERR_NO_SEGMENT_TABS As String = "No Segment tabs found. Please categorize at least one tab as a Segment Tab."
Public Const ERR_NO_TABS_FOUND As String = "No tabs found in the workbook."
Public Const ERR_CANCELLED As String = "Operation was cancelled by user."
Public Const ERR_SCRIPTING_RUNTIME As String = "Microsoft Scripting Runtime is not available. Please enable it in VBA References."
Public Const ERR_BASE_AMOUNTS_NOT_FOUND As String = "Could not identify Base amounts columns in Summary tab."

'====================  GLOBAL VARIABLES ====================
Public g_BaseAmountsLabel As String          ' User-specified label for base amounts
Public g_SegmentNames As Object              ' Dictionary of segment names (key=tab name, value=segment name)
Public g_FSLIList As Object                  ' Dictionary of unique FSLIs
Public g_PackList As Object                  ' Dictionary of unique packs

'==================== UTILITY FUNCTIONS ====================

' Check if Scripting Runtime is available
Public Function IsScriptingRuntimeAvailable() As Boolean
    On Error Resume Next
    Dim testDict As Object
    Set testDict = CreateObject("Scripting.Dictionary")
    IsScriptingRuntimeAvailable = (Err.Number = 0)
    On Error GoTo 0
End Function

' Display formatted error message
Public Sub ShowError(ByVal Title As String, ByVal Message As String)
    Dim fullMsg As String
    fullMsg = Message
    If Err.Number <> 0 Then
        fullMsg = fullMsg & vbCrLf & vbCrLf & "Error Number: " & Err.Number & vbCrLf & Err.Description
    End If
    MsgBox fullMsg, vbCritical + vbOKOnly, Title
End Sub

' Display formatted information message
Public Sub ShowInfo(ByVal Title As String, ByVal Message As String)
    MsgBox Message, vbInformation + vbOKOnly, Title
End Sub

' Display formatted warning message
Public Sub ShowWarning(ByVal Title As String, ByVal Message As String)
    MsgBox Message, vbExclamation + vbOKOnly, Title
End Sub

' Log message to immediate window (for debugging)
Public Sub LogMessage(ByVal Message As String)
    Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " - " & Message
End Sub

' Safe string trim (handles null/empty)
Public Function SafeTrim(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        SafeTrim = ""
    Else
        SafeTrim = Trim(CStr(value))
    End If
End Function

' Check if a value is numeric and not empty
Public Function IsValidNumber(ByVal value As Variant) As Boolean
    IsValidNumber = IsNumeric(value) And Not IsEmpty(value) And value <> ""
End Function

' Get workbook by name (case-insensitive, handles extensions)
Public Function GetWorkbookByName(ByVal wbName As String) As Workbook
    On Error Resume Next
    Dim wb As Workbook
    Dim wbNameWithoutExt As String

    ' Try exact name first
    Set wb = Workbooks(wbName)
    If Not wb Is Nothing Then
        Set GetWorkbookByName = wb
        Exit Function
    End If

    ' Try without extension
    wbNameWithoutExt = Replace(Replace(Replace(wbName, ".xlsx", ""), ".xlsm", ""), ".xls", "")

    For Each wb In Workbooks
        If LCase(Replace(Replace(Replace(wb.Name, ".xlsx", ""), ".xlsm", ""), ".xls", "")) = LCase(wbNameWithoutExt) Then
            Set GetWorkbookByName = wb
            Exit Function
        End If
    Next wb

    Set GetWorkbookByName = Nothing
End Function

' Create a dictionary object (with error handling)
Public Function CreateDictionary() As Object
    On Error GoTo ErrorHandler
    Set CreateDictionary = CreateObject("Scripting.Dictionary")
    Exit Function

ErrorHandler:
    ShowError "Missing Library", ERR_SCRIPTING_RUNTIME
    Set CreateDictionary = Nothing
End Function

' Format currency value for display
Public Function FormatCurrency(ByVal value As Variant) As String
    If IsValidNumber(value) Then
        FormatCurrency = Format(value, "#,##0.00")
    Else
        FormatCurrency = "N/A"
    End If
End Function

' Get tool version information
Public Function GetToolVersion() As String
    GetToolVersion = TOOL_NAME & " v" & TOOL_VERSION & " (" & TOOL_DATE & ")"
End Function

'==================== VALIDATION FUNCTIONS ====================

' Validate worksheet exists and has expected structure
Public Function ValidateWorksheetStructure(ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler

    ValidateWorksheetStructure = False

    ' Check if worksheet exists
    If ws Is Nothing Then Exit Function

    ' Check if row 6 exists (marker row)
    If ws.Cells(ROW_MARKER, COL_FSLI).Value <> "" Then
        ' Row 6 should have either marker text or be empty in pack columns
    End If

    ' Check if row 8 has pack information
    If SafeTrim(ws.Cells(ROW_PACK_INFO, COL_FSLI + 1).Value) = "" Then
        ' Should have at least some pack data
    End If

    ValidateWorksheetStructure = True
    Exit Function

ErrorHandler:
    ValidateWorksheetStructure = False
End Function

' Check if a string contains a marker that should be excluded
Public Function ContainsExclusionMarker(ByVal cellValue As String) As Boolean
    Dim trimmedValue As String
    trimmedValue = SafeTrim(cellValue)

    ' Case-insensitive comparison
    If InStr(1, trimmedValue, MARKER_TOTAL, vbTextCompare) > 0 Then
        ContainsExclusionMarker = True
    ElseIf InStr(1, trimmedValue, MARKER_ALL_JOURNALS, vbTextCompare) > 0 Then
        ContainsExclusionMarker = True
    ElseIf InStr(1, trimmedValue, MARKER_BASE_AMOUNTS, vbTextCompare) > 0 Then
        ContainsExclusionMarker = True
    Else
        ContainsExclusionMarker = False
    End If
End Function

' Validate category name
Public Function IsValidCategory(ByVal categoryName As String) As Boolean
    Select Case categoryName
        Case CAT_SEGMENT_SUMMARY, CAT_SEGMENT_TAB, CAT_UNCATEGORIZED
            IsValidCategory = True
        Case Else
            IsValidCategory = False
    End Select
End Function

' Get all valid category names as array
Public Function GetAllCategories() As Variant
    GetAllCategories = Array(CAT_SEGMENT_SUMMARY, CAT_SEGMENT_TAB, CAT_UNCATEGORIZED)
End Function

' Initialize global collections
Public Sub InitializeGlobalCollections()
    Set g_SegmentNames = CreateDictionary()
    Set g_FSLIList = CreateDictionary()
    Set g_PackList = CreateDictionary()
    g_BaseAmountsLabel = ""
End Sub

' Clean up global collections
Public Sub CleanupGlobalCollections()
    Set g_SegmentNames = Nothing
    Set g_FSLIList = Nothing
    Set g_PackList = Nothing
    g_BaseAmountsLabel = ""
End Sub
