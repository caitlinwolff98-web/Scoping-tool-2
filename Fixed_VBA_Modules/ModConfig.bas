Attribute VB_Name = "ModConfig"
Option Explicit

' ============================================================================
' MODULE: ModConfig (FIXED VERSION)
' PURPOSE: Centralized configuration and constants for the Bidvest Scoping Tool
' DESCRIPTION: Contains all global constants, configuration settings, and
'              shared utility functions used across modules
' FIXES APPLIED:
'   1. Added ParsePackCodeFromName - Extract "XX-####" from end of pack name
'   2. Added ParsePackNameWithoutCode - Get clean pack name without code
'   3. Added FixEncodingIssues - Fix "‰Û¢" to "•" character issues
'   4. Added ConvertRangeToTable - Convert ranges to Excel ListObjects
'   5. Fixed encoding in all string constants
' ============================================================================

' ==================== CATEGORY CONSTANTS ====================
' These constants define worksheet categories for processing
Public Const CAT_SEGMENT As String = "TGK Segment Tabs"
Public Const CAT_DISCONTINUED As String = "Discontinued Ops Tab"
Public Const CAT_INPUT_CONTINUING As String = "TGK Input Continuing Operations Tab"
Public Const CAT_JOURNALS_CONTINUING As String = "TGK Journals Continuing Tab"
Public Const CAT_CONSOLE_CONTINUING As String = "TGK Consol Continuing Tab"
Public Const CAT_BS As String = "TGK BS Tab"
Public Const CAT_IS As String = "TGK IS Tab"
Public Const CAT_PULL_WORKINGS As String = "Paul workings"
Public Const CAT_TRIAL_BALANCE As String = "Trial Balance"
Public Const CAT_UNCATEGORIZED As String = "Uncategorized"

' ==================== VERSION INFORMATION ====================
Public Const TOOL_VERSION As String = "8.2.0"
Public Const TOOL_NAME As String = "Bidvest Scoping Tool"
Public Const TOOL_DATE As String = "2025-11"

' ==================== PROCESSING CONSTANTS ====================
' Row indices for data structure
Public Const ROW_COLUMN_TYPE As Long = 6
Public Const ROW_PACK_NAME As Long = 7
Public Const ROW_PACK_CODE As Long = 8       ' NOTE: Now parsed from pack name, not read from row
Public Const ROW_DATA_START As Long = 9

' Column type identifiers
Public Const COLTYPE_ORIGINAL_ENTITY As String = "Original/Entity"
Public Const COLTYPE_CONSOLIDATION As String = "Consolidation/Consolidation"
Public Const COLTYPE_OTHER As String = "Other"

' ==================== POWER BI INTEGRATION CONSTANTS ====================
Public Const POWERBI_METADATA_SHEET As String = "PowerBI_Metadata"
Public Const POWERBI_SCOPING_SHEET As String = "PowerBI_Scoping"

' ==================== ERROR MESSAGES ====================
Public Const ERR_WORKBOOK_NOT_FOUND As String = "Could not find the specified workbook. Please ensure it is open."
Public Const ERR_REQUIRED_TAB_MISSING As String = "Required tabs are missing. At least one 'Input Continuing' tab must be categorized."
Public Const ERR_NO_TABS_FOUND As String = "No tabs found in the workbook."
Public Const ERR_CATEGORIZATION_CANCELLED As String = "Tab categorization was cancelled."
Public Const ERR_SCRIPTING_RUNTIME As String = "Microsoft Scripting Runtime is not available. Please enable it in VBA References."

' ==================== PACK CODE PARSING FUNCTIONS (NEW) ====================

' NEW: Parse pack code from end of pack name
' Pack names are in format: "Pack Name XX-####" where XX = 2 letters, #### = 4 digits
' Example: "Top Turf LS-0714" -> returns "LS-0714"
Public Function ParsePackCodeFromName(ByVal packNameWithCode As String) As String
    On Error Resume Next

    Dim trimmedName As String
    Dim lastWord As String
    Dim lastDashPos As Long
    Dim codeCandidate As String
    Dim letterPart As String
    Dim numberPart As String

    ' Trim input
    trimmedName = Trim(packNameWithCode)

    If trimmedName = "" Then
        ParsePackCodeFromName = ""
        Exit Function
    End If

    ' Find last dash position
    lastDashPos = InStrRev(trimmedName, "-")

    If lastDashPos = 0 Then
        ' No dash found, not a valid pack code
        ParsePackCodeFromName = ""
        Exit Function
    End If

    ' Extract potential code (everything after last space before dash to end)
    ' Look backwards from dash to find start of code (after space)
    Dim spaceBeforeDash As Long
    Dim i As Long

    ' Find space before the code
    For i = lastDashPos - 1 To 1 Step -1
        If Mid(trimmedName, i, 1) = " " Then
            spaceBeforeDash = i
            Exit For
        End If
    Next i

    If spaceBeforeDash > 0 Then
        ' Extract code candidate
        codeCandidate = Trim(Mid(trimmedName, spaceBeforeDash + 1))

        ' Validate format: XX-#### (2 letters, dash, 4 digits)
        If Len(codeCandidate) >= 6 Then
            letterPart = Left(codeCandidate, 2)
            numberPart = Mid(codeCandidate, 4, 4)

            ' Check if letters are actually letters and numbers are digits
            If IsLetter(Mid(letterPart, 1, 1)) And IsLetter(Mid(letterPart, 2, 1)) And _
               Mid(codeCandidate, 3, 1) = "-" And IsNumeric(numberPart) Then
                ParsePackCodeFromName = codeCandidate
                Exit Function
            End If
        End If
    End If

    ' If we get here, no valid code found
    ParsePackCodeFromName = ""

    On Error GoTo 0
End Function

' NEW: Parse pack name without the pack code at the end
' Example: "Top Turf LS-0714" -> returns "Top Turf"
Public Function ParsePackNameWithoutCode(ByVal packNameWithCode As String) As String
    On Error Resume Next

    Dim packCode As String
    Dim cleanName As String
    Dim trimmedName As String

    trimmedName = Trim(packNameWithCode)

    If trimmedName = "" Then
        ParsePackNameWithoutCode = ""
        Exit Function
    End If

    ' Get the pack code
    packCode = ParsePackCodeFromName(trimmedName)

    If packCode <> "" Then
        ' Remove the pack code from the end
        cleanName = Replace(trimmedName, packCode, "")
        cleanName = Trim(cleanName)
        ParsePackNameWithoutCode = cleanName
    Else
        ' No pack code found, return original
        ParsePackNameWithoutCode = trimmedName
    End If

    On Error GoTo 0
End Function

' NEW: Check if a character is a letter (A-Z or a-z)
Private Function IsLetter(ByVal char As String) As Boolean
    On Error Resume Next

    Dim asciiVal As Integer

    If Len(char) <> 1 Then
        IsLetter = False
        Exit Function
    End If

    asciiVal = Asc(UCase(char))

    ' A-Z in ASCII is 65-90
    IsLetter = (asciiVal >= 65 And asciiVal <= 90)

    On Error GoTo 0
End Function

' NEW: Fix encoding issues in text (convert weird symbols to proper ones)
' Fixes: "‰Û¢" -> "•"
Public Function FixEncodingIssues(ByVal text As String) As String
    On Error Resume Next

    Dim fixedText As String

    fixedText = text

    ' Replace common encoding issues
    fixedText = Replace(fixedText, "‰Û¢", "•")
    fixedText = Replace(fixedText, "â€¢", "•")
    fixedText = Replace(fixedText, "â€"", "–")
    fixedText = Replace(fixedText, "â€™", "'")
    fixedText = Replace(fixedText, "â€œ", """")
    fixedText = Replace(fixedText, "â€�", """")

    FixEncodingIssues = fixedText

    On Error GoTo 0
End Function

' NEW: Convert a range to an Excel ListObject table
' Parameters:
'   ws - Worksheet containing the range
'   dataRange - Range to convert (must include headers)
'   tableName - Name for the table
'   tableStyle - Optional table style (default: TableStyleMedium2)
' Returns: The created ListObject, or Nothing if failed
Public Function ConvertRangeToTable(ws As Worksheet, dataRange As Range, _
                                     tableName As String, _
                                     Optional tableStyle As String = "TableStyleMedium2") As ListObject
    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim safeTableName As String

    ' Make table name safe (no spaces, special chars)
    safeTableName = Replace(tableName, " ", "_")
    safeTableName = Replace(safeTableName, "-", "_")
    safeTableName = Replace(safeTableName, "(", "")
    safeTableName = Replace(safeTableName, ")", "")

    ' Delete existing table with same name
    On Error Resume Next
    ws.ListObjects(safeTableName).Delete
    On Error GoTo ErrorHandler

    ' Create new table
    Set tbl = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    tbl.Name = safeTableName
    tbl.TableStyle = tableStyle

    Set ConvertRangeToTable = tbl
    Exit Function

ErrorHandler:
    Debug.Print "Error converting range to table: " & Err.Description
    Set ConvertRangeToTable = Nothing
End Function

' ==================== UTILITY FUNCTIONS ====================

' Check if Scripting Runtime is available
Public Function IsScriptingRuntimeAvailable() As Boolean
    On Error Resume Next
    Dim testDict As Object
    Set testDict = CreateObject("Scripting.Dictionary")
    IsScriptingRuntimeAvailable = (Err.Number = 0)
    On Error GoTo 0
End Function

' Display formatted error message
Public Sub ShowError(ByVal errorTitle As String, ByVal errorMessage As String, Optional ByVal errorNumber As Long = 0)
    Dim msg As String
    msg = FixEncodingIssues(errorMessage)

    If errorNumber <> 0 Then
        msg = msg & vbCrLf & vbCrLf & "Error Number: " & errorNumber
    End If

    MsgBox msg, vbCritical, errorTitle
End Sub

' Display formatted information message
Public Sub ShowInfo(ByVal title As String, ByVal message As String)
    MsgBox FixEncodingIssues(message), vbInformation, title
End Sub

' Display formatted warning message
Public Function ShowWarning(ByVal title As String, ByVal message As String) As VbMsgBoxResult
    ShowWarning = MsgBox(FixEncodingIssues(message), vbExclamation + vbOKCancel, title)
End Function

' Log message to immediate window (for debugging)
Public Sub LogDebug(ByVal message As String)
    #If DEBUG_MODE Then
        Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " - " & message
    #End If
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
    IsValidNumber = False

    If Not IsEmpty(value) And Not IsNull(value) Then
        If IsNumeric(value) Then
            IsValidNumber = True
        End If
    End If
End Function

' Get workbook by name (case-insensitive, handles extensions)
Public Function GetWorkbookByName(ByVal workbookName As String) As Workbook
    On Error Resume Next
    Dim wb As Workbook
    Dim nameWithoutExt As String

    ' Try exact name first
    Set wb = Workbooks(workbookName)

    If wb Is Nothing Then
        ' Try without extension
        nameWithoutExt = Replace(Replace(workbookName, ".xlsx", ""), ".xlsm", "")
        nameWithoutExt = Replace(nameWithoutExt, ".xls", "")

        Set wb = Workbooks(nameWithoutExt & ".xlsx")
        If wb Is Nothing Then Set wb = Workbooks(nameWithoutExt & ".xlsm")
        If wb Is Nothing Then Set wb = Workbooks(nameWithoutExt & ".xls")
        If wb Is Nothing Then Set wb = Workbooks(nameWithoutExt)
    End If

    Set GetWorkbookByName = wb
    On Error GoTo 0
End Function

' Create a dictionary object (with error handling)
Public Function CreateDictionary() As Object
    On Error Resume Next
    Set CreateDictionary = CreateObject("Scripting.Dictionary")

    If Err.Number <> 0 Then
        ShowError "Missing Library", ERR_SCRIPTING_RUNTIME
        Set CreateDictionary = Nothing
    End If

    On Error GoTo 0
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

' ==================== VALIDATION FUNCTIONS ====================

' Validate workbook structure (checks for expected rows)
Public Function ValidateWorkbookStructure(ws As Worksheet) As Boolean
    On Error Resume Next

    ValidateWorkbookStructure = False

    ' Check if worksheet exists
    If ws Is Nothing Then Exit Function

    ' Check if row 6 has column type data
    If ws.Cells(ROW_COLUMN_TYPE, 2).value = "" Then Exit Function

    ' Check if row 7 has pack names
    If ws.Cells(ROW_PACK_NAME, 2).value = "" Then Exit Function

    ' NOTE: Row 8 (pack codes) is no longer required as we parse from pack names

    ValidateWorkbookStructure = True
    On Error GoTo 0
End Function

' Validate category name
Public Function IsValidCategory(ByVal categoryName As String) As Boolean
    Select Case categoryName
        Case CAT_SEGMENT, CAT_DISCONTINUED, CAT_INPUT_CONTINUING, _
             CAT_JOURNALS_CONTINUING, CAT_CONSOLE_CONTINUING, _
             CAT_BS, CAT_IS, CAT_PULL_WORKINGS, CAT_TRIAL_BALANCE, CAT_UNCATEGORIZED
            IsValidCategory = True
        Case Else
            IsValidCategory = False
    End Select
End Function

' Get all valid category names as array
Public Function GetAllCategories() As Variant
    GetAllCategories = Array( _
        CAT_SEGMENT, _
        CAT_DISCONTINUED, _
        CAT_INPUT_CONTINUING, _
        CAT_JOURNALS_CONTINUING, _
        CAT_CONSOLE_CONTINUING, _
        CAT_BS, _
        CAT_IS, _
        CAT_PULL_WORKINGS, _
        CAT_TRIAL_BALANCE, _
        CAT_UNCATEGORIZED _
    )
End Function

' Get required categories (must have at least one tab)
Public Function GetRequiredCategories() As Variant
    GetRequiredCategories = Array(CAT_INPUT_CONTINUING)
End Function

' Get single-tab categories (can only have one tab)
Public Function GetSingleTabCategories() As Variant
    GetSingleTabCategories = Array( _
        CAT_DISCONTINUED, _
        CAT_INPUT_CONTINUING, _
        CAT_JOURNALS_CONTINUING, _
        CAT_CONSOLE_CONTINUING, _
        CAT_BS, _
        CAT_IS, _
        CAT_TRIAL_BALANCE _
    )
End Function
