Attribute VB_Name = "ModDataProcessing"
Option Explicit

'==============================================================================
' MODULE: ModDataProcessing
' PURPOSE: Process segment data, analyze rows, extract FSLIs and pack information
' DESCRIPTION: Handles Row 6 analysis, FSLI extraction, pack name/code separation,
'              and data validation for segment tabs
'==============================================================================

' NOTE: PackInfo and ColumnAnalysis are now Class modules
' See clsPackInfo.cls and clsColumnAnalysis.cls

'==================== ROW 6 ANALYSIS ====================

' Analyze Row 6 to identify which columns should be included/excluded
Public Function AnalyzeRow6(ByVal ws As Worksheet, ByVal startCol As Long, ByVal endCol As Long) As Collection
    On Error GoTo ErrorHandler

    LogMessage "Analyzing Row 6 for worksheet: " & ws.Name

    Dim results As New Collection
    Dim colIdx As Long
    Dim analysis As clsColumnAnalysis

    For colIdx = startCol To endCol
        ' Create new analysis object for this column
        Set analysis = New clsColumnAnalysis

        ' Initialize analysis structure
        analysis.ColumnIndex = colIdx
        analysis.MarkerText = SafeTrim(ws.Cells(ROW_MARKER, colIdx).Value)

        ' Check if Row 6 contains exclusion marker
        If analysis.MarkerText = "" Then
            ' Empty Row 6 = valid column
            analysis.IsValid = True
        ElseIf ContainsExclusionMarker(analysis.MarkerText) Then
            ' Contains exclusion marker = skip this column
            analysis.IsValid = False
            LogMessage "  Column " & colIdx & " EXCLUDED (Row 6: '" & analysis.MarkerText & "')"
        Else
            ' Has text but not an exclusion marker - treat as valid
            analysis.IsValid = True
        End If

        ' Get pack information from Row 8 (only for valid columns)
        If analysis.IsValid Then
            Dim row8Value As String
            row8Value = SafeTrim(ws.Cells(ROW_PACK_INFO, colIdx).Value)

            If row8Value <> "" Then
                ' Separate pack name and pack code
                Call SeparatePackNameAndCode(row8Value, analysis.PackName, analysis.PackCode)
                LogMessage "  Column " & colIdx & " INCLUDED - Pack: " & analysis.PackName & " (Code: " & analysis.PackCode & ")"
            Else
                ' No pack data in Row 8
                analysis.PackName = ""
                analysis.PackCode = ""
                analysis.IsValid = False ' Skip columns without pack data
            End If
        Else
            analysis.PackName = ""
            analysis.PackCode = ""
        End If

        ' Add to results collection
        results.Add analysis
    Next colIdx

    LogMessage "Row 6 analysis complete. Found " & CountValidColumns(results) & " valid columns out of " & (endCol - startCol + 1)

    Set AnalyzeRow6 = results
    Exit Function

ErrorHandler:
    ShowError "Row 6 Analysis Error", "Error analyzing Row 6: " & Err.Description
    Set AnalyzeRow6 = New Collection
End Function

' Count valid columns from analysis results
Private Function CountValidColumns(ByVal results As Collection) As Long
    Dim count As Long
    Dim item As Variant

    count = 0
    For Each item In results
        If item.IsValid Then count = count + 1
    Next item

    CountValidColumns = count
End Function

'==================== PACK NAME/CODE SEPARATION ====================

' Separate pack name from pack code
' Expected format: "Pack Name (Code)" or "Pack Name - Code" or just "Pack Name"
Public Sub SeparatePackNameAndCode(ByVal fullText As String, ByRef packName As String, ByRef packCode As String)
    On Error GoTo ErrorHandler

    Dim trimmedText As String
    Dim openParenPos As Long
    Dim closeParenPos As Long
    Dim dashPos As Long

    trimmedText = SafeTrim(fullText)

    ' Try to extract code from parentheses first: "Pack Name (Code)"
    openParenPos = InStrRev(trimmedText, "(")
    closeParenPos = InStrRev(trimmedText, ")")

    If openParenPos > 0 And closeParenPos > openParenPos Then
        ' Found parentheses
        packCode = SafeTrim(Mid(trimmedText, openParenPos + 1, closeParenPos - openParenPos - 1))
        packName = SafeTrim(Left(trimmedText, openParenPos - 1))
    Else
        ' Try to extract code from dash: "Pack Name - Code"
        dashPos = InStrRev(trimmedText, "-")
        If dashPos > 0 Then
            packCode = SafeTrim(Mid(trimmedText, dashPos + 1))
            packName = SafeTrim(Left(trimmedText, dashPos - 1))
        Else
            ' No separator found - use entire text as pack name
            packName = trimmedText
            packCode = ""
        End If
    End If

    Exit Sub

ErrorHandler:
    packName = fullText
    packCode = ""
End Sub

'==================== FSLI EXTRACTION ====================

' Extract all FSLIs from a worksheet (Column B)
Public Function ExtractFSLIs(ByVal ws As Worksheet, ByVal startRow As Long) As Collection
    On Error GoTo ErrorHandler

    Dim fslis As New Collection
    Dim fsliDict As Object
    Set fsliDict = CreateDictionary()

    Dim rowIdx As Long
    Dim fsli As String
    Dim lastRow As Long

    ' Find last row with data in Column B
    lastRow = ws.Cells(ws.Rows.count, COL_FSLI).End(xlUp).Row

    LogMessage "Extracting FSLIs from " & ws.Name & " (Row " & startRow & " to " & lastRow & ")"

    For rowIdx = startRow To lastRow
        fsli = SafeTrim(ws.Cells(rowIdx, COL_FSLI).Value)

        If fsli <> "" Then
            ' Add to dictionary to ensure uniqueness
            If Not fsliDict.Exists(fsli) Then
                fsliDict.Add fsli, rowIdx
                fslis.Add fsli

                ' Also add to global FSLI list
                If Not g_FSLIList.Exists(fsli) Then
                    g_FSLIList.Add fsli, True
                End If
            End If
        End If
    Next rowIdx

    LogMessage "Found " & fslis.count & " unique FSLIs in " & ws.Name

    Set ExtractFSLIs = fslis
    Exit Function

ErrorHandler:
    ShowError "FSLI Extraction Error", "Error extracting FSLIs: " & Err.Description
    Set ExtractFSLIs = New Collection
End Function

'==================== BASE AMOUNTS IDENTIFICATION ====================

' Find all columns in Summary tab that contain base amounts label (Row 7)
Public Function FindBaseAmountsColumns(ByVal summaryWs As Worksheet, ByVal label As String) As Collection
    On Error GoTo ErrorHandler

    Dim columns As New Collection
    Dim colIdx As Long
    Dim cellValue As String
    Dim maxCol As Long

    ' Find last column with data
    maxCol = summaryWs.Cells(ROW_BASE_AMOUNTS, summaryWs.Columns.count).End(xlToLeft).Column

    LogMessage "Searching for base amounts columns with label: " & label

    For colIdx = 1 To maxCol
        cellValue = SafeTrim(summaryWs.Cells(ROW_BASE_AMOUNTS, colIdx).Value)

        ' Check if cell contains the base amounts label (case-insensitive)
        If InStr(1, cellValue, label, vbTextCompare) > 0 Then
            columns.Add colIdx
            LogMessage "  Found base amounts column at index: " & colIdx & " ('" & cellValue & "')"
        End If
    Next colIdx

    LogMessage "Found " & columns.count & " base amounts columns"

    Set FindBaseAmountsColumns = columns
    Exit Function

ErrorHandler:
    ShowError "Base Amounts Search Error", "Error finding base amounts columns: " & Err.Description
    Set FindBaseAmountsColumns = New Collection
End Function

' Calculate base amounts for each FSLI from Summary tab
Public Function CalculateBaseAmounts(ByVal summaryWs As Worksheet, ByVal baseAmountsColumns As Collection, ByVal fsliList As Collection) As Object
    On Error GoTo ErrorHandler

    Dim baseAmounts As Object
    Set baseAmounts = CreateDictionary()

    Dim fsli As Variant
    Dim colIdx As Variant
    Dim fsliRow As Long
    Dim total As Double
    Dim cellValue As Variant

    LogMessage "Calculating base amounts for " & fsliList.count & " FSLIs across " & baseAmountsColumns.count & " base amounts columns"

    For Each fsli In fsliList
        total = 0

        ' Find the row for this FSLI
        fsliRow = FindFSLIRow(summaryWs, CStr(fsli))

        If fsliRow > 0 Then
            ' Sum across all base amounts columns
            For Each colIdx In baseAmountsColumns
                cellValue = summaryWs.Cells(fsliRow, CLng(colIdx)).Value

                If IsValidNumber(cellValue) Then
                    total = total + CDbl(cellValue)
                End If
            Next colIdx

            ' Store the base amount for this FSLI
            baseAmounts.Add CStr(fsli), total
            LogMessage "  FSLI '" & fsli & "': Base amount = " & FormatCurrency(total)
        Else
            ' FSLI not found in summary tab
            LogMessage "  WARNING: FSLI '" & fsli & "' not found in summary tab"
            baseAmounts.Add CStr(fsli), 0
        End If
    Next fsli

    Set CalculateBaseAmounts = baseAmounts
    Exit Function

ErrorHandler:
    ShowError "Base Amounts Calculation Error", "Error calculating base amounts: " & Err.Description
    Set CalculateBaseAmounts = CreateDictionary()
End Function

' Find the row number for a specific FSLI in a worksheet
Public Function FindFSLIRow(ByVal ws As Worksheet, ByVal fsli As String) As Long
    On Error GoTo ErrorHandler

    Dim rowIdx As Long
    Dim lastRow As Long
    Dim cellValue As String

    lastRow = ws.Cells(ws.Rows.count, COL_FSLI).End(xlUp).Row

    For rowIdx = 1 To lastRow
        cellValue = SafeTrim(ws.Cells(rowIdx, COL_FSLI).Value)

        If StrComp(cellValue, fsli, vbTextCompare) = 0 Then
            FindFSLIRow = rowIdx
            Exit Function
        End If
    Next rowIdx

    ' FSLI not found
    FindFSLIRow = 0
    Exit Function

ErrorHandler:
    FindFSLIRow = 0
End Function

'==================== DATA EXTRACTION ====================

' Extract data for a specific FSLI and pack from a segment worksheet
Public Function GetSegmentPackFSLIValue(ByVal ws As Worksheet, ByVal fsli As String, ByVal packColumn As Long) As Variant
    On Error GoTo ErrorHandler

    Dim fsliRow As Long

    ' Find the row for this FSLI
    fsliRow = FindFSLIRow(ws, fsli)

    If fsliRow > 0 And packColumn > 0 Then
        GetSegmentPackFSLIValue = ws.Cells(fsliRow, packColumn).Value
    Else
        GetSegmentPackFSLIValue = 0
    End If

    Exit Function

ErrorHandler:
    GetSegmentPackFSLIValue = 0
End Function

'==================== DATA VALIDATION ====================

' Validate that a segment worksheet has the expected structure
Public Function ValidateSegmentWorksheet(ByVal ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler

    ValidateSegmentWorksheet = False

    ' Check that Column B has FSLI data
    If SafeTrim(ws.Cells(10, COL_FSLI).Value) = "" Then
        LogMessage "WARNING: No FSLI data found in Column B of " & ws.Name
        Exit Function
    End If

    ' Check that Row 6 exists
    If ws.Rows.count < ROW_MARKER Then
        LogMessage "WARNING: Worksheet " & ws.Name & " does not have enough rows"
        Exit Function
    End If

    ' Check that Row 8 exists
    If ws.Rows.count < ROW_PACK_INFO Then
        LogMessage "WARNING: Worksheet " & ws.Name & " does not have Row 8 (pack info)"
        Exit Function
    End If

    ValidateSegmentWorksheet = True
    Exit Function

ErrorHandler:
    LogMessage "ERROR: Failed to validate worksheet " & ws.Name & ": " & Err.Description
    ValidateSegmentWorksheet = False
End Function

' Get the range of data columns (excluding Column B which has FSLIs)
Public Function GetDataColumnRange(ByVal ws As Worksheet) As Long
    On Error GoTo ErrorHandler

    ' Find last column with data in Row 8 (pack info row)
    GetDataColumnRange = ws.Cells(ROW_PACK_INFO, ws.Columns.count).End(xlToLeft).Column

    Exit Function

ErrorHandler:
    GetDataColumnRange = 50 ' Default fallback
End Function

'==================== PACK COLLECTION ====================

' Build a collection of all unique packs across all segment tabs
Public Function CollectAllPacks(ByVal segmentWorksheets As Collection) As Collection
    On Error GoTo ErrorHandler

    Dim allPacks As New Collection
    Dim packDict As Object
    Set packDict = CreateDictionary()

    Dim ws As Worksheet
    Dim wsVar As Variant
    Dim colAnalysis As Collection
    Dim analysis As Variant
    Dim packKey As String

    LogMessage "Collecting all unique packs from " & segmentWorksheets.count & " segment worksheets"

    For Each wsVar In segmentWorksheets
        Set ws = wsVar

        ' Get last column
        Dim lastCol As Long
        lastCol = GetDataColumnRange(ws)

        ' Analyze Row 6 to get valid columns
        Set colAnalysis = AnalyzeRow6(ws, COL_FSLI + 1, lastCol)

        ' Extract packs from valid columns
        For Each analysis In colAnalysis
            If analysis.IsValid And analysis.PackName <> "" Then
                packKey = analysis.PackCode & "|" & analysis.PackName

                If Not packDict.Exists(packKey) Then
                    packDict.Add packKey, Array(analysis.PackName, analysis.PackCode)

                    ' Also add to global pack list
                    If Not g_PackList.Exists(packKey) Then
                        g_PackList.Add packKey, Array(analysis.PackName, analysis.PackCode)
                    End If
                End If
            End If
        Next analysis
    Next wsVar

    ' Convert dictionary to collection
    Dim key As Variant
    For Each key In packDict.Keys
        allPacks.Add packDict(key)
    Next key

    LogMessage "Found " & allPacks.count & " unique packs"

    Set CollectAllPacks = allPacks
    Exit Function

ErrorHandler:
    ShowError "Pack Collection Error", "Error collecting packs: " & Err.Description
    Set CollectAllPacks = New Collection
End Function
