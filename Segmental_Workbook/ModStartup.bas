Attribute VB_Name = "ModStartup"
Option Explicit

'==============================================================================
' MODULE: ModStartup
' PURPOSE: Create startup worksheet with launch button
' DESCRIPTION: Sets up a user-friendly start page with instructions and button
'==============================================================================

' Create the startup worksheet with button
Public Sub CreateStartupSheet()
    On Error Resume Next

    Dim ws As Worksheet
    Dim btn As Button
    Dim wb As Workbook

    Set wb = ThisWorkbook

    ' Delete existing Start sheet if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets("START HERE").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Create new Start sheet
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = "START HERE"

    ' Format the sheet
    With ws
        ' Title
        .Range("B2:H2").Merge
        .Range("B2").Value = "BIDVEST SEGMENTAL SCOPING TOOL"
        With .Range("B2")
            .Font.Size = 20
            .Font.Bold = True
            .Font.Color = RGB(0, 112, 192)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Range("B2:H2").RowHeight = 40

        ' Version info
        .Range("B3:H3").Merge
        .Range("B3").Value = "Version " & TOOL_VERSION & " - " & TOOL_DATE
        With .Range("B3")
            .Font.Size = 10
            .Font.Italic = True
            .HorizontalAlignment = xlCenter
        End With

        ' Welcome message
        .Range("B5").Value = "Welcome! This tool analyzes segmental financial reports and creates:"
        .Range("B5").Font.Size = 12
        .Range("B5").Font.Bold = True

        ' Features list
        Dim features As Variant
        features = Array( _
            "✓ FullTable - Aggregated data from all segments", _
            "✓ FullTablePercentage - Percentage of base amounts", _
            "✓ Interactive Scoping Dashboard", _
            "✓ Power BI-ready tables", _
            "✓ Threshold-based automatic scoping", _
            "✓ Manual scoping adjustments" _
        )

        Dim i As Long
        For i = 0 To UBound(features)
            .Cells(7 + i, 2).Value = features(i)
            .Cells(7 + i, 2).Font.Size = 11
        Next i

        ' Instructions section
        .Range("B14").Value = "BEFORE YOU START:"
        .Range("B14").Font.Size = 12
        .Range("B14").Font.Bold = True
        .Range("B14").Font.Color = RGB(192, 0, 0)

        Dim instructions As Variant
        instructions = Array( _
            "1. Open your segment report workbook", _
            "2. Ensure it has a Segment Summary tab with base amounts in Row 7", _
            "3. Ensure segment tabs have pack data in Row 8 (format: Pack Name (Code))", _
            "4. All tabs should have FSLIs in Column B", _
            "" _
        )

        For i = 0 To UBound(instructions)
            .Cells(16 + i, 2).Value = instructions(i)
            .Cells(16 + i, 2).Font.Size = 10
        Next i

        ' Button instruction
        .Range("B22").Value = "Click the button below to start the analysis:"
        .Range("B22").Font.Size = 12
        .Range("B22").Font.Bold = True

        ' Add the button
        Set btn = .Buttons.Add(100, 360, 200, 50)
        With btn
            .Caption = "▶ START ANALYSIS"
            .OnAction = "ModMain.RunSegmentalAnalysis"
            .Font.Size = 14
            .Font.Bold = True
        End With

        ' Additional buttons
        Set btn = .Buttons.Add(320, 360, 180, 50)
        With btn
            .Caption = "📊 DASHBOARD"
            .OnAction = "ModInteractiveDashboard.ShowInteractiveDashboard"
            .Font.Size = 12
            .Font.Bold = True
        End With

        Set btn = .Buttons.Add(520, 360, 180, 50)
        With btn
            .Caption = "🔄 REFRESH DATA"
            .OnAction = "ModMain.RefreshAllData"
            .Font.Size = 12
            .Font.Bold = True
        End With

        ' Help text
        .Range("B27").Value = "Need help? See the documentation:"
        .Range("B27").Font.Size = 10
        .Range("B27").Font.Italic = True

        .Range("B28").Value = "• INSTALLATION.md - Setup instructions"
        .Range("B29").Value = "• USAGE_GUIDE.md - How to use this tool"
        .Range("B30").Value = "• QUICK_START.md - Fast setup guide"

        For i = 28 To 30
            .Cells(i, 2).Font.Size = 9
            .Cells(i, 2).Font.Color = RGB(100, 100, 100)
        Next i

        ' Footer
        .Range("B32:H32").Merge
        .Range("B32").Value = "Bidvest © 2025 - Segmental Scoping Tool"
        With .Range("B32")
            .Font.Size = 8
            .Font.Italic = True
            .Font.Color = RGB(150, 150, 150)
            .HorizontalAlignment = xlCenter
        End With

        ' Format columns
        .Columns("A:A").ColumnWidth = 2
        .Columns("B:H").ColumnWidth = 15
        .Columns("B:H").AutoFit

        ' Add some color
        .Range("B2:H3").Interior.Color = RGB(217, 225, 242)
        .Range("B5:H13").Interior.Color = RGB(226, 239, 218)
        .Range("B14:H21").Interior.Color = RGB(252, 228, 214)

        ' Borders
        .Range("B2:H32").Borders.LineStyle = xlContinuous
        .Range("B2:H32").Borders.Weight = xlThin

        ' Freeze top rows
        .Range("B4").Select
        ActiveWindow.FreezePanes = True

        ' Select start position
        .Range("B2").Select

    End With

    MsgBox "Startup sheet created successfully!" & vbCrLf & vbCrLf & _
           "The 'START HERE' tab has been added with a button to launch the analysis.", _
           vbInformation, "Setup Complete"

End Sub

' Create a simple ribbon-like menu sheet
Public Sub CreateQuickAccessMenu()
    On Error Resume Next

    Dim ws As Worksheet
    Dim wb As Workbook

    Set wb = ThisWorkbook

    ' Delete existing menu if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets("QUICK MENU").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Create new menu sheet
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = "QUICK MENU"

    With ws
        ' Title
        .Range("A1:F1").Merge
        .Range("A1").Value = "⚡ QUICK ACCESS MENU"
        With .Range("A1")
            .Font.Size = 18
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(0, 112, 192)
            .Font.Color = RGB(255, 255, 255)
            .RowHeight = 35
        End With

        ' Main Actions
        Dim btn As Button
        Dim btnTop As Long
        btnTop = 50

        ' Run Analysis
        Set btn = .Buttons.Add(20, btnTop, 250, 45)
        With btn
            .Caption = "▶ RUN FULL ANALYSIS"
            .OnAction = "ModMain.RunSegmentalAnalysis"
            .Font.Size = 12
            .Font.Bold = True
        End With

        ' Dashboard
        Set btn = .Buttons.Add(290, btnTop, 250, 45)
        With btn
            .Caption = "📊 INTERACTIVE DASHBOARD"
            .OnAction = "ModInteractiveDashboard.ShowInteractiveDashboard"
            .Font.Size = 12
            .Font.Bold = True
        End With

        btnTop = btnTop + 60

        ' Apply Thresholds
        Set btn = .Buttons.Add(20, btnTop, 250, 45)
        With btn
            .Caption = "🎯 APPLY THRESHOLDS"
            .OnAction = "ModInteractiveDashboard.ApplyThresholdScoping"
            .Font.Size = 12
            .Font.Bold = True
        End With

        ' Clear Scoping
        Set btn = .Buttons.Add(290, btnTop, 250, 45)
        With btn
            .Caption = "🗑️ CLEAR SCOPING"
            .OnAction = "ModInteractiveDashboard.ClearAllScoping"
            .Font.Size = 12
            .Font.Bold = True
        End With

        btnTop = btnTop + 60

        ' Refresh Data
        Set btn = .Buttons.Add(20, btnTop, 250, 45)
        With btn
            .Caption = "🔄 REFRESH ALL DATA"
            .OnAction = "ModMain.RefreshAllData"
            .Font.Size = 12
            .Font.Bold = True
        End With

        ' Export to Power BI
        Set btn = .Buttons.Add(290, btnTop, 250, 45)
        With btn
            .Caption = "📤 EXPORT TO POWER BI"
            .OnAction = "ModMain.ExportToPowerBI"
            .Font.Size = 12
            .Font.Bold = True
        End With

        btnTop = btnTop + 60

        ' Regenerate Tables
        Set btn = .Buttons.Add(20, btnTop, 250, 45)
        With btn
            .Caption = "🔨 REGENERATE FULL TABLE"
            .OnAction = "ModMain.RegenerateFullTable"
            .Font.Size = 12
            .Font.Bold = True
        End With

        ' Regenerate Power BI
        Set btn = .Buttons.Add(290, btnTop, 250, 45)
        With btn
            .Caption = "🔧 REGENERATE PBI TABLES"
            .OnAction = "ModMain.RegeneratePowerBITables"
            .Font.Size = 12
            .Font.Bold = True
        End With

        ' Formatting
        .Columns("A:F").ColumnWidth = 12
        .Range("A1").Select

    End With

    MsgBox "Quick Access Menu created!" & vbCrLf & vbCrLf & _
           "Use the 'QUICK MENU' tab for easy access to all functions.", _
           vbInformation, "Menu Ready"

End Sub
