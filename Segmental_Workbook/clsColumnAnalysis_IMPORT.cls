VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColumnAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' CLASS: clsColumnAnalysis
' PURPOSE: Store column analysis results
'==============================================================================

Public ColumnIndex As Long
Public IsValid As Boolean
Public MarkerText As String
Public PackName As String
Public PackCode As String
