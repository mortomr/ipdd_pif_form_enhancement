VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ============================================================================
' CLASS MODULE: ThisWorkbook
' ============================================================================
' Purpose: Workbook-level event handlers
' Author: Data Architecture Team
' Date: 2025-11-13
'
' FEATURES:
'   - Auto-refresh Inflight table on workbook open
' ============================================================================

Option Explicit

' ----------------------------------------------------------------------------
' Event: Workbook_Open
' Purpose: Auto-refresh Inflight table when workbook opens
' Notes: Runs silently in background for seamless user experience
' ----------------------------------------------------------------------------
Private Sub Workbook_Open()
    On Error Resume Next  ' Don't block workbook open if refresh fails

    ' Auto-refresh Inflight table (silent, no message box)
    ' Call mod_WorksheetQuery.RefreshInflight(showMessage:=False)
    ThisWorkbook.RefreshAll
    


    On Error GoTo 0
End Sub

