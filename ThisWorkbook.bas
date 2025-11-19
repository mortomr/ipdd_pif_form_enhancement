Attribute VB_Name = "ThisWorkbook"
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
    Call mod_WorksheetQuery.RefreshInflight(showMessage:=False)

    On Error GoTo 0
End Sub
Attribute VB_Name = "ThisWorkbook"
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
    Call mod_WorksheetQuery.RefreshInflight(showMessage:=False)

    On Error GoTo 0
End Sub
