Attribute VB_Name = "mod_WorksheetQuery"
' ============================================================================
' MODULE: mod_WorksheetQuery (REFRESH-ONLY MODE)
' ============================================================================
' Purpose: Refresh user-created Archive and Inflight tables
' Author: Data Architecture Team
' Date: 2025-11-18
'
' FEATURES:
'   - REFRESH-ONLY mode - does NOT recreate tables
'   - User manually creates tables with data connections
'   - Code only refreshes existing connections
'   - NO header text modifications
'   - NO pane freezing
'   - Preserves all user formatting and customization
' ============================================================================

Option Explicit

' Sheet and table name constants
Private Const SHEET_ARCHIVE As String = "PIF_Archive"
Private Const SHEET_INFLIGHT As String = "PIF_Inflight"
Private Const TABLE_ARCHIVE As String = "tbl_PIF_Archive"
Private Const TABLE_INFLIGHT As String = "tbl_PIF_Inflight"

' ============================================================================
' PUBLIC FUNCTIONS - TABLE REFRESH ONLY
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: RefreshArchive
' Purpose: Refresh user-created ArchiveTable (does NOT recreate!)
' Usage: Call from button or manually
' Notes: User must manually create ArchiveTable with data connection first
' ----------------------------------------------------------------------------
Public Sub RefreshArchive()
    On Error GoTo ErrHandler

    Dim startTime As Double
    Dim recordCount As Long

    startTime = Timer
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing Archive..."

    ' Refresh the existing table
    recordCount = RefreshExistingTable(SHEET_ARCHIVE, TABLE_ARCHIVE)

    Application.ScreenUpdating = True
    Application.StatusBar = False

    Dim elapsed As Double
    elapsed = Timer - startTime

    MsgBox "Archive refreshed!" & vbCrLf & vbCrLf & _
           "Records: " & recordCount & vbCrLf & _
           "Time: " & Format(elapsed, "0.0") & " seconds", _
           vbInformation, "Refresh Complete"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Failed to refresh Archive:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "Make sure you've manually created '" & TABLE_ARCHIVE & "' on the '" & SHEET_ARCHIVE & "' sheet " & _
           "with a data connection to the database view.", _
           vbCritical, "Refresh Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: RefreshInflight
' Purpose: Refresh user-created InflightTable (does NOT recreate!)
' Usage: Call from button, manually, or Workbook_Open event
' Notes: User must manually create InflightTable with data connection first
' ----------------------------------------------------------------------------
Public Sub RefreshInflight(Optional ByVal showMessage As Boolean = True)
    On Error GoTo ErrHandler

    Dim startTime As Double
    Dim recordCount As Long

    startTime = Timer
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing Inflight..."

    ' Refresh the existing table
    recordCount = RefreshExistingTable(SHEET_INFLIGHT, TABLE_INFLIGHT)

    Application.ScreenUpdating = True
    Application.StatusBar = False

    If showMessage Then
        Dim elapsed As Double
        elapsed = Timer - startTime

        MsgBox "Inflight refreshed!" & vbCrLf & vbCrLf & _
               "Records: " & recordCount & vbCrLf & _
               "Time: " & Format(elapsed, "0.0") & " seconds", _
               vbInformation, "Refresh Complete"
    End If

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If showMessage Then
        MsgBox "Failed to refresh Inflight:" & vbCrLf & vbCrLf & _
               "Error: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
               "Make sure you've manually created '" & TABLE_INFLIGHT & "' on the '" & SHEET_INFLIGHT & "' sheet " & _
               "with a data connection to the database view.", _
               vbCritical, "Refresh Error"
    End If
End Sub

' ----------------------------------------------------------------------------
' Sub: RefreshAll
' Purpose: Refresh both Archive and Inflight tables
' ----------------------------------------------------------------------------
Public Sub RefreshAll(Optional ByVal showSuccessMessage As Boolean = True)
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    Call RefreshArchiveSilent
    Call RefreshInflightSilent

    Application.ScreenUpdating = True

    If showSuccessMessage Then
        MsgBox "Both Archive and Inflight tables refreshed.", _
               vbInformation, "Refresh Complete"
    End If

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error refreshing tables:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Refresh Error"
End Sub

' ============================================================================
' PRIVATE HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: RefreshExistingTable
' Purpose: Refresh an existing Excel Table or QueryTable (does NOT create!)
' Parameters:
'   sheetName - Name of worksheet containing the table
'   tableName - Name of the Excel Table to refresh
' Returns: Number of records in the table after refresh
' Notes: Looks for QueryTable connections and refreshes them
'        Preserves all user formatting and customization
' ----------------------------------------------------------------------------
Private Function RefreshExistingTable(ByVal sheetName As String, _
                                     ByVal tableName As String) As Long
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim qt As QueryTable
    Dim i As Integer
    Dim refreshed As Boolean

    ' Get worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Try to find and refresh QueryTable first (most common for data connections)
    refreshed = False
    For i = 1 To ws.QueryTables.count
        Set qt = ws.QueryTables(i)
        ' Refresh any QueryTable on this sheet
        qt.Refresh BackgroundQuery:=False
        refreshed = True
    Next i

    ' If no QueryTable found, try to refresh ListObject data connection
    If Not refreshed Then
        On Error Resume Next
        Set tbl = ws.ListObjects(tableName)
        On Error GoTo ErrHandler

        If Not tbl Is Nothing Then
            If tbl.QueryTable Is Nothing Then
                ' Table exists but has no data connection
                Err.Raise vbObjectError + 1001, "RefreshExistingTable", _
                    "Table '" & tableName & "' exists but has no data connection. " & _
                    "Please create a table with a data connection to the database view."
            Else
                ' Refresh via ListObject's QueryTable
                tbl.QueryTable.Refresh BackgroundQuery:=False
                refreshed = True
            End If
        End If
    End If

    ' Verify we found something to refresh
    If Not refreshed Then
        Err.Raise vbObjectError + 1002, "RefreshExistingTable", _
            "No table or data connection found on sheet '" & sheetName & "'. " & _
            "Please manually create '" & tableName & "' with a data connection first."
    End If

    ' Count records in the table (if it's a ListObject)
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    If Not tbl Is Nothing Then
        RefreshExistingTable = tbl.ListRows.count
    Else
        ' Fallback: count rows in QueryTable result range
        For i = 1 To ws.QueryTables.count
            If Not ws.QueryTables(i).ResultRange Is Nothing Then
                RefreshExistingTable = ws.QueryTables(i).ResultRange.Rows.count - 1 ' Subtract header
                Exit For
            End If
        Next i
    End If
    On Error GoTo 0

    Exit Function

ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ----------------------------------------------------------------------------
' Sub: RefreshArchiveSilent
' Purpose: Refresh Archive without messagebox (for batch operations)
' ----------------------------------------------------------------------------
Private Sub RefreshArchiveSilent()
    On Error Resume Next
    Call RefreshExistingTable(SHEET_ARCHIVE, TABLE_ARCHIVE)
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: RefreshInflightSilent
' Purpose: Refresh Inflight without messagebox (for batch operations)
' ----------------------------------------------------------------------------
Private Sub RefreshInflightSilent()
    On Error Resume Next
    Call RefreshExistingTable(SHEET_INFLIGHT, TABLE_INFLIGHT)
    On Error GoTo 0
End Sub
