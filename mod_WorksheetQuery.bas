<<<<<<< HEAD
Attribute VB_Name = "mod_WorksheetQuery"
' ============================================================================
' MODULE: mod_WorksheetQuery (STREAMLINED & OPTIMIZED)
' ============================================================================
' Purpose: Create and refresh Archive and Inflight FLEET-WIDE views
' Author: Data Architecture Team
' Date: 2025-11-13
'
' FEATURES:
'   - Fleet-wide views (ALL SITES visible)
'   - Excel Tables (ListObjects) for native filtering/sorting
'   - No site-specific filtering
'   - No freeze panes
'   - Auto-refresh on workbook open (Inflight only)
' ============================================================================

Option Explicit

' Sheet name constants
Private Const SHEET_ARCHIVE As String = "PIF_Archive"
Private Const SHEET_INFLIGHT As String = "PIF_Inflight"

' ============================================================================
' PUBLIC FUNCTIONS - WORKSHEET CREATION AND REFRESH
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: RefreshArchive
' Purpose: Create or refresh the Archive worksheet (ALL SITES)
' Usage: Call from button or manually
' ----------------------------------------------------------------------------
Public Sub RefreshArchive()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim sql As String
    Dim startTime As Double
    Dim connStr As String
    Dim recordCount As Long

    startTime = Timer
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing Archive..."

    ' Get or create worksheet
    Set ws = GetOrCreateWorksheet(SHEET_ARCHIVE)

    ' Build SQL query - NO SITE FILTER (fleet-wide view)
    sql = "SELECT * FROM dbo.vw_pif_approved_wide ORDER BY approval_date DESC, pif_id, project_id"

    ' Build connection string
    connStr = "OLEDB;Provider=SQLOLEDB;Data Source=" & mod_Database.SQL_SERVER & _
              ";Initial Catalog=" & mod_Database.SQL_DATABASE & _
              ";Integrated Security=SSPI;"

    ' Create Excel Table from query
    Call CreateExcelTableFromQuery(ws, "ArchiveTable", sql, connStr, "PIF Archive - All Sites")

    ' Count records
    If ws.ListObjects.Count > 0 Then
        recordCount = ws.ListObjects(1).ListRows.Count
    End If

    ws.Activate
    Application.ScreenUpdating = True
    Application.StatusBar = False

    Dim elapsed As Double
    elapsed = Timer - startTime

    MsgBox "Archive refreshed!" & vbCrLf & vbCrLf & _
           "Records: " & recordCount & vbCrLf & _
           "Time: " & Format(elapsed, "0.0") & " seconds" & vbCrLf & vbCrLf & _
           "Use native Excel filters and sorting on the table.", _
           vbInformation, "Refresh Complete"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Failed to refresh Archive:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Refresh Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: RefreshInflight
' Purpose: Create or refresh the Inflight worksheet (ALL SITES)
' Usage: Call from button, manually, or Workbook_Open event
' ----------------------------------------------------------------------------
Public Sub RefreshInflight(Optional ByVal showMessage As Boolean = True)
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim sql As String
    Dim startTime As Double
    Dim connStr As String
    Dim recordCount As Long

    startTime = Timer
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing Inflight..."

    ' Get or create worksheet
    Set ws = GetOrCreateWorksheet(SHEET_INFLIGHT)

    ' Build SQL query - NO SITE FILTER (fleet-wide view)
    sql = "SELECT * FROM dbo.vw_pif_inflight_wide ORDER BY submission_date DESC, pif_id, project_id"

    ' Build connection string
    connStr = "OLEDB;Provider=SQLOLEDB;Data Source=" & mod_Database.SQL_SERVER & _
              ";Initial Catalog=" & mod_Database.SQL_DATABASE & _
              ";Integrated Security=SSPI;"

    ' Create Excel Table from query
    Call CreateExcelTableFromQuery(ws, "InflightTable", sql, connStr, "PIF Inflight - All Sites")

    ' Count records
    If ws.ListObjects.Count > 0 Then
        recordCount = ws.ListObjects(1).ListRows.Count
    End If

    ws.Activate
    Application.ScreenUpdating = True
    Application.StatusBar = False

    If showMessage Then
        Dim elapsed As Double
        elapsed = Timer - startTime

        MsgBox "Inflight refreshed!" & vbCrLf & vbCrLf & _
               "Records: " & recordCount & vbCrLf & _
               "Time: " & Format(elapsed, "0.0") & " seconds" & vbCrLf & vbCrLf & _
               "Use native Excel filters and sorting on the table.", _
               vbInformation, "Refresh Complete"
    End If

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If showMessage Then
        MsgBox "Failed to refresh Inflight:" & vbCrLf & vbCrLf & _
               "Error: " & Err.Number & " - " & Err.Description, _
               vbCritical, "Refresh Error"
    End If
End Sub

' ----------------------------------------------------------------------------
' Sub: RefreshAll
' Purpose: Refresh both Archive and Inflight worksheets
' ----------------------------------------------------------------------------
Public Sub RefreshAll(Optional ByVal showSuccessMessage As Boolean = True)
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    Call RefreshArchiveSilent
    Call RefreshInflightSilent

    Application.ScreenUpdating = True

    If showSuccessMessage Then
        MsgBox "Both Archive and Inflight worksheets refreshed.", _
               vbInformation, "Refresh Complete"
    End If

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error refreshing worksheets:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Refresh Error"
End Sub

' ============================================================================
' PRIVATE HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: CreateExcelTableFromQuery
' Purpose: Create Excel Table with QueryTable connection (REFRESHABLE!)
' Parameters:
'   ws - Target worksheet
'   tableName - Name for the Excel Table
'   sql - SQL query string
'   connStr - Connection string
'   title - Title for header
' Notes: Keeps QueryTable connection for native Excel refresh capability
'        Right-click table > Refresh to update data from database
' ----------------------------------------------------------------------------
Private Sub CreateExcelTableFromQuery(ByVal ws As Worksheet, _
                                     ByVal tableName As String, _
                                     ByVal sql As String, _
                                     ByVal connStr As String, _
                                     ByVal title As String)
    On Error GoTo ErrHandler

    Dim qt As QueryTable
    Dim tbl As ListObject
    Dim i As Integer

    ' Clear worksheet
    ws.Cells.Clear

    ' Delete existing ListObjects (Tables) first
    For i = ws.ListObjects.Count To 1 Step -1
        ws.ListObjects(i).Delete
    Next i

    ' Delete existing QueryTables
    For i = ws.QueryTables.Count To 1 Step -1
        ws.QueryTables(i).Delete
    Next i

    ' Add title in row 1
    ws.Range("B1").Value = title
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 14

    ' Add instructions in row 2
    ws.Range("B2").Value = "Right-click table > Refresh to update from database"
    ws.Range("B2").Font.Italic = True
    ws.Range("B2").Font.Size = 9
    ws.Range("B2").Font.Color = RGB(0, 128, 0)

    ' Create QueryTable starting at B4 with Table format
    Set qt = ws.QueryTables.Add( _
        Connection:=connStr, _
        Destination:=ws.Range("B4"), _
        sql:=sql)

    ' Configure QueryTable properties for optimal behavior
    With qt
        .Name = tableName & "_Query"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True  ' KEY: Preserve formatting on refresh!
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True

        ' Refresh to populate data
        .Refresh BackgroundQuery:=False
    End With

    ' Convert QueryTable result to Excel Table (keeps QueryTable connection!)
    If Not qt.ResultRange Is Nothing Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, qt.ResultRange, , xlYes)
        tbl.Name = tableName
        tbl.TableStyle = "TableStyleMedium2"

        ' Format the table header
        With tbl.HeaderRowRange
            .Font.Bold = True
            .Font.Size = 11
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

        ' Auto-fit columns
        tbl.Range.Columns.AutoFit
    End If

    ' DON'T delete the QueryTable - keep it for refresh capability!
    ' The ListObject and QueryTable work together now

    Exit Sub

ErrHandler:
    MsgBox "Error creating Excel Table:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbExclamation, "Table Creation Error"
End Sub

' ----------------------------------------------------------------------------
' Function: GetOrCreateWorksheet
' Purpose: Get existing worksheet or create new one
' Parameters:
'   sheetName - Name of worksheet
' Returns: Worksheet object
' ----------------------------------------------------------------------------
Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    ' Try to get existing worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    ' Create if doesn't exist
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If

    ' Unprotect if protected
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    Set GetOrCreateWorksheet = ws
End Function

' ----------------------------------------------------------------------------
' Sub: RefreshArchiveSilent
' Purpose: Refresh Archive without messagebox (for batch operations)
' ----------------------------------------------------------------------------
Private Sub RefreshArchiveSilent()
    On Error Resume Next

    Dim ws As Worksheet
    Dim sql As String
    Dim connStr As String

    Set ws = GetOrCreateWorksheet(SHEET_ARCHIVE)

    sql = "SELECT * FROM dbo.vw_pif_approved_wide ORDER BY approval_date DESC, pif_id, project_id"

    connStr = "OLEDB;Provider=SQLOLEDB;Data Source=" & mod_Database.SQL_SERVER & _
              ";Initial Catalog=" & mod_Database.SQL_DATABASE & _
              ";Integrated Security=SSPI;"

    Call CreateExcelTableFromQuery(ws, "ArchiveTable", sql, connStr, "PIF Archive - All Sites")

    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: RefreshInflightSilent
' Purpose: Refresh Inflight without messagebox (for batch operations)
' ----------------------------------------------------------------------------
Private Sub RefreshInflightSilent()
    On Error Resume Next

    Dim ws As Worksheet
    Dim sql As String
    Dim connStr As String

    Set ws = GetOrCreateWorksheet(SHEET_INFLIGHT)

    sql = "SELECT * FROM dbo.vw_pif_inflight_wide ORDER BY submission_date DESC, pif_id, project_id"

    connStr = "OLEDB;Provider=SQLOLEDB;Data Source=" & mod_Database.SQL_SERVER & _
              ";Initial Catalog=" & mod_Database.SQL_DATABASE & _
              ";Integrated Security=SSPI;"

    Call CreateExcelTableFromQuery(ws, "InflightTable", sql, connStr, "PIF Inflight - All Sites")

    On Error GoTo 0
End Sub
=======
Attribute VB_Name = "mod_WorksheetQuery"
' ============================================================================
' MODULE: mod_WorksheetQuery (STREAMLINED & OPTIMIZED)
' ============================================================================
' Purpose: Create and refresh Archive and Inflight FLEET-WIDE views
' Author: Data Architecture Team
' Date: 2025-11-13
'
' FEATURES:
'   - Fleet-wide views (ALL SITES visible)
'   - Excel Tables (ListObjects) for native filtering/sorting
'   - No site-specific filtering
'   - No freeze panes
'   - Auto-refresh on workbook open (Inflight only)
' ============================================================================

Option Explicit

' Sheet name constants
Private Const SHEET_ARCHIVE As String = "PIF_Archive"
Private Const SHEET_INFLIGHT As String = "PIF_Inflight"

' ============================================================================
' PUBLIC FUNCTIONS - WORKSHEET CREATION AND REFRESH
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: RefreshArchive
' Purpose: Create or refresh the Archive worksheet (ALL SITES)
' Usage: Call from button or manually
' ----------------------------------------------------------------------------
Public Sub RefreshArchive()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim sql As String
    Dim startTime As Double
    Dim connStr As String
    Dim recordCount As Long

    startTime = Timer
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing Archive..."

    ' Get or create worksheet
    Set ws = GetOrCreateWorksheet(SHEET_ARCHIVE)

    ' Build SQL query - NO SITE FILTER (fleet-wide view)
    sql = "SELECT * FROM dbo.vw_pif_approved_wide ORDER BY approval_date DESC, pif_id, project_id"

    ' Build connection string
    connStr = "OLEDB;Provider=SQLOLEDB;Data Source=" & mod_Database.SQL_SERVER & _
              ";Initial Catalog=" & mod_Database.SQL_DATABASE & _
              ";Integrated Security=SSPI;"

    ' Create Excel Table from query
    Call CreateExcelTableFromQuery(ws, "ArchiveTable", sql, connStr, "PIF Archive - All Sites")

    ' Count records
    If ws.ListObjects.Count > 0 Then
        recordCount = ws.ListObjects(1).ListRows.Count
    End If

    ws.Activate
    Application.ScreenUpdating = True
    Application.StatusBar = False

    Dim elapsed As Double
    elapsed = Timer - startTime

    MsgBox "Archive refreshed!" & vbCrLf & vbCrLf & _
           "Records: " & recordCount & vbCrLf & _
           "Time: " & Format(elapsed, "0.0") & " seconds" & vbCrLf & vbCrLf & _
           "Use native Excel filters and sorting on the table.", _
           vbInformation, "Refresh Complete"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Failed to refresh Archive:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Refresh Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: RefreshInflight
' Purpose: Create or refresh the Inflight worksheet (ALL SITES)
' Usage: Call from button, manually, or Workbook_Open event
' ----------------------------------------------------------------------------
Public Sub RefreshInflight(Optional ByVal showMessage As Boolean = True)
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim sql As String
    Dim startTime As Double
    Dim connStr As String
    Dim recordCount As Long

    startTime = Timer
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing Inflight..."

    ' Get or create worksheet
    Set ws = GetOrCreateWorksheet(SHEET_INFLIGHT)

    ' Build SQL query - NO SITE FILTER (fleet-wide view)
    sql = "SELECT * FROM dbo.vw_pif_inflight_wide ORDER BY submission_date DESC, pif_id, project_id"

    ' Build connection string
    connStr = "OLEDB;Provider=SQLOLEDB;Data Source=" & mod_Database.SQL_SERVER & _
              ";Initial Catalog=" & mod_Database.SQL_DATABASE & _
              ";Integrated Security=SSPI;"

    ' Create Excel Table from query
    Call CreateExcelTableFromQuery(ws, "InflightTable", sql, connStr, "PIF Inflight - All Sites")

    ' Count records
    If ws.ListObjects.Count > 0 Then
        recordCount = ws.ListObjects(1).ListRows.Count
    End If

    ws.Activate
    Application.ScreenUpdating = True
    Application.StatusBar = False

    If showMessage Then
        Dim elapsed As Double
        elapsed = Timer - startTime

        MsgBox "Inflight refreshed!" & vbCrLf & vbCrLf & _
               "Records: " & recordCount & vbCrLf & _
               "Time: " & Format(elapsed, "0.0") & " seconds" & vbCrLf & vbCrLf & _
               "Use native Excel filters and sorting on the table.", _
               vbInformation, "Refresh Complete"
    End If

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If showMessage Then
        MsgBox "Failed to refresh Inflight:" & vbCrLf & vbCrLf & _
               "Error: " & Err.Number & " - " & Err.Description, _
               vbCritical, "Refresh Error"
    End If
End Sub

' ----------------------------------------------------------------------------
' Sub: RefreshAll
' Purpose: Refresh both Archive and Inflight worksheets
' ----------------------------------------------------------------------------
Public Sub RefreshAll(Optional ByVal showSuccessMessage As Boolean = True)
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    Call RefreshArchiveSilent
    Call RefreshInflightSilent

    Application.ScreenUpdating = True

    If showSuccessMessage Then
        MsgBox "Both Archive and Inflight worksheets refreshed.", _
               vbInformation, "Refresh Complete"
    End If

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error refreshing worksheets:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Refresh Error"
End Sub

' ============================================================================
' PRIVATE HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: CreateExcelTableFromQuery
' Purpose: Create Excel Table with QueryTable connection (REFRESHABLE!)
' Parameters:
'   ws - Target worksheet
'   tableName - Name for the Excel Table
'   sql - SQL query string
'   connStr - Connection string
'   title - Title for header
' Notes: Keeps QueryTable connection for native Excel refresh capability
'        Right-click table > Refresh to update data from database
' ----------------------------------------------------------------------------
Private Sub CreateExcelTableFromQuery(ByVal ws As Worksheet, _
                                     ByVal tableName As String, _
                                     ByVal sql As String, _
                                     ByVal connStr As String, _
                                     ByVal title As String)
    On Error GoTo ErrHandler

    Dim qt As QueryTable
    Dim tbl As ListObject
    Dim i As Integer

    ' Clear worksheet
    ws.Cells.Clear

    ' Delete existing ListObjects (Tables) first
    For i = ws.ListObjects.Count To 1 Step -1
        ws.ListObjects(i).Delete
    Next i

    ' Delete existing QueryTables
    For i = ws.QueryTables.Count To 1 Step -1
        ws.QueryTables(i).Delete
    Next i

    ' Add title in row 1
    ws.Range("B1").Value = title
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 14

    ' Add instructions in row 2
    ws.Range("B2").Value = "Right-click table > Refresh to update from database"
    ws.Range("B2").Font.Italic = True
    ws.Range("B2").Font.Size = 9
    ws.Range("B2").Font.Color = RGB(0, 128, 0)

    ' Create QueryTable starting at B4 with Table format
    Set qt = ws.QueryTables.Add( _
        Connection:=connStr, _
        Destination:=ws.Range("B4"), _
        sql:=sql)

    ' Configure QueryTable properties for optimal behavior
    With qt
        .Name = tableName & "_Query"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True  ' KEY: Preserve formatting on refresh!
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True

        ' Refresh to populate data
        .Refresh BackgroundQuery:=False
    End With

    ' Convert QueryTable result to Excel Table (keeps QueryTable connection!)
    If Not qt.ResultRange Is Nothing Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, qt.ResultRange, , xlYes)
        tbl.Name = tableName
        tbl.TableStyle = "TableStyleMedium2"

        ' Format the table header
        With tbl.HeaderRowRange
            .Font.Bold = True
            .Font.Size = 11
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

        ' Auto-fit columns
        tbl.Range.Columns.AutoFit
    End If

    ' DON'T delete the QueryTable - keep it for refresh capability!
    ' The ListObject and QueryTable work together now

    Exit Sub

ErrHandler:
    MsgBox "Error creating Excel Table:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbExclamation, "Table Creation Error"
End Sub

' ----------------------------------------------------------------------------
' Function: GetOrCreateWorksheet
' Purpose: Get existing worksheet or create new one
' Parameters:
'   sheetName - Name of worksheet
' Returns: Worksheet object
' ----------------------------------------------------------------------------
Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    ' Try to get existing worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    ' Create if doesn't exist
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If

    ' Unprotect if protected
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    Set GetOrCreateWorksheet = ws
End Function

' ----------------------------------------------------------------------------
' Sub: RefreshArchiveSilent
' Purpose: Refresh Archive without messagebox (for batch operations)
' ----------------------------------------------------------------------------
Private Sub RefreshArchiveSilent()
    On Error Resume Next

    Dim ws As Worksheet
    Dim sql As String
    Dim connStr As String

    Set ws = GetOrCreateWorksheet(SHEET_ARCHIVE)

    sql = "SELECT * FROM dbo.vw_pif_approved_wide ORDER BY approval_date DESC, pif_id, project_id"

    connStr = "OLEDB;Provider=SQLOLEDB;Data Source=" & mod_Database.SQL_SERVER & _
              ";Initial Catalog=" & mod_Database.SQL_DATABASE & _
              ";Integrated Security=SSPI;"

    Call CreateExcelTableFromQuery(ws, "ArchiveTable", sql, connStr, "PIF Archive - All Sites")

    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: RefreshInflightSilent
' Purpose: Refresh Inflight without messagebox (for batch operations)
' ----------------------------------------------------------------------------
Private Sub RefreshInflightSilent()
    On Error Resume Next

    Dim ws As Worksheet
    Dim sql As String
    Dim connStr As String

    Set ws = GetOrCreateWorksheet(SHEET_INFLIGHT)

    sql = "SELECT * FROM dbo.vw_pif_inflight_wide ORDER BY submission_date DESC, pif_id, project_id"

    connStr = "OLEDB;Provider=SQLOLEDB;Data Source=" & mod_Database.SQL_SERVER & _
              ";Initial Catalog=" & mod_Database.SQL_DATABASE & _
              ";Integrated Security=SSPI;"

    Call CreateExcelTableFromQuery(ws, "InflightTable", sql, connStr, "PIF Inflight - All Sites")

    On Error GoTo 0
End Sub
>>>>>>> 254c5f99beb56416a0a77e8c6100e047b5b2e663
