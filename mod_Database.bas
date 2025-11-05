Attribute VB_Name = "mod_Database"
' ============================================================================
' MODULE: mod_Database
' ============================================================================
' Purpose: Handle database connections and SQL execution for PIF submission
' Author: Data Architecture Team
' Date: 2025-11-05
'
' Dependencies:
'   - Requires Microsoft ActiveX Data Objects Library (ADODB)
'     Tools > References > Microsoft ActiveX Data Objects 6.1 Library
'
' Configuration:
'   - Update SQL_SERVER and SQL_DATABASE constants below
'   - Set SQL_TRUSTED = True for Windows Authentication
'   - Set SQL_TRUSTED = False for SQL Authentication (not recommended)
' ============================================================================

Option Explicit

' ============================================================================
' CONFIGURATION - UPDATE THESE VALUES
' ============================================================================
Private Const SQL_SERVER As String = "YOUR_SQL_SERVER_NAME"
Private Const SQL_DATABASE As String = "YOUR_DATABASE_NAME"
Private Const SQL_TRUSTED As Boolean = True  ' True = Windows Auth, False = SQL Auth

' Connection timeout settings
Private Const CONNECTION_TIMEOUT As Integer = 30    ' seconds
Private Const COMMAND_TIMEOUT As Integer = 300      ' 5 minutes for large operations

' ============================================================================
' PUBLIC FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: GetDBConnection
' Purpose: Establish and return an ADODB connection to SQL Server
' Returns: ADODB.Connection object (or Nothing if failed)
' ----------------------------------------------------------------------------
Public Function GetDBConnection() As ADODB.Connection
    On Error GoTo ErrHandler
    
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    
    ' Build connection string
    Dim connStr As String
    connStr = "Provider=SQLOLEDB;" & _
              "Data Source=" & SQL_SERVER & ";" & _
              "Initial Catalog=" & SQL_DATABASE & ";"
    
    If SQL_TRUSTED Then
        ' Windows Authentication (recommended)
        connStr = connStr & "Integrated Security=SSPI;"
    Else
        ' SQL Authentication (not recommended - credentials exposed)
        ' If using SQL Auth, store credentials securely (e.g., encrypted or prompted)
        MsgBox "SQL Authentication not implemented for security reasons." & vbCrLf & _
               "Please use Windows Authentication (SQL_TRUSTED = True).", vbExclamation
        Set GetDBConnection = Nothing
        Exit Function
    End If
    
    ' Open connection
    conn.ConnectionTimeout = CONNECTION_TIMEOUT
    conn.CommandTimeout = COMMAND_TIMEOUT
    conn.Open connStr
    
    Set GetDBConnection = conn
    Exit Function
    
ErrHandler:
    MsgBox "Database connection failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "Server: " & SQL_SERVER & vbCrLf & _
           "Database: " & SQL_DATABASE, _
           vbCritical, "Connection Error"
    Set GetDBConnection = Nothing
End Function

' ----------------------------------------------------------------------------
' Function: ExecuteSQL
' Purpose: Execute a SQL command that doesn't return results (INSERT, UPDATE, DELETE, etc.)
' Parameters:
'   sql - SQL statement to execute
' Returns: True if successful, False if failed
' ----------------------------------------------------------------------------
Public Function ExecuteSQL(ByVal sql As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim conn As ADODB.Connection
    Dim recordsAffected As Long
    
    Set conn = GetDBConnection()
    
    If conn Is Nothing Then
        ExecuteSQL = False
        Exit Function
    End If
    
    ' Execute the SQL
    conn.Execute sql, recordsAffected
    
    ' Close connection
    conn.Close
    Set conn = Nothing
    
    ExecuteSQL = True
    Exit Function
    
ErrHandler:
    MsgBox "SQL execution failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "SQL: " & Left(sql, 500), _
           vbCritical, "SQL Error"
    
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    
    ExecuteSQL = False
End Function

' ----------------------------------------------------------------------------
' Function: GetRecordset
' Purpose: Execute a SQL query and return results as recordset
' Parameters:
'   sql - SELECT statement to execute
' Returns: ADODB.Recordset (or Nothing if failed)
' Note: Caller is responsible for closing recordset and connection
' ----------------------------------------------------------------------------
Public Function GetRecordset(ByVal sql As String) As ADODB.Recordset
    On Error GoTo ErrHandler
    
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set conn = GetDBConnection()
    If conn Is Nothing Then
        Set GetRecordset = Nothing
        Exit Function
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    Set GetRecordset = rs
    ' Note: Connection remains open for recordset to use
    Exit Function
    
ErrHandler:
    MsgBox "Query failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "SQL: " & Left(sql, 500), _
           vbCritical, "Query Error"
    Set GetRecordset = Nothing
End Function

' ----------------------------------------------------------------------------
' Function: BulkInsertToStaging
' Purpose: Efficiently upload Excel range to SQL staging table
' Parameters:
'   dataRange - Excel range containing data (assumes row 1 is headers)
'   tableName - Name of staging table (without schema)
'   columnMapping - Dictionary mapping Excel column index to SQL field name
' Returns: True if successful, False if failed
' ----------------------------------------------------------------------------
Public Function BulkInsertToStaging(ByVal dataRange As Range, _
                                    ByVal tableName As String, _
                                    Optional ByVal schemaName As String = "dbo") As Boolean
    On Error GoTo ErrHandler
    
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim i As Long, j As Long
    Dim sql As String
    Dim rowCount As Long
    Dim startTime As Double
    
    startTime = Timer
    
    Set conn = GetDBConnection()
    If conn Is Nothing Then
        BulkInsertToStaging = False
        Exit Function
    End If
    
    ' Truncate staging table first
    Application.StatusBar = "Truncating " & tableName & "..."
    sql = "TRUNCATE TABLE " & schemaName & "." & tableName
    conn.Execute sql
    
    ' Open recordset on staging table
    sql = "SELECT * FROM " & schemaName & "." & tableName & " WHERE 1=0"
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    
    ' Get field names from recordset
    Dim fieldNames() As String
    ReDim fieldNames(0 To rs.Fields.Count - 1)
    For j = 0 To rs.Fields.Count - 1
        fieldNames(j) = rs.Fields(j).Name
    Next j
    
    ' Loop through Excel range and add records
    Application.StatusBar = "Uploading to " & tableName & "..."
    Application.ScreenUpdating = False
    rowCount = 0
    
    Dim colCount As Integer
    colCount = dataRange.Columns.Count
    If colCount > rs.Fields.Count Then colCount = rs.Fields.Count
    
    ' Skip header row, process data rows
    For i = 2 To dataRange.Rows.Count
        ' Check if row has data (skip empty rows)
        If Not IsEmpty(dataRange.Cells(i, 1).Value) Then
            rs.AddNew
            
            ' Map Excel columns to recordset fields
            For j = 1 To colCount
                Dim cellValue As Variant
                cellValue = dataRange.Cells(i, j).Value
                
                ' Handle empty cells
                If IsEmpty(cellValue) Then
                    rs.Fields(j - 1).Value = Null
                Else
                    rs.Fields(j - 1).Value = cellValue
                End If
            Next j
            
            rs.Update
            rowCount = rowCount + 1
            
            ' Progress indicator every 100 rows
            If rowCount Mod 100 = 0 Then
                Application.StatusBar = "Uploaded " & rowCount & " rows to " & tableName & "..."
            End If
        End If
    Next i
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Dim elapsed As Double
    elapsed = Timer - startTime
    
    MsgBox "Successfully uploaded " & rowCount & " rows to " & tableName & vbCrLf & _
           "Elapsed time: " & Format(elapsed, "0.0") & " seconds", _
           vbInformation, "Upload Complete"
    
    BulkInsertToStaging = True
    Exit Function
    
ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Bulk insert failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & _
           "Table: " & tableName & vbCrLf & _
           "Rows processed: " & rowCount, _
           vbCritical, "Upload Error"
    
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    
    BulkInsertToStaging = False
End Function

' ----------------------------------------------------------------------------
' Function: TestConnection
' Purpose: Test database connection and display server info
' Returns: True if connection successful
' ----------------------------------------------------------------------------
Public Function TestConnection() As Boolean
    On Error GoTo ErrHandler
    
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim serverVersion As String
    Dim dbName As String
    Dim userName As String
    
    Set conn = GetDBConnection()
    
    If conn Is Nothing Then
        TestConnection = False
        Exit Function
    End If
    
    ' Get server information
    Set rs = New ADODB.Recordset
    rs.Open "SELECT @@VERSION AS Version, DB_NAME() AS DatabaseName, SYSTEM_USER AS UserName", conn
    
    If Not rs.EOF Then
        serverVersion = Left(rs.Fields("Version").Value, 100)
        dbName = rs.Fields("DatabaseName").Value
        userName = rs.Fields("UserName").Value
    End If
    
    rs.Close
    conn.Close
    
    MsgBox "Connection successful!" & vbCrLf & vbCrLf & _
           "Server: " & SQL_SERVER & vbCrLf & _
           "Database: " & dbName & vbCrLf & _
           "User: " & userName & vbCrLf & vbCrLf & _
           "Version: " & serverVersion, _
           vbInformation, "Connection Test"
    
    TestConnection = True
    Exit Function
    
ErrHandler:
    MsgBox "Connection test failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Connection Test Failed"
    TestConnection = False
End Function

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: SQLSafe
' Purpose: Escape single quotes in strings for SQL injection prevention
' Parameters:
'   text - String to make SQL-safe
' Returns: SQL-safe string
' ----------------------------------------------------------------------------
Public Function SQLSafe(ByVal text As String) As String
    SQLSafe = Replace(text, "'", "''")
End Function

' ----------------------------------------------------------------------------
' Function: GetRecordCount
' Purpose: Get count of records in a table
' Parameters:
'   tableName - Table name (without schema)
'   schemaName - Schema name (default: dbo)
' Returns: Record count (or -1 if error)
' ----------------------------------------------------------------------------
Public Function GetRecordCount(ByVal tableName As String, _
                                Optional ByVal schemaName As String = "dbo") As Long
    On Error GoTo ErrHandler
    
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim count As Long
    
    sql = "SELECT COUNT(*) AS RecordCount FROM " & schemaName & "." & tableName
    Set rs = GetRecordset(sql)
    
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            count = rs.Fields("RecordCount").Value
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    GetRecordCount = count
    Exit Function
    
ErrHandler:
    GetRecordCount = -1
End Function
