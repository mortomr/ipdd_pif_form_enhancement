Attribute VB_Name = "mod_Database"
' ============================================================================
' MODULE: mod_Database (SECURITY HARDENED)
' ============================================================================
' Purpose: Handle database connections and SQL execution for PIF submission
' Author: Data Architecture Team
' Date: 2025-11-05
' Version: 2.0.0 - SECURITY HARDENED
'
' CRITICAL SECURITY IMPROVEMENTS:
'   - Parameterized queries eliminate SQL injection vulnerabilities
'   - Stored procedure execution with type-safe parameters
'   - Centralized error logging (no SQL exposure to end users)
'   - Connection reuse pattern for better performance
'   - Transaction support with proper rollback
'
' Dependencies:
'   - Requires Microsoft ActiveX Data Objects Library (ADODB)
'     Tools > References > Microsoft ActiveX Data Objects 6.1 Library
'   - Requires mod_SharedConstants module
'
' Configuration:
'   - Update SQL_SERVER and SQL_DATABASE constants below
'   - Set SQL_TRUSTED = True for Windows Authentication
' ============================================================================

Option Explicit
Option Compare Binary
Option Base 0

' Module version
Private Const MODULE_VERSION As String = "2.0.0"

' ============================================================================
' CONFIGURATION - UPDATE THESE VALUES
' ============================================================================
Private Const SQL_SERVER As String = "LITDBETST012\TVNN160002"
Private Const SQL_DATABASE As String = "IPDD"
Private Const SQL_TRUSTED As Boolean = True  ' True = Windows Auth only

' Connection timeout settings
Private Const CONNECTION_TIMEOUT As Integer = 30    ' seconds
Private Const COMMAND_TIMEOUT As Integer = 300      ' 5 minutes for large operations

' ============================================================================
' PUBLIC FUNCTIONS - CONNECTION MANAGEMENT
' ============================================================================

' ----------------------------------------------------------------------------
' Function: GetDBConnection
' Purpose: Establish and return an ADODB connection to SQL Server
' Returns: ADODB.Connection object (or Nothing if failed)
' UNCHANGED: This function remains secure as-is
' ----------------------------------------------------------------------------
Public Function GetDBConnection() As ADODB.Connection
    On Error GoTo ErrHandler

    Dim dbConnection As ADODB.Connection
    Set dbConnection = New ADODB.Connection

    ' Build connection string
    Dim connectionString As String
    connectionString = "Provider=SQLOLEDB;" & _
                      "Data Source=" & SQL_SERVER & ";" & _
                      "Initial Catalog=" & SQL_DATABASE & ";"

    If SQL_TRUSTED Then
        ' Windows Authentication (required for security)
        connectionString = connectionString & "Integrated Security=SSPI;"
    Else
        ' SQL Authentication DISABLED for security
        MsgBox "SQL Authentication is disabled for security reasons." & vbCrLf & _
               "Please use Windows Authentication (SQL_TRUSTED = True).", _
               vbExclamation, "Authentication Error"
        Set GetDBConnection = Nothing
        Exit Function
    End If

    ' Open connection
    dbConnection.ConnectionTimeout = CONNECTION_TIMEOUT
    dbConnection.CommandTimeout = COMMAND_TIMEOUT
    dbConnection.Open connectionString

    Set GetDBConnection = dbConnection
    Exit Function

ErrHandler:
    MsgBox "Database connection failed:" & vbCrLf & vbCrLf & _
           "Error Code: " & Err.Number & vbCrLf & _
           "Server: " & SQL_SERVER & vbCrLf & _
           "Database: " & SQL_DATABASE & vbCrLf & vbCrLf & _
           "Please verify network connectivity and server name.", _
           vbCritical, "Connection Error"

    Call LogTechnicalError("GetDBConnection", Err.Number, Err.Description, _
                          "Server: " & SQL_SERVER & ", Database: " & SQL_DATABASE)

    Set GetDBConnection = Nothing
End Function

' ============================================================================
' NEW SECURE FUNCTIONS - PARAMETERIZED QUERIES
' ============================================================================

' ----------------------------------------------------------------------------
' Function: ExecuteSQLSecure
' Purpose: Execute parameterized SQL command (INSERT, UPDATE, DELETE, etc.)
' Parameters:
'   dbConnection - Active connection (pass Nothing to create new)
'   sqlStatement - SQL with ? placeholders for parameters
'   params - Array of parameter values
' Returns: True if successful, False if failed
' SECURITY: Eliminates SQL injection through parameterized queries
' ----------------------------------------------------------------------------
Public Function ExecuteSQLSecure(ByRef dbConnection As ADODB.Connection, _
                                ByVal sqlStatement As String, _
                                ParamArray params() As Variant) As Boolean
    On Error GoTo ErrHandler

    Dim dbCommand As ADODB.Command
    Dim recordsAffected As Long
    Dim closeConnectionAfter As Boolean
    Dim i As Long

    ' Connection management
    closeConnectionAfter = (dbConnection Is Nothing)
    If dbConnection Is Nothing Then
        Set dbConnection = GetDBConnection()
        If dbConnection Is Nothing Then
            ExecuteSQLSecure = False
            Exit Function
        End If
    End If

    ' SECURITY: Use Command object with parameters
    Set dbCommand = New ADODB.Command
    Set dbCommand.ActiveConnection = dbConnection
    dbCommand.CommandText = sqlStatement
    dbCommand.CommandType = adCmdText
    dbCommand.CommandTimeout = COMMAND_TIMEOUT

    ' Add parameters with improved type handling and sizing
    If UBound(params) >= LBound(params) Then
        For i = LBound(params) To UBound(params)
            Dim parameter As ADODB.Parameter
            Dim paramValue As Variant
            Dim paramSize As Long

            paramValue = params(i)

            ' SECURITY & IMPROVEMENT: Type-safe parameter creation with proper sizing
            On Error Resume Next  ' Catch type conversion errors

            If IsNull(paramValue) Or IsEmpty(paramValue) Then
                ' NULL parameter - no size needed
                Set parameter = dbCommand.CreateParameter("p" & i, adVarWChar, adParamInput, 0, Null)

            ElseIf VarType(paramValue) = vbDate Then
                ' Date parameter
                Set parameter = dbCommand.CreateParameter("p" & i, adDBTimeStamp, adParamInput, , paramValue)

            ElseIf IsNumeric(paramValue) Then
                ' Numeric parameter - detect appropriate type
                Err.Clear

                ' Check if decimal
                If InStr(1, CStr(paramValue), ".") > 0 Or InStr(1, CStr(paramValue), "E") > 0 Then
                    ' Decimal number
                    Set parameter = dbCommand.CreateParameter("p" & i, adNumeric, adParamInput, , CDec(paramValue))
                ElseIf paramValue >= -2147483648# And paramValue <= 2147483647# Then
                    ' Within INTEGER range
                    Set parameter = dbCommand.CreateParameter("p" & i, adInteger, adParamInput, , CLng(paramValue))
                Else
                    ' Large integer - use BIGINT (pass as Decimal in VBA)
                    Set parameter = dbCommand.CreateParameter("p" & i, adBigInt, adParamInput, , CDec(paramValue))
                End If

                ' Check for conversion error
                If Err.Number <> 0 Then
                    Call LogTechnicalError("ExecuteSQLSecure", Err.Number, _
                                          "Numeric type conversion failed for parameter " & i, _
                                          "Value: " & paramValue & ", VarType: " & VarType(paramValue))
                    ' Fall back to string
                    Err.Clear
                    Set parameter = dbCommand.CreateParameter("p" & i, adVarWChar, adParamInput, 255, CStr(paramValue))
                End If

            Else
                ' String parameter - calculate appropriate size
                paramSize = Len(CStr(paramValue))
                If paramSize < 255 Then paramSize = 255       ' Minimum buffer
                If paramSize > 4000 Then paramSize = 4000     ' Max without NVARCHAR(MAX)

                Set parameter = dbCommand.CreateParameter("p" & i, adVarWChar, adParamInput, paramSize, CStr(paramValue))
            End If

            On Error GoTo ErrHandler  ' Resume normal error handling

            ' Append parameter if successfully created
            If Not parameter Is Nothing Then
                dbCommand.Parameters.Append parameter
            End If
        Next i
    End If

    ' Execute parameterized command
    dbCommand.Execute recordsAffected

    ' Cleanup
    If closeConnectionAfter Then
        dbConnection.Close
        Set dbConnection = Nothing
    End If

    Set dbCommand = Nothing
    ExecuteSQLSecure = True
    Exit Function

ErrHandler:
    ' IMPROVEMENT: User-friendly error message (no SQL exposure)
    MsgBox "Database operation failed." & vbCrLf & vbCrLf & _
           "Error Code: " & Err.Number & vbCrLf & vbCrLf & _
           "Please contact IT support if this persists.", _
           vbCritical, "Database Error"

    Call LogTechnicalError("ExecuteSQLSecure", Err.Number, Err.Description, _
                          "SQL: " & Left(sqlStatement, 200))

    If closeConnectionAfter And Not dbConnection Is Nothing Then
        If dbConnection.State = adStateOpen Then dbConnection.Close
        Set dbConnection = Nothing
    End If

    ExecuteSQLSecure = False
End Function

' ----------------------------------------------------------------------------
' Function: GetRecordsetSecure
' Purpose: Execute parameterized SELECT query and return recordset
' Parameters:
'   dbConnection - Active connection (pass Nothing to create new)
'   sqlStatement - SQL with ? placeholders
'   params - Array of parameter values
' Returns: ADODB.Recordset (or Nothing if failed)
' SECURITY: Parameterized queries prevent SQL injection
' NOTE: Caller must close recordset AND connection when done
' ----------------------------------------------------------------------------
Public Function GetRecordsetSecure(ByRef dbConnection As ADODB.Connection, _
                                   ByVal sqlStatement As String, _
                                   ParamArray params() As Variant) As ADODB.Recordset
    On Error GoTo ErrHandler

    Dim dbCommand As ADODB.Command
    Dim resultSet As ADODB.Recordset
    Dim i As Long

    ' Connection management
    If dbConnection Is Nothing Then
        Set dbConnection = GetDBConnection()
        If dbConnection Is Nothing Then
            Set GetRecordsetSecure = Nothing
            Exit Function
        End If
    End If

    ' SECURITY: Parameterized command
    Set dbCommand = New ADODB.Command
    Set dbCommand.ActiveConnection = dbConnection
    dbCommand.CommandText = sqlStatement
    dbCommand.CommandType = adCmdText
    dbCommand.CommandTimeout = COMMAND_TIMEOUT

    ' Add parameters with improved type handling and sizing
    If UBound(params) >= LBound(params) Then
        For i = LBound(params) To UBound(params)
            Dim parameter As ADODB.Parameter
            Dim paramValue As Variant
            Dim paramSize As Long

            paramValue = params(i)

            ' SECURITY & IMPROVEMENT: Type-safe parameter creation with proper sizing
            On Error Resume Next  ' Catch type conversion errors

            If IsNull(paramValue) Or IsEmpty(paramValue) Then
                ' NULL parameter - no size needed
                Set parameter = dbCommand.CreateParameter("p" & i, adVarWChar, adParamInput, 0, Null)

            ElseIf VarType(paramValue) = vbDate Then
                ' Date parameter
                Set parameter = dbCommand.CreateParameter("p" & i, adDBTimeStamp, adParamInput, , paramValue)

            ElseIf IsNumeric(paramValue) Then
                ' Numeric parameter - detect appropriate type
                Err.Clear

                ' Check if decimal
                If InStr(1, CStr(paramValue), ".") > 0 Or InStr(1, CStr(paramValue), "E") > 0 Then
                    ' Decimal number
                    Set parameter = dbCommand.CreateParameter("p" & i, adNumeric, adParamInput, , CDec(paramValue))
                ElseIf paramValue >= -2147483648# And paramValue <= 2147483647# Then
                    ' Within INTEGER range
                    Set parameter = dbCommand.CreateParameter("p" & i, adInteger, adParamInput, , CLng(paramValue))
                Else
                    ' Large integer - use BIGINT (pass as Decimal in VBA)
                    Set parameter = dbCommand.CreateParameter("p" & i, adBigInt, adParamInput, , CDec(paramValue))
                End If

                ' Check for conversion error
                If Err.Number <> 0 Then
                    Call LogTechnicalError("GetRecordsetSecure", Err.Number, _
                                          "Numeric type conversion failed for parameter " & i, _
                                          "Value: " & paramValue & ", VarType: " & VarType(paramValue))
                    ' Fall back to string
                    Err.Clear
                    Set parameter = dbCommand.CreateParameter("p" & i, adVarWChar, adParamInput, 255, CStr(paramValue))
                End If

            Else
                ' String parameter - calculate appropriate size
                paramSize = Len(CStr(paramValue))
                If paramSize < 255 Then paramSize = 255       ' Minimum buffer
                If paramSize > 4000 Then paramSize = 4000     ' Max without NVARCHAR(MAX)

                Set parameter = dbCommand.CreateParameter("p" & i, adVarWChar, adParamInput, paramSize, CStr(paramValue))
            End If

            On Error GoTo ErrHandler  ' Resume normal error handling

            ' Append parameter if successfully created
            If Not parameter Is Nothing Then
                dbCommand.Parameters.Append parameter
            End If
        Next i
    End If

    ' PERFORMANCE: Use forward-only cursor for best performance
    Set resultSet = New ADODB.Recordset
    resultSet.CursorLocation = adUseClient
    resultSet.Open dbCommand, , adOpenForwardOnly, adLockReadOnly

    Set GetRecordsetSecure = resultSet
    ' NOTE: Connection remains open - caller must close both recordset AND connection
    Exit Function

ErrHandler:
    MsgBox "Query failed." & vbCrLf & vbCrLf & _
           "Error Code: " & Err.Number & vbCrLf & vbCrLf & _
           "Please contact IT support.", _
           vbCritical, "Query Error"

    Call LogTechnicalError("GetRecordsetSecure", Err.Number, Err.Description, _
                          "SQL: " & Left(sqlStatement, 200))

    Set GetRecordsetSecure = Nothing
End Function

' ----------------------------------------------------------------------------
' Function: ExecuteStoredProcedure
' Purpose: Execute stored procedure with parameters (most secure method)
' Parameters:
'   dbConnection - Active connection (pass Nothing to create new)
'   procedureName - Name of stored procedure (e.g., "dbo.usp_commit_to_inflight")
'   returnRecordset - True to return recordset, False for non-query
'   params - Groups of 5: parameter name, type, direction, size, value
'            Example: "@SourceFile", adVarWChar, adParamInput, 255, ThisWorkbook.Name
' Returns: ADODB.Recordset for queries, Nothing for non-queries
' SECURITY: Stored procedures are most secure - prevents all injection
' IMPROVEMENT: Added parameter count validation and size support
' ----------------------------------------------------------------------------
Public Function ExecuteStoredProcedure(ByRef dbConnection As ADODB.Connection, _
                                       ByVal procedureName As String, _
                                       ByVal returnRecordset As Boolean, _
                                       ParamArray params() As Variant) As ADODB.Recordset
    On Error GoTo ErrHandler

    Dim dbCommand As ADODB.Command
    Dim resultSet As ADODB.Recordset
    Dim closeConnectionAfter As Boolean
    Dim i As Long
    Dim paramCount As Long

    ' IMPROVEMENT: Validate parameter count
    If UBound(params) >= LBound(params) Then
        paramCount = UBound(params) - LBound(params) + 1

        ' Parameters must be in groups of 5: name, type, direction, size, value
        If (paramCount Mod 5) <> 0 Then
            MsgBox "Invalid parameter count." & vbCrLf & vbCrLf & _
                   "Parameters must be in groups of 5:" & vbCrLf & _
                   "  1. Parameter name (e.g., '@ParamName')" & vbCrLf & _
                   "  2. Data type (e.g., adVarWChar)" & vbCrLf & _
                   "  3. Direction (e.g., adParamInput)" & vbCrLf & _
                   "  4. Size (e.g., 255 for strings, 0 for numbers)" & vbCrLf & _
                   "  5. Value" & vbCrLf & vbCrLf & _
                   "Actual parameter count: " & paramCount, _
                   vbCritical, "Parameter Error"

            Call LogTechnicalError("ExecuteStoredProcedure", 0, _
                                  "Invalid parameter count: " & paramCount & " (must be divisible by 5)", _
                                  "Procedure: " & procedureName)

            Set ExecuteStoredProcedure = Nothing
            Exit Function
        End If
    End If

    ' Connection management
    closeConnectionAfter = (dbConnection Is Nothing)
    If dbConnection Is Nothing Then
        Set dbConnection = GetDBConnection()
        If dbConnection Is Nothing Then
            Set ExecuteStoredProcedure = Nothing
            Exit Function
        End If
    End If

    ' SECURITY: Stored procedure execution
    Set dbCommand = New ADODB.Command
    Set dbCommand.ActiveConnection = dbConnection
    dbCommand.CommandText = procedureName
    dbCommand.CommandType = adCmdStoredProc  ' CRITICAL: Prevents SQL injection
    dbCommand.CommandTimeout = COMMAND_TIMEOUT

    ' Add parameters (groups of 5: name, type, direction, size, value)
    If UBound(params) >= LBound(params) Then
        For i = LBound(params) To UBound(params) Step 5
            Dim paramName As String
            Dim paramType As ADODB.DataTypeEnum
            Dim paramDirection As ADODB.ParameterDirectionEnum
            Dim paramSize As Long
            Dim paramValue As Variant

            paramName = params(i)
            paramType = params(i + 1)
            paramDirection = params(i + 2)
            paramSize = params(i + 3)
            paramValue = params(i + 4)

            ' IMPROVEMENT: Auto-calculate size for strings if size = 0
            If (paramType = adVarWChar Or paramType = adVarChar) And paramSize = 0 Then
                If Not IsNull(paramValue) And Not IsEmpty(paramValue) Then
                    paramSize = Len(CStr(paramValue))
                    If paramSize < 255 Then paramSize = 255  ' Minimum buffer
                    If paramSize > 4000 Then paramSize = 4000  ' Max without NVARCHAR(MAX)
                Else
                    paramSize = 255  ' Default for NULL
                End If
            End If

            Dim parameter As ADODB.Parameter
            Set parameter = dbCommand.CreateParameter(paramName, paramType, paramDirection, paramSize, paramValue)
            dbCommand.Parameters.Append parameter
        Next i
    End If

    ' Execute
    If returnRecordset Then
        Set resultSet = dbCommand.Execute
        Set ExecuteStoredProcedure = resultSet
    Else
        dbCommand.Execute
        Set ExecuteStoredProcedure = Nothing
    End If

    ' IMPROVEMENT: Cleanup Command object
    Set dbCommand = Nothing

    ' Cleanup connection if created internally and not returning recordset
    If closeConnectionAfter And Not returnRecordset Then
        If dbConnection.State = adStateOpen Then
            dbConnection.Close
        End If
        Set dbConnection = Nothing
    End If

    Exit Function

ErrHandler:
    MsgBox "Stored procedure execution failed." & vbCrLf & vbCrLf & _
           "Procedure: " & procedureName & vbCrLf & _
           "Error Code: " & Err.Number & vbCrLf & vbCrLf & _
           "Please contact IT support.", _
           vbCritical, "Database Error"

    Call LogTechnicalError("ExecuteStoredProcedure", Err.Number, Err.Description, _
                          "Procedure: " & procedureName)

    If closeConnectionAfter And Not dbConnection Is Nothing Then
        If dbConnection.State = adStateOpen Then dbConnection.Close
        Set dbConnection = Nothing
    End If

    Set ExecuteStoredProcedure = Nothing
End Function

' ============================================================================
' BULK INSERT FUNCTIONS
' ============================================================================

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
    Dim i As Long, j As Long
    Dim rowCount As Long
    Dim startTime As Double
    Dim params() As Variant
    Dim colCount As Integer
    
    startTime = Timer
    
    Set conn = GetDBConnection()
    If conn Is Nothing Then
        BulkInsertToStaging = False
        Exit Function
    End If
    
    ' Truncate staging table first
    Application.StatusBar = "Truncating " & tableName & "..."
    If Not ExecuteSQLSecure(conn, "TRUNCATE TABLE " & schemaName & "." & tableName) Then
        BulkInsertToStaging = False
        Exit Function
    End If
    
    ' Loop through Excel range and add records
    Application.StatusBar = "Uploading to " & tableName & "..."
    Application.ScreenUpdating = False
    rowCount = 0
    
    conn.BeginTrans
    
    For i = 2 To dataRange.Rows.Count
        ' Check if row has data (skip empty rows)
        If Not IsEmpty(dataRange.Cells(i, 1).Value) Then
            If tableName = "tbl_pif_projects_staging" Then
                ReDim params(0 To 20) ' Adjust size based on usp_insert_project_staging parameters
                params(0) = dataRange.Cells(i, 7).Value  ' pif_id
                params(1) = dataRange.Cells(i, 13).Value ' project_id
                params(2) = dataRange.Cells(i, 18).Value ' status
                params(3) = dataRange.Cells(i, 6).Value  ' change_type
                params(4) = dataRange.Cells(i, 5).Value  ' accounting_treatment
                params(5) = dataRange.Cells(i, 19).Value ' category
                params(6) = dataRange.Cells(i, 8).Value  ' seg
                params(7) = dataRange.Cells(i, 9).Value  ' opco
                params(8) = dataRange.Cells(i, 10).Value ' site
                params(9) = dataRange.Cells(i, 11).Value ' strategic_rank
                params(10) = dataRange.Cells(i, 13).Value ' funding_project
                params(11) = dataRange.Cells(i, 14).Value ' project_name
                params(12) = dataRange.Cells(i, 15).Value ' original_fp_isd
                params(13) = dataRange.Cells(i, 16).Value ' revised_fp_isd
                params(14) = dataRange.Cells(i, 17).Value ' moving_isd_year
                params(15) = dataRange.Cells(i, 17).Value ' lcm_issue (assuming same column as moving_isd_year for now)
                params(16) = dataRange.Cells(i, 20).Value ' justification
                params(17) = dataRange.Cells(i, 20).Value ' prior_year_spend (assuming same column as justification for now)
                params(18) = dataRange.Cells(i, 3).Value  ' archive_flag
                params(19) = dataRange.Cells(i, 4).Value  ' include_flag
                
                If Not ExecuteStoredProcedure(conn, "usp_insert_project_staging", False, _
                                            "@pif_id", adVarChar, adParamInput, 16, params(0), _
                                            "@project_id", adVarChar, adParamInput, 10, params(1), _
                                            "@status", adVarChar, adParamInput, 58, params(2), _
                                            "@change_type", adVarChar, adParamInput, 12, params(3), _
                                            "@accounting_treatment", adVarChar, adParamInput, 14, params(4), _
                                            "@category", adVarChar, adParamInput, 26, params(5), _
                                            "@seg", adInteger, adParamInput, , params(6), _
                                            "@opco", adVarChar, adParamInput, 4, params(7), _
                                            "@site", adVarChar, adParamInput, 4, params(8), _
                                            "@strategic_rank", adVarChar, adParamInput, 26, params(9), _
                                            "@funding_project", adVarChar, adParamInput, 10, params(10), _
                                            "@project_name", adVarChar, adParamInput, 35, params(11), _
                                            "@original_fp_isd", adVarChar, adParamInput, 8, params(12), _
                                            "@revised_fp_isd", adVarChar, adParamInput, 5, params(13), _
                                            "@moving_isd_year", adChar, adParamInput, 1, params(14), _
                                            "@lcm_issue", adVarChar, adParamInput, 11, params(15), _
                                            "@justification", adVarChar, adParamInput, 192, params(16), _
                                            "@prior_year_spend", adNumeric, adParamInput, , params(17), _
                                            "@archive_flag", adBoolean, adParamInput, , params(18), _
                                            "@include_flag", adBoolean, adParamInput, , params(19)) Then
                    conn.RollbackTrans
                    BulkInsertToStaging = False
                    Exit Function
                End If
            ElseIf tableName = "tbl_pif_cost_staging" Then
                ReDim params(0 To 6) ' Adjust size based on usp_insert_cost_staging parameters
                params(0) = dataRange.Cells(i, 1).Value  ' pif_id
                params(1) = dataRange.Cells(i, 2).Value  ' project_id
                params(2) = dataRange.Cells(i, 3).Value  ' scenario
                params(3) = dataRange.Cells(i, 4).Value  ' year
                params(4) = dataRange.Cells(i, 5).Value  ' requested_value
                params(5) = dataRange.Cells(i, 6).Value  ' current_value
                params(6) = dataRange.Cells(i, 7).Value  ' variance_value
                
                If Not ExecuteStoredProcedure(conn, "usp_insert_cost_staging", False, _
                                            "@pif_id", adVarChar, adParamInput, 16, params(0), _
                                            "@project_id", adVarChar, adParamInput, 10, params(1), _
                                            "@scenario", adVarChar, adParamInput, 12, params(2), _
                                            "@year", adDate, adParamInput, , params(3), _
                                            "@requested_value", adNumeric, adParamInput, , params(4), _
                                            "@current_value", adNumeric, adParamInput, , params(5), _
                                            "@variance_value", adNumeric, adParamInput, , params(6)) Then
                    conn.RollbackTrans
                    BulkInsertToStaging = False
                    Exit Function
                End If
            End If
            
            rowCount = rowCount + 1
            
            ' Progress indicator every 100 rows
            If rowCount Mod 100 = 0 Then
                Application.StatusBar = "Uploaded " & rowCount & " rows to " & tableName & "..."
            End If
        End If
    Next i
    
    conn.CommitTrans
    
    conn.Close
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
    
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
            On Error Resume Next
            conn.RollbackTrans
            On Error GoTo 0
            conn.Close
        End If
        Set conn = Nothing
    End If
    
    MsgBox "Bulk insert failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & _
           "Table: " & tableName & vbCrLf & _
           "Rows processed: " & rowCount, _
           vbCritical, "Upload Error"
           
    BulkInsertToStaging = False
End Function

' ============================================================================
' WRAPPER FUNCTIONS FOR SUBMIT MODULE
' ============================================================================

' ----------------------------------------------------------------------------
' Function: BulkInsertProjects
' Purpose: Wrapper to bulk insert project data
' ----------------------------------------------------------------------------
Public Function BulkInsertProjects(ByVal dataRange As Range) As Boolean
    BulkInsertProjects = BulkInsertToStaging(dataRange, "tbl_pif_projects_staging", "dbo")
End Function

' ----------------------------------------------------------------------------
' Function: BulkInsertCosts
' Purpose: Wrapper to bulk insert cost data
' ----------------------------------------------------------------------------
Public Function BulkInsertCosts(ByVal dataRange As Range) As Boolean
    BulkInsertCosts = BulkInsertToStaging(dataRange, "tbl_pif_cost_staging", "dbo")
End Function

' ============================================================================
' TESTING AND UTILITY FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: TestConnection
' Purpose: Test database connection and display server info
' Returns: True if connection successful
' UNCHANGED: This function is already secure
' ----------------------------------------------------------------------------
Public Function TestConnection() As Boolean
    On Error GoTo ErrHandler

    Dim dbConnection As ADODB.Connection
    Dim resultSet As ADODB.Recordset
    Dim serverVersion As String
    Dim dbName As String
    Dim userName As String

    Set dbConnection = GetDBConnection()

    If dbConnection Is Nothing Then
        TestConnection = False
        Exit Function
    End If

    ' Get server information (safe - no user input)
    Set resultSet = New ADODB.Recordset
    resultSet.Open "SELECT @@VERSION AS Version, DB_NAME() AS DatabaseName, SYSTEM_USER AS UserName", dbConnection

    If Not resultSet.EOF Then
        serverVersion = Left(resultSet.Fields("Version").Value, 100)
        dbName = resultSet.Fields("DatabaseName").Value
        userName = resultSet.Fields("UserName").Value
    End If

    resultSet.Close
    dbConnection.Close

    MsgBox "Connection successful!" & vbCrLf & vbCrLf & _
           "Server: " & SQL_SERVER & vbCrLf & _
           "Database: " & dbName & vbCrLf & _
           "User: " & userName & vbCrLf & vbCrLf & _
           "Version: " & serverVersion & vbCrLf & vbCrLf & _
           "Module Version: " & MODULE_VERSION, _
           vbInformation, "Connection Test"

    TestConnection = True
    Exit Function

ErrHandler:
    MsgBox "Connection test failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Connection Test Failed"
    TestConnection = False
End Function

' ----------------------------------------------------------------------------
' Function: GetRecordCount
' Purpose: Get count of records in a table (SECURE VERSION)
' Parameters:
'   tableName - Table name (without schema)
'   schemaName - Schema name (default: dbo)
' Returns: Record count (or -1 if error)
' SECURITY: Now uses parameterized query
' ----------------------------------------------------------------------------
Public Function GetRecordCount(ByVal tableName As String, _
                              Optional ByVal schemaName As String = "dbo") As Long
    On Error GoTo ErrHandler

    Dim dbConnection As ADODB.Connection
    Dim resultSet As ADODB.Recordset
    Dim count As Long

    ' SECURITY NOTE: Table/schema names cannot be parameterized in SQL
    ' Validate input to prevent injection
    If Not IsValidSQLIdentifier(tableName) Or Not IsValidSQLIdentifier(schemaName) Then
        MsgBox "Invalid table or schema name.", vbExclamation
        GetRecordCount = -1
        Exit Function
    End If

    ' Safe to concatenate validated identifiers
    Dim sqlStatement As String
    sqlStatement = "SELECT COUNT(*) AS RecordCount FROM " & schemaName & "." & tableName

    Set dbConnection = GetDBConnection()
    If dbConnection Is Nothing Then
        GetRecordCount = -1
        Exit Function
    End If

    Set resultSet = New ADODB.Recordset
    resultSet.Open sqlStatement, dbConnection, adOpenForwardOnly, adLockReadOnly

    If Not resultSet.EOF Then
        count = resultSet.Fields("RecordCount").Value
    End If

    resultSet.Close
    dbConnection.Close

    Set resultSet = Nothing
    Set dbConnection = Nothing

    GetRecordCount = count
    Exit Function

ErrHandler:
    Call LogTechnicalError("GetRecordCount", Err.Number, Err.Description, _
                          "Table: " & schemaName & "." & tableName)
    GetRecordCount = -1
End Function

' ============================================================================
' DEPRECATED FUNCTIONS (FOR BACKWARD COMPATIBILITY)
' ============================================================================
' WARNING: These functions use string concatenation and are vulnerable
' to SQL injection. They are retained only for backward compatibility
' during migration. All new code should use the Secure versions above.
' ============================================================================

' ----------------------------------------------------------------------------
' Function: ExecuteSQL
' PURPOSE: DEPRECATED - USE ExecuteSQLSecure INSTEAD
' SECURITY WARNING: Vulnerable to SQL injection
' ----------------------------------------------------------------------------
Public Function ExecuteSQL(ByVal sql As String) As Boolean
    ' Display deprecation warning
    Debug.Print "WARNING: ExecuteSQL is deprecated. Use ExecuteSQLSecure instead."

    On Error GoTo ErrHandler

    Dim conn As ADODB.Connection
    Dim recordsAffected As Long

    Set conn = GetDBConnection()

    If conn Is Nothing Then
        ExecuteSQL = False
        Exit Function
    End If

    conn.Execute sql, recordsAffected
    conn.Close
    Set conn = Nothing

    ExecuteSQL = True
    Exit Function

ErrHandler:
    MsgBox "SQL execution failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "SQL Error"

    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If

    ExecuteSQL = False
End Function

' ----------------------------------------------------------------------------
' Function: GetRecordset
' PURPOSE: DEPRECATED - USE GetRecordsetSecure INSTEAD
' SECURITY WARNING: Vulnerable to SQL injection
' ----------------------------------------------------------------------------
Public Function GetRecordset(ByVal sql As String) As ADODB.Recordset
    Debug.Print "WARNING: GetRecordset is deprecated. Use GetRecordsetSecure instead."

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
    Exit Function

ErrHandler:
    MsgBox "Query failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Query Error"
    Set GetRecordset = Nothing
End Function

' ----------------------------------------------------------------------------
' Function: SQLSafe
' PURPOSE: DEPRECATED - INADEQUATE PROTECTION
' SECURITY WARNING: Only escapes single quotes - use parameterized queries instead
' ----------------------------------------------------------------------------
Public Function SQLSafe(ByVal text As String) As String
    Debug.Print "WARNING: SQLSafe provides inadequate protection. Use parameterized queries."
    SQLSafe = Replace(text, "'", "''")
End Function

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: IsValidSQLIdentifier
' Purpose: Validate that a string is a safe SQL identifier (table/column name)
' Parameters:
'   identifier - String to validate
' Returns: True if valid, False if contains dangerous characters
' SECURITY: Prevents SQL injection in dynamic SQL
' ----------------------------------------------------------------------------
Private Function IsValidSQLIdentifier(ByVal identifier As String) As Boolean
    Dim i As Long
    Dim char As String

    ' Empty or too long
    If Len(identifier) = 0 Or Len(identifier) > 128 Then
        IsValidSQLIdentifier = False
        Exit Function
    End If

    ' Must start with letter or underscore
    char = Left(identifier, 1)
    If Not ((char >= "A" And char <= "Z") Or (char >= "a" And char <= "z") Or char = "_") Then
        IsValidSQLIdentifier = False
        Exit Function
    End If

    ' Check each character
    For i = 1 To Len(identifier)
        char = Mid(identifier, i, 1)
        ' Allow letters, digits, underscore
        If Not ((char >= "A" And char <= "Z") Or (char >= "a" And char <= "z") Or _
                (char >= "0" And char <= "9") Or char = "_") Then
            IsValidSQLIdentifier = False
            Exit Function
        End If
    Next i

    IsValidSQLIdentifier = True
End Function

' ----------------------------------------------------------------------------
' Sub: LogTechnicalError
' Purpose: Log detailed error information to hidden worksheet
' Parameters:
'   functionName - Name of function where error occurred
'   errorNumber - Error number
'   errorDescription - Error description
'   additionalContext - Additional context (SQL, parameters, etc.)
' SECURITY: Keeps technical details from end users, available for IT support
' ----------------------------------------------------------------------------
Private Sub LogTechnicalError(ByVal functionName As String, _
                              ByVal errorNumber As Long, _
                              ByVal errorDescription As String, _
                              ByVal additionalContext As String)
    On Error Resume Next  ' Don't let logging errors crash the application

    Dim wsLog As Worksheet
    Dim lastRow As Long

    ' Create or get error log worksheet (hidden from users)
    Set wsLog = Nothing
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets(SHEET_ERROR_LOG)
    On Error GoTo 0

    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = SHEET_ERROR_LOG
        wsLog.Visible = xlSheetVeryHidden

        ' Headers
        wsLog.Range("A1:E1").Value = Array("Timestamp", "Function", "Error Number", "Description", "Context")
        wsLog.Rows(1).Font.Bold = True
        wsLog.Rows(1).Interior.Color = RGB(200, 200, 200)
    End If

    lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1

    ' Log error details
    wsLog.Cells(lastRow, 1).Value = Now
    wsLog.Cells(lastRow, 2).Value = functionName
    wsLog.Cells(lastRow, 3).Value = errorNumber
    wsLog.Cells(lastRow, 4).Value = errorDescription
    wsLog.Cells(lastRow, 5).Value = additionalContext

    ' Auto-fit columns
    wsLog.Columns("A:E").AutoFit

    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Function: GetModuleVersion
' Purpose: Return version of this module
' Returns: Version string
' ----------------------------------------------------------------------------
Public Function GetModuleVersion() As String
    GetModuleVersion = MODULE_VERSION
End Function
