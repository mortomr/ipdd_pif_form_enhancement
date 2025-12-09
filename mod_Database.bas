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
Public Const SQL_SERVER As String = "JDCDBETSP1000\PVRN120002" ' prod
'Public Const SQL_SERVER As String = "LITDBETST012\TVNN160002" ' test
Public Const SQL_DATABASE As String = "IPDD"
Public Const SQL_TRUSTED As Boolean = True  ' True = Windows Auth only

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
            Dim parameter As ADODB.parameter
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
    Application.ScreenUpdating = True  ' Ensure error message is visible

    ' IMPROVEMENT: User-friendly error message (no SQL exposure)
    Dim errMsg As String
    errMsg = "Database operation failed." & vbCrLf & vbCrLf & _
             "Error: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf

    ' Add specific guidance for common errors
    If Err.Number = -2147217900 Or InStr(1, Err.Description, "Invalid object name", vbTextCompare) > 0 Then
        errMsg = errMsg & "LIKELY CAUSE: Missing database tables" & vbCrLf & _
                         "ACTION: Run PIF_Database_DDL.sql to create required tables"
    ElseIf Err.Number = -2147467259 Or InStr(1, Err.Description, "Login failed", vbTextCompare) > 0 Then
        errMsg = errMsg & "LIKELY CAUSE: Database connection failure" & vbCrLf & _
                         "ACTION: Check connection settings in mod_Database.bas" & vbCrLf & _
                         "Server: " & SQL_SERVER & vbCrLf & _
                         "Database: " & SQL_DATABASE
    Else
        errMsg = errMsg & "Please contact IT support if this persists."
    End If

    MsgBox errMsg, vbCritical, "Database Error"

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
            Dim parameter As ADODB.parameter
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

    ' CRITICAL IMPROVEMENT: Refresh parameters to ensure correct mapping by name
    dbCommand.Parameters.Refresh

    ' Assign values to existing parameters in the refreshed collection
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

            ' Find the parameter by name and assign its value
            On Error Resume Next  ' Handle case where parameter name might not be found (shouldn't happen with Refresh)
            With dbCommand.Parameters(paramName)
                .value = paramValue
                ' Explicitly set type for BIT fields to ensure 0/1 is sent as TinyInt
                If paramName = "@archive_flag" Or paramName = "@include_flag" Then
                    .Type = adTinyInt
                End If
            End With
            On Error GoTo ErrHandler  ' Resume normal error handling
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

' ----------------------------------------------------------------------------
' Function: ExecuteStoredProcedureNonQuery
' Purpose: Execute stored procedure that doesn't return a recordset (INSERT, UPDATE, DELETE)
' Parameters:
'   dbConnection - Active connection (pass Nothing to create new)
'   procedureName - Name of stored procedure
'   params - Groups of 5: parameter name, type, direction, size, value
' Returns: Boolean - True if successful, False if failed
' SECURITY: Stored procedures are most secure - prevents all injection
' ----------------------------------------------------------------------------
Public Function ExecuteStoredProcedureNonQuery(ByRef dbConnection As ADODB.Connection, _
                                               ByVal procedureName As String, _
                                               ParamArray params() As Variant) As Boolean
    On Error GoTo ErrHandler

    Dim dbCommand As ADODB.Command
    Dim recordsAffected As Long
    Dim closeConnectionAfter As Boolean
    Dim adBit As Boolean
    Dim i As Long
    Dim totalParams As Long
    Dim detailedErrorLog As String

    ' Validate parameter count
    totalParams = UBound(params) - LBound(params) + 1
    If (totalParams Mod 5) <> 0 Then
        Debug.Print "CRITICAL ERROR: Invalid parameter count for " & procedureName
        Debug.Print "  Expected groups of 5 (name, type, direction, size, value)"
        Debug.Print "  Total parameters: " & totalParams
        ExecuteStoredProcedureNonQuery = False
        Exit Function
    End If

    ' Connection management
    closeConnectionAfter = (dbConnection Is Nothing)
    If dbConnection Is Nothing Then
        Set dbConnection = GetDBConnection()
        If dbConnection Is Nothing Then
            ExecuteStoredProcedureNonQuery = False
            Exit Function
        End If
    End If

    ' Setup stored procedure command
    Set dbCommand = New ADODB.Command
    Set dbCommand.ActiveConnection = dbConnection
    dbCommand.CommandText = procedureName
    dbCommand.CommandType = adCmdStoredProc
    dbCommand.CommandTimeout = 300  ' 5 minutes timeout

    ' Add parameters with extensive error checking
    For i = LBound(params) To UBound(params) Step 5
        Dim paramName As String
        Dim paramType As ADODB.DataTypeEnum
        Dim paramDirection As ADODB.ParameterDirectionEnum
        Dim paramSize As Long
        Dim paramValue As Variant

        ' Extract parameter details
        paramName = params(i)
        paramType = params(i + 1)
        paramDirection = params(i + 2)
        paramSize = params(i + 3)
        paramValue = params(i + 4)

        ' Detailed parameter logging
        Debug.Print "  Parameter Details:"
        Debug.Print "    Name: " & paramName
        Debug.Print "    Type: " & paramType
        Debug.Print "    Direction: " & paramDirection
        Debug.Print "    Size: " & paramSize
        Debug.Print "    Value: " & IIf(IsNull(paramValue), "NULL", CStr(paramValue))
        Debug.Print "    VarType: " & VarType(paramValue)

        ' Create parameter
        ' Dim parameter As ADODB.Parameter
        ' On Error Resume Next
        ' Set parameter = dbCommand.CreateParameter(paramName, paramType, paramDirection, paramSize, paramValue)
        
        ' Modify parameter creation logic in ExecuteStoredProcedureNonQuery
        Dim parameter As ADODB.Parameter
        On Error Resume Next
        Select Case paramType
            Case adVarChar
                ' Special handling for justification (VARCHAR(MAX))
                If paramName = "@justification" Then
                    Set parameter = dbCommand.CreateParameter(paramName, paramType, paramDirection, -1, CStr(paramValue))
                Else
                    Set parameter = dbCommand.CreateParameter(paramName, paramType, paramDirection, paramSize, CStr(paramValue))
                End If            
            Case adInteger, adSmallInt, adTinyInt
                ' Integer parameters
                Set parameter = dbCommand.CreateParameter(paramName, paramType, paramDirection, paramSize, CLng(paramValue))
            
            Case adDecimal, adNumeric
                ' Decimal parameters with explicit precision and scale
                Dim decValue As Variant
                If IsNull(paramValue) Then
                    Set parameter = dbCommand.CreateParameter(paramName, paramType, paramDirection, 18, Null)
                Else
                    decValue = CDec(paramValue)
                    Set parameter = dbCommand.CreateParameter(paramName, paramType, paramDirection, 18, decValue)
                End If
                parameter.Precision = 18
                parameter.NumericScale = 2

            Case adBit
                ' Bit (boolean) parameters
                If IsNull(paramValue) Then
                    Set parameter = dbCommand.CreateParameter(paramName, paramType, paramDirection, 0, Null)
                Else
                    Set parameter = dbCommand.CreateParameter(paramName, paramType, paramDirection, 0, CBool(paramValue))
                End If
            Case Else
                ' Default case
                Set parameter = dbCommand.CreateParameter(paramName, paramType, paramDirection, paramSize, paramValue)
        End Select


        ' Capture detailed error if parameter creation fails
        If Err.Number <> 0 Then
            detailedErrorLog = "Parameter Creation Error:" & vbCrLf & _
                               "  Name: " & paramName & vbCrLf & _
                               "  Error: " & Err.Number & " - " & Err.Description & vbCrLf & _
                               "  Type: " & paramType & vbCrLf & _
                               "  Value: " & IIf(IsNull(paramValue), "NULL", CStr(paramValue))
            
            Debug.Print detailedErrorLog
            Err.Clear
            ExecuteStoredProcedureNonQuery = False
            GoTo ErrHandler
        End If
        On Error GoTo ErrHandler

        ' Append parameter
        dbCommand.Parameters.Append parameter
    Next i

    ' Execute the command
    dbCommand.Execute recordsAffected

    ' Cleanup
    If closeConnectionAfter Then
        If dbConnection.State = adStateOpen Then
            dbConnection.Close
        End If
        Set dbConnection = Nothing
    End If

    Set dbCommand = Nothing
    ExecuteStoredProcedureNonQuery = True
    Exit Function

ErrHandler:
    ' Comprehensive error logging
    Dim finalErrorMsg As String
    finalErrorMsg = "Stored Procedure Execution Error:" & vbCrLf & _
                    "  Procedure: " & procedureName & vbCrLf & _
                    "  Error Number: " & Err.Number & vbCrLf & _
                    "  Error Description: " & Err.Description & vbCrLf & _
                    IIf(detailedErrorLog <> "", vbCrLf & "Parameter Details:" & vbCrLf & detailedErrorLog, "")

    Debug.Print finalErrorMsg

    ' Additional error checking for ADO errors
    If Not dbConnection Is Nothing Then
        Dim adoErr As ADODB.Error
        For Each adoErr In dbConnection.Errors
            Debug.Print "ADO Error:"
            Debug.Print "  Number: " & adoErr.Number
            Debug.Print "  Description: " & adoErr.Description
            Debug.Print "  Source: " & adoErr.Source
            Debug.Print "  SQL State: " & adoErr.SQLState
        Next adoErr
    End If

    ' Cleanup
    If Not dbConnection Is Nothing Then
        If dbConnection.State = adStateOpen Then
            On Error Resume Next
            dbConnection.RollbackTrans
            dbConnection.Close
            On Error GoTo 0
        End If
        Set dbConnection = Nothing
    End If

    ExecuteStoredProcedureNonQuery = False
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
' Add a new function to validate site
Public Function ValidateSiteConsistency(ByVal selectedSite As String, ByVal rowSite As String, ByVal pifId As String) As Boolean
    ' Convert both to uppercase for case-insensitive comparison
    selectedSite = UCase(Trim(selectedSite))
    rowSite = UCase(Trim(rowSite))
    
    ' If no site is selected, allow all sites
    If selectedSite = "" Or selectedSite = "FLEET" Then
        ValidateSiteConsistency = True
        Exit Function
    End If
    
    ' First, check if the row's site matches the selected site
    If rowSite <> selectedSite Then
        Debug.Print "SITE MISMATCH:" & vbCrLf & _
                    "  Selected Site: " & selectedSite & vbCrLf & _
                    "  Row Site: " & rowSite & vbCrLf & _
                    "  PIF ID: " & pifId & vbCrLf & _
                    "  Skipping this row."
        ValidateSiteConsistency = False
        Exit Function
    End If
    
    ' Additional PIF ID validation
    If InStr(1, pifId, selectedSite) = 0 Then
        Debug.Print "PIF ID SITE MISMATCH WARNING:" & vbCrLf & _
                    "  PIF ID: " & pifId & vbCrLf & _
                    "  Does not contain site: " & selectedSite
    End If
    
    ValidateSiteConsistency = True
End Function

Public Function BulkInsertToStaging(ByVal dataRange As Range, _
                                    ByVal tableName As String, _
                                    Optional ByVal schemaName As String = "dbo", _
                                    Optional ByVal selectedSite As String = "") As Boolean
    On Error GoTo LogError

    Dim conn As ADODB.Connection
    'Dim selectedSite As String
    Dim i As Long, j As Long
    Dim rowCount As Long
    Dim startTime As Double
    Dim params() As Variant
    Dim wsData As Worksheet
    Dim actualRow As Long
    Dim errorDetailLog As String  ' Variable to capture detailed error info
    Dim rowSite As String
    Dim rowPifId As String

    startTime = Timer

    Debug.Print "=== BulkInsertToStaging STARTED for " & tableName & " ==="
    Debug.Print "Data range rows: " & dataRange.Rows.count

    ' Get the worksheet reference for absolute column access
    Set wsData = dataRange.Worksheet

    Set conn = GetDBConnection()
    If conn Is Nothing Then
        Debug.Print "ERROR: GetDBConnection returned Nothing!"
        errorDetailLog = "Connection failed: GetDBConnection returned Nothing"
        BulkInsertToStaging = False
        GoTo LogError
    End If
    Debug.Print "Database connection established"

    ' Truncate staging table first
    Application.StatusBar = "Truncating " & tableName & "..."
    Debug.Print "Truncating " & tableName & "..."
    If Not ExecuteSQLSecure(conn, "TRUNCATE TABLE " & schemaName & "." & tableName) Then
        Debug.Print "ERROR: Failed to truncate table"
        errorDetailLog = "Failed to truncate table: " & tableName
        BulkInsertToStaging = False
        GoTo LogError
    End If
    Debug.Print "Table truncated successfully"

    ' Loop through Excel range and add records
    Application.StatusBar = "Uploading to " & tableName & "..."
    Application.ScreenUpdating = False
    rowCount = 0

    conn.BeginTrans
    Debug.Print "Transaction started"

    For j = 1 To dataRange.Rows.count
        ' Calculate actual worksheet row
        actualRow = dataRange.row + j - 1

        PrintDetailedRowData wsData, actualRow

        ' Check if row has data (skip empty rows) - use PIF_ID column (H=8)
        If Not IsEmpty(wsData.Cells(actualRow, 8).value) Then
            If tableName = "tbl_pif_projects_staging" Then
                ' Get row-specific site and PIF ID for validation
                rowSite = Trim(wsData.Cells(actualRow, 11).value)    ' Column K = Site
                rowPifId = Trim(wsData.Cells(actualRow, 8).value)    ' Column H = PIF_ID

                ' Validate site consistency if a site is provided
                If selectedSite <> "" Then
                    If Not ValidateSiteConsistency(selectedSite, rowSite, rowPifId) Then
                        Debug.Print "  ERROR: Site validation failed for row " & actualRow
                        
                        errorDetailLog = "Site Validation Failed:" & vbCrLf & _
                                         "Selected Site: " & selectedSite & vbCrLf & _
                                         "Row Site: " & rowSite & vbCrLf & _
                                         "PIF ID: " & rowPifId
                        
                        conn.RollbackTrans
                        BulkInsertToStaging = False
                        GoTo LogError
                    End If
                End If

                ' In BulkInsertToStaging function, modify parameter preparation
                ReDim params(0 To 20) ' 21 parameters for usp_insert_project_staging (added line_item)

                ' Use absolute column references with proper type conversion and NULL handling
                params(0) = SafeString(wsData.Cells(actualRow, 8).Value, 16)   ' pif_id (H) - VARCHAR
                params(1) = SafeString(wsData.Cells(actualRow, 14).Value, 10)  ' project_id (N) - VARCHAR
                params(2) = SafeInteger(wsData.Cells(actualRow, 7).Value)      ' line_item (G) - INT (NEW)
                params(3) = SafeString(wsData.Cells(actualRow, 19).Value, 58)  ' status (S) - VARCHAR

                ' Robust handling of change type
                Dim changeType As Variant
                changeType = SafeString(wsData.Cells(actualRow, 6).Value, 12)
                If IsNull(changeType) Then
                    errorDetailLog = "ERROR: Change Type is NULL for row " & actualRow
                    BulkInsertToStaging = False
                    GoTo LogError
                End If
                params(4) = changeType

                ' Accounting treatment
                params(5) = SafeString(wsData.Cells(actualRow, 5).Value, 14)   ' accounting_treatment (E) - VARCHAR

                ' Category handling
                Dim category As Variant
                category = SafeString(wsData.Cells(actualRow, 20).Value, 26)
                If IsNull(category) Then
                    ' Skip rows with no category
                    Debug.Print "  SKIPPING row " & actualRow & " (No Category)"
                    GoTo NextRow
                End If
                params(6) = category

                params(7) = SafeInteger(wsData.Cells(actualRow, 9).Value)      ' seg (I) - INT
                params(8) = SafeString(wsData.Cells(actualRow, 10).Value, 4)   ' opco (J) - VARCHAR
                params(9) = SafeString(wsData.Cells(actualRow, 11).Value, 4)   ' site (K) - VARCHAR
                params(10) = SafeString(wsData.Cells(actualRow, 12).Value, 26) ' strategic_rank (L) - VARCHAR
                params(11) = SafeString(wsData.Cells(actualRow, 14).Value, 10) ' funding_project (N) - VARCHAR
                params(12) = SafeString(wsData.Cells(actualRow, 15).Value, 35) ' project_name (O) - VARCHAR

                ' Date handling
                params(13) = FormatDateISO(wsData.Cells(actualRow, 16).Value)  ' original_fp_isd (P) - VARCHAR
                params(14) = FormatDateISO(wsData.Cells(actualRow, 17).Value)  ' revised_fp_isd (Q) - VARCHAR

                ' Moving ISD Year handling
                Dim movingIsdYear As Variant
                movingIsdYear = wsData.Cells(actualRow, 40).Value
                If IsEmpty(movingIsdYear) Or IsNull(movingIsdYear) Or Trim(CStr(movingIsdYear)) = "" Or movingIsdYear = 0 Then
                    params(15) = "N"  ' Default to "N"
                Else
                    params(15) = Left(UCase(Trim(CStr(movingIsdYear))), 1)
                End If

                params(16) = SafeString(wsData.Cells(actualRow, 18).Value, 20) ' lcm_issue (R) - VARCHAR

                ' Justification handling
                Dim justification As Variant
                justification = SafeString(wsData.Cells(actualRow, 21).Value, 192)
                If IsNull(justification) Then
                    ' Skip rows with no justification if archived or included
                    If SafeBoolean(wsData.Cells(actualRow, 3).Value) = 1 Or SafeBoolean(wsData.Cells(actualRow, 4).Value) = 1 Then
                        Debug.Print "  SKIPPING row " & actualRow & " (No Justification for Archived/Included)"
                        GoTo NextRow
                    End If
                End If
                params(17) = justification

                params(18) = SafeDecimal(wsData.Cells(actualRow, 41).Value)    ' prior_year_spend (AO) - DECIMAL
                params(19) = SafeBoolean(wsData.Cells(actualRow, 3).Value)     ' archive_flag (C) - BIT
                params(20) = SafeBoolean(wsData.Cells(actualRow, 4).Value)     ' include_flag (D) - BIT
                
                PrintParameterDetails params


                Debug.Print "  Attempting to insert row " & actualRow & ": PIF=" & params(0) & ", Project=" & params(1) & ", Line Item=" & params(2)

                If Not ExecuteStoredProcedureNonQuery(conn, "usp_insert_project_staging", _
                    "@pif_id", adVarChar, adParamInput, 16, params(0), _
                    "@project_id", adVarChar, adParamInput, 10, params(1), _
                    "@line_item", adInteger, adParamInput, 0, params(2), _
                    "@status", adVarChar, adParamInput, 58, params(3), _
                    "@change_type", adVarChar, adParamInput, 12, params(4), _
                    "@accounting_treatment", adVarChar, adParamInput, 14, params(5), _
                    "@category", adVarChar, adParamInput, 26, params(6), _
                    "@seg", adInteger, adParamInput, 0, params(7), _
                    "@opco", adVarChar, adParamInput, 4, params(8), _
                    "@site", adVarChar, adParamInput, 4, params(9), _
                    "@strategic_rank", adVarChar, adParamInput, 26, params(10), _
                    "@funding_project", adVarChar, adParamInput, 10, params(11), _
                    "@project_name", adVarChar, adParamInput, 35, params(12), _
                    "@original_fp_isd", adVarChar, adParamInput, 20, params(13), _
                    "@revised_fp_isd", adVarChar, adParamInput, 20, params(14), _
                    "@moving_isd_year", adChar, adParamInput, 1, params(15), _
                    "@lcm_issue", adVarChar, adParamInput, 20, params(16), _
                    "@justification", adVarChar, adParamInput, 192, params(17), _
                    "@prior_year_spend", adNumeric, adParamInput, 0, params(18), _
                    "@archive_flag", adTinyInt, adParamInput, 0, params(19), _
                    "@include_flag", adTinyInt, adParamInput, 0, params(20)) Then
                    
                    Debug.Print "ERROR DETAILS FOR ROW " & actualRow & ":"
                    PrintDetailedRowData wsData, actualRow
                    
                    ' Capture detailed error information
                    errorDetailLog = "Failed to insert row " & actualRow & vbCrLf & _
                                     "PIF ID: " & params(0) & vbCrLf & _
                                     "Project ID: " & params(1) & vbCrLf & _
                                     "Line Item: " & params(2)
                    
                    conn.RollbackTrans
                    BulkInsertToStaging = False
                    GoTo LogError
                End If

                rowCount = rowCount + 1
            End If
        Else
            Debug.Print "Skipping row " & actualRow & " (PIF_ID is empty)"
        End If
    Next j

    Debug.Print "Loop completed. Total rows processed: " & rowCount
    Debug.Print "Committing transaction..."
    conn.CommitTrans
    Debug.Print "Transaction committed"

    conn.Close
    Set conn = Nothing

    Application.StatusBar = False
    Application.ScreenUpdating = True

    Dim elapsed As Double
    elapsed = Timer - startTime

    Debug.Print "Successfully uploaded " & rowCount & " rows to " & tableName & " in " & Format(elapsed, "0.0") & " seconds"
    Debug.Print "=== BulkInsertToStaging COMPLETED SUCCESSFULLY ==="

    BulkInsertToStaging = True
    Exit Function

LogError:
    ' Enhanced error logging
    Dim finalErrorMsg As String
    finalErrorMsg = "Bulk insert failed:" & vbCrLf & _
                    "Table: " & tableName & vbCrLf & _
                    "Error Details: " & errorDetailLog

    MsgBox finalErrorMsg, vbCritical, "Upload Error"
    
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
            On Error Resume Next
            conn.RollbackTrans
            conn.Close
            On Error GoTo 0
        End If
        Set conn = Nothing
    End If

    BulkInsertToStaging = False
End Function

Private Sub PrintParameterDetails(ByRef params() As Variant)
    Dim i As Long
    Dim paramNames() As Variant
    
    paramNames = Array("pif_id", "project_id", "line_item", "status", "change_type", _
                       "accounting_treatment", "category", "seg", "opco", "site", _
                       "strategic_rank", "funding_project", "project_name", _
                       "original_fp_isd", "revised_fp_isd", "moving_isd_year", _
                       "lcm_issue", "justification", "prior_year_spend", _
                       "archive_flag", "include_flag")
    
    Debug.Print "PARAMETER DETAILS:"
    For i = 0 To UBound(params)
        Debug.Print "  " & paramNames(i) & ": " & _
                    IIf(IsNull(params(i)), "NULL", CStr(params(i)))
    Next i
End Sub

' Public Function BulkInsertToStaging(ByVal dataRange As Range, _
'                                     ByVal tableName As String, _
'                                     Optional ByVal schemaName As String = "dbo") As Boolean
'     On Error GoTo ErrHandler

'     Dim conn As ADODB.Connection
'     Dim i As Long, j As Long
'     Dim rowCount As Long
'     Dim startTime As Double
'     Dim params() As Variant
'     Dim colCount As Integer
'     Dim wsData As Worksheet
'     Dim actualRow As Long

'     startTime = Timer

'     Debug.Print "=== BulkInsertToStaging STARTED for " & tableName & " ==="
'     Debug.Print "Data range rows: " & dataRange.Rows.Count

'     ' Get the worksheet reference for absolute column access
'     Set wsData = dataRange.Worksheet

'     Set conn = GetDBConnection()
'     If conn Is Nothing Then
'         Debug.Print "ERROR: GetDBConnection returned Nothing!"
'         BulkInsertToStaging = False
'         Exit Function
'     End If
'     Debug.Print "Database connection established"
'     If conn Is Nothing Then
'         Debug.Print "CRITICAL: Database connection is NULL!"
'         MsgBox "Failed to establish database connection. Check your connection settings.", vbCritical
'         BulkInsertToStaging = False
'         Exit Function
'     End If

'     If conn.State <> adStateOpen Then
'         Debug.Print "CRITICAL: Database connection is not open!"
'         MsgBox "Database connection is not open. Attempting to reconnect.", vbCritical
'         conn.Open  ' Try to reopen
'         If conn.State <> adStateOpen Then
'             BulkInsertToStaging = False
'             Exit Function
'         End If
'     End If

'     ' Truncate staging table first
'     Application.StatusBar = "Truncating " & tableName & "..."
'     Debug.Print "Truncating " & tableName & "..."
'     If Not ExecuteSQLSecure(conn, "TRUNCATE TABLE " & schemaName & "." & tableName) Then
'         Debug.Print "ERROR: Failed to truncate table"
'         BulkInsertToStaging = False
'         Exit Function
'     End If
'     Debug.Print "Table truncated successfully"

'     ' Loop through Excel range and add records
'     Application.StatusBar = "Uploading to " & tableName & "..."
'     Application.ScreenUpdating = False
'     rowCount = 0

'     conn.BeginTrans
'     Debug.Print "Transaction started"

'     For i = 1 To dataRange.Rows.Count
'         ' Calculate actual worksheet row
'         actualRow = dataRange.Row + i - 1

'         ' Check if row has data (skip empty rows) - use PIF_ID column (H=8)
'         If Not IsEmpty(wsData.Cells(actualRow, 8).Value) Then
'             If tableName = "tbl_pif_projects_staging" Then
'     ReDim params(0 To 20) ' 21 parameters for usp_insert_project_staging (added line_item)
'     ' Use absolute column references with proper type conversion and formatting
'     params(0) = SafeString(wsData.Cells(actualRow, 8).Value)   ' pif_id (H) - VARCHAR
'     params(1) = SafeString(wsData.Cells(actualRow, 14).Value)  ' project_id (N) - VARCHAR
'     params(2) = SafeInteger(wsData.Cells(actualRow, 7).Value)  ' line_item (G) - INT (NEW)
'     params(3) = SafeString(wsData.Cells(actualRow, 19).Value)  ' status (S) - VARCHAR
'     params(4) = SafeString(wsData.Cells(actualRow, 6).Value)   ' change_type (F) - VARCHAR
'     params(5) = SafeString(wsData.Cells(actualRow, 5).Value)   ' accounting_treatment (E) - VARCHAR
'     params(6) = SafeString(wsData.Cells(actualRow, 20).Value)  ' category (T) - VARCHAR
'     params(7) = SafeInteger(wsData.Cells(actualRow, 9).Value)  ' seg (I) - INT
'     params(8) = SafeString(wsData.Cells(actualRow, 10).Value)  ' opco (J) - VARCHAR
'     params(9) = SafeString(wsData.Cells(actualRow, 11).Value)  ' site (K) - VARCHAR
'     params(10) = SafeString(wsData.Cells(actualRow, 12).Value) ' strategic_rank (L) - VARCHAR
'     params(11) = SafeString(wsData.Cells(actualRow, 14).Value) ' funding_project (N) - VARCHAR
'     params(12) = SafeString(wsData.Cells(actualRow, 15).Value) ' project_name (O) - VARCHAR
'     params(13) = FormatDateISO(wsData.Cells(actualRow, 16).Value)  ' original_fp_isd (P) - VARCHAR
'     params(14) = FormatDateISO(wsData.Cells(actualRow, 17).Value)  ' revised_fp_isd (Q) - VARCHAR
'     params(15) = SafeString(wsData.Cells(actualRow, 39).Value) ' moving_isd_year (AN) - CHAR
'     params(16) = SafeString(wsData.Cells(actualRow, 18).Value) ' lcm_issue (R) - VARCHAR
'     params(17) = SafeString(wsData.Cells(actualRow, 21).Value) ' justification (U) - VARCHAR
'     params(18) = SafeDecimal(wsData.Cells(actualRow, 41).Value) ' prior_year_spend (AO) - DECIMAL
'     params(19) = SafeBoolean(wsData.Cells(actualRow, 3).Value)  ' archive_flag (C) - BIT
'     params(20) = SafeBoolean(wsData.Cells(actualRow, 4).Value)  ' include_flag (D) - BIT

'     Debug.Print "  Calling stored procedure with params: pif_id=" & params(0) & ", project_id=" & params(1) & ", line_item=" & params(2)

'     If Not ExecuteStoredProcedureNonQuery(conn, "usp_insert_project_staging", _
'                                 "@pif_id", adVarChar, adParamInput, 16, params(0), _
'                                 "@project_id", adVarChar, adParamInput, 10, params(1), _
'                                 "@line_item", adInteger, adParamInput, 0, params(2), _
'                                 "@status", adVarChar, adParamInput, 58, params(3), _
'                                 "@change_type", adVarChar, adParamInput, 12, params(4), _
'                                 "@accounting_treatment", adVarChar, adParamInput, 14, params(5), _
'                                 "@category", adVarChar, adParamInput, 26, params(6), _
'                                 "@seg", adInteger, adParamInput, 0, params(7), _
'                                 "@opco", adVarChar, adParamInput, 4, params(8), _
'                                 "@site", adVarChar, adParamInput, 4, params(9), _
'                                 "@strategic_rank", adVarChar, adParamInput, 26, params(10), _
'                                 "@funding_project", adVarChar, adParamInput, 10, params(11), _
'                                 "@project_name", adVarChar, adParamInput, 35, params(12), _
'                                 "@original_fp_isd", adVarChar, adParamInput, 20, params(13), _
'                                 "@revised_fp_isd", adVarChar, adParamInput, 20, params(14), _
'                                 "@moving_isd_year", adChar, adParamInput, 1, params(15), _
'                                 "@lcm_issue", adVarChar, adParamInput, 20, params(16), _
'                                 "@justification", adVarChar, adParamInput, 192, params(17), _
'                                 "@prior_year_spend", adNumeric, adParamInput, 0, params(18), _
'                                 "@archive_flag", adTinyInt, adParamInput, 0, params(19), _
'                                 "@include_flag", adTinyInt, adParamInput, 0, params(20)) Then
'         LogDetailedError "BulkInsertToStaging", actualRow, _
'         "Failed to insert project: PIF=" & params(0) & ", Project=" & params(1) & ", Line Item=" & params(2)
'         Debug.Print "  ERROR: ExecuteStoredProcedureNonQuery returned False!"
'         conn.RollbackTrans
'         BulkInsertToStaging = False
'         Exit Function
'     End If
'                 Debug.Print "  Row inserted successfully"
'             ElseIf tableName = "tbl_pif_cost_staging" Then
'                 ReDim params(0 To 7) ' 8 parameters for usp_insert_cost_staging (added line_item)
'                 ' Cost_Unpivoted sheet has columns A-H with proper type conversion
'                 params(0) = SafeString(wsData.Cells(actualRow, 1).Value)  ' pif_id (A) - VARCHAR
'                 params(1) = SafeString(wsData.Cells(actualRow, 2).Value)  ' project_id (B) - VARCHAR
'                 params(2) = SafeInteger(wsData.Cells(actualRow, 3).Value) ' line_item (C) - INT (NEW)
'                 params(3) = SafeString(wsData.Cells(actualRow, 4).Value)  ' scenario (D) - VARCHAR
'                 params(4) = SafeDate(wsData.Cells(actualRow, 5).Value)    ' year (E) - DATE
'                 params(5) = SafeDecimal(wsData.Cells(actualRow, 6).Value) ' requested_value (F) - DECIMAL
'                 params(6) = SafeDecimal(wsData.Cells(actualRow, 7).Value) ' current_value (G) - DECIMAL
'                 params(7) = SafeDecimal(wsData.Cells(actualRow, 8).Value) ' variance_value (H) - DECIMAL

'                 If Not ExecuteStoredProcedureNonQuery(conn, "usp_insert_cost_staging", _
'                                             "@pif_id", 200, 1, 16, params(0), _
'                                             "@project_id", 200, 1, 10, params(1), _
'                                             "@line_item", 3, 1, 0, params(2), _
'                                             "@scenario", 200, 1, 12, params(3), _
'                                             "@year", 7, 1, 0, params(4), _
'                                             "@requested_value", 131, 1, 0, params(5), _
'                                             "@current_value", 131, 1, 0, params(6), _
'                                             "@variance_value", 131, 1, 0, params(7)) Then
'                     conn.RollbackTrans
'                     BulkInsertToStaging = False
'                     Exit Function
'                 End If
'             End If
            
'             rowCount = rowCount + 1

'             ' Progress indicator every 100 rows
'             If rowCount Mod 100 = 0 Then
'                 Application.StatusBar = "Uploaded " & rowCount & " rows to " & tableName & "..."
'                 Debug.Print "Progress: " & rowCount & " rows uploaded"
'             End If
'         Else
'             Debug.Print "Skipping row " & actualRow & " (PIF_ID is empty)"
'         End If
'     Next i

'     Debug.Print "Loop completed. Total rows processed: " & rowCount
'     Debug.Print "Committing transaction..."
'     conn.CommitTrans
'     Debug.Print "Transaction committed"

'     conn.Close
'     Set conn = Nothing

'     Application.StatusBar = False
'     Application.ScreenUpdating = True

'     Dim elapsed As Double
'     elapsed = Timer - startTime

'     Debug.Print "Successfully uploaded " & rowCount & " rows to " & tableName & " in " & Format(elapsed, "0.0") & " seconds"
'     Debug.Print "=== BulkInsertToStaging COMPLETED SUCCESSFULLY ==="

'     BulkInsertToStaging = True
'     Exit Function
    
' ErrHandler:
'     Debug.Print "=== ERROR in BulkInsertToStaging ==="
'     Debug.Print "Error Number: " & Err.Number
'     Debug.Print "Error Description: " & Err.Description
'     Debug.Print "Table: " & tableName
'     Debug.Print "Rows processed before error: " & rowCount

'     Application.StatusBar = False
'     Application.ScreenUpdating = True

'     If Not conn Is Nothing Then
'         If conn.State = adStateOpen Then
'             On Error Resume Next
'             Debug.Print "Rolling back transaction..."
'             conn.RollbackTrans
'             On Error GoTo 0
'             conn.Close
'         End If
'         Set conn = Nothing
'     End If

'     ' Enhanced error message with diagnostic information
'     Dim errMsg As String
'     errMsg = "Bulk insert failed:" & vbCrLf & vbCrLf & _
'              "Error: " & Err.Number & " - " & Err.Description & vbCrLf & _
'              "Table: " & tableName & vbCrLf & _
'              "Rows processed: " & rowCount & vbCrLf & vbCrLf

'     ' Add specific guidance for common errors
'     If Err.Number = -2147217900 Or InStr(1, Err.Description, "Could not find stored procedure", vbTextCompare) > 0 Then
'         errMsg = errMsg & "LIKELY CAUSE: Missing stored procedures" & vbCrLf & _
'                          "ACTION: Run PIF_Database_DDL.sql to create required stored procedures" & vbCrLf & _
'                          "See VERIFY_STORED_PROC.sql to check if procedures exist"
'     ElseIf Err.Number = -2147467259 Or InStr(1, Err.Description, "Login failed", vbTextCompare) > 0 Or _
'            InStr(1, Err.Description, "connect", vbTextCompare) > 0 Then
'         errMsg = errMsg & "LIKELY CAUSE: Database connection failure" & vbCrLf & _
'                          "ACTION: Check database connection settings in mod_Database.bas" & vbCrLf & _
'                          "Server: " & SQL_SERVER & vbCrLf & _
'                          "Database: " & SQL_DATABASE
'     End If

'     MsgBox errMsg, vbCritical, "Upload Error"

'     BulkInsertToStaging = False
' End Function

' ============================================================================
' WRAPPER FUNCTIONS FOR SUBMIT MODULE
' ============================================================================

' ----------------------------------------------------------------------------
' Function: BulkInsertProjects
' Purpose: Wrapper to bulk insert project data
' ----------------------------------------------------------------------------
Public Function BulkInsertProjects(ByVal dataRange As Range, Optional ByVal site As String = "") As Boolean
    ' Pass the site to BulkInsertToStaging
    BulkInsertProjects = mod_Database.BulkInsertToStaging(dataRange, "tbl_pif_projects_staging", "dbo", site)
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
        serverVersion = Left(resultSet.Fields("Version").value, 100)
        dbName = resultSet.Fields("DatabaseName").value
        userName = resultSet.Fields("UserName").value
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
        count = resultSet.Fields("RecordCount").value
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
    SQLSafe = Replace(text, "'", "''")
End Function

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================
Private Sub LogDetailedError(ByVal procedureName As String, ByVal rowNumber As Long, ByVal errorMsg As String)
    Debug.Print "ERROR in " & procedureName & " at row " & rowNumber & ": " & errorMsg
    ' Optionally log to a file or error sheet if needed
End Sub

' Helper function to convert Excel date to ISO format
Public Function FormatDateISO(ByVal dateValue As Variant) As Variant
    If IsEmpty(dateValue) Or IsNull(dateValue) Or Trim(CStr(dateValue)) = "" Then
        FormatDateISO = Null
        Exit Function
    End If

    ' Handle text cases
    Dim strDate As String
    strDate = UCase(Trim(CStr(dateValue)))
    
    Select Case strDate
        Case "ANNUALLY", "ANNUAL", "QUARTERLY", "SEMI-ANNUALLY"
            ' Return the original text for these special cases
            FormatDateISO = strDate
            Exit Function
    End Select

    ' Try to convert to date
    Dim convertedDate As Date
    On Error Resume Next
    convertedDate = CDate(dateValue)
    On Error GoTo 0

    If convertedDate = 0 Then
        ' If not a recognizable date, return original text
        FormatDateISO = strDate
    Else
        FormatDateISO = Format(convertedDate, "yyyy-mm-dd")
    End If
End Function
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
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsLog.Name = SHEET_ERROR_LOG
        wsLog.Visible = xlSheetVeryHidden

        ' Headers
        wsLog.Range("A1:E1").value = Array("Timestamp", "Function", "Error Number", "Description", "Context")
        wsLog.Rows(1).Font.Bold = True
        wsLog.Rows(1).Interior.Color = RGB(200, 200, 200)
    End If

    lastRow = wsLog.Cells(wsLog.Rows.count, 1).End(xlUp).row + 1

    ' Log error details
    wsLog.Cells(lastRow, 1).value = Now
    wsLog.Cells(lastRow, 2).value = functionName
    wsLog.Cells(lastRow, 3).value = errorNumber
    wsLog.Cells(lastRow, 4).value = errorDescription
    wsLog.Cells(lastRow, 5).value = additionalContext

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

' ============================================================================
' DATA TYPE CONVERSION HELPERS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: SafeString
' Purpose: Convert cell value to string, handling NULL/Empty
' Returns: String value or NULL for SQL
' ----------------------------------------------------------------------------
Public Function SafeString(ByVal cellValue As Variant, Optional ByVal maxLength As Integer = 0) As Variant
    If IsEmpty(cellValue) Or IsNull(cellValue) Then
        SafeString = Null
        Exit Function
    End If

    Dim strValue As String
    strValue = Trim(CStr(cellValue))
    
    ' Return NULL if empty after trimming
    If strValue = "" Then
        SafeString = Null
        Exit Function
    End If
    
    ' Special handling for justification (use VARCHAR(MAX))
    If maxLength = 192 Then
        ' For very long text, just ensure it's not truncated
        SafeString = strValue
    ElseIf maxLength > 0 And Len(strValue) > maxLength Then
        SafeString = Left(strValue, maxLength)
    Else
        SafeString = strValue
    End If
End Function

' ----------------------------------------------------------------------------
' Function: SafeInteger
' Purpose: Convert cell value to Integer, handling NULL/Empty
' Returns: Integer value or NULL for SQL
' ----------------------------------------------------------------------------
Public Function SafeInteger(ByVal cellValue As Variant) As Variant
    If IsEmpty(cellValue) Or IsNull(cellValue) Then
        SafeInteger = Null
        Exit Function
    End If

    ' Handle different input types
    Select Case VarType(cellValue)
        Case vbString
            ' Try to convert string to integer
            If IsNumeric(cellValue) Then
                SafeInteger = CLng(cellValue)
            Else
                SafeInteger = Null
            End If
        Case vbInteger, vbLong, vbByte
            SafeInteger = CLng(cellValue)
        Case vbDouble, vbSingle, vbDecimal
            ' Round to nearest integer
            SafeInteger = CLng(Round(CDbl(cellValue), 0))
        Case Else
            SafeInteger = Null
    End Select
End Function

' ----------------------------------------------------------------------------
' Function: SafeDecimal
' Purpose: Convert cell value to Decimal, handling NULL/Empty
' Returns: Decimal value or NULL for SQL
' ----------------------------------------------------------------------------
' Modify SafeDecimal to handle precision
Public Function SafeDecimal(ByVal cellValue As Variant) As Variant
    If IsEmpty(cellValue) Or IsNull(cellValue) Then
        SafeDecimal = Null
        Exit Function
    End If

    ' Handle different input types
    Select Case VarType(cellValue)
        Case vbString
            ' Try to convert string to decimal
            If IsNumeric(cellValue) Then
                SafeDecimal = CDec(Round(CDbl(cellValue), 2))
            Else
                SafeDecimal = Null
            End If
        Case vbInteger, vbLong, vbByte
            SafeDecimal = CDec(cellValue)
        Case vbDouble, vbSingle, vbDecimal
            SafeDecimal = CDec(Round(CDbl(cellValue), 2))
        Case Else
            SafeDecimal = Null
    End Select
End Function

' ----------------------------------------------------------------------------
' Function: SafeBoolean
' Purpose: Convert cell value to Boolean for SQL Server BIT type
' Returns: 1 (True), 0 (False), or NULL for SQL
' Notes: Accepts TRUE/FALSE, 1/0, Y/N, Yes/No, T/F
' ----------------------------------------------------------------------------
Public Function SafeBoolean(ByVal cellValue As Variant) As Variant
    If IsEmpty(cellValue) Or IsNull(cellValue) Then
        SafeBoolean = Null
        Exit Function
    End If

    ' Handle different input types
    Select Case VarType(cellValue)
        Case vbBoolean
            SafeBoolean = IIf(cellValue, 1, 0)
        Case vbInteger, vbLong, vbByte
            SafeBoolean = IIf(cellValue <> 0, 1, 0)
        Case vbString
            Dim strValue As String
            strValue = UCase(Trim(cellValue))
            Select Case strValue
                Case "TRUE", "1", "Y", "YES", "T"
                    SafeBoolean = 1
                Case "FALSE", "0", "N", "NO", "F"
                    SafeBoolean = 0
                Case Else
                    SafeBoolean = Null
            End Select
        Case Else
            SafeBoolean = Null
    End Select
End Function

' ----------------------------------------------------------------------------
' Function: SafeDate
' Purpose: Convert cell value to Date, handling NULL/Empty
' Returns: Date value or NULL for SQL
' ----------------------------------------------------------------------------
Public Function SafeDate(ByVal cellValue As Variant) As Variant
    If IsEmpty(cellValue) Or IsNull(cellValue) Then
        SafeDate = Null
    ElseIf Trim(CStr(cellValue)) = "" Then
        SafeDate = Null
    ElseIf IsDate(cellValue) Then
        SafeDate = CDate(cellValue)
    Else
        SafeDate = Null
    End If
End Function


' Add a new helper function for additional debugging
Private Sub PrintDetailedRowData(ByVal wsData As Worksheet, ByVal actualRow As Long)
    Debug.Print "Detailed Row Data for Row " & actualRow & ":"
    Debug.Print "  PIF ID: " & wsData.Cells(actualRow, 8).Value
    Debug.Print "  Project ID: " & wsData.Cells(actualRow, 14).Value
    Debug.Print "  Line Item: " & wsData.Cells(actualRow, 7).Value
    Debug.Print "  Status: " & wsData.Cells(actualRow, 19).Value
    Debug.Print "  Change Type: " & wsData.Cells(actualRow, 6).Value
    Debug.Print "  Accounting Treatment: " & wsData.Cells(actualRow, 5).Value
    Debug.Print "  Category: " & wsData.Cells(actualRow, 20).Value
    Debug.Print "  SEG: " & wsData.Cells(actualRow, 9).Value
    Debug.Print "  OPCO: " & wsData.Cells(actualRow, 10).Value
    Debug.Print "  Site: " & wsData.Cells(actualRow, 11).Value
    Debug.Print "  Strategic Rank: " & wsData.Cells(actualRow, 12).Value
    Debug.Print "  Funding Project: " & wsData.Cells(actualRow, 14).Value
    Debug.Print "  Project Name: " & wsData.Cells(actualRow, 15).Value
    Debug.Print "  Original FP ISD: " & wsData.Cells(actualRow, 16).Value
    Debug.Print "  Revised FP ISD: " & wsData.Cells(actualRow, 17).Value
    Debug.Print "  Moving ISD Year: " & wsData.Cells(actualRow, 39).Value
    Debug.Print "  LCM Issue: " & wsData.Cells(actualRow, 18).Value
    Debug.Print "  Justification: " & wsData.Cells(actualRow, 21).Value
    Debug.Print "  Prior Year Spend: " & wsData.Cells(actualRow, 41).Value
    Debug.Print "  Archive Flag: " & wsData.Cells(actualRow, 3).Value
    Debug.Print "  Include Flag: " & wsData.Cells(actualRow, 4).Value
End Sub