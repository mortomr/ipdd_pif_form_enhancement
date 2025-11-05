Attribute VB_Name = "mod_SharedConstants"
' ============================================================================
' MODULE: mod_SharedConstants
' ============================================================================
' Purpose: Centralized constants for entire PIF application
' Author: Data Architecture Team
' Date: 2025-11-05
' Version: 2.0.0
'
' IMPROVEMENT: Single source of truth for all constants
' MAINTAINABILITY: Changes to values only need to be made in one location
' ============================================================================

Option Explicit
Option Compare Binary
Option Base 0

' Module version
Private Const MODULE_VERSION As String = "2.0.0"

' ============================================================================
' SHEET NAMES
' ============================================================================
Public Const SHEET_DATA As String = "PIF_Data"
Public Const SHEET_COST_UNPIVOTED As String = "Cost_Unpivoted"
Public Const SHEET_VALIDATION_REPORT As String = "Validation_Report"
Public Const SHEET_ERROR_LOG As String = "Error_Log"

' ============================================================================
' PIF STATUS VALUES
' ============================================================================
Public Const STATUS_APPROVED As String = "Approved"
Public Const STATUS_DISPOSITIONED As String = "Dispositioned"
Public Const STATUS_PENDING As String = "Pending"
Public Const STATUS_PARKED As String = "Parked"
Public Const STATUS_ACTIVE As String = "Active"

' ============================================================================
' CHANGE TYPES
' ============================================================================
Public Const CHANGE_TYPE_NEW As String = "New"
Public Const CHANGE_TYPE_REVISION As String = "Revision"
Public Const CHANGE_TYPE_CANCELLATION As String = "Cancellation"

' ============================================================================
' SCENARIOS
' ============================================================================
Public Const SCENARIO_TARGET As String = "Target"
Public Const SCENARIO_CLOSINGS As String = "Closings"

' ============================================================================
' COLUMN MAPPINGS (EXCEL LAYOUT)
' ============================================================================
' IMPROVEMENT: Enum for type safety and IntelliSense support
Public Enum PIFDataColumns
    colArchive = 3           ' C
    colInclude = 4           ' D
    colAccounting = 5        ' E
    colChangeType = 6        ' F
    colPIFID = 7            ' G
    colSEG = 8              ' H
    colOpco = 9             ' I
    colSite = 10            ' J
    colStrategicRank = 11   ' K
    colFromBlanket = 12     ' L
    colFundingProject = 13  ' M
    colProjectName = 14     ' N
    colOriginalISD = 15     ' O
    colRevisedISD = 16      ' P
    colLCMIssue = 17        ' Q
    colStatus = 18          ' R
    colCategory = 19        ' S
    colJustification = 20   ' T
End Enum

' ============================================================================
' VALIDATION RULES
' ============================================================================
Public Const MAX_PIF_ID_LENGTH As Integer = 16
Public Const MAX_PROJECT_ID_LENGTH As Integer = 10
Public Const MAX_JUSTIFICATION_LENGTH As Integer = 192
Public Const MAX_PROJECT_NAME_LENGTH As Integer = 35
Public Const MAX_STATUS_LENGTH As Integer = 58

' ============================================================================
' DATABASE SETTINGS
' ============================================================================
Public Const DEFAULT_SCHEMA As String = "dbo"
Public Const BATCH_SIZE As Long = 500  ' For bulk operations

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: GetApprovedStatusList
' Purpose: Return SQL-safe list of approved statuses for WHERE IN clauses
' Returns: String formatted for SQL: 'Approved','Dispositioned'
' ----------------------------------------------------------------------------
Public Function GetApprovedStatusList() As String
    ' IMPROVEMENT: Single source of truth for approved statuses
    GetApprovedStatusList = "'" & STATUS_APPROVED & "','" & STATUS_DISPOSITIONED & "'"
End Function

' ----------------------------------------------------------------------------
' Function: GetCurrentFiscalYear
' Purpose: Calculate current fiscal year dynamically
' Returns: Integer year (e.g., 2025)
' ----------------------------------------------------------------------------
Public Function GetCurrentFiscalYear() As Integer
    ' IMPROVEMENT: Dynamic year calculation replaces hard-coded years
    ' Assumes fiscal year = calendar year (adjust if different)
    GetCurrentFiscalYear = Year(Date)
End Function

' ----------------------------------------------------------------------------
' Function: GetFiscalYearEnd
' Purpose: Get fiscal year end date for specified year offset
' Parameters:
'   yearOffset - 0 for current year, 1 for CY+1, etc.
' Returns: Date (e.g., 12/31/2025)
' ----------------------------------------------------------------------------
Public Function GetFiscalYearEnd(ByVal yearOffset As Integer) As Date
    Dim baseYear As Integer
    baseYear = GetCurrentFiscalYear()
    GetFiscalYearEnd = DateSerial(baseYear + yearOffset, 12, 31)
End Function

' ----------------------------------------------------------------------------
' Function: IsApprovedStatus
' Purpose: Check if a status value is considered "approved"
' Parameters:
'   statusValue - Status string to check
' Returns: True if status is approved or dispositioned
' ----------------------------------------------------------------------------
Public Function IsApprovedStatus(ByVal statusValue As String) As Boolean
    Dim upperStatus As String
    upperStatus = UCase(Trim(statusValue))

    IsApprovedStatus = (upperStatus = UCase(STATUS_APPROVED)) Or _
                       (upperStatus = UCase(STATUS_DISPOSITIONED))
End Function

' ----------------------------------------------------------------------------
' Function: GetModuleVersion
' Purpose: Return version of this constants module
' Returns: Version string
' ----------------------------------------------------------------------------
Public Function GetModuleVersion() As String
    GetModuleVersion = MODULE_VERSION
End Function
