<<<<<<< HEAD
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
' ADODB CONSTANTS (in case they're missing from your ADODB library)
' ============================================================================
' If you get "Variable not defined" errors for these, uncomment the ones you need
' Public Const adBit As Integer = 128
' Public Const adTinyInt As Integer = 16
' Public Const adUnsignedTinyInt As Integer = 17

' ============================================================================
' SHEET NAMES
' ============================================================================
' NOTE: Update these if your Excel workbook uses different sheet names
Public Const SHEET_DATA As String = "PIF"
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
    colFundingProject = 13  ' M (also used as project_id in database)
    colProjectName = 14     ' N
    colOriginalISD = 15     ' O
    colRevisedISD = 16      ' P
    colLCMIssue = 17        ' Q
    colStatus = 18          ' R
    colCategory = 19        ' S
    colJustification = 20   ' T
    ' Cost columns U-BF (21-58) - see cost column constants below
    colMovingISDYear = 39   ' AM
    colPriorYearSpend = 40  ' AN
End Enum

' ============================================================================
' COST COLUMN MAPPINGS (EXCEL LAYOUT)
' ============================================================================
' Target Scenario columns
Public Const COL_TARGET_REQ_CY As Integer = 21      ' U (2025)
Public Const COL_TARGET_REQ_CY1 As Integer = 22     ' V (2026)
Public Const COL_TARGET_REQ_CY2 As Integer = 23     ' W (2027)
Public Const COL_TARGET_REQ_CY3 As Integer = 24     ' X (2028)
Public Const COL_TARGET_REQ_CY4 As Integer = 25     ' Y (2029)
Public Const COL_TARGET_REQ_CY5 As Integer = 26     ' Z (2030)

Public Const COL_TARGET_APPR_CY As Integer = 27     ' AA (2025)
Public Const COL_TARGET_APPR_CY1 As Integer = 28    ' AB (2026)
Public Const COL_TARGET_APPR_CY2 As Integer = 29    ' AC (2027)
Public Const COL_TARGET_APPR_CY3 As Integer = 30    ' AD (2028)
Public Const COL_TARGET_APPR_CY4 As Integer = 31    ' AE (2029)
Public Const COL_TARGET_APPR_CY5 As Integer = 32    ' AF (2030)

Public Const COL_TARGET_VAR_CY As Integer = 33      ' AG (2025)
Public Const COL_TARGET_VAR_CY1 As Integer = 34     ' AH (2026)
Public Const COL_TARGET_VAR_CY2 As Integer = 35     ' AI (2027)
Public Const COL_TARGET_VAR_CY3 As Integer = 36     ' AJ (2028)
Public Const COL_TARGET_VAR_CY4 As Integer = 37     ' AK (2029)
Public Const COL_TARGET_VAR_CY5 As Integer = 38     ' AL (2030)

' Closings Scenario columns
Public Const COL_CLOSINGS_REQ_CY As Integer = 41    ' AO (2025)
Public Const COL_CLOSINGS_REQ_CY1 As Integer = 42   ' AP (2026)
Public Const COL_CLOSINGS_REQ_CY2 As Integer = 43   ' AQ (2027)
Public Const COL_CLOSINGS_REQ_CY3 As Integer = 44   ' AR (2028)
Public Const COL_CLOSINGS_REQ_CY4 As Integer = 45   ' AS (2029)
Public Const COL_CLOSINGS_REQ_CY5 As Integer = 46   ' AT (2030)

Public Const COL_CLOSINGS_APPR_CY As Integer = 47   ' AU (2025)
Public Const COL_CLOSINGS_APPR_CY1 As Integer = 48  ' AV (2026)
Public Const COL_CLOSINGS_APPR_CY2 As Integer = 49  ' AW (2027)
Public Const COL_CLOSINGS_APPR_CY3 As Integer = 50  ' AX (2028)
Public Const COL_CLOSINGS_APPR_CY4 As Integer = 51  ' AY (2029)
Public Const COL_CLOSINGS_APPR_CY5 As Integer = 52  ' AZ (2030)

Public Const COL_CLOSINGS_VAR_CY As Integer = 53    ' BA (2025)
Public Const COL_CLOSINGS_VAR_CY1 As Integer = 54   ' BB (2026)
Public Const COL_CLOSINGS_VAR_CY2 As Integer = 55   ' BC (2027)
Public Const COL_CLOSINGS_VAR_CY3 As Integer = 56   ' BD (2028)
Public Const COL_CLOSINGS_VAR_CY4 As Integer = 57   ' BE (2029)
Public Const COL_CLOSINGS_VAR_CY5 As Integer = 58   ' BF (2030)

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
=======
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
' ADODB CONSTANTS (in case they're missing from your ADODB library)
' ============================================================================
' If you get "Variable not defined" errors for these, uncomment the ones you need
' Public Const adBit As Integer = 128
' Public Const adTinyInt As Integer = 16
' Public Const adUnsignedTinyInt As Integer = 17

' ============================================================================
' SHEET NAMES
' ============================================================================
' NOTE: Update these if your Excel workbook uses different sheet names
Public Const SHEET_DATA As String = "PIF"
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
    colFundingProject = 13  ' M (also used as project_id in database)
    colProjectName = 14     ' N
    colOriginalISD = 15     ' O
    colRevisedISD = 16      ' P
    colLCMIssue = 17        ' Q
    colStatus = 18          ' R
    colCategory = 19        ' S
    colJustification = 20   ' T
    ' Cost columns U-BF (21-58) - see cost column constants below
    colMovingISDYear = 39   ' AM
    colPriorYearSpend = 40  ' AN
End Enum

' ============================================================================
' COST COLUMN MAPPINGS (EXCEL LAYOUT)
' ============================================================================
' Target Scenario columns
Public Const COL_TARGET_REQ_CY As Integer = 21      ' U (2025)
Public Const COL_TARGET_REQ_CY1 As Integer = 22     ' V (2026)
Public Const COL_TARGET_REQ_CY2 As Integer = 23     ' W (2027)
Public Const COL_TARGET_REQ_CY3 As Integer = 24     ' X (2028)
Public Const COL_TARGET_REQ_CY4 As Integer = 25     ' Y (2029)
Public Const COL_TARGET_REQ_CY5 As Integer = 26     ' Z (2030)

Public Const COL_TARGET_APPR_CY As Integer = 27     ' AA (2025)
Public Const COL_TARGET_APPR_CY1 As Integer = 28    ' AB (2026)
Public Const COL_TARGET_APPR_CY2 As Integer = 29    ' AC (2027)
Public Const COL_TARGET_APPR_CY3 As Integer = 30    ' AD (2028)
Public Const COL_TARGET_APPR_CY4 As Integer = 31    ' AE (2029)
Public Const COL_TARGET_APPR_CY5 As Integer = 32    ' AF (2030)

Public Const COL_TARGET_VAR_CY As Integer = 33      ' AG (2025)
Public Const COL_TARGET_VAR_CY1 As Integer = 34     ' AH (2026)
Public Const COL_TARGET_VAR_CY2 As Integer = 35     ' AI (2027)
Public Const COL_TARGET_VAR_CY3 As Integer = 36     ' AJ (2028)
Public Const COL_TARGET_VAR_CY4 As Integer = 37     ' AK (2029)
Public Const COL_TARGET_VAR_CY5 As Integer = 38     ' AL (2030)

' Closings Scenario columns
Public Const COL_CLOSINGS_REQ_CY As Integer = 41    ' AO (2025)
Public Const COL_CLOSINGS_REQ_CY1 As Integer = 42   ' AP (2026)
Public Const COL_CLOSINGS_REQ_CY2 As Integer = 43   ' AQ (2027)
Public Const COL_CLOSINGS_REQ_CY3 As Integer = 44   ' AR (2028)
Public Const COL_CLOSINGS_REQ_CY4 As Integer = 45   ' AS (2029)
Public Const COL_CLOSINGS_REQ_CY5 As Integer = 46   ' AT (2030)

Public Const COL_CLOSINGS_APPR_CY As Integer = 47   ' AU (2025)
Public Const COL_CLOSINGS_APPR_CY1 As Integer = 48  ' AV (2026)
Public Const COL_CLOSINGS_APPR_CY2 As Integer = 49  ' AW (2027)
Public Const COL_CLOSINGS_APPR_CY3 As Integer = 50  ' AX (2028)
Public Const COL_CLOSINGS_APPR_CY4 As Integer = 51  ' AY (2029)
Public Const COL_CLOSINGS_APPR_CY5 As Integer = 52  ' AZ (2030)

Public Const COL_CLOSINGS_VAR_CY As Integer = 53    ' BA (2025)
Public Const COL_CLOSINGS_VAR_CY1 As Integer = 54   ' BB (2026)
Public Const COL_CLOSINGS_VAR_CY2 As Integer = 55   ' BC (2027)
Public Const COL_CLOSINGS_VAR_CY3 As Integer = 56   ' BD (2028)
Public Const COL_CLOSINGS_VAR_CY4 As Integer = 57   ' BE (2029)
Public Const COL_CLOSINGS_VAR_CY5 As Integer = 58   ' BF (2030)

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
>>>>>>> 254c5f99beb56416a0a77e8c6100e047b5b2e663
