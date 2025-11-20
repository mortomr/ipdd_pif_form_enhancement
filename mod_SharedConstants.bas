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
