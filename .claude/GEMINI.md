# Gemini Code Assistant Context

## Project: PIF Form Enhancement

This project is a VBA application for Microsoft Excel that automates the submission of "Project Impact Form" (PIF) data to a SQL Server database.

### Project Overview

The system is designed to replace a manual, Excel-based PIF process with a more robust, database-driven solution. It uses VBA macros within an Excel workbook to validate and transmit data to a SQL Server backend. The process involves several steps, including data validation, unpivoting cost data, staging, and final commitment to the database.

### Key Technologies

*   **Frontend:** Microsoft Excel (.xlsm)
*   **Backend:** Microsoft SQL Server
*   **Language:** Visual Basic for Applications (VBA)
*   **Database Connectivity:** ActiveX Data Objects (ADODB)

### Architecture

The submission process follows these steps:
1.  **Unpivot Data:** Cost data, which is in a wide format in Excel, is transformed into a long, normalized format.
2.  **Backup:** Backups of the current "inflight" data are created in the database.
3.  **Staging:** The data from Excel is uploaded to staging tables in the SQL Server database.
4.  **Validation:** A series of validation rules are applied, both within Excel and on the SQL Server side via a stored procedure.
5.  **Commit:** If validation passes, the data is moved from the staging tables to the "inflight" tables.
6.  **Archive:** PIFs marked as "Approved" or "Dispositioned" are moved from the "inflight" tables to the permanent "approved" tables.
7.  **Logging:** The submission event is recorded in a log table.

## Key Files

*   `PIF_Database_DDL.sql`: SQL script to create the entire database schema, including tables, views, and stored procedures.
*   `mod_Database.bas`: VBA module responsible for all database connection and communication logic. It uses ADODB to connect to the SQL Server.
*   `mod_Validation.bas`: VBA module that contains all data validation logic. It performs checks within Excel and also executes a validation stored procedure on the SQL server.
*   `mod_Submit.bas`: The main VBA module that orchestrates the entire submission process, calling functions from the other modules in the correct order.
*   `README.md`: Detailed documentation covering setup, usage, troubleshooting, and maintenance.
*   `PIF_Configuration_Document.md`: A template for recording the specific configuration details of a deployment.
*   `PIF_Monthly_Checklist.md`: A user-facing checklist for the month-end submission process.

## Building and Running

This project is not compiled in a traditional sense. It is run from within the Microsoft Excel application.

### Setup

1.  **Database Setup:** Execute the `PIF_Database_DDL.sql` script on a SQL Server instance to create the necessary tables, views, and stored procedures.
2.  **Excel Configuration:**
    *   Open the `.xlsm` workbook.
    *   Import the three `.bas` modules (`mod_Database.bas`, `mod_Validation.bas`, `mod_Submit.bas`).
    *   In the VBA editor, go to `Tools > References` and enable `Microsoft ActiveX Data Objects 6.1 Library`.
    *   In `mod_Database.bas`, update the `SQL_SERVER` and `SQL_DATABASE` constants with the correct connection details.
3.  **Testing:**
    *   Run the `TestConnection` macro in `mod_Database.bas` to verify the database connection.
    *   Use the "Validate Data" button on the Excel sheet to check the data against the validation rules.
    *   Use the "Submit to Database" button to run the full submission process.

## Development Conventions

*   **Modular Code:** The VBA code is separated into modules based on functionality (Database, Validation, Submission).
*   **Constants:** Constants are used for sheet names, column numbers, and other configuration values to improve maintainability.
*   **Transactional Process:** The database submission is designed to be atomic. It occurs within a transaction and includes backup and logging steps for robustness.
*   **Dual Validation:** Validation is performed both on the client-side (Excel/VBA) for immediate feedback and on the server-side (SQL) for data integrity.
*   **Clear Naming:** VBA subroutines and functions, as well as SQL tables and columns, have descriptive names.
*   **Error Handling:** The VBA code includes error handling to catch and report issues during the submission process.
