' ==========================================================================
' Module      : mod_510k_Processor
' Author      : [Original Author - Unknown]
' Date        : [Original Date - Unknown]
' Maintainer  : Cline (AI Assistant)
' Version     : See mod_Config.VERSION_INFO
' ==========================================================================
' Description : This module serves as the central orchestrator for the FDA
'               510(k) lead scoring process. It coordinates the workflow,
'               including data refresh, parameter loading (weights, keywords,
'               cache), row-by-row scoring calculation, data write-back,
'               formatting application, and conditional data archiving.
'               It relies heavily on helper modules for specific tasks,
'               promoting modularity and maintainability.
'
' Key Function: ProcessMonthly510k() - The main public subroutine that
'               initiates and manages the entire processing pipeline.
'               Typically called from Workbook_Open or a UI button.
'
' Dependencies: - mod_Logger: For logging events to the RunLog sheet.
'               - mod_DebugTraceHelpers: For detailed debug tracing.
'               - mod_Config: For global constants (sheet names, version).
'               - mod_Schema: For data structure definitions and column mapping.
'               - mod_DataIO: For data input/output operations (PQ refresh, table I/O).
'               - mod_Weights: For loading and accessing scoring weights/keywords.
'               - mod_Cache: For managing the company recap cache.
'               - mod_Score: For calculating the 510(k) score for each record.
'               - mod_Format: For applying formatting to the output data.
'               - mod_Archive: For archiving processed data to monthly sheets.
'               - mod_Utils: For miscellaneous utility functions (getting sheets, UI handling).
'
' Assumptions : - Specific Excel Tables exist on the "Weights" sheet.
'               - A "CompanyCache" sheet exists for caching.
'               - A "CurrentMonthData" sheet exists for Power Query output.
'               - Required helper modules are present in the VBA project.
'
' Revision History:
' --------------------------------------------------------------------------
' Date        Author          Description
' ----------- --------------- ----------------------------------------------
' 2025-04-30  Cline (AI)      - Added detailed module header comment block.
' 2025-04-30  Cline (AI)      - Corrected undefined variable 'lgCRITICAL' to 'lgERROR'
'                             in ProcessErrorHandler (using LogLevel enum).
' 2025-04-30  Cline (AI)      - Corrected undefined variable 'lvlFATAL' to 'lvlERROR'
'                             in ProcessErrorHandler (using eTraceLvl enum).
' 2025-04-30  Cline (AI)      - Removed cleanup lines in CleanExit for module-level
'                               variables now managed in other modules (e.g.,
'                               dictACWeights, dictCache, keyword lists).
' 2025-04-30  Cline (AI)      - Corrected undefined variable 'VERSION_INFO' by adding
'                               module qualifier 'mod_Config.' in initial logging calls.
' 2025-04-30  Cline (AI)      - Added FlushLogBuf call in ProcessErrorHandler to ensure
'                               log buffer is written even when errors occur.
' 2025-04-30  Cline (AI)      - Added robust connection refresh fallback if table missing.
'                             - Tightened error handler to capture Err details immediately.
' 2025-04-30  Cline (AI)      - Moved errNum/errDesc declaration to top of sub for scope.
' 2025-04-30  Cline (AI)      - Added detailed Debug.Print statements around table checks
'                               and fallback logic for enhanced tracing.
' 2025-04-30  Cline (AI)      - Removed all temporary Debug.Print statements.
' [Previous dates/authors/changes unknown]
' ==========================================================================
'--- Code for Module: mod_510k_Processor ---
Option Explicit
Attribute VB_Name = "mod_510k_Processor"

#Const REFACTOR_MODE = 0 '<<< Refactoring complete for this module's orchestration

' ==========================================================================
' ==========================================================================
' ===                                                                    ===
' ===          FDA 510(k) LEAD SCORING PROCESSING MODULE                 ===
' ===                 (mod_510k_Processor)                               ===
' ===                                                                    ===
' ==========================================================================
' ==========================================================================
' --- Requires companion module: mod_Logger                              ---
' --- Requires companion module: mod_DebugTraceHelpers                   ---
' --- Requires companion module: mod_Config                              ---
' --- Requires companion module: mod_Schema                              ---
' --- Requires companion module: mod_DataIO                              ---
' --- Requires companion module: mod_Weights                             ---
' --- Requires companion module: mod_Cache                               ---
' --- Requires companion module: mod_Score                               ---
' --- Requires companion module: mod_Format                              ---
' --- Requires companion module: mod_Archive                             ---
' --- Requires companion module: mod_Utils                               ---
' --- Assumes Excel Tables named: tblACWeights, tblSTWeights,          ---
' --- tblPCWeights, tblKeywords on sheet named "Weights"               ---
' --- Assumes Cache sheet named "CompanyCache" with headers            ---
' --- Assumes Data sheet named "CurrentMonthData" with PQ output table ---

' ==========================================================================
' ===               MODULE-LEVEL VARIABLES / OBJECTS                   ===
' ==========================================================================
' --- Module-level variables removed as logic moved to dedicated modules ---

' ==========================================================================
' ===                   MAIN ORCHESTRATION SUB                         ===
' ==========================================================================
Public Sub ProcessMonthly510k()
    ' Purpose: Main control routine orchestrating the entire 510(k) processing workflow.
    '          Called by Workbook_Open or a button. Handles setup, data refresh,
    '          parameter loading, scoring, formatting, caching, and archiving.

    ' --- Variable Declarations ---
    Dim wsData As Worksheet, wsWeights As Worksheet, wsCache As Worksheet, wsLog As Worksheet
    Dim startMonth As Date, targetMonthName As String, archiveSheetName As String
    Dim startTime As Double: startTime = Timer ' Start timing the process
    Dim tblData As ListObject ' Represents the main data table
    Dim recordCount As Long   ' Number of records to process
    Dim dataArr As Variant    ' Array to hold data for fast processing
    Dim i As Long             ' Loop counter
    Dim scoreResult As Variant ' Array holding results from Calculate510kScore
    Dim currentRecap As String ' Holds the company recap text
    Dim useOpenAI As Boolean   ' Flag indicating if OpenAI should be attempted
    Dim colIndices As Object   ' Dictionary mapping column names to indices
    Dim proceed As Boolean     ' Flag indicating if processing should run
    Dim mustArchive As Boolean ' Flag indicating if archiving is needed
    Dim originalCalcBeforeSave As Boolean ' For performance tweak
    Dim originalCalcState As XlCalculation ' For performance tweak
    ' --- Error variables declared at procedure level ---
    Dim errNum  As Long
    Dim errDesc As String
    ' -------------------------------------------------

    ' --- Error Handling Setup ---
    On Error GoTo ProcessErrorHandler

    ' --- Initial Setup & Screen Handling ---
    Application.ScreenUpdating = False
    If ActiveWindow.FreezePanes Then ActiveWindow.FreezePanes = False ' Ensure off at start
    originalCalcState = Application.Calculation ' Store initial calc state
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.Cursor = xlWait
    Application.StatusBar = "Initializing 510(k) processing..."

    ' --- Initialize Logging ---
    ' Uses lg... constants from mod_Logger's Public Enum LogLevel
    ' Uses lvl... constants from mod_DebugTraceHelpers' Public Enum eTraceLvl
    LogEvt "ProcessStart", lgINFO, "ProcessMonthly510k Started", "Version=" & mod_Config.VERSION_INFO ' Qualified constant
    TraceEvt lvlINFO, "ProcessMonthly510k", "Process Start", "Version=" & mod_Config.VERSION_INFO ' Qualified constant; Use enum

    ' --- Get Worksheet Objects Safely (Moved to mod_Utils) ---
    If Not mod_Utils.GetWorksheets(wsData, wsWeights, wsCache) Then
        GoTo CleanExit ' Exit handled by EnsureUIOn
    End If

    ' --- Ensure Data Table Exists (Guard Rail) ---
    On Error Resume Next ' Check for table existence
    Set tblData = wsData.ListObjects(1) ' Try to get the first table
    On Error GoTo ProcessErrorHandler ' Restore error handler

    If tblData Is Nothing Then
        Dim rng As Range
        On Error Resume Next ' Handle errors during table creation
        Set rng = wsData.Range("A1").CurrentRegion ' Assume data starts at A1
        If Not rng Is Nothing And rng.Cells.Count > 1 Then ' Check if there's data
            Set tblData = wsData.ListObjects.Add(SourceType:=xlSrcRange, _
                                                 Source:=rng, _
                                                 XlListObjectHasHeaders:=xlYes)
            If Not tblData Is Nothing Then
                tblData.Name = "pgGet510kData_" & Format(Now, "yyyymmddhhmmss") ' Give unique name
                LogEvt "DataTable", lgWARN, "Table was missing – recreated from current region as '" & tblData.Name & "'."
                TraceEvt lvlWARN, "ProcessMonthly510k", "Data table missing, recreated as '" & tblData.Name & "'"
            Else
                LogEvt "DataTable", lgERROR, "Table was missing and failed to recreate from current region."
                TraceEvt lvlERROR, "ProcessMonthly510k", "Data table missing, failed to recreate."
                GoTo ProcessErrorHandler ' Cannot proceed without a table
            End If
        Else
             LogEvt "DataTable", lgERROR, "Table was missing and no data found in CurrentRegion of A1 to recreate it."
             TraceEvt lvlERROR, "ProcessMonthly510k", "Data table missing, no data in A1 CurrentRegion."
             GoTo ProcessErrorHandler ' Cannot proceed without a table
        End If
        On Error GoTo ProcessErrorHandler ' Restore error handler
    End If
    ' --- End Table Guard Rail ---

    '--- FINAL FALLBACK: build the table from any query that targets CurrentMonthData ---
    If tblData Is Nothing Then
        Dim pq As WorkbookConnection, sql As String ' Changed conn to pq for clarity
        Dim foundConnection As Boolean: foundConnection = False
        For Each pq In ThisWorkbook.Connections
            ' Check if it's an OLEDB connection (typical for Power Query)
            If TypeName(pq.OLEDBConnection) = "OLEDBConnection" Then
                On Error Resume Next ' Handle connections without CommandText
                sql = pq.OLEDBConnection.CommandText
                Dim cmdTextErr As Long: cmdTextErr = Err.Number ' Capture error
                On Error GoTo ProcessErrorHandler ' Restore handler

                ' Check if the command text (often contains sheet name) targets our sheet
                If cmdTextErr = 0 And InStr(1, sql, DATA_SHEET_NAME, vbTextCompare) > 0 Then
                    LogEvt "DataTable", lgWARN, _
                           "Table not found; attempting to recreate it by refreshing connection: " & pq.Name
                    TraceEvt lvlWARN, "ProcessMonthly510k", _
                             "Table missing – refreshing connection", "Conn=" & pq.Name

                    pq.Refresh                         ' loads data to sheet, should recreate ListObject
                    foundConnection = True
                    Exit For ' Found and refreshed the relevant connection
                End If
            End If
        Next pq

        ' Try setting the table object again after potential refresh
        If foundConnection Then
            On Error Resume Next
            Set tblData = wsData.ListObjects(1)           ' try again
            On Error GoTo ProcessErrorHandler
        End If
    End If
    ' --- End Robust Fallback ---

    ' --- Final check if table exists after all attempts (with improved message) ---
    If tblData Is Nothing Then
        Dim msg As String: msg = "No list-object found on '" & DATA_SHEET_NAME & _
                         "' and none could be recreated automatically via connection refresh."
        LogEvt "DataTable", lgERROR, msg
        TraceEvt lvlERROR, "ProcessMonthly510k", msg
        errDesc = msg        ' push a useful message into the handler (errDesc declared at top)
        GoTo ProcessErrorHandler ' Critical failure if table still missing
    End If
    ' --- End Final Check ---


    ' --- Determine Target Month & Check Guard Conditions ---
    startMonth = DateSerial(Year(DateAdd("m", -1, Date)), Month(DateAdd("m", -1, Date)), 1)
    targetMonthName = Format$(startMonth, "MMM-yyyy")
    archiveSheetName = targetMonthName
    mustArchive = Not mod_DataIO.SheetExists(archiveSheetName) ' Use mod_DataIO
    proceed = mustArchive Or Day(Date) <= 5 Or mod_Utils.IsMaintainerUser() ' Use mod_Utils

    LogEvt "ArchiveCheck", IIf(proceed, lgINFO, lgWARN), _
           "Guard conditions: Archive needed=" & mustArchive & _
           ", Day of month=" & Day(Date) & ", Is maintainer=" & mod_Utils.IsMaintainerUser() & _
           ", Will proceed=" & proceed
    TraceEvt IIf(proceed, lvlINFO, lvlWARN), "ProcessMonthly510k", "Guard Check", "Proceed=" & proceed & ", ArchiveNeeded=" & mustArchive & ", Day=" & Day(Date) & ", Maintainer=" & mod_Utils.IsMaintainerUser()

    If Not proceed Then
        LogEvt "ProcessSkip", lgINFO, "Processing skipped: Archive exists, not day 1-5, not maintainer."
        TraceEvt lvlINFO, "ProcessMonthly510k", "Processing Skipped (Guard Conditions Met)"
        Application.StatusBar = "Month " & targetMonthName & " already archived. Refreshing current view only."
        ' Attempt refresh even if skipping full process
        On Error Resume Next: Set tblData = wsData.ListObjects(1): On Error GoTo ProcessErrorHandler
        If tblData Is Nothing Then
            LogEvt "Refresh", lgERROR, "Data table not found on " & DATA_SHEET_NAME & " during skipped run check."
            TraceEvt lvlERROR, "ProcessMonthly510k", "Data table not found during skipped run refresh check"
        Else
            ' Use mod_DataIO for refresh during skipped run check
            If Not mod_DataIO.RefreshPowerQuery(tblData) Then
                LogEvt "Refresh", lgERROR, "PQ Refresh failed during skipped run check (via mod_DataIO)."
                TraceEvt lvlERROR, "ProcessMonthly510k", "PQ Refresh failed during skipped run check (mod_DataIO)"
                 ' GoTo ProcessErrorHandler ' Option to make it critical
            End If
        End If
        Set tblData = Nothing
        GoTo CleanExit ' Exit handled by EnsureUIOn
    End If
    Application.StatusBar = "Processing for month: " & targetMonthName

    ' --- Get Data Table & Check for Data (Redundant check, but safe) ---
    If tblData Is Nothing Then
        On Error Resume Next: Set tblData = wsData.ListObjects(1): On Error GoTo ProcessErrorHandler
    End If
    If tblData Is Nothing Then
        LogEvt "DataTable", lgERROR, "Data table not found on " & DATA_SHEET_NAME
        TraceEvt lvlERROR, "ProcessMonthly510k", "Data table object lost or not found before refresh"
        GoTo ProcessErrorHandler
    End If

    ' --- Refresh Power Query Data (using the table object) ---
    Application.StatusBar = "Refreshing FDA data from Power Query..."
    LogEvt "Refresh", lgINFO, "Attempting PQ refresh for table: " & tblData.Name
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Refresh Power Query Start", "Table=" & tblData.Name
    ' Use mod_DataIO for main refresh
    If Not mod_DataIO.RefreshPowerQuery(tblData) Then ' Includes post-refresh lock
         LogEvt "Refresh", lgERROR, "PQ Refresh failed via mod_DataIO. Processing stopped."
         TraceEvt lvlERROR, "ProcessMonthly510k", "PQ Refresh Failed (mod_DataIO) - Halting Process"
         GoTo ProcessErrorHandler ' Stop on critical PQ error
    End If
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Refresh Power Query End"

    ' --- Re-check table and data after refresh (PQ might empty it) ---
    If tblData Is Nothing Then ' Should not happen if RefreshPowerQuery succeeded, but defensive check
         LogEvt "DataTable", lgERROR, "Data table object became Nothing after refresh."
         TraceEvt lvlERROR, "ProcessMonthly510k", "Data table object lost after refresh"
         GoTo ProcessErrorHandler
    End If
    If tblData.ListRows.Count = 0 Then
        LogEvt "DataTable", lgWARN, "No data returned by Power Query for " & targetMonthName & "."
        TraceEvt lvlWARN, "ProcessMonthly510k", "No data after PQ refresh", "Month=" & targetMonthName
        MsgBox "No data returned by Power Query for " & targetMonthName & ". Nothing to process.", vbInformation, "No Data"
        GoTo CleanExit ' Exit handled by EnsureUIOn
    End If
    recordCount = tblData.ListRows.Count
    LogEvt "DataTable", lgINFO, "Table contains " & recordCount & " rows post-refresh."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Data Rows Post-Refresh", "Count=" & recordCount

    ' --- Add/Verify Output Columns ---
    LogEvt "Columns", lgINFO, "Checking/Adding scoring output columns..."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Add/Verify Columns Start"
    ' Use mod_Format
    If Not mod_Format.AddScoreColumnsIfNeeded(tblData) Then GoTo ProcessErrorHandler
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Add/Verify Columns End"

    ' --- Map Column Headers to Indices ---
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Map Columns Start"
    ' Use mod_Schema
    Set colIndices = mod_Schema.GetColumnIndices(tblData.HeaderRowRange) ' Now handles duplicates
    If colIndices Is Nothing Then GoTo ProcessErrorHandler
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Map Columns End", "MappedKeys=" & colIndices.Count

    ' --- Load Weights, Keywords, and Cache ---
    Application.StatusBar = "Loading scoring parameters and cache..."
    LogEvt "LoadParams", lgINFO, "Loading weights, keywords, and cache..."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Load Parameters Start"
    ' Use mod_Weights and mod_Cache
    If Not mod_Weights.LoadAll(wsWeights) Then GoTo ProcessErrorHandler ' Renamed per roadmap example
    Call mod_Cache.LoadCompanyCache(wsCache)
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Load Parameters End"

    ' --- Read Data into Array for Fast Processing ---
    Application.StatusBar = "Reading data into memory (" & recordCount & " rows)..."
    LogEvt "ReadData", lgINFO, "Reading data into array..."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Read Data to Array Start", "Rows=" & recordCount
    ' Ensure dataArr is correctly dimensioned, especially for single row case
    If recordCount = 1 Then
        ' Handle single row specifically to ensure 2D array
        Dim singleRowData As Variant
        singleRowData = tblData.DataBodyRange.Value ' Read value first
        ' Check if it's already a 2D array (1 row, N columns)
        If Not IsArray(singleRowData) Then
            ' If not an array, it's a single value; create a 1x1 2D array
            ReDim dataArr(1 To 1, 1 To 1)
            dataArr(1, 1) = singleRowData
        ElseIf UBound(singleRowData, 1) = 1 And UBound(singleRowData, 2) >= 1 Then
             ' It's already a 1xN 2D array
             dataArr = singleRowData
        ElseIf UBound(singleRowData) >= 1 And UBound(singleRowData, 1) > 1 Then
             ' It's a Nx1 2D array (less likely for table read, but handle)
             dataArr = singleRowData ' Keep as is
        Else
             ' It's a 1D array (e.g., 1 row, 1 column read as 1D) - Convert to 1xN 2D
             Dim numCols As Long: numCols = tblData.ListColumns.Count
             ReDim dataArr(1 To 1, 1 To numCols)
             For i = 1 To numCols
                 dataArr(1, i) = singleRowData(i)
             Next i
        End If
    Else
        ' Multiple rows, read normally
        dataArr = tblData.DataBodyRange.Value2 ' Use Value2 for performance
    End If
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Read Data to Array End"

    ' --- Determine if OpenAI should be used ---
    useOpenAI = mod_Utils.IsMaintainerUser() ' Use mod_Utils
    LogEvt "OpenAICheck", lgINFO, "OpenAI usage check", "IsMaintainer=" & useOpenAI

    ' --- Process Each Row ---
    Application.StatusBar = "Calculating scores and fetching recaps (0% complete)..."
    LogEvt "ProcessRows", lgINFO, "Starting row processing loop..."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Process Rows Start", "TotalRows=" & recordCount

    Dim companyName As String
    Dim scoreCol As Long, categoryCol As Long, recapCol As Long ' Indices for writing back
    scoreCol = mod_Schema.SafeGetColIndex(colIndices, "Final_Score")
    categoryCol = mod_Schema.SafeGetColIndex(colIndices, "Category")
    recapCol = mod_Schema.SafeGetColIndex(colIndices, "CompanyRecap")
    Dim acWtCol As Long: acWtCol = mod_Schema.SafeGetColIndex(colIndices, "AC_Wt")
    Dim pcWtCol As Long: pcWtCol = mod_Schema.SafeGetColIndex(colIndices, "PC_Wt")
    Dim kwWtCol As Long: kwWtCol = mod_Schema.SafeGetColIndex(colIndices, "KW_Wt")
    Dim stWtCol As Long: stWtCol = mod_Schema.SafeGetColIndex(colIndices, "ST_Wt")
    Dim ptWtCol As Long: ptWtCol = mod_Schema.SafeGetColIndex(colIndices, "PT_Wt")
    Dim glWtCol As Long: glWtCol = mod_Schema.SafeGetColIndex(colIndices, "GL_Wt")
    Dim nfCalcCol As Long: nfCalcCol = mod_Schema.SafeGetColIndex(colIndices, "NF_Calc")
    Dim synCalcCol As Long: synCalcCol = mod_Schema.SafeGetColIndex(colIndices, "Synergy_Calc")
    Dim scorePctCol As Long: scorePctCol = mod_Schema.SafeGetColIndex(colIndices, "Score_Percent")

    ' Check if essential output columns were found
    If scoreCol = 0 Or categoryCol = 0 Or recapCol = 0 Then
        LogEvt "ProcessRows", lgERROR, "Could not find essential output column indices (Final_Score, Category, CompanyRecap). Halting processing."
        TraceEvt lvlERROR, "ProcessMonthly510k", "Missing essential output column indices", "Score=" & scoreCol & ", Cat=" & categoryCol & ", Recap=" & recapCol
        GoTo ProcessErrorHandler
    End If

    For i = 1 To recordCount
        ' Calculate Score (Use mod_Score)
        scoreResult = mod_Score.Calculate510kScore(dataArr, i, colIndices)

        ' Get Company Recap (Use mod_Cache)
        companyName = mod_Schema.SafeGetString(dataArr, i, colIndices, "Applicant")
        currentRecap = mod_Cache.GetCompanyRecap(companyName, useOpenAI)

        ' Write results back to the array
        dataArr(i, scoreCol) = scoreResult(0) ' Final_Score_Raw
        dataArr(i, categoryCol) = scoreResult(1) ' Category
        dataArr(i, acWtCol) = scoreResult(2) ' AC_Wt
        dataArr(i, pcWtCol) = scoreResult(3) ' PC_Wt
        dataArr(i, kwWtCol) = scoreResult(4) ' KW_Wt
        dataArr(i, stWtCol) = scoreResult(5) ' ST_Wt
        dataArr(i, ptWtCol) = scoreResult(6) ' PT_Wt
        dataArr(i, glWtCol) = scoreResult(7) ' GL_Wt
        dataArr(i, nfCalcCol) = scoreResult(8) ' NF_Calc
        dataArr(i, synCalcCol) = scoreResult(9) ' Synergy_Calc
        dataArr(i, scorePctCol) = scoreResult(0) ' Score_Percent (same as raw score before formatting)
        dataArr(i, recapCol) = currentRecap ' CompanyRecap

        ' Update Status Bar periodically
        If i Mod 50 = 0 Then ' Update every 50 rows
            Application.StatusBar = "Calculating scores and fetching recaps (" & Format(i / recordCount, "0%") & " complete)..."
            DoEvents ' Allow UI to update and prevent freezing
        End If
    Next i
    LogEvt "ProcessRows", lgINFO, "Finished row processing loop."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Process Rows End"

    ' --- Write Processed Data Back to Table ---
    Application.StatusBar = "Writing processed data back to table..."
    LogEvt "WriteData", lgINFO, "Writing processed data back to table..."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Write Data Back Start"
    ' Use mod_DataIO
    If Not mod_DataIO.ArrayToTable(dataArr, tblData) Then GoTo ProcessErrorHandler
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Write Data Back End"

    ' --- Save Company Cache (Use mod_Cache) ---
    Application.StatusBar = "Saving company cache..."
    LogEvt "SaveCache", lgINFO, "Saving company cache..."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Save Cache Start"
    Call mod_Cache.SaveCompanyCache(wsCache)
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Save Cache End"

    ' --- Apply Formatting ---
    Application.StatusBar = "Applying formatting..."
    LogEvt "Formatting", lgINFO, "Applying formatting..."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Formatting Start"
    ' Use mod_Format
    If Not mod_Format.ApplyAll(tblData, wsData) Then GoTo ProcessErrorHandler ' Passing wsData for Freeze
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Formatting End"

    ' --- Archive Month if Necessary ---
    If mustArchive Then
        Application.StatusBar = "Archiving previous month's data..."
        LogEvt "Archiving", lgINFO, "Archiving month: " & archiveSheetName
        TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Archiving Start", "Sheet=" & archiveSheetName
        ' Use mod_Archive
        If Not mod_Archive.ArchiveIfNeeded(tblData, archiveSheetName) Then GoTo ProcessErrorHandler
        TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Archiving End"
    Else
        LogEvt "Archiving", lgINFO, "Archiving not required for this run."
        TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Archiving Skipped"
    End If

    ' --- Finalization ---
    Application.StatusBar = "Processing complete."
    LogEvt "ProcessComplete", lgINFO, "510(k) processing completed successfully."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Process Complete"
    MsgBox "510(k) processing complete for " & targetMonthName & "." & vbCrLf & _
           "Duration: " & Format(Timer - startTime, "0.0") & " seconds.", vbInformation, "Processing Complete"

    ' --- Error Handling ---
    GoTo CleanExit ' Skip error handler if successful

ProcessErrorHandler:
    ' --- Capture error details IMMEDIATELY ---
    ' Dim errNum  As Long:  errNum  = Err.Number ' Moved declaration to top
    ' Dim errDesc As String: errDesc = Err.Description ' Moved declaration to top
    errNum = Err.Number ' Capture current error number
    If errDesc = "" Then errDesc = Err.Description ' Capture description if not already set by fallback
    ' --- End Capture ---
    LogEvt "ProcessError", lgERROR, "Unhandled Error #" & errNum & " in ProcessMonthly510k: " & errDesc ' Use lgERROR instead of lgCRITICAL
    TraceEvt lvlERROR, "ProcessMonthly510k", "FATAL ERROR", "Err=" & errNum & " - " & errDesc ' Use lvlERROR instead of lvlFATAL
    ' --- Explicitly flush log buffer on error ---
    On Error Resume Next ' Prevent error in Flush from masking original error
    FlushLogBuf
    On Error GoTo 0 ' Restore default error handling (though we are exiting)
    ' --- End Flush ---
    MsgBox "An unexpected error occurred: " & vbCrLf & errDesc & vbCrLf & "Please check the RunLog sheet for details.", vbCritical, "Processing Error"
    ' Fall through to CleanExit

CleanExit:
    LogEvt "ProcessEnd", lgINFO, "ProcessMonthly510k Ended", "Duration=" & Format(Timer - startTime, "0.00") & "s"
    TraceEvt lvlINFO, "ProcessMonthly510k", "Process End", "Duration=" & Format(Timer - startTime, "0.00") & "s"
    ' --- Explicitly flush log buffer on successful completion ---
    On Error Resume Next ' Prevent error in Flush from causing issues here
    FlushLogBuf
    On Error GoTo 0 ' Restore default error handling
    ' --- End Flush ---
    Call mod_Utils.EnsureUIOn(originalCalcState) ' Use mod_Utils, restore original calc state
    ' --- Clean up local objects ---
    Set wsData = Nothing
    Set wsWeights = Nothing
    Set wsCache = Nothing
    Set wsLog = Nothing
    Set tblData = Nothing
    Set colIndices = Nothing
    ' --- Module-level objects (like dictionaries in mod_Weights/mod_Cache) ---
    ' --- are managed within their respective modules and don't need cleanup here ---
End Sub

' ==========================================================================
' ===                   HELPER FUNCTIONS / SUBS                        ===
' ==========================================================================
' --- All helper functions previously in this module have been moved     ---
' --- to their respective dedicated modules (mod_DataIO, mod_Schema,    ---
' --- mod_Weights, mod_Cache, mod_Score, mod_Format, mod_Archive,       ---
' --- mod_Utils). This module now only contains the main orchestration   ---
' --- subroutine `ProcessMonthly510k`.                                   ---

</final_file_content>

IMPORTANT: For any future changes to this file, use the final_file_content shown above as your reference. This content reflects the current state of the file, including any auto-formatting (e.g., if you used single quotes but the formatter converted them to double quotes). Always base your SEARCH/REPLACE operations on this final version to ensure accuracy.<environment_details>
# VSCode Visible Files
src/vba/mod_510k_Processor.bas

# VSCode Open Tabs
../../Users/Ryan/AppData/Local/Programs/Microsoft VS Code/Untitled-1
src/vba/mod_Logger.bas
src/vba/ThisWorkbook.cls
src/vba/mod_DataIO.bas
src/vba/mod_Format.bas
src/vba/mod_510k_Processor.bas
commit_message.txt

# Current Time
4/30/2025, 6:21:21 PM (America/Phoenix, UTC-7:00)

# Context Window Usage
261,775 / 1,048.576K tokens used (25%)

# Current Mode
ACT MODE
</environment_details>
