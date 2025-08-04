' ==========================================================================
' Module      : mod_510k_Processor_Update
' Author      : Cline (AI Assistant)
' Date        : May 7, 2025
' Purpose     : Contains the code changes needed to fix the Power Query
'               refresh issue in mod_510k_Processor by implementing the
'               timer-based refresh solution.
'
' HOW TO USE  : Copy the ProcessMonthly510k subroutine below to replace the
'               existing subroutine in mod_510k_Processor or use the
'               instructions provided in POWER_QUERY_REFRESH_FIX_IMPLEMENTATION.txt.
'
' Changes Made: 
' - The original mod_DataIO.RefreshPowerQuery call is replaced with
'   mod_DataIO_Enhanced.RefreshPowerQuery, which uses a timer-based approach
'   to avoid the Error 1004 issue.
' - Log messages are updated to reflect the enhanced implementation.
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
    mod_DebugTraceHelpers.TraceEvt lvlINFO, "ProcessMonthly510k", "Process Start", "Version=" & mod_Config.VERSION_INFO ' Qualified constant; Use enum

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
                mod_DebugTraceHelpers.TraceEvt lvlWARN, "ProcessMonthly510k", "Data table missing, recreated as '" & tblData.Name & "'"
            Else
                LogEvt "DataTable", lgERROR, "Table was missing and failed to recreate from current region."
                mod_DebugTraceHelpers.TraceEvt lvlERROR, "ProcessMonthly510k", "Data table missing, failed to recreate."
                GoTo ProcessErrorHandler ' Cannot proceed without a table
            End If
        Else
             LogEvt "DataTable", lgERROR, "Table was missing and no data found in CurrentRegion of A1 to recreate it."
             mod_DebugTraceHelpers.TraceEvt lvlERROR, "ProcessMonthly510k", "Data table missing, no data in A1 CurrentRegion."
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
                    mod_DebugTraceHelpers.TraceEvt lvlWARN, "ProcessMonthly510k", _
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
        mod_DebugTraceHelpers.TraceEvt lvlERROR, "ProcessMonthly510k", msg
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
    mod_DebugTraceHelpers.TraceEvt IIf(proceed, lvlINFO, lvlWARN), "ProcessMonthly510k", "Guard Check", "Proceed=" & proceed & ", ArchiveNeeded=" & mustArchive & ", Day=" & Day(Date) & ", Maintainer=" & mod_Utils.IsMaintainerUser()

    If Not proceed Then
        LogEvt "ProcessSkip", lgINFO, "Processing skipped: Archive exists, not day 1-5, not maintainer."
        mod_DebugTraceHelpers.TraceEvt lvlINFO, "ProcessMonthly510k", "Processing Skipped (Guard Conditions Met)"
        Application.StatusBar = "Month " & targetMonthName & " already archived. Refreshing current view only."
        ' Attempt refresh even if skipping full process
        On Error Resume Next: Set tblData = wsData.ListObjects(1): On Error GoTo ProcessErrorHandler
        If tblData Is Nothing Then
            LogEvt "Refresh", lgERROR, "Data table not found on " & DATA_SHEET_NAME & " during skipped run check."
            mod_DebugTraceHelpers.TraceEvt lvlERROR, "ProcessMonthly510k", "Data table not found during skipped run refresh check"
        Else
            ' Use mod_DataIO_Enhanced for refresh during skipped run check
            If Not mod_DataIO_Enhanced.RefreshPowerQuery(tblData) Then
                LogEvt "Refresh", lgERROR, "PQ Refresh failed during skipped run check (via mod_DataIO_Enhanced)."
                mod_DebugTraceHelpers.TraceEvt lvlERROR, "ProcessMonthly510k", "PQ Refresh failed during skipped run check (Enhanced)"
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
        mod_DebugTraceHelpers.TraceEvt lvlERROR, "ProcessMonthly510k", "Data table object lost or not found before refresh"
        GoTo ProcessErrorHandler
    End If

    ' --- Ask User About Refresh ---
    Dim refreshNeeded As Boolean
    refreshNeeded = False ' Default to skip refresh
    If MsgBox("Do you want to attempt refreshing the FDA data?" & vbCrLf & _
              "(Selecting 'No' will use existing data without refreshing)", _
              vbQuestion + vbYesNo, "Refresh Options") = vbYes Then
        refreshNeeded = True
    End If

    ' --- Refresh Power Query Data (Conditional) --- [MODIFIED SECTION]
    If refreshNeeded Then
        Application.StatusBar = "Refreshing FDA data from Power Query..."
        LogEvt "Refresh", lgINFO, "Attempting PQ refresh for table: " & tblData.Name
        mod_DebugTraceHelpers.TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Refresh Power Query Start", "Table=" & tblData.Name
        
        ' Use mod_DataIO_Enhanced for more reliable refresh
        If Not mod_DataIO_Enhanced.RefreshPowerQuery(tblData) Then ' Timer-based refresh with retry
             LogEvt "Refresh", lgERROR, "PQ Refresh failed via mod_DataIO_Enhanced. Processing stopped."
             mod_DebugTraceHelpers.TraceEvt lvlERROR, "ProcessMonthly510k", "PQ Refresh Failed (Enhanced) - Halting Process"
             ' Set an appropriate error message to show to the user
             errDesc = "Power Query refresh failed. Processing cannot continue without fresh data."
             MsgBox errDesc, vbCritical, "Refresh Error - Processing Halted"
             ' Exit process
             GoTo CleanExit ' Changed from ProcessErrorHandler to CleanExit to ensure clean exit
        End If
        
        mod_DebugTraceHelpers.TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Refresh Power Query End"
    Else
        LogEvt "Refresh", lgINFO, "Refresh skipped by user request. Using existing data."
        mod_DebugTraceHelpers.TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Refresh Power Query Skipped (User Request)"
        ' Ensure status bar is cleared or updated if refresh is skipped
        Application.StatusBar = "Using existing data (refresh skipped)..."
    End If
    ' --- End of Modified Section ---

    ' --- Re-check table and data after refresh OR skip ---
    If tblData Is Nothing Then ' Should not happen if RefreshPowerQuery succeeded, but defensive check
         LogEvt "DataTable", lgERROR, "Data table object became Nothing after refresh."
         mod_DebugTraceHelpers.TraceEvt lvlERROR, "ProcessMonthly510k", "Data table object lost after refresh"
         GoTo ProcessErrorHandler
    End If
    If tblData.ListRows.Count = 0 Then
        LogEvt "DataTable", lgWARN, "No data returned by Power Query for " & targetMonthName & "."
        mod_DebugTraceHelpers.TraceEvt lvlWARN, "ProcessMonthly510k", "No data after PQ refresh", "Month=" & targetMonthName
        MsgBox "No data returned by Power Query for " & targetMonthName & ". Nothing to process.", vbInformation, "No Data"
        GoTo CleanExit ' Exit handled by EnsureUIOn
    End If
    recordCount = tblData.ListRows.Count
    LogEvt "DataTable", lgINFO, "Table contains " & recordCount & " rows post-refresh."
    mod_DebugTraceHelpers.TraceEvt lvlINFO, "ProcessMonthly510k", "Data Rows Post-Refresh", "Count=" & recordCount

    ' --- Add/Verify Output Columns ---
    LogEvt "Columns", lgINFO, "Checking/Adding scoring output columns..."
    mod_DebugTraceHelpers.TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Add/Verify Columns Start"
    ' Use mod_Format
    If Not mod_Format.AddScoreColumnsIfNeeded(tblData) Then GoTo ProcessErrorHandler
    mod_DebugTraceHelpers.TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Add/Verify Columns End"

    ' --- Map Column Headers to Indices ---
    mod_DebugTraceHelpers.TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Map Columns Start"
    ' Use mod_Schema
    Set colIndices = mod_Schema.GetColumnIndices(tblData.HeaderRowRange) ' Now handles duplicates
    If colIndices Is Nothing Then GoTo ProcessErrorHandler
    mod_DebugTraceHelpers.TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Map Columns End", "MappedKeys=" & colIndices.Count

    ' --- Load Weights, Keywords, and Cache ---
    Application.StatusBar = "Loading scoring parameters and cache..."
    LogEvt "LoadParams", lgINFO, "Loading weights, keywords, and cache..."
    mod_DebugTraceHelpers.TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Load Parameters Start"
    ' Use mod_Weights and mod_Cache
    If Not mod_Weights.LoadAll(wsWeights) Then GoTo ProcessErrorHandler ' Renamed per roadmap example
    Call mod_Cache.LoadCompanyCache(wsCache)
    mod_DebugTraceHelpers.TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Load Parameters End"

    ' --- Read Data into Array for Fast Processing ---
    Application.StatusBar = "Reading data into memory (" & recordCount & " rows)..."
    LogEvt "ReadData", lgINFO, "Reading data into array..."
    mod_DebugTraceHelpers.TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Read Data to Array Start", "Rows=" & recordCount
    '
