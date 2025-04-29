'--- Code for Module: mod_510k_Processor ---
Option Explicit

' ==========================================================================
' ==========================================================================
' ===                                                                    ===
' ===          FDA 510(k) LEAD SCORING PROCESSING MODULE                 ===
' ===                 (mod_510k_Processor)                               ===
' ===                                                                    ===
' ==========================================================================
' ==========================================================================
' --- Requires companion module: mod_Logger (Optional but Recommended) ---
' --- Assumes Excel Tables named: tblACWeights, tblSTWeights,          ---
' --- tblPCWeights, tblKeywords on sheet named "Weights"               ---
' --- Assumes Cache sheet named "CompanyCache" with headers            ---
' --- Assumes Data sheet named "CurrentMonthData" with PQ output table ---

' ==========================================================================
' ===                     CONFIGURATION CONSTANTS                        ===
' ==========================================================================
' --- Essential Configuration ---
' *** IMPORTANT: SET YOUR WINDOWS USERNAME FOR MAINTAINER FEATURES (e.g., OpenAI, DebugMode) ***
Private Const MAINTAINER_USERNAME As String = "YourWindowsUsername" ' <<< UPDATE THIS

' *** Double-check these names match your Excel objects ***
Public Const PQ_CONNECTION_NAME As String = "Query - pqGet510kData" ' Name of the Power Query connection/query (Public for ThisWorkbook)
Private Const DATA_SHEET_NAME As String = "CurrentMonthData"  ' Sheet where Power Query loads data
Private Const WEIGHTS_SHEET_NAME As String = "Weights"        ' Sheet containing weight/keyword tables
Private Const CACHE_SHEET_NAME As String = "CompanyCache"      ' Sheet for persistent company recap cache
Private Const LOG_SHEET_NAME As String = "LogSheet"             ' Optional: Name for the log sheet

' *** Path to file containing ONLY your OpenAI API Key ***
' *** Uses %APPDATA% environment variable for user-specific location ***
Private Const API_KEY_FILE_PATH As String = "%APPDATA%\510k_Tool\openai_key.txt" ' <<< ENSURE THIS PATH IS CORRECT & FILE EXISTS

' --- Scoring Defaults & Parameters (Used if lookup fails or as base values) ---
' *** REVIEW AND CONFIRM THESE VALUES BASED ON YOUR SCORING MODEL ***
Private Const DEFAULT_AC_WEIGHT As Double = 0.2
Private Const DEFAULT_PC_WEIGHT As Double = 0.2
Private Const DEFAULT_ST_WEIGHT As Double = 0.6 ' Default to Traditional if SubmType not found
Private Const DEFAULT_PT_WEIGHT As Double = 0.5 ' Default if ProcTimeDays is invalid or <162
Private Const HIGH_KW_WEIGHT As Double = 0.85
Private Const LOW_KW_WEIGHT As Double = 0.2
Private Const US_GL_WEIGHT As Double = 0.6
Private Const OTHER_GL_WEIGHT As Double = 0.5
Private Const NF_COSMETIC As Double = -2#  ' Negative Factor for purely cosmetic devices (CONFIRM VALUE)
Private Const NF_DIAGNOSTIC As Double = -0.2 ' Negative Factor for purely diagnostic software (CONFIRM VALUE)
Private Const SYNERGY_BONUS As Double = 0.15 ' Bonus for specific AC + High KW match (CONFIRM VALUE/LOGIC)

' --- OpenAI Configuration (Optional) ---
Private Const OPENAI_API_URL As String = "https://api.openai.com/v1/chat/completions"
Private Const OPENAI_MODEL As String = "gpt-3.5-turbo" ' Or "gpt-4o-mini" etc. - check pricing/availability
Private Const OPENAI_MAX_TOKENS As Long = 100 ' Limit response length
Private Const OPENAI_TIMEOUT_MS As Long = 60000 ' 60 seconds timeout for API call

' --- UI & Formatting ---
Public Const VERSION_INFO As String = "v1.0 - Initial Build" ' Simple version tracking (Public for Logger)
Private Const RECAP_MAX_LEN As Long = 32760 ' Max characters for cell / recap text to avoid overflow
Private Const DEFAULT_RECAP_TEXT = "Needs Research" ' Default text when recap is missing

' ==========================================================================
' ===               MODULE-LEVEL VARIABLES / OBJECTS                   ===
' ==========================================================================
' Dictionaries and Collections for efficient lookups during processing
Private dictACWeights As Object       ' Scripting.Dictionary: Key=AC Code, Value=Weight
Private dictSTWeights As Object       ' Scripting.Dictionary: Key=SubmType Code, Value=Weight
Private dictPCWeights As Object       ' Scripting.Dictionary: Key=PC Code, Value=Weight
Private highValKeywordsList As Collection ' Collection of high-value keywords (Strings)
Private nfCosmeticKeywordsList As Collection
Private nfDiagnosticKeywordsList As Collection
Private therapeuticKeywordsList As Collection
Private dictCache As Object           ' Scripting.Dictionary: Key=CompanyName, Value=RecapText (In-memory for current run)


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

    ' --- Error Handling Setup ---
    On Error GoTo ProcessErrorHandler

    ' --- Initial Setup & Screen Handling ---
    ' Turning these off significantly speeds up processing
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual ' Prevent formulas recalculating mid-process
    Application.EnableEvents = False ' Prevent other event code interfering
    Application.Cursor = xlWait     ' Indicate busy status
    Application.StatusBar = "Initializing 510(k) processing..."

    ' --- Initialize Logging (Optional but recommended) ---
    ' Assumes a logging module 'mod_Logger' exists with InitializeLog, LogEvt, FlushLogBuf
    ' On Error Resume Next ' In case logger module missing
    ' InitializeLog LOG_SHEET_NAME, lvlINFO ' Set default logging level (e.g., INFO, DEBUG, WARN, ERROR)
    ' If Err.Number <> 0 Then Debug.Print Time & " - WARNING: Logger module/sheet not found. Logging disabled.": Err.Clear
    ' On Error GoTo ProcessErrorHandler ' Restore error handling
    ' LogEvt "ProcessStart", lvlINFO, "ProcessMonthly510k Started", "Version=" & VERSION_INFO

    ' --- Get Worksheet Objects Safely ---
    If Not GetWorksheets(wsData, wsWeights, wsCache) Then GoTo CleanExit ' Critical failure if sheets missing

    ' --- Determine Target Month & Check if Already Archived ---
    startMonth = DateSerial(Year(DateAdd("m", -1, Date)), Month(DateAdd("m", -1, Date)), 1)
    targetMonthName = Format$(startMonth, "MMM-yyyy") ' e.g., "Apr-2025" - Avoid using only month name if multi-year
    archiveSheetName = targetMonthName

    proceed = Not SheetExists(archiveSheetName)
    ' LogEvt "ArchiveCheck", IIf(proceed, lvlINFO, lvlWARN), "Archive check for " & archiveSheetName & ". Exists=" & Not proceed

    If Not proceed Then
        ' LogEvt "ProcessSkip", lvlINFO, "Processing skipped as archive sheet '" & archiveSheetName & "' already exists."
        Application.StatusBar = "Month " & targetMonthName & " already archived. Refreshing current view only."
        ' Even if skipping main process, refresh current data for user
        If Not RefreshPowerQuery(PQ_CONNECTION_NAME) Then
            ' Log error, but allow cleanup
             ' LogEvt "Refresh", lvlERROR, "PQ Refresh failed during skipped run check."
        End If
        GoTo CleanExit ' Skip main processing
    End If
    Application.StatusBar = "Processing for month: " & targetMonthName

    ' --- Refresh Power Query Data ---
    Application.StatusBar = "Refreshing FDA data from Power Query..."
     ' LogEvt "Refresh", lvlINFO, "Attempting PQ refresh for " & PQ_CONNECTION_NAME
    If Not RefreshPowerQuery(PQ_CONNECTION_NAME) Then GoTo ErrorHandler ' Stop on critical PQ error

    ' --- Get Data Table & Check for Data ---
    On Error Resume Next: Set tblData = wsData.ListObjects(1): On Error GoTo ProcessErrorHandler
    If tblData Is Nothing Then LogEvt "DataTable", lvlERROR, "Data table not found on " & DATA_SHEET_NAME: GoTo ErrorHandler
     ' LogEvt "DataTable", lvlINFO, "Found data table: " & tblData.Name
    If tblData.ListRows.Count = 0 Then
         ' LogEvt "DataTable", lvlWARN, "No data returned by Power Query for " & targetMonthName & "."
        MsgBox "No data returned by Power Query for " & targetMonthName & ". Nothing to process.", vbInformation, "No Data"
        GoTo CleanExit
    End If
    recordCount = tblData.ListRows.Count
     ' LogEvt "DataTable", lvlINFO, "Table contains " & recordCount & " rows."

    ' --- Add/Verify Output Columns ---
     ' LogEvt "Columns", lvlINFO, "Checking/Adding scoring output columns..."
    If Not AddScoreColumnsIfNeeded(tblData) Then GoTo ErrorHandler ' Exit if columns can't be added/verified

    ' --- Map Column Headers to Indices (CRITICAL after adding columns) ---
    Set colIndices = GetColumnIndices(tblData.HeaderRowRange)
    If colIndices Is Nothing Then GoTo ErrorHandler ' Exit if required columns verification fails

    ' --- Load Weights, Keywords, and Cache ---
    Application.StatusBar = "Loading scoring parameters and cache..."
     ' LogEvt "LoadParams", lvlINFO, "Loading weights, keywords, and cache..."
    If Not LoadWeightsAndKeywords(wsWeights) Then GoTo ErrorHandler ' Treat failure to load weights as critical
    Call LoadCompanyCache(wsCache) ' Load cache (handles errors internally)

    ' --- Read Data into Array for Fast Processing ---
    Application.StatusBar = "Reading data into memory (" & recordCount & " rows)..."
     ' LogEvt "ReadData", lvlINFO, "Reading data into array..."
    dataArr = tblData.DataBodyRange.Value2 ' Read all data at once for speed
     ' LogEvt "ReadData", lvlINFO, "Read " & recordCount & " records into array."

    ' --- Main Processing Loop ---
    Application.StatusBar = "Calculating scores and fetching recaps (0% Complete)..."
     ' LogEvt "ScoreLoop", lvlINFO, "Starting main processing loop for " & recordCount & " records."
    useOpenAI = IsMaintainerUser() ' Determine if OpenAI can be used for this run

    For i = 1 To recordCount ' Loop through rows in the array (1-based)
        ' Calculate score components for the current row
        scoreResult = Calculate510kScore(dataArr, i, colIndices) ' Pass array, row index, column map

        ' Get company recap (check cache first, then optionally call OpenAI)
        Dim companyName As String
        On Error Resume Next: companyName = Trim(CStr(dataArr(i, colIndices("Applicant")))): On Error GoTo ProcessErrorHandler
        If Len(companyName) > 0 Then
            currentRecap = GetCompanyRecap(companyName, useOpenAI)
        Else
            currentRecap = "Invalid Applicant Name"
             ' LogEvt "ScoreLoop", lvlWARN, "Row " & i & ": Invalid/blank Applicant name."
        End If

        ' Write calculated results back into the data array (in memory)
        WriteResultsToArray dataArr, i, colIndices, scoreResult, currentRecap

        ' Update status bar periodically
        If i Mod 50 = 0 Or i = recordCount Then ' Update every 50 rows and at the end
            Application.StatusBar = "Calculating scores and fetching recaps (" & Format(i / recordCount, "0%") & " Complete)..."
            DoEvents ' Allow Excel to remain responsive on large datasets
        End If
    Next i
     ' LogEvt "ScoreLoop", lvlINFO, "Main processing loop complete."

    ' --- Write Processed Array Back to Sheet ---
    Application.StatusBar = "Writing results back to Excel sheet..."
     ' LogEvt "WriteBack", lvlINFO, "Writing " & recordCount & " rows back to table '" & tblData.Name & "'."
    tblData.DataBodyRange.Value = dataArr ' Use .Value to preserve date formats better than .Value2 sometimes
     ' LogEvt "WriteBack", lvlINFO, "Array write complete."

    ' --- Apply Number Formats (Essential before sorting/display) ---
     ' LogEvt "Formatting", lvlINFO, "Applying number formats."
    ApplyNumberFormats tblData

    ' --- Sort Table by DecisionDate (Newest First) ---
    Application.StatusBar = "Sorting data..."
     ' LogEvt "Sort", lvlINFO, "Sorting table by Decision Date (Descending)."
    SortDataTable tblData, "DecisionDate", xlDescending

    ' --- Save Updated Company Cache ---
    Application.StatusBar = "Saving company cache..."
    If Not wsCache Is Nothing And Not dictCache Is Nothing Then
        If dictCache.Count > 0 Then
             ' LogEvt "SaveCache", lvlINFO, "Saving " & dictCache.Count & " items to cache sheet '" & wsCache.Name & "'."
            Call SaveCompanyCache(wsCache)
        Else
             ' LogEvt "SaveCache", lvlINFO, "In-memory cache is empty, skipping save to sheet."
        End If
    End If

    ' --- Final Layout, Formatting & Visual Polish ---
    Application.StatusBar = "Applying final layout and formatting..."
     ' LogEvt "Formatting", lvlINFO, "Applying final layout and formatting."
    ' Call ReorganizeColumns(tblData) ' Uncomment and adjust array in function if needed
    Call FormatTableLook(wsData)
    Call FormatCategoryColors(tblData)
    Call CreateShortNamesAndComments(tblData) ' Shorten long device names
    Call FreezeHeaderAndKeyCols(wsData) ' Freeze panes
     ' LogEvt "Formatting", lvlINFO, "Final formatting applied."

    ' --- Archive Month ---
    Application.StatusBar = "Archiving month: " & targetMonthName & "..."
     ' LogEvt "Archive", lvlINFO, "Starting archive for " & targetMonthName & "."
    Call ArchiveMonth(wsData, archiveSheetName) ' Pass wsData and the calculated name

    ' --- Completion Message ---
    Dim endTime As Double: endTime = Timer
    Dim elapsed As String: elapsed = Format(endTime - startTime, "0.00")
     ' LogEvt "ProcessEnd", lvlINFO, "Processing completed successfully.", "Records=" & recordCount & ", Elapsed=" & elapsed & "s"
    Application.StatusBar = "Processing complete for " & targetMonthName & "."
    MsgBox "Monthly 510(k) data processed and archived for " & targetMonthName & "." & vbCrLf & vbCrLf & _
           "Processed " & recordCount & " records in " & elapsed & " seconds.", vbInformation, "Processing Complete"

CleanExit: ' Label for cleanup code (reached on success or after error handling)
     ' LogEvt "Cleanup", lvlINFO, "CleanExit reached. Releasing objects and restoring settings."
    ' Release objects
    Set dictACWeights = Nothing: Set dictSTWeights = Nothing: Set dictPCWeights = Nothing
    Set keywordsList = Nothing: Set nfCosmeticKeywordsList = Nothing
    Set nfDiagnosticKeywordsList = Nothing: Set therapeuticKeywordsList = Nothing
    Set dictCache = Nothing: Set colIndices = Nothing
    Set wsData = Nothing: Set wsWeights = Nothing: Set wsCache = Nothing: Set wsLog = Nothing: Set tblData = Nothing
    ' Restore application settings
    If Not Application.ScreenUpdating Then Application.ScreenUpdating = True
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculation = xlCalculationAutomatic
    If Not Application.EnableEvents Then Application.EnableEvents = True
    Application.StatusBar = False
    Application.Cursor = xlDefault
    Debug.Print Time & " - ProcessMonthly510k Finished. Objects released."
     ' FlushLogBuf ' Call final log flush if using logger module
    Exit Sub ' End of the Sub

ProcessErrorHandler: ' Central error handler for the main sub
      Dim errNum As Long: errNum = Err.Number
      Dim errDesc As String: errDesc = Err.Description
      Dim errSource As String: errSource = Err.Source
      ' Log the error
       ' LogEvt "ProcessError", lvlFATAL, "Error #" & errNum & " in " & errSource & ": " & errDesc
       ' FlushLogBuf ' Attempt immediate flush
      ' Display detailed error info
      Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
      Debug.Print Time & " - FATAL ERROR #" & errNum & " in " & errSource & " (ProcessMonthly510k): " & errDesc
      Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
      MsgBox "A critical error occurred during 510(k) processing:" & vbCrLf & vbCrLf & _
             "Error Number: " & errNum & vbCrLf & _
             "Module: mod_510k_Processor" & vbCrLf & _
             "Procedure: ProcessMonthly510k" & vbCrLf & _
             "Source: " & errSource & vbCrLf & _
             "Description: " & errDesc & vbCrLf & vbCrLf & _
             "Processing has been stopped.", vbCritical, "Processing Error"
      Resume CleanExit ' Jump to cleanup after fatal error
End Sub ' End of ProcessMonthly510k


' ==========================================================================
' ===                    CORE SCORING FUNCTION                         ===
' ==========================================================================
Private Function Calculate510kScore(dataArr As Variant, rowIdx As Long, ByVal cols As Object) As Variant
    ' Purpose: Calculates the 510(k) score based on various factors for a single record.
    ' Inputs:  dataArr - The 2D variant array holding all data.
    '          rowIdx - The current row number being processed in the array.
    '          cols - Dictionary mapping column names to their indices.
    ' Returns: A Variant array containing score components:
    '          Array(0=FinalScore, 1=Category, 2=AC_Wt, 3=PC_Wt, 4=KW_Wt, 5=ST_Wt,
    '                6=PT_Wt, 7=GL_Wt, 8=NF_Calc, 9=Synergy_Calc)

    ' --- Variable Declarations ---
    Dim AC As String, PC As String, DeviceName As String, Statement As String, SubmType As String, Country As String
    Dim ProcTimeDays As Variant, combinedText As String
    Dim AC_Wt As Double, PC_Wt As Double, KW_Wt As Double, ST_Wt As Double, PT_Wt As Double, GL_Wt As Double
    Dim NF_Calc As Double, Synergy_Calc As Double, Final_Score_Raw As Double
    Dim Category As String
    Dim HasHighValueKW As Boolean, IsCosmetic As Boolean, IsDiagnostic As Boolean, HasTherapeuticMention As Boolean
    Dim kw As Variant, kNum As String ' For loops and error logging

    ' --- Error Handling for this Function ---
    On Error GoTo ScoreErrorHandler

    ' --- 1. Extract Data Using Column Indices ---
    ' Use helper function for safe extraction
    AC = SafeGetString(dataArr, rowIdx, cols, "AC")
    PC = SafeGetString(dataArr, rowIdx, cols, "PC")
    DeviceName = SafeGetString(dataArr, rowIdx, cols, "DeviceName")
    Statement = SafeGetString(dataArr, rowIdx, cols, "Statement")
    SubmType = SafeGetString(dataArr, rowIdx, cols, "SubmType")
    Country = UCase(SafeGetString(dataArr, rowIdx, cols, "Country"))
    ProcTimeDays = SafeGetVariant(dataArr, rowIdx, cols, "ProcTimeDays")
    combinedText = DeviceName & " " & Statement ' For keyword searching
    kNum = SafeGetString(dataArr, rowIdx, cols, "K_Number") ' For logging context

    ' --- 2. Calculate Individual Weights ---
    ' Lookup weights using dictionaries, applying defaults if key not found
    AC_Wt = GetWeightFromDict(dictACWeights, AC, DEFAULT_AC_WEIGHT)
    PC_Wt = GetWeightFromDict(dictPCWeights, PC, DEFAULT_PC_WEIGHT)
    ST_Wt = GetWeightFromDict(dictSTWeights, SubmType, DEFAULT_ST_WEIGHT)

    ' Processing Time Weight (using Select Case for clarity)
    If IsNumeric(ProcTimeDays) Then
        Select Case CDbl(ProcTimeDays)
            Case Is > 172: PT_Wt = 0.65
            Case 162 To 172: PT_Wt = 0.6
            Case Else: PT_Wt = 0.5
        End Select
    Else: PT_Wt = DEFAULT_PT_WEIGHT
    End If

    ' Geographic Location Weight
    If Country = "US" Then GL_Wt = US_GL_WEIGHT Else GL_Wt = OTHER_GL_WEIGHT

    ' Keyword Weight & High KW Flag (using helper function)
    HasHighValueKW = CheckKeywords(combinedText, highValKeywordsList)
    If HasHighValueKW Then KW_Wt = HIGH_KW_WEIGHT Else KW_Wt = LOW_KW_WEIGHT

    ' --- 3. Negative Factors (NF) & Synergy Logic (Using Keyword Lists) ---
    NF_Calc = 0: Synergy_Calc = 0
    ' Check keyword categories using helper function
    IsCosmetic = CheckKeywords(combinedText, nfCosmeticKeywordsList)
    IsDiagnostic = CheckKeywords(combinedText, nfDiagnosticKeywordsList)
    HasTherapeuticMention = CheckKeywords(combinedText, therapeuticKeywordsList)

    ' Apply NF based on refined rules - *** REVIEW/CONFIRM THIS LOGIC ***
    ' Reference: FDA510k_AI_Refined_Methodologies..., FDA510k_AI_Alignment... docs
    If IsCosmetic And Not HasTherapeuticMention Then NF_Calc = NF_COSMETIC
    If IsDiagnostic And Not HasTherapeuticMention Then
        ' Check if NF_COSMETIC was already applied; avoid double penalty unless intended
        If NF_Calc = 0 Then NF_Calc = NF_DIAGNOSTIC Else NF_Calc = NF_Calc + NF_DIAGNOSTIC ' Additive? Confirm requirement
    End If

    ' Apply Synergy - *** CONFIRM THIS LOGIC ***
    ' Reference: FDA510k_AI_Synergy_Implementation... doc
    If (AC = "OR" Or AC = "NE") And HasHighValueKW Then Synergy_Calc = SYNERGY_BONUS

    ' --- 4. Final Score Calculation ---
    ' Reference: FDA510k_AI_Granular_Weights... doc for formula
    Final_Score_Raw = (AC_Wt + PC_Wt + KW_Wt + ST_Wt + PT_Wt + GL_Wt + NF_Calc + Synergy_Calc) / 6 ' <<< CONFIRM divisor (6?)
    If Final_Score_Raw < 0 Then Final_Score_Raw = 0 ' Ensure score doesn't go below zero

    ' --- 5. Determine Category ---
    ' Reference: FDA510k_AI_Granular_Weights... doc for thresholds
    Select Case Final_Score_Raw
        Case Is > 0.6: Category = "High"
        Case 0.5 To 0.6: Category = "Moderate" ' Includes 0.5 and 0.6 exactly
        Case 0.4 To 0.499999999999: Category = "Low" ' Up to (but not including) 0.5
        Case Else: Category = "Almost None"
    End Select

    ' --- 6. Return Results ---
    ' Return all calculated components for potential analysis or display
    Calculate510kScore = Array(Final_Score_Raw, Category, AC_Wt, PC_Wt, KW_Wt, ST_Wt, PT_Wt, GL_Wt, NF_Calc, Synergy_Calc)
    Exit Function ' Normal successful exit

ScoreErrorHandler:
    ' --- Error Handling specific to this function ---
    Dim errDesc As String: errDesc = Err.Description
     ' LogEvt "ScoreError", lvlERROR, "Error scoring row " & rowIdx & " (K#: " & kNum & "): " & errDesc, "AC=" & AC & ", PC=" & PC
    Debug.Print Time & " - ERROR scoring row " & rowIdx & " (K#: " & kNum & "): " & errDesc
    ' Return a default/error array structure consistent with the expected return type
    Calculate510kScore = Array(0, "Error", 0, 0, 0, 0, 0, 0, 0, 0)
End Function


' ==========================================================================
' ===                COMPANY RECAP & CACHING FUNCTIONS                   ===
' ==========================================================================

Private Function GetCompanyRecap(companyName As String, useOpenAI As Boolean) As String
    ' Purpose: Retrieves company recap. Checks memory cache first, then optionally calls OpenAI (if maintainer).
    '          Updates memory cache. Persistent saving happens separately.
    Dim finalRecap As String

    ' Initialize dictionary if it hasn't been created yet (first call)
    If dictCache Is Nothing Then Set dictCache = CreateObject("Scripting.Dictionary"): dictCache.CompareMode = vbTextCompare

    ' Handle blank company names
    If Len(Trim(companyName)) = 0 Then GetCompanyRecap = "Invalid Applicant Name": Exit Function

    ' 1. Check In-Memory Cache for the current run
    If dictCache.Exists(companyName) Then
        finalRecap = dictCache(companyName) ' Cache HIT
         ' LogEvt "CacheCheck", lvlDEBUG, "Memory Cache HIT.", "Company=" & companyName
    Else
        ' 2. Cache MISS - Need to determine the recap
         ' LogEvt "CacheCheck", lvlDEBUG, "Memory Cache MISS.", "Company=" & companyName
        finalRecap = DEFAULT_RECAP_TEXT ' Assume default initially

        ' 3. Attempt OpenAI *only if* flag is set (meaning maintainer is running)
        If useOpenAI Then
            Dim openAIResult As String
             ' LogEvt "OpenAI", lvlINFO, "Attempting OpenAI call.", "Company=" & companyName
            openAIResult = GetCompanyRecapOpenAI(companyName) ' Call the API function

            If openAIResult <> "" And Not LCase(openAIResult) Like "error:*" Then
                finalRecap = openAIResult ' Success - Use AI result
                 ' LogEvt "OpenAI", lvlINFO, "OpenAI SUCCESS.", "Company=" & companyName
            Else
                 ' Log API Error or skipped call (handled inside GetCompanyRecapOpenAI logging)
                 ' Keep the default "Needs Research" recap if OpenAI fails or is skipped
                 ' LogEvt "OpenAI", IIf(LCase(openAIResult) Like "error:*", lvlERROR, lvlWARN), "OpenAI Failed or Skipped. Result: " & openAIResult, "Company=" & companyName
            End If
        Else
             ' LogEvt "OpenAI", lvlINFO, "OpenAI call skipped (Not Maintainer).", "Company=" & companyName
             ' Keep default "Needs Research" as OpenAI wasn't attempted
        End If

        ' 4. Add the determined recap (Default or AI Result) to the memory cache for this run
        On Error Resume Next ' Handle potential error adding to dictionary (e.g., unexpected key issue)
        dictCache(companyName) = finalRecap
        If Err.Number <> 0 Then LogEvt "CacheUpdate", lvlERROR, "Error adding '" & companyName & "' to memory cache: " & Err.Description: Err.Clear
        On Error GoTo 0
    End If

    GetCompanyRecap = finalRecap ' Return the final recap string
End Function

Private Function GetCompanyRecapOpenAI(companyName As String) As String
    ' Purpose: Calls OpenAI API to get a company summary.
    ' WARNING: Requires robust JSON parsing library/implementation for production use.
    '          Includes basic error handling and secure key retrieval structure.

    Dim apiKey As String, result As String, http As Object, url As String, jsonPayload As String, jsonResponse As String
    GetCompanyRecapOpenAI = "" ' Default return

    ' --- Pre-checks ---
    If Not IsMaintainerUser() Then
         ' LogEvt "OpenAI_Skip", lvlINFO, "Skipped OpenAI Call: Not Maintainer User.", "Company=" & companyName
        Exit Function ' Silently exit if not the designated maintainer
    End If

    apiKey = GetAPIKey() ' Read key using helper function
    If apiKey = "" Then
         ' LogEvt "OpenAI_Skip", lvlERROR, "Skipped OpenAI Call: API Key Not Found/Configured.", "Company=" & companyName
        GetCompanyRecapOpenAI = "Error: API Key Not Configured"
        Exit Function
    End If

    On Error GoTo OpenAIErrorHandler

    ' --- Construct API Request ---
    url = OPENAI_API_URL
    Dim modelName As String: modelName = OPENAI_MODEL
    Dim systemPrompt As String, userPrompt As String
    ' Craft a specific prompt for concise, relevant summaries
    systemPrompt = "You are an analyst summarizing medical device related companies based *only* on publicly available information. " & _
                   "Provide a *neutral*, *very concise* (1 sentence ideally, 2 max) summary of the company '" & Replace(companyName, """", "'") & "' " & _
                   "identifying its primary business sector or main product type (e.g., orthopedics, diagnostics, surgical tools, contract manufacturer). " & _
                   "If unsure or company is generic, state 'General medical device company'. Avoid speculation or marketing language."
    userPrompt = "Summarize: " & companyName ' Keep user prompt simple

    ' Construct JSON Payload (Handle quotes carefully)
    jsonPayload = "{""model"": """ & modelName & """, ""messages"": [" & _
                  "{""role"": ""system"", ""content"": """ & JsonEscape(systemPrompt) & """}," & _
                  "{""role"": ""user"", ""content"": """ & JsonEscape(userPrompt) & """}" & _
                  "], ""temperature"": 0.3, ""max_tokens"": " & OPENAI_MAX_TOKENS & "}" ' Lower temp for factual summary

    ' --- Make HTTP Request ---
     ' LogEvt "OpenAI_HTTP", lvlINFO, "Sending request...", "Company=" & companyName & ", Model=" & modelName
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False ' Synchronous call
    http.setTimeouts OPENAI_TIMEOUT_MS, OPENAI_TIMEOUT_MS, OPENAI_TIMEOUT_MS, OPENAI_TIMEOUT_MS ' Set timeouts
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send jsonPayload

     ' LogEvt "OpenAI_HTTP", lvlINFO, "Response Received.", "Company=" & companyName & ", Status=" & http.Status

    ' --- Process Response ---
    If http.Status = 200 Then ' Success
        jsonResponse = http.responseText
        ' --- Basic JSON Parsing (Replace with robust parser if available) ---
        Const CONTENT_TAG As String = """content"":"""
        Dim contentStart As Long, contentEnd As Long, searchStart As Long
        searchStart = InStr(1, jsonResponse, """role"":""assistant""") ' Find assistant message block
        If searchStart > 0 Then
            contentStart = InStr(searchStart, jsonResponse, CONTENT_TAG)
            If contentStart > 0 Then
                contentStart = contentStart + Len(CONTENT_TAG)
                contentEnd = InStr(contentStart, jsonResponse, """") ' Find closing quote
                If contentEnd > contentStart Then
                    result = Mid$(jsonResponse, contentStart, contentEnd - contentStart)
                    result = JsonUnescape(result) ' Basic unescape function
                Else: result = "Error: Parse Fail (End Quote)": LogEvt "OpenAI_Parse", lvlERROR, result, Left(jsonResponse, 500)
                End If
            Else ' Check for API error structure before declaring parse fail
                If InStr(1, jsonResponse, """error""", vbTextCompare) > 0 Then
                     result = "Error: API returned error object."
                     LogEvt "OpenAI_APIError", lvlERROR, result, Left(jsonResponse, 500)
                Else
                     result = "Error: Parse Fail (Start Tag)"
                     LogEvt "OpenAI_Parse", lvlERROR, result, Left(jsonResponse, 500)
                End If
            End If
        Else: result = "Error: Parse Fail (No Assistant Role)": LogEvt "OpenAI_Parse", lvlERROR, result, Left(jsonResponse, 500)
        End If
    Else ' HTTP Error
        result = "Error: API Call Failed - Status " & http.Status & " - " & http.statusText
         ' LogEvt "OpenAI_HTTP", lvlERROR, result, Left(http.responseText, 500)
    End If

    Set http = Nothing
    ' Truncate if too long for Excel cell (should be rare with max_tokens)
    If Len(result) > RECAP_MAX_LEN Then result = Left$(result, RECAP_MAX_LEN) & "..."
    GetCompanyRecapOpenAI = Trim(result)
    Exit Function ' Normal exit after processing response

OpenAIErrorHandler:
    ' --- Error Handling specific to this function ---
    Dim errDesc As String: errDesc = Err.Description
     ' LogEvt "OpenAI_VBAError", lvlERROR, "VBA Exception during OpenAI Call: " & errDesc, "Company=" & companyName
    GetCompanyRecapOpenAI = "Error: VBA Exception - " & errDesc
    If Not http Is Nothing Then Set http = Nothing ' Ensure object release on error
End Function

Private Sub LoadCompanyCache(wsCache As Worksheet)
    ' Loads persistent cache from sheet into the in-memory dictionary.
    Dim lastRow As Long, i As Long, cacheData As Variant, loadedCount As Long
    Set dictCache = CreateObject("Scripting.Dictionary"): dictCache.CompareMode = vbTextCompare
    On Error Resume Next ' Handle sheet/table errors gracefully
    lastRow = wsCache.Cells(wsCache.Rows.Count, "A").End(xlUp).Row
    If Err.Number <> 0 Or lastRow < 2 Then GoTo ExitLoadCache ' Error finding last row or Empty sheet
    On Error GoTo CacheLoadError ' Proper handler for read errors
    cacheData = wsCache.Range("A2:B" & lastRow).Value2 ' Read CompanyName and RecapText
    If Err.Number <> 0 Then GoTo CacheLoadError ' Error reading range

    If IsArray(cacheData) Then
        For i = 1 To UBound(cacheData, 1)
            Dim k As String: k = Trim(CStr(cacheData(i, 1)))
            Dim v As String: v = CStr(cacheData(i, 2))
            If Len(k) > 0 Then ' Ensure key is not blank
                If Not dictCache.Exists(k) Then dictCache.Add k, v Else dictCache(k) = v ' Add or overwrite
            End If
        Next i
    ElseIf lastRow = 2 Then ' Handle potential single data row case
         Dim kS As String: kS = Trim(CStr(wsCache.Range("A2").Value2))
         Dim vS As String: vS = CStr(wsCache.Range("B2").Value2)
         If Len(kS) > 0 Then dictCache(kS) = vS
    End If
ExitLoadCache:
    loadedCount = dictCache.Count
     ' LogEvt "LoadCache", IIf(Err.Number <> 0 And loadedCount = 0, lvlWARN, lvlINFO), "Loaded " & loadedCount & " items into memory cache.", IIf(Err.Number <> 0 And loadedCount = 0, "Sheet might be empty or error occurred finding data.", "")
    Err.Clear ' Clear any residual error state from loading attempts
    Exit Sub
CacheLoadError:
     ' LogEvt "LoadCache", lvlERROR, "Error reading cache data from sheet: " & Err.Description
     Err.Clear ' Clear error and exit with potentially partial/empty cache
     Resume ExitLoadCache
End Sub

Private Sub SaveCompanyCache(wsCache As Worksheet)
    ' Saves the in-memory cache dictionary back to the persistent sheet cache.
    Dim key As Variant, i As Long, outputArr() As Variant, saveCount As Long
    If dictCache Is Nothing Or dictCache.Count = 0 Then LogEvt "SaveCache", lvlINFO, "In-memory cache empty, skipping save.": Exit Sub

    On Error GoTo CacheSaveError
    saveCount = dictCache.Count
    ReDim outputArr(1 To saveCount, 1 To 3) ' CompanyName, RecapText, LastUpdated
    i = 1
    For Each key In dictCache.Keys
        outputArr(i, 1) = key
        outputArr(i, 2) = dictCache(key)
        outputArr(i, 3) = Now ' Update timestamp on save
        i = i + 1
    Next key

    ' Turn off events/calc during write for speed/stability
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    With wsCache
        .Range("A2:C" & .Rows.Count).ClearContents ' Clear old cache data (keeping header)
        If saveCount > 0 Then
            .Range("A2").Resize(saveCount, 3).Value = outputArr ' Write new data
            .Range("C2").Resize(saveCount, 1).NumberFormat = "m/d/yyyy h:mm AM/PM" ' Format date
        End If
    End With
     ' LogEvt "SaveCache", lvlINFO, "Saved " & saveCount & " items to cache sheet."
    GoTo CacheSaveExit ' Jump to cleanup

CacheSaveError:
     ' LogEvt "SaveCache", lvlERROR, "Error saving cache to sheet '" & wsCache.Name & "': " & Err.Description
    MsgBox "Error saving company cache to sheet '" & wsCache.Name & "': " & Err.Description, vbExclamation, "Cache Save Error"
CacheSaveExit:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub


' ==========================================================================
' ===                WEIGHTS & KEYWORDS LOADING FUNCTIONS                ===
' ==========================================================================

Private Function LoadWeightsAndKeywords(wsWeights As Worksheet) As Boolean
    ' Purpose: Loads scoring weights and keywords from named tables on the Weights sheet.
    '          Initializes module-level dictionaries/collections. Returns False on critical failure.
    Dim success As Boolean: success = True ' Assume success unless critical error

    On Error GoTo LoadErrorHandler ' Central handler for this function

    ' LogEvt "LoadParams", lvlINFO, "Attempting to load weights and keywords from sheet: " & wsWeights.Name

    ' Load required weights/keywords
    Set dictACWeights = LoadTableToDict(wsWeights, "tblACWeights")
    Set dictSTWeights = LoadTableToDict(wsWeights, "tblSTWeights")
    Set dictPCWeights = LoadTableToDict(wsWeights, "tblPCWeights") ' Optional table
    Set highValKeywordsList = LoadTableToList(wsWeights, "tblKeywords")

    ' Load optional keyword lists for NF logic (create empty collections if tables don't exist)
    Set nfCosmeticKeywordsList = LoadTableToList(wsWeights, "tblNFCosmeticKeywords") ' Assumes this table exists
    Set nfDiagnosticKeywordsList = LoadTableToList(wsWeights, "tblNFDiagnosticKeywords") ' Assumes this table exists
    Set therapeuticKeywordsList = LoadTableToList(wsWeights, "tblTherapeuticKeywords") ' Assumes this table exists

    ' --- Validation ---
    ' Check if critical objects were created successfully (Load helpers return empty objects on non-critical errors)
    If dictACWeights Is Nothing Or dictSTWeights Is Nothing Or highValKeywordsList Is Nothing Then
         ' LogEvt "LoadParams", lvlFATAL, "Critical failure: Could not load AC/ST weights or HighValue Keywords."
        GoTo LoadErrorCritical ' Critical if essential weights/keywords fail
    End If

    ' Log counts (provides visibility into loaded data)
     ' LogEvt "LoadParams", IIf(dictACWeights.Count = 0, lvlWARN, lvlDETAIL), "Loaded " & dictACWeights.Count & " AC Weights."
     ' LogEvt "LoadParams", IIf(dictSTWeights.Count = 0, lvlWARN, lvlDETAIL), "Loaded " & dictSTWeights.Count & " ST Weights."
     ' LogEvt "LoadParams", IIf(dictPCWeights Is Nothing Or dictPCWeights.Count = 0, lvlINFO, lvlDETAIL), "Loaded " & IIf(dictPCWeights Is Nothing, 0, dictPCWeights.Count) & " PC Weights (Optional)."
     ' LogEvt "LoadParams", IIf(highValKeywordsList.Count = 0, lvlWARN, lvlDETAIL), "Loaded " & highValKeywordsList.Count & " HighVal Keywords."
     ' LogEvt "LoadParams", IIf(nfCosmeticKeywordsList.Count = 0, lvlINFO, lvlDETAIL), "Loaded " & nfCosmeticKeywordsList.Count & " Cosmetic Keywords (Optional)."
     ' LogEvt "LoadParams", IIf(nfDiagnosticKeywordsList.Count = 0, lvlINFO, lvlDETAIL), "Loaded " & nfDiagnosticKeywordsList.Count & " Diagnostic Keywords (Optional)."
     ' LogEvt "LoadParams", IIf(therapeuticKeywordsList.Count = 0, lvlINFO, lvlDETAIL), "Loaded " & therapeuticKeywordsList.Count & " Therapeutic Keywords (Optional)."

    LoadWeightsAndKeywords = True ' Success (even if optional lists are empty)
    Exit Function ' Normal Exit

LoadErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
     ' LogEvt "LoadParams", lvlWARN, "Non-critical error loading one or more weight/keyword tables: " & errDesc & ". Defaults will be used.", "Sheet=" & wsWeights.Name
    MsgBox "Warning: Error loading weights/keywords from '" & wsWeights.Name & "'. " & vbCrLf & _
           "Check tables exist and are named correctly (e.g., tblACWeights, tblKeywords, tblNFCosmeticKeywords etc.)." & vbCrLf & _
           "Default values will be used where possible.", vbExclamation, "Load Warning"
    ' Ensure objects are initialized even on non-critical error to prevent crashes later
    If dictACWeights Is Nothing Then Set dictACWeights = CreateObject("Scripting.Dictionary")
    If dictSTWeights Is Nothing Then Set dictSTWeights = CreateObject("Scripting.Dictionary")
    If dictPCWeights Is Nothing Then Set dictPCWeights = CreateObject("Scripting.Dictionary")
    If highValKeywordsList Is Nothing Then Set highValKeywordsList = New Collection
    If nfCosmeticKeywordsList Is Nothing Then Set nfCosmeticKeywordsList = New Collection
    If nfDiagnosticKeywordsList Is Nothing Then Set nfDiagnosticKeywordsList = New Collection
    If therapeuticKeywordsList Is Nothing Then Set therapeuticKeywordsList = New Collection
    LoadWeightsAndKeywords = True ' Allow continue with defaults
    Exit Function

LoadErrorCritical:
    ' LogEvt "LoadParams", lvlFATAL, "CRITICAL Error: Failed to initialize essential weight/keyword objects. Cannot continue processing."
    MsgBox "Critical Error: Could not create necessary objects for AC/ST weights or HighValue Keywords. Processing cannot continue.", vbCritical, "Load Failure"
    ' Clean up any partially created objects
    Set dictACWeights = Nothing: Set dictSTWeights = Nothing: Set dictPCWeights = Nothing
    Set highValKeywordsList = Nothing: Set nfCosmeticKeywordsList = Nothing
    Set nfDiagnosticKeywordsList = Nothing: Set therapeuticKeywordsList = Nothing
    LoadWeightsAndKeywords = False ' Signal critical failure
End Function

Private Function LoadTableToDict(ws As Worksheet, tableName As String) As Object ' Scripting.Dictionary
    ' Helper: Loads 2-column table to dictionary. Handles errors gracefully.
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Case-insensitive keys
    Dim tbl As ListObject, dataRange As Range, dataArr As Variant, i As Long, key As String, val As Variant

    On Error Resume Next ' Temporarily suppress errors for object checking
    Set tbl = ws.ListObjects(tableName)
    If Err.Number <> 0 Then GoTo ExitLoadTableDict ' Table not found
    If tbl.ListRows.Count = 0 Then GoTo ExitLoadTableDict ' Table empty
    Set dataRange = tbl.DataBodyRange
    If dataRange Is Nothing Or dataRange.Columns.Count < 2 Then GoTo ExitLoadTableDict ' Invalid structure
    dataArr = dataRange.Value2 ' Read data into array
    If Err.Number <> 0 Then GoTo ExitLoadTableDict ' Error reading data
    On Error GoTo 0 ' Restore normal error handling

    ' Process the array data
    If IsArray(dataArr) Then
        For i = 1 To UBound(dataArr, 1) ' Loop through rows
            key = Trim(CStr(dataArr(i, 1))) ' Key from Col 1
            val = dataArr(i, 2) ' Value from Col 2
            If Len(key) > 0 Then
                If Not dict.Exists(key) Then dict.Add key, val Else dict(key) = val ' Add or overwrite
            End If
        Next i
    ElseIf Not IsEmpty(dataArr) Then ' Handle single row data read (returns single value if 1x1, 1D array otherwise)
        ' This case needs careful handling depending on how Excel reads a single row table range
         key = Trim(CStr(tbl.DataBodyRange.Cells(1, 1).Value2))
         val = tbl.DataBodyRange.Cells(1, 2).Value2
         If Len(key) > 0 Then If Not dict.Exists(key) Then dict.Add key, val Else dict(key) = val
    End If

ExitLoadTableDict:
    If Err.Number <> 0 Then
         ' LogEvt "LoadHelper", lvlWARN, "Error loading table '" & tableName & "' to Dict: " & Err.Description
        Debug.Print Time & " - Error loading table '" & tableName & "' to Dict: " & Err.Description: Err.Clear
    End If
    Set LoadTableToDict = dict ' Return dictionary (possibly empty)
End Function

Private Function LoadTableToList(ws As Worksheet, tableName As String) As Collection
    ' Helper: Loads first column of table to collection. Handles errors gracefully.
    Dim coll As New Collection, tbl As ListObject, dataRange As Range, dataArr As Variant, i As Long, item As String
    On Error Resume Next ' Suppress errors for object checks
    Set tbl = ws.ListObjects(tableName)
    If Err.Number <> 0 Then GoTo ExitLoadTableList ' Table not found
    If tbl.ListRows.Count = 0 Then GoTo ExitLoadTableList ' Table empty
    Set dataRange = tbl.ListColumns(1).DataBodyRange ' Get first column data range
    If dataRange Is Nothing Then GoTo ExitLoadTableList ' Column/Range invalid
    dataArr = dataRange.Value2 ' Read data
    If Err.Number <> 0 Then GoTo ExitLoadTableList ' Error reading data
    On Error GoTo 0 ' Restore error handling

    ' Process array data
    If IsArray(dataArr) Then
        For i = 1 To UBound(dataArr, 1)
            item = Trim(CStr(dataArr(i, 1)))
            If Len(item) > 0 Then On Error Resume Next: coll.Add item, item: On Error GoTo 0 ' Add unique items
        Next i
    ElseIf Not IsEmpty(dataArr) Then ' Handle single row data
        item = Trim(CStr(dataArr))
        If Len(item) > 0 Then On Error Resume Next: coll.Add item, item: On Error GoTo 0
    End If

ExitLoadTableList:
    If Err.Number <> 0 Then
         ' LogEvt "LoadHelper", lvlWARN, "Error loading table '" & tableName & "' to List: " & Err.Description
        Debug.Print Time & " - Error loading table '" & tableName & "' to List: " & Err.Description: Err.Clear
    End If
    Set LoadTableToList = coll ' Return collection (possibly empty)
End Function

Private Function GetWeightFromDict(dict As Object, key As String, defaultWeight As Double) As Double
    ' Safely gets weight from dictionary, returning default if key not found or dict invalid.
    If dict Is Nothing Then GetWeightFromDict = defaultWeight: Exit Function
    If dict.Exists(key) Then GetWeightFromDict = dict(key) Else GetWeightFromDict = defaultWeight
End Function

Private Function CheckKeywords(textToCheck As String, keywordColl As Collection) As Boolean
    ' Checks if any keyword from the collection exists in the text (case-insensitive).
    Dim kw As Variant
    CheckKeywords = False ' Default to false
    If keywordColl Is Nothing Or keywordColl.Count = 0 Or Len(Trim(textToCheck)) = 0 Then Exit Function
    On Error Resume Next ' Ignore errors during loop (e.g., unexpected item type)
    For Each kw In keywordColl
        If InStr(1, textToCheck, CStr(kw), vbTextCompare) > 0 Then CheckKeywords = True: Exit For ' Found a match
    Next kw
    On Error GoTo 0
End Function


' ==========================================================================
' ===                   COLUMN MANAGEMENT FUNCTIONS                      ===
' ==========================================================================

Private Function AddScoreColumnsIfNeeded(tblData As ListObject) As Boolean
    ' Purpose: Adds the necessary output columns to the table if they don't exist. Returns True on success.
    Dim requiredCols As Variant, colName As Variant, col As ListColumn, addedCol As Boolean: addedCol = False
    On Error GoTo AddColErrorHandler

    ' Define all columns VBA might add/write to
    requiredCols = Array("AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", "GL_Wt", "NF_Calc", "Synergy_Calc", "Final_Score", "Score_Percent", "Category", "CompanyRecap")

     ' LogEvt "Columns", lvlDETAIL, "Verifying existence of calculated columns..."
    Dim currentHeaders As String, hCell As Range
    For Each hCell In tblData.HeaderRowRange: currentHeaders = currentHeaders & "|" & UCase(Trim(hCell.Value)) & "|": Next hCell ' Build header map

    For Each colName In requiredCols
        ' Check if column exists using the header map (faster than ListColumns object check in loop)
        If InStr(1, currentHeaders, "|" & UCase(colName) & "|", vbTextCompare) = 0 Then
            On Error Resume Next ' Temporarily handle error if column add fails
            Set col = tblData.ListColumns.Add()
            If Err.Number <> 0 Then GoTo AddColErrorHandler ' Critical failure adding column
            col.Name = colName ' Set name after adding
            addedCol = True
             ' LogEvt "Columns", lvlINFO, "Added missing column: " & colName
            On Error GoTo AddColErrorHandler ' Restore handler
        End If
        Set col = Nothing ' Release object
    Next colName

    ' If columns were added, need to resize array read later? No, read after adding.
    ' If addedCol Then Debug.Print Time & " - Columns added. Table structure potentially changed."

    AddScoreColumnsIfNeeded = True ' Success
    Exit Function

AddColErrorHandler:
     ' LogEvt "Columns", lvlERROR, "Error checking/adding columns to table '" & tblData.Name & "': " & Err.Description
    MsgBox "Error verifying or adding required columns to table '" & tblData.Name & "': " & vbCrLf & Err.Description, vbCritical, "Column Setup Error"
    AddScoreColumnsIfNeeded = False ' Signal failure
End Function

Private Sub WriteResultsToArray(ByRef dataArr As Variant, ByVal rowIdx As Long, ByVal cols As Object, ByVal scoreResult As Variant, ByVal recap As String)
    ' Purpose: Writes calculated score components and recap back into the main data array.
    ' Inputs: dataArr (passed ByRef to modify), row index, column index dictionary, scoreResult array, recap string.
    On Error Resume Next ' Suppress errors during write if a column index is bad (should be caught by GetColumnIndices)

    ' Write score components from scoreResult array
    dataArr(rowIdx, cols("Final_Score")) = scoreResult(0) ' Index 0 = FinalScore_Raw
    dataArr(rowIdx, cols("Category")) = scoreResult(1)    ' Index 1 = Category
    dataArr(rowIdx, cols("AC_Wt")) = scoreResult(2)
    dataArr(rowIdx, cols("PC_Wt")) = scoreResult(3)
    dataArr(rowIdx, cols("KW_Wt")) = scoreResult(4)
    dataArr(rowIdx, cols("ST_Wt")) = scoreResult(5)
    dataArr(rowIdx, cols("PT_Wt")) = scoreResult(6)
    dataArr(rowIdx, cols("GL_Wt")) = scoreResult(7)
    dataArr(rowIdx, cols("NF_Calc")) = scoreResult(8)
    dataArr(rowIdx, cols("Synergy_Calc")) = scoreResult(9)
    ' Calculate Percent on the fly for storage
    dataArr(rowIdx, cols("Score_Percent")) = scoreResult(0) * 100

    ' Write Company Recap
    dataArr(rowIdx, cols("CompanyRecap")) = recap

    If Err.Number <> 0 Then
         ' LogEvt "WriteToArray", lvlERROR, "Error writing results to array row " & rowIdx & ": " & Err.Description
        Debug.Print Time & " - Error writing results to array row " & rowIdx & ": " & Err.Description: Err.Clear
    End If
    On Error GoTo 0 ' Restore default error handling
End Sub

Private Function GetColumnIndices(headerRange As Range) As Object ' Scripting.Dictionary
    ' Purpose: Creates dictionary mapping header names to column indices; validates required columns exist.
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Case-insensitive keys
    Dim cell As Range, colNum As Long: colNum = 1, missingCols As String

    ' Define ALL columns needed (from PQ + VBA additions)
    Dim requiredCols As Variant
    requiredCols = Array("K_Number", "Applicant", "ContactName", "DeviceName", "DecisionDate", "DateReceived", "ProcTimeDays", "AC", "PC", "SubmType", "City", "State", "Country", "Statement", "FDA_Link", _
                         "AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", "GL_Wt", "NF_Calc", "Synergy_Calc", "Final_Score", "Score_Percent", "Category", "CompanyRecap")

    ' Map header names to column index numbers
    For Each cell In headerRange.Cells
        Dim h As String: h = Trim(cell.Value)
        If Len(h) > 0 Then
            If Not dict.Exists(h) Then
                dict.Add h, colNum
            Else
                ' Log duplicate header only once
                If colNum = dict(h) + 1 Then ' Check if it's the immediately following column
                     ' LogEvt "ColumnMap", lvlWARN, "Duplicate header found and ignored: '" & h & "' at column " & colNum
                    Debug.Print Time & " - Warning: Duplicate header found and ignored: '" & h & "' at column " & colNum
                End If
            End If
        End If
        colNum = colNum + 1
    Next cell

    ' --- Validation: Check if all required columns exist ---
    Dim reqCol As Variant
    For Each reqCol In requiredCols
        If Not dict.Exists(reqCol) Then
            missingCols = missingCols & vbCrLf & " - " & reqCol
        End If
    Next reqCol

    If Len(missingCols) > 0 Then
         ' LogEvt "ColumnMap", lvlERROR, "Required columns missing in table header:" & Replace(missingCols, vbCrLf, ", ")
        MsgBox "Error: The following required columns were not found in sheet '" & headerRange.Parent.Name & "':" & missingCols & vbCrLf & "Please ensure Power Query output and VBA column additions match.", vbCritical, "Missing Columns"
        Set GetColumnIndices = Nothing ' Signal error
    Else
         ' LogEvt "ColumnMap", lvlINFO, "Column indices mapped successfully for " & dict.Count & " columns."
        Set GetColumnIndices = dict ' Success
    End If
End Function

' ==========================================================================
' ===             LAYOUT, FORMATTING & PRESENTATION FUNCTIONS            ===
' ==========================================================================

Private Sub ReorganizeColumns(tbl As ListObject)
    ' Purpose: Rearranges table columns into the final desired presentation order.
    ' Note: Run AFTER all data processing and BEFORE formatting that relies on column position.
    On Error GoTo ReorgErrorHandler
    ' *** Define the desired final column order ***
    Dim targetOrder As Variant: targetOrder = Array( _
        "K_Number", "DecisionDate", "Applicant", "DeviceName", "ContactName", "CompanyRecap", "Score_Percent", "Category", "FDA_Link", _
        "DateReceived", "ProcTimeDays", "AC", "PC", "SubmType", "City", "State", "Country", "Statement", _
        "Final_Score", "NF_Calc", "Synergy_Calc", "AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", "GL_Wt" _
        ) ' <<< ADJUST THIS ARRAY TO YOUR PREFERRED FINAL LAYOUT >>>

    Dim targetPosition As Long, currentPosition As Long, col As ListColumn, currentColName As String

    Application.ScreenUpdating = False
     ' LogEvt "Formatting", lvlDETAIL, "Starting column reorganization..."
    Dim currentOrder As String, i As Long
    For i = LBound(targetOrder) To UBound(targetOrder)
        currentColName = targetOrder(i)
        targetPosition = i + 1 ' 1-based index for columns

        On Error Resume Next ' Check if column exists
        Set col = tbl.ListColumns(currentColName)
        Dim errNum As Long: errNum = Err.Number
        On Error GoTo ReorgErrorHandler ' Restore handler

        If errNum = 0 And Not col Is Nothing Then
            currentPosition = col.Index
            If currentPosition <> targetPosition Then
                col.Range.Move Destination:=tbl.HeaderRowRange.Cells(1, targetPosition)
                 ' LogEvt "Formatting", lvlDETAIL, "Moved column '" & currentColName & "' from " & currentPosition & " to " & targetPosition
            End If
            Set col = Nothing ' Release object
        Else
             ' LogEvt "Formatting", lvlWARN, "Column '" & currentColName & "' not found during reorganization. Skipping."
            Debug.Print Time & " - Warning: Column '" & currentColName & "' not found during reorg."
        End If
    Next i

    Application.ScreenUpdating = True
     ' LogEvt "Formatting", lvlINFO, "Column reorganization complete."
    Exit Sub

ReorgErrorHandler:
    Application.ScreenUpdating = True
     ' LogEvt "Formatting", lvlERROR, "Error during column reorganization: " & Err.Description
    MsgBox "Error occurred during column reorganization: " & Err.Description, vbCritical, "Column Reorder Error"
End Sub

Private Sub FormatTableLook(ws As Worksheet)
    ' Purpose: Applies consistent table style, alignment, widths, and borders.
    On Error GoTo FormatLookErrorHandler
    Dim tbl As ListObject: If ws.ListObjects.Count = 0 Then Exit Sub: Set tbl = ws.ListObjects(1)
    Dim listCol As ListColumn
    Dim centerCols As Variant: centerCols = Array("ProcTimeDays", "AC", "PC", "Category") ' Columns to center align
    Dim wideCols As Variant: wideCols = Array("DeviceName", "Applicant", "CompanyRecap") ' Columns needing more width
    Dim colName As Variant

     ' LogEvt "Formatting", lvlDETAIL, "Applying table style, alignment, widths, borders..."

    ' Apply Table Style
    tbl.TableStyle = "TableStyleMedium2" ' Choose a style (Medium2 is clean blue)

    ' Center specific columns
    For Each colName In centerCols
        On Error Resume Next: Set listCol = tbl.ListColumns(colName): If Not listCol Is Nothing Then listCol.DataBodyRange.HorizontalAlignment = xlCenter: listCol.HeaderRowRange.HorizontalAlignment = xlCenter: Set listCol = Nothing: On Error GoTo FormatLookErrorHandler
    Next colName

    ' Autofit all columns first
    On Error Resume Next: tbl.Range.Columns.AutoFit: On Error GoTo FormatLookErrorHandler

    ' Set specific widths for wider columns (adjust widths as needed)
    On Error Resume Next: If Not tbl.ListColumns("DeviceName") Is Nothing Then tbl.ListColumns("DeviceName").Range.ColumnWidth = 45
    On Error Resume Next: If Not tbl.ListColumns("Applicant") Is Nothing Then tbl.ListColumns("Applicant").Range.ColumnWidth = 30
    On Error Resume Next: If Not tbl.ListColumns("CompanyRecap") Is Nothing Then tbl.ListColumns("CompanyRecap").Range.ColumnWidth = 50
    On Error GoTo FormatLookErrorHandler

    ' Apply basic thin borders to all cells in the table range
    With tbl.Range.Borders: .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = xlAutomatic: End With

    ' Optional: Add a slightly thicker right border after key info columns if desired
    ' On Error Resume Next: If Not tbl.ListColumns("Category") Is Nothing Then With tbl.ListColumns("Category").Range.Borders(xlEdgeRight): .LineStyle = xlContinuous: .Weight = xlMedium: End With
    ' On Error Resume Next: If Not tbl.ListColumns("FDA_Link") Is Nothing Then With tbl.ListColumns("FDA_Link").Range.Borders(xlEdgeRight): .LineStyle = xlContinuous: .Weight = xlMedium: End With
    On Error GoTo FormatLookErrorHandler

     ' LogEvt "Formatting", lvlDETAIL, "Applied FormatTableLook."
    Exit Sub

FormatLookErrorHandler:
     ' LogEvt "Formatting", lvlERROR, "Error applying table look formatting: " & Err.Description
    Debug.Print Time & " - Error applying table look formatting: " & Err.Description
End Sub

Private Sub FormatCategoryColors(tblData As ListObject)
    ' Purpose: Applies conditional background/font colors based on the Category text.
    Dim rng As Range
    On Error Resume Next ' Skip if Category column doesn't exist
    Set rng = tblData.ListColumns("Category").DataBodyRange
    If rng Is Nothing Then Exit Sub
    On Error GoTo CatColorError

     ' LogEvt "Formatting", lvlDETAIL, "Applying category conditional formatting..."
    rng.FormatConditions.Delete ' Clear existing rules first

    ' Define colors (adjust RGB values as desired)
    Dim highColor As Long: highColor = RGB(198, 239, 206) ' Green fill
    Dim modColor As Long: modColor = RGB(255, 235, 156) ' Yellow fill
    Dim lowColor As Long: lowColor = RGB(255, 221, 179) ' Orange fill (or light orange)
    Dim noneColor As Long: noneColor = RGB(242, 242, 242) ' Light grey fill
    Dim errorColor As Long: errorColor = RGB(255, 199, 206) ' Light red fill

    ' Apply rules (order can matter if categories overlap, but shouldn't here)
    ApplyCondColor rng, "High", highColor
    ApplyCondColor rng, "Moderate", modColor
    ApplyCondColor rng, "Low", lowColor
    ApplyCondColor rng, "Almost None", noneColor
    ApplyCondColor rng, "Error", errorColor ' Highlight scoring errors

     ' LogEvt "Formatting", lvlDETAIL, "Applied category colors."
    Exit Sub

CatColorError:
     ' LogEvt "Formatting", lvlERROR, "Error applying category colors: " & Err.Description
    Debug.Print Time & " - Error applying category colors: " & Err.Description
End Sub

Private Sub ApplyCondColor(rng As Range, categoryText As String, fillColor As Long)
    ' Helper for FormatCategoryColors - Adds one rule.
    With rng.FormatConditions.Add(Type:=xlTextString, String:=categoryText, TextOperator:=xlEqual)
        .Interior.Color = fillColor
        .Font.Color = IIf(GetBrightness(fillColor) < 130, vbWhite, vbBlack) ' White font on dark, Black on light
        .StopIfTrue = False ' Allow multiple rules if needed, though typically not for exact text match
    End With
End Sub

Private Function GetBrightness(clr As Long) As Double
    ' Calculates perceived brightness (0-255) for font color contrast. Using standard formula.
    On Error Resume Next ' Handle potential invalid color value
    GetBrightness = ((clr Mod 256) * 0.299) + (((clr \ 256) Mod 256) * 0.587) + (((clr \ 65536) Mod 256) * 0.114)
    If Err.Number <> 0 Then GetBrightness = 128 ' Default brightness if error
    On Error GoTo 0
End Function

Private Sub ApplyNumberFormats(tblData As ListObject)
    ' Purpose: Applies consistent number formats to scoring columns.
    On Error Resume Next ' Ignore errors if a column doesn't exist
     ' LogEvt "Formatting", lvlDETAIL, "Applying number formats..."
    If Not tblData.ListColumns("Score_Percent") Is Nothing Then tblData.ListColumns("Score_Percent").DataBodyRange.NumberFormat = "0.0%"
    If Not tblData.ListColumns("Final_Score") Is Nothing Then tblData.ListColumns("Final_Score").DataBodyRange.NumberFormat = "0.000"
    Dim scoreWtCols As Variant: scoreWtCols = Array("AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", "GL_Wt", "NF_Calc", "Synergy_Calc")
    Dim colName As Variant
    For Each colName In scoreWtCols
        If Not tblData.ListColumns(colName) Is Nothing Then tblData.ListColumns(colName).DataBodyRange.NumberFormat = "0.00"
    Next colName
    If Not tblData.ListColumns("ProcTimeDays") Is Nothing Then tblData.ListColumns("ProcTimeDays").DataBodyRange.NumberFormat = "0"
     ' LogEvt "Formatting", lvlDETAIL, "Number formats applied."
    On Error GoTo 0 ' Restore default error handling
End Sub

Private Sub CreateShortNamesAndComments(tblData As ListObject)
    ' Purpose: Shortens long Device Names in the cell and adds the full name as a comment.
    Dim devNameCol As ListColumn, devNameRange As Range, cell As Range
    Dim originalName As String, shortName As String

    On Error Resume Next ' Skip if "DeviceName" column doesn't exist
    Set devNameCol = tblData.ListColumns("DeviceName")
    If devNameCol Is Nothing Then Exit Sub
    Set devNameRange = devNameCol.DataBodyRange
    If devNameRange Is Nothing Then Exit Sub
    On Error GoTo ShortNameErrorHandler

     ' LogEvt "Formatting", lvlDETAIL, "Applying short names and comments..."
    Application.ScreenUpdating = False ' Speed up loop

    For Each cell In devNameRange.Cells
        ' Clear existing comment first to prevent duplicates/errors
        If Not cell.Comment Is Nothing Then cell.Comment.Delete

        originalName = Trim(CStr(cell.Value))

        ' Only shorten and add comment if name exceeds max length
        If Len(originalName) > SHORT_NAME_MAX_LEN Then
            shortName = Left$(originalName, SHORT_NAME_MAX_LEN - Len(SHORT_NAME_ELLIPSIS)) & SHORT_NAME_ELLIPSIS

            ' Check if the cell *already* contains the shortened version (idempotency)
            If cell.Value <> shortName Then cell.Value = shortName ' Update cell value ONLY if necessary

            ' Add the full original name as a comment
            On Error Resume Next ' Handle potential error adding comment
            cell.AddComment Text:=originalName
            If Err.Number = 0 Then
                 cell.Comment.Shape.TextFrame.AutoSize = True ' Resize comment box
            Else
                 ' LogEvt "Formatting", lvlWARN, "Could not add comment to " & cell.Address & ": " & Err.Description
                Debug.Print Time & " - Warning: Could not add comment to cell " & cell.Address & ": " & Err.Description: Err.Clear
            End If
            On Error GoTo ShortNameErrorHandler ' Restore handler
        End If
    Next cell

    Application.ScreenUpdating = True
     ' LogEvt "Formatting", lvlDETAIL, "Short names/comments processing complete."
    Exit Sub

ShortNameErrorHandler:
    Application.ScreenUpdating = True
     ' LogEvt "Formatting", lvlERROR, "Error applying short names/comments: " & Err.Description
    MsgBox "Error applying short device names/comments: " & Err.Description, vbExclamation, "Short Name Error"
End Sub

Private Sub FreezeHeaderAndKeyCols(ws As Worksheet)
    ' Purpose: Freezes header row and key columns for better navigation.
    '          Freezes columns up to "Category" by default, or adjust as needed.
    On Error GoTo FreezeErrorHandler
    Dim tbl As ListObject, targetCol As ListColumn, freezeColIndex As Long
    Const COL_TO_FREEZE_AFTER As String = "Category" ' <<< ADJUST: Freeze columns up to and including this one
    Const FALLBACK_FREEZE_COL As Long = 4 ' Fallback if COL_TO_FREEZE_AFTER not found

    If ws.ListObjects.Count = 0 Then Exit Sub: Set tbl = ws.ListObjects(1)

     ' LogEvt "Formatting", lvlDETAIL, "Applying freeze panes..."
    On Error Resume Next ' Find the column index
    Set targetCol = tbl.ListColumns(COL_TO_FREEZE_AFTER)
    If targetCol Is Nothing Then freezeColIndex = FALLBACK_FREEZE_COL + 1 Else freezeColIndex = targetCol.Index + 1
    On Error GoTo FreezeErrorHandler ' Restore handler

    ' Ensure freeze index is valid
    If freezeColIndex < 2 Then freezeColIndex = 2
    If freezeColIndex > ws.Columns.Count Then freezeColIndex = ws.Columns.Count

    Dim targetCell As Range
    Set targetCell = ws.Cells(tbl.HeaderRowRange.Row + 1, freezeColIndex) ' Cell below header, right of last frozen column

    ' Apply Freeze Panes
    ws.Activate ' Sheet must be active
    If ActiveWindow.FreezePanes Then ActiveWindow.FreezePanes = False ' Unfreeze first
    targetCell.Select ' Select the cell to freeze relative to
    ActiveWindow.FreezePanes = True
    ws.Cells(1, 1).Select ' Select A1 for visual tidiness

     ' LogEvt "Formatting", lvlDETAIL, "Freeze Panes applied after column: " & COL_TO_FREEZE_AFTER & " (Index " & freezeColIndex - 1 & ")"
    Exit Sub

FreezeErrorHandler:
     ' LogEvt "Formatting", lvlERROR, "Error applying freeze panes: " & Err.Description
    Debug.Print Time & " - Error applying freeze panes: " & Err.Description
    ' Non-critical error, don't stop execution, but maybe notify user
    ' MsgBox "Could not apply freeze panes: " & Err.Description, vbExclamation, "Freeze Panes Error"
End Sub


' ==========================================================================
' ===                       ARCHIVING FUNCTION                         ===
' ==========================================================================
Private Sub ArchiveMonth(wsDataSource As Worksheet, archiveSheetName As String)
    ' Purpose: Archives data: Copies sheet, Renames, Converts to Values, Unlists Table, Protects.
    Dim wsArchive As Worksheet

    On Error GoTo ArchiveErrorHandler
    Application.DisplayAlerts = False ' Prevent overwrite prompts etc.
     ' LogEvt "Archive", lvlINFO, "Starting archive process for: " & archiveSheetName

    ' --- 1. Copy the Source Sheet ---
    wsDataSource.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set wsArchive = ActiveSheet ' The newly copied sheet

    ' --- 2. Rename the Copied Sheet ---
    On Error Resume Next ' Attempt rename
    wsArchive.Name = Left(archiveSheetName, 31) ' Enforce 31 char limit
    If Err.Number <> 0 Then ' Handle potential error (e.g., invalid chars, duplicate name somehow)
        Dim fallbackName As String: fallbackName = "Archive_Error_" & Format(Now(), "yyyyMMdd_HHmmss")
         ' LogEvt "Archive", lvlWARN, "Rename to '" & archiveSheetName & "' failed. Using fallback: " & fallbackName
        wsArchive.Name = fallbackName: Err.Clear
    End If
    On Error GoTo ArchiveErrorHandler ' Restore normal error handling

    ' --- 3. Convert Formulas to Static Values ---
    If wsArchive.UsedRange.Cells.CountLarge > 1 Then ' Check if sheet is not empty
        wsArchive.UsedRange.Value = wsArchive.UsedRange.Value
         ' LogEvt "Archive", lvlDETAIL, "Converted formulas to values on sheet: " & wsArchive.Name
    End If

    ' --- 4. Remove Table Object (Unlist) ---
    If wsArchive.ListObjects.Count > 0 Then
        On Error Resume Next ' Handle potential error during unlist
        wsArchive.ListObjects(1).Unlist
         ' LogEvt "Archive", lvlDETAIL, "Unlisted table on archive sheet: " & wsArchive.Name
        On Error GoTo ArchiveErrorHandler
    End If

    ' --- 5. Optional: Clear Comments from Archive ---
    ' On Error Resume Next: wsArchive.Cells.ClearComments: LogEvt "Archive", lvlDETAIL, "Cleared comments from archive." : On Error GoTo ArchiveErrorHandler

    ' --- 6. Optional: Protect the Archived Sheet ---
    ' On Error Resume Next
    ' wsArchive.Protect Password:="YourPassword", UserInterfaceOnly:=True ' Allow VBA to modify if needed later
    ' LogEvt "Archive", lvlDETAIL, "Protected archive sheet: " & wsArchive.Name
    ' On Error GoTo ArchiveErrorHandler

     ' LogEvt "Archive", lvlINFO, "Successfully archived data to sheet: " & wsArchive.Name
    Application.DisplayAlerts = True ' Restore alerts
    Exit Sub ' Successful Exit

ArchiveErrorHandler:
    ' --- Error Handling for Archiving ---
    Dim errDesc As String: errDesc = Err.Description: Dim errNum As Long: errNum = Err.Number
    Application.DisplayAlerts = True ' Restore alerts immediately
     ' LogEvt "Archive", lvlERROR, "Error during archiving for '" & archiveSheetName & "': " & errDesc & " (#" & errNum & ")"
    MsgBox "Error during archiving process for sheet '" & archiveSheetName & "': " & vbCrLf & errDesc, vbCritical, "Archive Error"
    ' Attempt to delete the partially created/failed archive sheet to avoid clutter
    If Not wsArchive Is Nothing Then
        If wsArchive.Name <> wsDataSource.Name Then ' Ensure we don't delete the original
            On Error Resume Next ' Ignore error if delete fails
            wsArchive.Delete
            On Error GoTo 0
             ' LogEvt "Archive", lvlWARN, "Attempted delete of partial archive sheet due to error."
        End If
    End If
End Sub


' ==========================================================================
' ===                   HELPER & UTILITY FUNCTIONS                     ===
' ==========================================================================

Private Function GetWorksheets(ByRef wsData As Worksheet, ByRef wsWeights As Worksheet, ByRef wsCache As Worksheet) As Boolean
    ' Safely gets required worksheet objects by name. Logs errors and returns False on failure.
    Dim success As Boolean: success = True
    Const PROC_NAME As String = "GetWorksheets"

    On Error Resume Next ' Check each sheet individually
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    If Err.Number <> 0 Then LogEvt PROC_NAME, lvlERROR, "Data sheet '" & DATA_SHEET_NAME & "' not found.": success = False: Err.Clear Else LogEvt PROC_NAME, lvlDEBUG, "Found wsData: " & wsData.Name

    Set wsWeights = ThisWorkbook.Sheets(WEIGHTS_SHEET_NAME)
    If Err.Number <> 0 Then LogEvt PROC_NAME, lvlERROR, "Weights sheet '" & WEIGHTS_SHEET_NAME & "' not found.": success = False: Err.Clear Else LogEvt PROC_NAME, lvlDEBUG, "Found wsWeights: " & wsWeights.Name

    Set wsCache = ThisWorkbook.Sheets(CACHE_SHEET_NAME)
    If Err.Number <> 0 Then LogEvt PROC_NAME, lvlERROR, "Cache sheet '" & CACHE_SHEET_NAME & "' not found.": success = False: Err.Clear Else LogEvt PROC_NAME, lvlDEBUG, "Found wsCache: " & wsCache.Name
    On Error GoTo 0 ' Restore default error handling

    If Not success Then
        MsgBox "Critical Error: One or more required worksheets could not be found." & vbCrLf & _
               "Ensure sheets named '" & DATA_SHEET_NAME & "', '" & WEIGHTS_SHEET_NAME & "', and '" & CACHE_SHEET_NAME & "' exist.", vbCritical, "Sheet Missing"
        ' Clean up partially assigned objects if error occurred
        Set wsData = Nothing: Set wsWeights = Nothing: Set wsCache = Nothing
    End If
    GetWorksheets = success
End Function

Public Function RefreshPowerQuery(connectionName As String) As Boolean ' Changed pattern to exact name
    ' Attempts synchronous Power Query refresh using exact connection name. Returns True on success.
    Dim conn As WorkbookConnection
    Dim startTime As Double: startTime = Timer
    Dim success As Boolean: success = False
    Const TIMEOUT_SECONDS As Long = 180 ' 3 minutes

    On Error GoTo RefreshErrorHandler
     ' LogEvt "Refresh", lvlINFO, "Attempting PQ refresh for connection: '" & connectionName & "'"

    ' Find the specific connection
    On Error Resume Next ' Check if connection exists
    Set conn = ThisWorkbook.Connections(connectionName)
    If Err.Number <> 0 Then
         ' LogEvt "Refresh", lvlERROR, "Connection '" & connectionName & "' not found."
        MsgBox "Error: Power Query connection named '" & connectionName & "' was not found." & vbCrLf & _
               "Please check the connection name in VBA constant 'PQ_CONNECTION_NAME'.", vbCritical, "Connection Not Found"
        GoTo RefreshExit ' Exit function, indicating failure
    End If
    On Error GoTo RefreshErrorHandler ' Restore handler

    ' Set to foreground refresh
    If Not conn.OLEDBConnection Is Nothing Then conn.OLEDBConnection.BackgroundQuery = False
    If Not conn.ODBCConnection Is Nothing Then conn.ODBCConnection.BackgroundQuery = False

     ' LogEvt "Refresh", lvlDETAIL, "Refreshing connection '" & conn.Name & "' synchronously..."
    conn.Refresh ' Synchronous call

    ' --- Simple check loop (removed complex timer) ---
    ' Power Query synchronous refresh should block VBA execution, but add a small safety loop.
    Dim waitLoops As Integer
    Do While conn.Refreshing And waitLoops < TIMEOUT_SECONDS * 2 ' Check twice per second approx
        DoEvents
        Application.Wait Now + TimeValue("00:00:01") ' Wait 1 second
        waitLoops = waitLoops + 1
    Loop

    If conn.Refreshing Then ' Still refreshing after timeout
         ' LogEvt "Refresh", lvlERROR, "Power Query refresh timed out after " & TIMEOUT_SECONDS & " seconds.", "Connection=" & connectionName
        MsgBox "Power Query refresh timed out after " & TIMEOUT_SECONDS & " seconds.", vbExclamation, "Refresh Timeout"
        ' Optionally attempt to cancel: conn.CancelRefresh
    Else
        success = True ' Refresh completed without timeout
         ' LogEvt "Refresh", lvlINFO, "Connection '" & connectionName & "' refresh completed."
    End If

RefreshExit:
    Set conn = Nothing
    RefreshPowerQuery = success
    Exit Function

RefreshErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
     ' LogEvt "Refresh", lvlERROR, "Error during PQ refresh for '" & connectionName & "': " & errDesc
    MsgBox "Error refreshing Power Query connection '" & connectionName & "': " & vbCrLf & errDesc, vbCritical, "Power Query Refresh Error"
    success = False
    Resume RefreshExit ' Go to cleanup
End Function

Public Function IsMaintainerUser() As Boolean ' Public for potential use elsewhere
    ' Checks if current user is the defined maintainer (case-insensitive).
    On Error Resume Next ' Handle error getting environment variable
    IsMaintainerUser = (LCase(Environ("USERNAME")) = LCase(MAINTAINER_USERNAME))
    If Err.Number <> 0 Then LogEvt "Util", lvlERROR, "Error checking MAINTAINER_USERNAME: " & Err.Description: IsMaintainerUser = False
    On Error GoTo 0
End Function

Private Function GetAPIKey() As String
    ' Reads API key from external file using %APPDATA%. Returns empty string on failure.
    Dim fso As Object, ts As Object, keyPath As String, WshShell As Object, fileContent As String: fileContent = ""
    On Error GoTo KeyError

    keyPath = API_KEY_FILE_PATH ' Get path from constant
    ' Expand environment variables like %APPDATA%
    Set WshShell = CreateObject("WScript.Shell")
    keyPath = WshShell.ExpandEnvironmentStrings(keyPath)
    Set WshShell = Nothing

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(keyPath) Then
        Set ts = fso.OpenTextFile(keyPath, 1) ' 1 = ForReading
        If Not ts.AtEndOfStream Then fileContent = ts.ReadAll ' Read entire file content
        ts.Close
         ' LogEvt "APIKey", lvlINFO, "API Key read successfully from: " & keyPath
    Else
         ' LogEvt "APIKey", lvlWARN, "API Key file not found at: " & keyPath
        Debug.Print Time & " - WARNING: API Key file not found at specified path: " & keyPath
    End If
    GoTo KeyExit ' Jump to cleanup

KeyError:
     ' LogEvt "APIKey", lvlERROR, "Error reading API Key from '" & keyPath & "': " & Err.Description
    Debug.Print Time & " - ERROR reading API Key from '" & keyPath & "': " & Err.Description
KeyExit:
    GetAPIKey = Trim(fileContent) ' Return trimmed content (or empty string)
    ' Cleanup objects
    If Not ts Is Nothing Then On Error Resume Next: ts.Close: Set ts = Nothing
    If Not fso Is Nothing Then Set fso = Nothing
    On Error GoTo 0
End Function

Private Function SheetExists(sheetName As String) As Boolean
    ' Case-insensitive check if a sheet exists in the current workbook.
    Dim ws As Worksheet
    On Error Resume Next ' Suppress error if sheet doesn't exist
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = (Err.Number = 0)
    On Error GoTo 0 ' Restore default error handling
    Set ws = Nothing ' Release object
End Function

' --- Safe Data Extraction Helpers ---
Private Function SafeGetString(arr As Variant, r As Long, ByVal cols As Object, colName As String) As String
    On Error Resume Next: SafeGetString = Trim(CStr(arr(r, cols(colName)))): If Err.Number <> 0 Then SafeGetString = "": Err.Clear
End Function
Private Function SafeGetVariant(arr As Variant, r As Long, ByVal cols As Object, colName As String) As Variant
    On Error Resume Next: SafeGetVariant = arr(r, cols(colName)): If Err.Number <> 0 Then SafeGetVariant = Null: Err.Clear
End Function

' --- JSON Helper Functions (Basic - Consider a robust library) ---
Private Function JsonEscape(strInput As String) As String
    ' Basic escaping for JSON strings within VBA strings
    strInput = Replace(strInput, "\", "\\") ' Escape backslashes FIRST
    strInput = Replace(strInput, """", "\""") ' Escape double quotes
    strInput = Replace(strInput, vbCrLf, "\n") ' Escape newlines
    strInput = Replace(strInput, vbCr, "\n")
    strInput = Replace(strInput, vbLf, "\n")
    strInput = Replace(strInput, vbTab, "\t") ' Escape tabs
    JsonEscape = strInput
End Function

Private Function JsonUnescape(strInput As String) As String
    ' Basic unescaping for JSON strings received from API
    strInput = Replace(strInput, "\""", """") ' Unescape double quotes
    strInput = Replace(strInput, "\\", "\") ' Unescape backslashes LAST
    strInput = Replace(strInput, "\n", vbCrLf) ' Unescape newlines
    strInput = Replace(strInput, "\t", vbTab) ' Unescape tabs
    JsonUnescape = strInput
End Function


' --- Placeholder for Logger Module Calls (If Used) ---
' Replace these with actual calls to your logging module if you implement one
Private Sub LogEvt(eventCode As String, eventLevel As Integer, eventDesc As String, Optional eventDetail As String = "")
    Const lvlDEBUG As Integer = 1
    Const lvlINFO As Integer = 2
    Const lvlWARN As Integer = 3
    Const lvlERROR As Integer = 4
    Const lvlFATAL As Integer = 5
    
    ' Simple Debug.Print implementation if no logger module exists
    Dim levelStr As String
    Select Case eventLevel
        Case lvlDEBUG: levelStr = "DEBUG"
        Case lvlINFO: levelStr = "INFO"
        Case lvlWARN: levelStr = "WARN"
        Case lvlERROR: levelStr = "ERROR"
        Case lvlFATAL: levelStr = "FATAL"
        Case Else: levelStr = "UNKNOWN"
    End Select
    Debug.Print Time & " [" & levelStr & "] " & eventCode & ": " & eventDesc & IIf(eventDetail <> "", " | " & eventDetail, "")
    
    ' --- Call your actual Logger sub here ---
    ' Example: Call mod_Logger.LogEvent(eventCode, eventLevel, eventDesc, eventDetail)
End Sub

Private Sub FlushLogBuf()
     ' Call your actual log flush sub here
     ' Example: Call mod_Logger.FlushBuffer
     Debug.Print Time & " [INFO] LogFlush: (Placeholder) Log buffer flushed."
End Sub

' ==========================================================================
' ===                        END OF MODULE                               ===
' ==========================================================================