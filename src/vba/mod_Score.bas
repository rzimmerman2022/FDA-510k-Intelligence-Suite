' =========  mod_Score.bas  =========
' Purpose: Handles the calculation of the 510(k) score for individual records,
'          including keyword checks and weight lookups.
' Key APIs exposed: Calculate510kScore
' Maintainer: [Your Name/Team]
' Dependencies: mod_Logger, mod_DebugTraceHelpers, mod_Config, mod_Weights, mod_Schema
' =====================================
Option Explicit

' --- Module-Level Object for Regular Expressions (Late Binding) ---
' Moved here as CheckKeywords uses it
Private regex As Object

' ==========================================================================
' ===                PUBLIC SCORING FUNCTION                         ===
' ==========================================================================
Public Function Calculate510kScore(dataArr As Variant, rowIdx As Long, ByVal cols As Object) As Variant
    ' Purpose: Calculates the 510(k) score based on various factors for a single record.
    ' Inputs:  dataArr - The 2D variant array holding all data.
    '          rowIdx - The current row number being processed in the array.
    '          cols - Dictionary mapping column names (including Name#Index for duplicates) to indices.
    ' Returns: A Variant array containing score components:
    '          Array(0=FinalScore, 1=Category, 2=AC_Wt, 3=PC_Wt, 4=KW_Wt, 5=ST_Wt,
    '                6=PT_Wt, 7=GL_Wt, 8=NF_Calc, 9=Synergy_Calc)
    ' Dependencies: Uses constants from mod_Config, safe getters from mod_Schema,
    '               and loaded data accessors from mod_Weights.

    ' --- Variable Declarations ---
    Dim AC As String, PC As String, DeviceName As String, Statement As String, SubmType As String, Country As String
    Dim ProcTimeDays As Variant, combinedText As String
    Dim AC_Wt As Double, PC_Wt As Double, KW_Wt As Double, ST_Wt As Double, PT_Wt As Double, GL_Wt As Double
    Dim NF_Calc As Double, Synergy_Calc As Double, Final_Score_Raw As Double
    Dim Category As String
    Dim HasHighValueKW As Boolean, IsCosmetic As Boolean, IsDiagnostic As Boolean, HasTherapeuticMention As Boolean
    Dim kw As Variant, kNum As String ' For loops and error logging
    Const PROC_NAME As String = "mod_Score.Calculate510kScore" ' Updated PROC_NAME

    ' --- Error Handling for this Function ---
    On Error GoTo ScoreErrorHandler

    ' --- 1. Extract Data Using Column Indices (Use SafeGetString/Variant from mod_Schema) ---
    AC = mod_Schema.SafeGetString(dataArr, rowIdx, cols, "AC")
    PC = mod_Schema.SafeGetString(dataArr, rowIdx, cols, "PC")
    DeviceName = mod_Schema.SafeGetString(dataArr, rowIdx, cols, "DeviceName")
    Statement = mod_Schema.SafeGetString(dataArr, rowIdx, cols, "Statement")
    SubmType = mod_Schema.SafeGetString(dataArr, rowIdx, cols, "SubmType")
    Country = UCase(mod_Schema.SafeGetString(dataArr, rowIdx, cols, "Country"))
    ProcTimeDays = mod_Schema.SafeGetVariant(dataArr, rowIdx, cols, "ProcTimeDays")
    combinedText = DeviceName & " " & Statement ' For keyword searching
    kNum = mod_Schema.SafeGetString(dataArr, rowIdx, cols, "K_Number") ' For logging context

    ' --- 2. Calculate Individual Weights (Use GetWeightFromDict helper below and constants from mod_Config) ---
    ' Access loaded weights via mod_Weights accessor functions
    AC_Wt = GetWeightFromDict(mod_Weights.GetACWeights(), AC, DEFAULT_AC_WEIGHT)
    PC_Wt = GetWeightFromDict(mod_Weights.GetPCWeights(), PC, DEFAULT_PC_WEIGHT)
    ST_Wt = GetWeightFromDict(mod_Weights.GetSTWeights(), SubmType, DEFAULT_ST_WEIGHT)

    ' Processing Time Weight
    If IsNumeric(ProcTimeDays) Then
        Select Case CDbl(ProcTimeDays)
            Case Is > 172: PT_Wt = 0.65
            Case 162 To 172: PT_Wt = 0.6
            Case Else: PT_Wt = DEFAULT_PT_WEIGHT ' Includes < 162 and non-positive/invalid
        End Select
    Else: PT_Wt = DEFAULT_PT_WEIGHT ' Default if ProcTimeDays is not numeric
    End If

    ' Geographic Location Weight
    If Country = "US" Then GL_Wt = US_GL_WEIGHT Else GL_Wt = OTHER_GL_WEIGHT

    ' Keyword Weight (using CheckKeywords helper below and collections from mod_Weights)
    HasHighValueKW = CheckKeywords(combinedText, mod_Weights.GetHighValueKeywords())
    If HasHighValueKW Then KW_Wt = HIGH_KW_WEIGHT Else KW_Wt = LOW_KW_WEIGHT

    ' --- 3. Negative Factors (NF) & Synergy Logic ---
    NF_Calc = 0: Synergy_Calc = 0
    IsCosmetic = CheckKeywords(combinedText, mod_Weights.GetNFCosmeticKeywords())
    IsDiagnostic = CheckKeywords(combinedText, mod_Weights.GetNFDiagnosticKeywords())
    HasTherapeuticMention = CheckKeywords(combinedText, mod_Weights.GetTherapeuticKeywords())

    ' Apply Negative Factors (Ensure Therapeutic overrides NF)
    If IsCosmetic And Not HasTherapeuticMention Then NF_Calc = NF_COSMETIC
    If IsDiagnostic And Not HasTherapeuticMention Then
        ' Additive logic: If both Cosmetic and Diagnostic (and not Therapeutic) apply both NFs
        If NF_Calc = 0 Then NF_Calc = NF_DIAGNOSTIC Else NF_Calc = NF_Calc + NF_DIAGNOSTIC
    End If

    ' Apply Synergy Bonus
    If (AC = "OR" Or AC = "NE") And HasHighValueKW Then Synergy_Calc = SYNERGY_BONUS

    ' --- 4. Final Score Calculation ---
    ' Ensure divisor matches the number of components being summed (adjust if logic changes)
    Final_Score_Raw = (AC_Wt + PC_Wt + KW_Wt + ST_Wt + PT_Wt + GL_Wt + NF_Calc + Synergy_Calc) / 6
    If Final_Score_Raw < 0 Then Final_Score_Raw = 0 ' Floor score at 0

    ' --- 5. Determine Category ---
    Select Case Final_Score_Raw
        Case Is > 0.6: Category = "High"
        Case 0.5 To 0.6: Category = "Moderate"
        Case 0.4 To 0.499999999999: Category = "Low" ' Explicit upper bound for Low
        Case Else: Category = "Almost None" ' Includes scores < 0.4 and exactly 0
    End Select

    ' --- 6. Return Results ---
    Calculate510kScore = Array(Final_Score_Raw, Category, AC_Wt, PC_Wt, KW_Wt, ST_Wt, PT_Wt, GL_Wt, NF_Calc, Synergy_Calc)
    Exit Function

ScoreErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    LogEvt PROC_NAME, lgERROR, "Error scoring row " & rowIdx & " (K#: " & kNum & "): " & errDesc, "AC=" & AC & ", PC=" & PC ' Use lgERROR
    TraceEvt lvlERROR, PROC_NAME, "Error scoring row", "Row=" & rowIdx & ", K#=" & kNum & ", Err=" & Err.Number & " - " & errDesc
    Debug.Print Time & " - ERROR scoring row " & rowIdx & " (K#: " & kNum & "): " & errDesc
    Calculate510kScore = Array(0, "Error", 0, 0, 0, 0, 0, 0, 0, 0) ' Return error state array
End Function

' ==========================================================================
' ===                PRIVATE HELPER FUNCTIONS                        ===
' ==========================================================================

' --- CheckKeywords (Using RegExp) ---
Private Function CheckKeywords(textToCheck As String, keywordColl As Collection) As Boolean
    Dim kw As Variant
    CheckKeywords = False
    If keywordColl Is Nothing Or keywordColl.Count = 0 Or Len(Trim(textToCheck)) = 0 Then Exit Function
    Const PROC_NAME As String = "mod_Score.CheckKeywords" ' Updated PROC_NAME

    ' --- Initialize RegExp object (Late Binding) ---
    If regex Is Nothing Then Set regex = CreateObject("VBScript.RegExp")

    ' --- Build pattern and test ---
    On Error GoTo CheckKeywordsErrorHandler ' Handle errors during RegExp or pattern building

    ' Build the pattern: (keyword1|keyword2|keyword3)
    ' Need to escape any special regex characters within keywords if they exist
    Dim patternBuilder As Object: Set patternBuilder = CreateObject("System.Collections.ArrayList") ' Use ArrayList for dynamic add
    For Each kw In keywordColl
        ' Basic escaping for common characters, might need more robust escaping if keywords are complex
        Dim escapedKw As String: escapedKw = CStr(kw)
        escapedKw = Replace(escapedKw, "\", "\\")
        escapedKw = Replace(escapedKw, ".", "\.")
        escapedKw = Replace(escapedKw, "|", "\|")
        escapedKw = Replace(escapedKw, "(", "\(")
        escapedKw = Replace(escapedKw, ")", "\)")
        escapedKw = Replace(escapedKw, "[", "\[")
        escapedKw = Replace(escapedKw, "]", "\]")
        escapedKw = Replace(escapedKw, "*", "\*")
        escapedKw = Replace(escapedKw, "+", "\+")
        escapedKw = Replace(escapedKw, "?", "\?")
        escapedKw = Replace(escapedKw, "{", "\{")
        escapedKw = Replace(escapedKw, "}", "\}")
        escapedKw = Replace(escapedKw, "^", "\^")
        escapedKw = Replace(escapedKw, "$", "\$")
        patternBuilder.Add escapedKw
    Next kw

    If patternBuilder.Count = 0 Then GoTo CheckKeywordsExit ' No valid keywords to build pattern

    regex.Pattern = Join(patternBuilder.ToArray(), "|") ' Join keywords with OR operator
    regex.IgnoreCase = True ' Case-insensitive match
    regex.Global = False    ' Only need to find one match

    ' Test the input string against the pattern
    CheckKeywords = regex.Test(textToCheck)

CheckKeywordsExit:
    Set patternBuilder = Nothing
    Exit Function

CheckKeywordsErrorHandler:
    LogEvt PROC_NAME, lgERROR, "Error during RegExp keyword check: " & Err.Description ' Use lgERROR
    TraceEvt lvlERROR, PROC_NAME, "RegExp Error", "Err=" & Err.Number & " - " & Err.Description
    Debug.Print Time & " - ERROR in CheckKeywords RegExp: " & Err.Description
    CheckKeywords = False ' Return False on error
    Resume CheckKeywordsExit ' Go to cleanup
End Function

' --- GetWeightFromDict (Helper) ---
Private Function GetWeightFromDict(dict As Object, key As String, defaultWeight As Double) As Double
    ' Purpose: Safely retrieves a weight (Double) from a dictionary, using default if key not found or value is invalid.
    Const PROC_NAME As String = "mod_Score.GetWeightFromDict" ' Updated PROC_NAME
    If dict Is Nothing Then GetWeightFromDict = defaultWeight: Exit Function ' Handle Nothing dictionary object

    Dim value As Variant
    On Error Resume Next ' Suppress errors during dictionary access/conversion

    If dict.Exists(key) Then
        value = dict(key)
        If IsNumeric(value) Then
            GetWeightFromDict = CDbl(value) ' Convert valid numeric value to Double
            If Err.Number <> 0 Then GetWeightFromDict = defaultWeight: Err.Clear ' Use default if CDbl fails (overflow?)
        Else
            GetWeightFromDict = defaultWeight ' Value exists but is not numeric
        End If
    Else
        GetWeightFromDict = defaultWeight ' Key does not exist
    End If

    On Error GoTo 0 ' Restore default error handling
End Function
