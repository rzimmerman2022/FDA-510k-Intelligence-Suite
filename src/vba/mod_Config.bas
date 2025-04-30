' =========  mod_Config.bas  =========
' Purpose: Holds all global configuration constants for the 510k Intelligence Suite.
' Key APIs exposed: Public constants for sheet names, paths, scoring defaults, API settings, etc.
' Maintainer: [Your Name/Team]
' =====================================
Option Explicit
Attribute VB_Name = "mod_Config"

' --- Essential Configuration ---
' *** IMPORTANT: SET YOUR WINDOWS USERNAME FOR MAINTAINER FEATURES (e.g., OpenAI, DebugMode) ***
Public Const MAINTAINER_USERNAME As String = "YourWindowsUsername" ' <<< UPDATE THIS

' *** Double-check these names match your Excel objects ***
Public Const DATA_SHEET_NAME As String = "CurrentMonthData"  ' Sheet where Power Query loads data
Public Const WEIGHTS_SHEET_NAME As String = "Weights"        ' Sheet containing weight/keyword tables
Public Const CACHE_SHEET_NAME As String = "CompanyCache"      ' Sheet for persistent company recap cache
Public Const LOG_SHEET_NAME As String = "RunLog"             ' Sheet for logging events
Public Const SHORT_NAME_MAX_LEN As Long = 75 ' Maximum length for shortened device names
Public Const SHORT_NAME_ELLIPSIS As String = "..." ' Text to append to shortened names

' *** Path to file containing ONLY your OpenAI API Key ***
' *** Uses %APPDATA% environment variable for user-specific location ***
Public Const API_KEY_FILE_PATH As String = "%APPDATA%\510k_Tool\openai_key.txt" ' <<< ENSURE THIS PATH IS CORRECT & FILE EXISTS

' --- Scoring Defaults & Parameters (Used if lookup fails or as base values) ---
' *** REVIEW AND CONFIRM THESE VALUES BASED ON YOUR SCORING MODEL ***
Public Const DEFAULT_AC_WEIGHT As Double = 0.2
Public Const DEFAULT_PC_WEIGHT As Double = 0.2
Public Const DEFAULT_ST_WEIGHT As Double = 0.6 ' Default to Traditional if SubmType not found
Public Const DEFAULT_PT_WEIGHT As Double = 0.5 ' Default if ProcTimeDays is invalid or <162
Public Const HIGH_KW_WEIGHT As Double = 0.85
Public Const LOW_KW_WEIGHT As Double = 0.2
Public Const US_GL_WEIGHT As Double = 0.6
Public Const OTHER_GL_WEIGHT As Double = 0.5
Public Const NF_COSMETIC As Double = -2#  ' Negative Factor for purely cosmetic devices (CONFIRM VALUE)
Public Const NF_DIAGNOSTIC As Double = -0.2 ' Negative Factor for purely diagnostic software (CONFIRM VALUE)
Public Const SYNERGY_BONUS As Double = 0.15 ' Bonus for specific AC + High KW match (CONFIRM VALUE/LOGIC)

' --- OpenAI Configuration (Optional) ---
Public Const OPENAI_API_URL As String = "https://api.openai.com/v1/chat/completions"
Public Const OPENAI_MODEL As String = "gpt-3.5-turbo" ' Or "gpt-4o-mini" etc. - check pricing/availability
Public Const OPENAI_MAX_TOKENS As Long = 100 ' Limit response length
Public Const OPENAI_TIMEOUT_MS As Long = 60000 ' 60 seconds timeout for API call

' --- UI & Formatting ---
Public Const VERSION_INFO As String = "v1.9 - Split Code Gen" ' Simple version tracking (Public for Logger)
Public Const RECAP_MAX_LEN As Long = 32760 ' Max characters for cell / recap text to avoid overflow
Public Const DEFAULT_RECAP_TEXT = "Needs Research" ' Default text when recap is missing
