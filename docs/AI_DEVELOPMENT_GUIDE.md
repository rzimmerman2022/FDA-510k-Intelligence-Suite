# AI Development Guide - FDA 510(k) Intelligence Suite

This document outlines best practices and specific requirements for developing and maintaining the VBA code within this project, particularly when utilizing AI assistance (like Cline). Adhering to these guidelines ensures code quality, maintainability, and consistency.

## 1. Core Principles

*   **Modularity:** Respect the existing modular structure. Place new functionality in the most appropriate existing module or create a new, well-defined module if necessary.
*   **Clarity:** Write clear, readable code. Use meaningful variable and procedure names.
*   **Robustness:** Implement appropriate error handling (`On Error GoTo`, logging) for key operations. Use safe data access methods (e.g., `mod_Schema.SafeGetString`).
*   **Configuration:** Centralize configuration settings (sheet names, defaults, API keys/paths) in `mod_Config` using `Public Const`. Avoid hardcoding values directly in procedural code.
*   **Documentation:** Maintain high-quality documentation through standardized header comments and meaningful inline comments.

## 2. Coding Standards (VBA)

*   **`Option Explicit`:** Mandatory at the top of every module to enforce variable declaration.
*   **Variable Declaration:** Declare all variables using `Dim` (or `Private`/`Public` for module-level scope) with the most specific data type possible (e.g., `Dim ws As Worksheet`, `Dim count As Long`, `Dim name As String`). Avoid excessive use of `Variant` unless necessary (e.g., iterating arrays, handling diverse return types).
*   **Naming Conventions:**
    *   Procedures (Subs/Functions): `PascalCase` (e.g., `ProcessMonthly510k`, `LoadTableToDict`).
    *   Local Variables: `camelCase` (e.g., `procStartTime`, `missingSheets`).
    *   Constants: `ALL_CAPS_SNAKE_CASE` (e.g., `DATA_SHEET_NAME`, `API_KEY_FILE_PATH`).
    *   Module-Level Private Variables: Use a prefix like `m_` or suffix (less common in VBA) if needed for clarity, otherwise `camelCase` is often sufficient if usage is clear (e.g., `dictCache` in `mod_Cache`).
    *   Enums: Use a prefix for members (e.g., `lgINFO` for `LogLevel`, `lvlERROR` for `eTraceLvl`).
*   **Indentation:** Use consistent indentation (e.g., 4 spaces) to clearly show code blocks (loops, conditionals, `With` blocks).
*   **Scope:** Declare variables in the narrowest possible scope. Prefer local variables within procedures over module-level variables unless state needs to be maintained across calls.
*   **Argument Passing:** Prefer passing arguments `ByVal` unless the procedure explicitly needs to modify the original variable passed by the caller (then use `ByRef`). Be explicit (`ByVal` or `ByRef`) rather than relying on the default (`ByRef`).
*   **Error Handling:** Use `On Error GoTo <Label>` for structured error handling in significant procedures. Log errors using `mod_Logger.LogEvt` and/or `mod_DebugTraceHelpers.TraceEvt`. Use `On Error Resume Next` sparingly and only when you intend to check `Err.Number` immediately afterward or when failure of a specific line is non-critical and expected. Always restore default error handling with `On Error GoTo 0`.
*   **Object Variables:** Always set object variables to `Nothing` when finished to release memory (e.g., `Set ws = Nothing`).

## 3. Commenting Requirements

### 3.1. Module Header Comment Block

**This is mandatory for every `.bas` and `.cls` module.** It provides essential context for anyone (human or AI) reading the code.

**Structure:**

```vba
' ==========================================================================
' Module      : [Module Name (e.g., mod_DataIO, ThisWorkbook)]
' Author      : [Original Author - Unknown, unless known]
' Date        : [Original Date - Unknown, unless known]
' Maintainer  : [Your Name (e.g., Cline (AI Assistant))]
' Version     : [Link to version info, e.g., See mod_Config.VERSION_INFO or N/A]
' ==========================================================================
' Description : [Concise paragraph explaining the module's overall purpose,
'               responsibilities, and role within the application architecture.]
'
' Key Functions/Procedures:
'               - [Public Function/Sub Name 1]: [Brief description of its purpose.]
'               - [Public Function/Sub Name 2]: [Brief description...]
'               - ... (List primary public interfaces)
'
' Private Helpers: (Optional section if notable private functions exist)
'               - [Private Function/Sub Name 1]: [Brief description...]
'
' Dependencies: - [List other custom modules this module relies on (e.g., mod_Logger, mod_Config).]
'               - [List required external libraries/references (e.g., Scripting.Dictionary, MSXML2.ServerXMLHTTP).]
'               - [Mention key assumptions (e.g., specific table names, sheet names).]
'
' Revision History:
' --------------------------------------------------------------------------
' Date        Author          Description
' ----------- --------------- ----------------------------------------------
' [YYYY-MM-DD]  [Your Name]     - [Clear, concise description of the change made.]
' ... (Previous entries maintained chronologically)
' ==========================================================================
Option Explicit
Attribute VB_Name = "[Module Name]" ' If applicable
```

**Maintenance:**

*   **Read First:** Before modifying *any* code in a module, read its header comment thoroughly.
*   **Update Revision History:** **Crucially, every time you make a functional change or fix within a module, add a new entry to the `Revision History` table.** Include the date (YYYY-MM-DD), your identifier ("Cline (AI)" or subsequent AI name), and a brief but clear description of the change. Keep entries chronological (newest usually added at the top or bottom depending on preference, but be consistent).
*   **Update Other Sections:** If your changes alter the module's purpose, add/remove key functions, change dependencies, or modify assumptions, update the relevant sections (`Description`, `Key Functions`, `Dependencies`) accordingly.

### 3.2. Inline Comments

*   Use the single quote (`'`) for inline comments.
*   Focus on explaining the **why**, not the **what**. Explain complex logic, assumptions, workarounds, or the reasoning behind a particular implementation choice.
    *   *Bad:* `' Increment i` (Code `i = i + 1` is self-explanatory)
    *   *Good:* `' Use late binding to avoid dependency on specific library version`
    *   *Good:* `' Handle potential error if sheet doesn't exist before attempting delete`
*   Keep comments concise and close to the code they describe.
*   Update or remove comments when the code they refer to changes. Stale comments are misleading.

## 4. AI Assistant Workflow

1.  **Understand Task:** Clearly define the goal or problem to be solved.
2.  **Analyze Context:** Review relevant module header comments, inline comments, and related code provided by the user or previous steps.
3.  **Plan Changes:** Outline the necessary code modifications. Identify which modules/functions will be affected.
4.  **Implement Changes:** Use the appropriate tools (`replace_in_file` preferred for targeted edits, `write_to_file` for new files or major rewrites).
5.  **Update Documentation:**
    *   Add a new entry to the `Revision History` in the header of **each modified module**.
    *   Update other header sections (`Description`, `Dependencies`, etc.) if necessary.
    *   Add/update inline comments for complex or non-obvious changes.
6.  **Verify (Simulated):** Mentally review the changes against the requirements and coding standards.
7.  **Present Result:** Use `attempt_completion` to report the changes made, including the documentation updates.

By following these guidelines, AI assistants can contribute effectively to the project while maintaining code quality and documentation standards.
