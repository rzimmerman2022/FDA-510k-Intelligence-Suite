# Root-Cause Scan – May 8, 2025

## 🟥 Critical

### src/vba/mod_Archive.bas:158-159 – PasteSpecial could strip comments and formatting

```vb
    tblData.Range.Copy
    wsArchive.Range("A1").PasteSpecial xlPasteValues
    wsArchive.Range("A1").PasteSpecial xlPasteFormats
```

While there is an attempt to preserve formats with a second PasteSpecial, this implementation could still strip comments and conditional formatting rules. This is the only active (non-legacy) instance of value-pasting that could affect data integrity.

### src/vba/ModuleManager.bas – Multiple unhandled On Error Resume Next blocks

```vb
    On Error Resume Next
    MakeNestedFolders EXPORT_FOLDER

    If Err.Number <> 0 Then
        Debug.Print "ERROR creating folders: " & Err.Description
        MsgBox "Error creating export folders: " & Err.Description, vbCritical
        Exit Sub
    End If
```

The ModuleManager module contains at least 12 instances of On Error Resume Next, with several missing proper error checks. Critical examples include:

```vb
    On Error Resume Next
    fso.CopyFile fileItem.Path, SOURCE_FOLDER & fileItem.Name, True ' True = Overwrite
    ' No Err.Number check after file copy operation
```

```vb
    On Error Resume Next ' Handle potential import errors
    Set newComp = ThisWorkbook.VBProject.VBComponents.Import(fullPath)
    ' No direct Err.Number check after crucial import operation
```

These operations could silently fail without proper error handling, potentially leading to corrupted module states.

## 🟧 Warnings

### src/vba/mod_TestRefresh.bas:53 – On Error Resume Next without complete error handling block

```vb
    On Error Resume Next ' Isolate error specifically on the refresh line

    ' Ensure BackgroundQuery is False (Important!)
    wbConn.OLEDBConnection.BackgroundQuery = False ' Set before refreshing
    If Err.Number <> 0 Then
         Debug.Print "TestRefreshOnly: Error setting BackgroundQuery=False for '" & wbConn.Name & "'. Err: " & Err.Description
         Err.Clear ' Clear error and try refresh anyway
    End If
```

This code checks the error after setting BackgroundQuery, but doesn't properly handle the error state for the subsequent Refresh operation in the same error-handling block.

### src/vba/mod_Format.bas – Multiple unguarded column operations that could fail silently

```vb
    On Error Resume Next ' Ignore if column doesn't exist (should have been added)
    tbl.ListColumns(colName).DataBodyRange.NumberFormat = "0.0"
```

The mod_Format.bas module contains 17 instances of On Error Resume Next. While many are intentional to handle missing columns, they lack error logging which could make debugging difficult. Critical examples include:

```vb
    On Error Resume Next
    tbl.ListColumns("Score_Percent").DataBodyRange.NumberFormat = "0.0%"
    ' No verification if operation succeeded
```

```vb
    On Error Resume Next ' Check if column exists
    Set catCol = tblData.ListColumns("Category")
    ' No check if catCol Is Nothing after attempting to set it
```

Several operations apply formatting without verifying the success of the operation, which could lead to inconsistent visual presentation.

## 🟩 Info / Observations

### ModuleManager.bas contains Option Explicit

Contrary to what initial scanning suggested, ModuleManager.bas does include Option Explicit at the top of the module, which helps prevent undeclared variable bugs.

### No Table.RemoveColumns found in Power Query files

The Power Query file (FDA_510k_Query.pq) does not contain any instances of Table.RemoveColumns, which is good as this operation could potentially delete fields that VBA code might later expect.

### Extensive use of On Error Resume Next throughout codebase

The codebase contains 247 instances of On Error Resume Next. While many are followed by appropriate error handling, the sheer number suggests a maintenance risk that could lead to silent failures if code is changed without understanding error handling flows.

## Next Steps Checklist

1. ✅ **Enhance PasteSpecial in mod_Archive.bas** - Consider using a multi-step approach to preserve all formatting elements and comments, such as:
   - First copying the entire range structure
   - Then applying values separately
   - Consider calling a comprehensive ApplyAll function after paste operations

2. ✅ **Review error handling in ModuleManager.bas** - Add proper Err.Number checks after all On Error Resume Next blocks:
   - Focus on file operations (CopyFile)
   - Add error checking after component Import/Export operations
   - Consider adding a centralized error handling pattern for file operations

3. ✅ **Implement consistent error handling pattern** - Standardize the approach to error handling across modules:
   - Always check Err.Number after On Error Resume Next when critical operations are performed
   - Either use Try/Trap pattern (On Error Resume Next + If Err check) or use On Error GoTo consistently
   - Add debug logging even for non-critical errors

4. ✅ **Add debug logging for silent failures** - For operations where On Error Resume Next is necessary but failures are acceptable, add debug logging so issues don't go completely unnoticed:
   - In mod_Format.bas, add If Err.Number <> 0 Then LogEvt... blocks after formatting operations
   - Consider a helper function like TryFormat() that encapsulates the error handling pattern
