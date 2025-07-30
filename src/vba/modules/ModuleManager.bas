' ==========================================================================
' Module      : ModuleManager
' Author      : [Original Author - Unknown]
' Date        : [Original Date - Unknown]
' Maintainer  : Cline (AI Assistant)
' Version     : N/A (Utility Module)
' ==========================================================================
' Description : This module provides utility functions for managing the VBA
'               code modules within the project itself. It facilitates exporting
'               modules from the VBE to external files (for version control like Git)
'               and importing them back into the VBE from those files. It includes
'               logic to handle potential auto-renaming issues by the VBE during
'               import and allows skipping specific folders (e.g., "_legacy").
'               It also provides a function to copy exported files back to the
'               source directory after review.
'
'               NOTE: This module relies on hardcoded folder paths (SOURCE_FOLDER,
'               EXPORT_FOLDER) which may need adjustment depending on the
'               development environment setup. It also requires a reference to
'               'Microsoft Scripting Runtime' for Dictionary and FileSystemObject.
'
' Key Functions:
'               - ExportAllModules: Exports .bas, .cls, .frm files to EXPORT_FOLDER.
'               - ImportAllModules: Removes existing modules and imports from
'                 SOURCE_FOLDER, attempting to preserve original VB_Names.
'               - UpdateSourceFromExport: Copies files from EXPORT_FOLDER to
'                 SOURCE_FOLDER (overwrites).
'
' Private Helpers:
'               - MakeNestedFolders: Creates directory structures.
'               - GatherFiles: Recursively finds module files, extracts VB_Name.
'               - ReadText: Reads file content with specified encoding.
'               - ParseVBName: Extracts the VB_Name attribute from file content.
'
' Dependencies: - Requires reference to 'Microsoft Scripting Runtime'.
'               - Requires reference to 'Microsoft Visual Basic for Applications Extensibility 5.3'.
'               - Relies on hardcoded SOURCE_FOLDER and EXPORT_FOLDER paths.
'
' Revision History:
' --------------------------------------------------------------------------
' Date        Author          Description
' ----------- --------------- ----------------------------------------------
' 2025-04-30  Cline (AI)      - Added detailed module header comment block.
' [Previous dates/authors/changes unknown]
' ==========================================================================
Option Explicit
Attribute VB_Name = "ModuleManager"

' Updated paths to match your actual folder structure
Private Const SOURCE_FOLDER As String = "C:\FDA\FDA-510k-Intelligence-Suite\src\vba\"
Private Const EXPORT_FOLDER As String = "C:\FDA\FDA-510k-Intelligence-Suite\src\vba-export\"

' Exports all modules to a separate folder for safety
Public Sub ExportAllModules()
    Dim component As VBIDE.VBComponent
    Dim fileName As String
    Dim extension As String
    Dim exportCount As Integer

    exportCount = 0

    ' Clear the immediate window for better debugging visibility
    Debug.Print String(50, "-")
    Debug.Print "EXPORT DIAGNOSTIC: " & Now()
    Debug.Print "Current directory: " & CurDir()
    Debug.Print "SOURCE_FOLDER = " & SOURCE_FOLDER
    Debug.Print "EXPORT_FOLDER = " & EXPORT_FOLDER

    ' Create full folder structure
    On Error Resume Next
    MakeNestedFolders EXPORT_FOLDER

    If Err.Number <> 0 Then
        Debug.Print "ERROR creating folders: " & Err.Description
        MsgBox "Error creating export folders: " & Err.Description, vbCritical
        Exit Sub
    End If

    ' Double-check folder exists
    If Dir(EXPORT_FOLDER, vbDirectory) = "" Then
        Debug.Print "ERROR: Export folder wasn't created: " & EXPORT_FOLDER
        MsgBox "Failed to create export folder!", vbCritical
        Exit Sub
    Else
        Debug.Print "Export folder exists: " & EXPORT_FOLDER
    End If
    On Error GoTo 0

    ' Loop through all components in the project
    For Each component In ThisWorkbook.VBProject.VBComponents
        ' Determine the file extension based on component type
        Select Case component.Type
            Case vbext_ct_ClassModule
                extension = ".cls"
            Case vbext_ct_MSForm
                extension = ".frm"
            Case vbext_ct_StdModule
                extension = ".bas"
            Case vbext_ct_Document
                ' Skip document modules
                GoTo NextComponent
            Case Else
                ' Skip other types
                GoTo NextComponent
        End Select

        ' Construct the full file path
        fileName = EXPORT_FOLDER & component.Name & extension

        ' Export with error handling
        On Error Resume Next
        component.Export fileName

        If Err.Number <> 0 Then
            Debug.Print "ERROR exporting " & component.Name & ": " & Err.Description
            Err.Clear
        Else
            exportCount = exportCount + 1
            Debug.Print "Exported: " & fileName
        End If
        On Error GoTo 0

NextComponent:
    Next component

    Debug.Print "Export completed. " & exportCount & " modules exported."
    Debug.Print String(50, "-")

    MsgBox "All modules exported to: " & EXPORT_FOLDER & vbCrLf & _
           "Review changes before copying to source folder." & vbCrLf & _
           "Modules exported: " & exportCount, vbInformation
End Sub

' Creates nested folders, handling the full path
Private Sub MakeNestedFolders(ByVal folderPath As String)
    Dim pathParts As Variant
    Dim currentPath As String
    Dim i As Integer

    ' Remove trailing backslash if present
    If Right(folderPath, 1) = "\" Then
        folderPath = Left(folderPath, Len(folderPath) - 1)
    End If

    ' Split the path into parts
    pathParts = Split(folderPath, "\")

    ' Start with the drive root
    currentPath = pathParts(0) & "\"

    ' Create each folder level
    For i = 1 To UBound(pathParts)
        currentPath = currentPath & pathParts(i) & "\"

        ' Create directory if it doesn't exist
        If Dir(currentPath, vbDirectory) = "" Then
            Debug.Print "Creating folder: " & currentPath
            MkDir currentPath
        End If
    Next i
End Sub

'===========================================================
'  IMPORT *with* auto-rename and _legacy skip
'===========================================================
' Requires reference to 'Microsoft Scripting Runtime' for Dictionary object
Public Sub ImportAllModules()
    Dim fso As Object:        Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dict As Object:       Set dict = CreateObject("Scripting.Dictionary") ' Stores VB_Name -> FullPath
    Dim importCount As Long
    Dim key As Variant, newComp As VBIDE.VBComponent ' Declare here

    '-------------------------------------------------------
    ' 1) Gather files (recursively), skipping _legacy, using internal VB_Name as key
    '-------------------------------------------------------
    GatherFiles SOURCE_FOLDER, dict, fso ' Pass dict and fso

    If dict.Count = 0 Then
        MsgBox "No valid modules found to import in " & SOURCE_FOLDER & " (or subfolders, excluding _legacy).", vbExclamation
        GoTo Cleanup ' Ensure cleanup happens
    End If
    Debug.Print "Found " & dict.Count & " modules to potentially import."

    '-------------------------------------------------------
    ' 2) Remove existing components that match gathered VB_Names
    '    (Iterate backwards for safe removal)
    '-------------------------------------------------------
    Dim comp As VBIDE.VBComponent
    Dim i As Integer
    Debug.Print "Starting removal check..."
    For i = ThisWorkbook.VBProject.VBComponents.Count To 1 Step -1
        Set comp = ThisWorkbook.VBProject.VBComponents(i)
        ' Check if the component's actual name exists in our dictionary of VB_Names to import
        If dict.Exists(comp.Name) Then
            ' Only remove standard/class/form modules, skip document modules etc.
            Select Case comp.Type
                Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm
                    Debug.Print "Attempting to remove existing component: " & comp.Name
                    On Error Resume Next
                    ThisWorkbook.VBProject.VBComponents.Remove comp
                    If Err.Number <> 0 Then
                        Debug.Print "ERROR removing existing component '" & comp.Name & "': " & Err.Description
                        Err.Clear
                    Else
                        Debug.Print "Successfully removed existing component: " & comp.Name
                    End If
                    On Error GoTo 0
                Case Else
                    Debug.Print "Skipping removal check for non-module component: " & comp.Name & " (Type: " & comp.Type & ")"
            End Select
        End If
    Next i
    Debug.Print "Removal check finished."

    '-------------------------------------------------------
    ' 3) Import and force-rename if VBE auto-renamed
    '-------------------------------------------------------
    importCount = 0 ' Reset import count
    Debug.Print "Starting import process..."
    For Each key In dict.Keys ' key is the correct VB_Name
        Dim fullPath As String
        fullPath = dict(key)
        Debug.Print "Attempting to import: " & key & " from " & fullPath

        On Error Resume Next ' Handle potential import errors
        Set newComp = ThisWorkbook.VBProject.VBComponents.Import(fullPath)
        If Err.Number <> 0 Then
            Debug.Print "ERROR importing file '" & fso.GetFileName(fullPath) & "' (expected VB_Name '" & key & "'): " & Err.Description
            Err.Clear
            GoTo NextImport ' Skip rename if import failed
        End If
        On Error GoTo 0 ' Restore default error handling

        ' Force-rename only if VBE auto-renamed it AND it's not a Document module
        If Not newComp Is Nothing Then
            If newComp.Type <> vbext_ct_Document And newComp.Name <> key Then
                Debug.Print "VBE auto-renamed to '" & newComp.Name & "'. Forcing rename to correct VB_Name: '" & key & "'"
                On Error Resume Next ' Handle potential rename errors (e.g., name still taken)
                newComp.Name = key
                If Err.Number <> 0 Then
                     Debug.Print "ERROR force-renaming component '" & newComp.Name & "' to '" & key & "': " & Err.Description
                     Err.Clear
                Else
                     Debug.Print "Successfully force-renamed component to: " & key
                End If
                On Error GoTo 0
            Else
                Debug.Print "Component imported/kept correct name: " & newComp.Name
            End If
            importCount = importCount + 1
        Else
            Debug.Print "ERROR: Import returned Nothing for file: " & fullPath
        End If

NextImport:
        Set newComp = Nothing ' Clear for next loop
    Next key
    Debug.Print "Import process finished."

Cleanup:
    ' Cleanup Scripting objects
    Set fso = Nothing
    Set dict = Nothing
    Set comp = Nothing
    ' newComp is cleared in the loop

    MsgBox "Import completed." & vbCrLf & "Processed " & importCount & " modules from: " & SOURCE_FOLDER, vbInformation
End Sub

'-----------------------------------------------------------
' Populate dict(VB_Name or FileName) = fullPath, skipping _legacy folders
' Tries default encoding, then Unicode, then falls back to filename.
'-----------------------------------------------------------
Private Sub GatherFiles(startPath As String, ByRef d As Object, ByRef fso As Object)
    ' --- 1. Skip if path contains _legacy or folder name starts with _legacy ---
    If InStr(1, startPath, "\_legacy\", vbTextCompare) > 0 _
       Or LCase$(fso.GetFileName(startPath)) Like "_legacy*" Then
        Debug.Print "GatherFiles: SKIP folder (legacy path or name): " & startPath
        Exit Sub
    End If

    ' --- Check if folder exists before proceeding ---
    If Not fso.FolderExists(startPath) Then
        Debug.Print "GatherFiles: ERROR - Folder not found: " & startPath
        Exit Sub
    End If

    Debug.Print "GatherFiles: Scanning folder: " & startPath

    Dim fil As Object, txt As String, vbName As String
    Dim fld As Object ' Moved folder declaration up

    ' --- 2. Scan files in the current folder ---
    On Error Resume Next ' Handle errors enumerating files
    For Each fil In fso.GetFolder(startPath).Files
        If Err.Number <> 0 Then
            Debug.Print "GatherFiles: ERROR enumerating files in " & startPath & ": " & Err.Description
            Err.Clear
            Exit For ' Stop trying to process files in this folder if enumeration fails
        End If

        Select Case LCase$(fso.GetExtensionName(fil.Name))
            Case "bas", "cls", "frm"
                vbName = "" ' Reset for each file

                '--- Try default encoding (ANSI / UTF-8) first ---
                On Error Resume Next ' Handle potential read errors
                txt = ReadText(fil.Path, -2) ' TriStateUseDefault = -2
                If Err.Number = 0 Then
                    vbName = ParseVBName(txt) ' Attempt to parse
                Else
                    Debug.Print "GatherFiles: ERROR reading file (Default Encoding): " & fil.Path & " - " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0 ' Restore default error handling

                '--- If not found, try UTF-16 ---
                If vbName = "" Then
                    On Error Resume Next ' Handle potential read errors
                    txt = ReadText(fil.Path, -1) ' TriStateTrue = -1 (Unicode)
                     If Err.Number = 0 Then
                        vbName = ParseVBName(txt) ' Attempt to parse
                    Else
                        Debug.Print "GatherFiles: ERROR reading file (Unicode Encoding): " & fil.Path & " - " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo 0 ' Restore default error handling
                End If

                '--- Final fallback to file base name ---
                If vbName = "" Then
                    vbName = fso.GetBaseName(fil.Name)
                    Debug.Print "GatherFiles: WARNING - VB_Name missing/unreadable; using file name: [" & vbName & "] for " & fil.Path
                End If

                '--- Add to dictionary ---
                If Len(vbName) > 0 Then
                    If Not d.Exists(vbName) Then
                        d.Add vbName, fil.Path
                        Debug.Print "GatherFiles: ADDED [" & vbName & "] -> " & fil.Path
                    Else
                        Debug.Print "GatherFiles: DUPLICATE name [" & vbName & "] found. Ignoring " & fil.Path & " (Keeping existing: " & d(vbName) & ")"
                    End If
                Else
                     Debug.Print "GatherFiles: ERROR - Could not determine a valid module name for: " & fil.Path
                End If
        End Select
    Next fil
    On Error GoTo 0 ' Restore default error handling

    ' --- 3. Recurse into sub-folders ---
    On Error Resume Next ' Handle errors enumerating subfolders
    For Each fld In fso.GetFolder(startPath).SubFolders
         If Err.Number <> 0 Then
             Debug.Print "GatherFiles: ERROR enumerating subfolders in " & startPath & ": " & Err.Description
             Err.Clear
             Exit For ' Stop trying to process subfolders if enumeration fails
         End If
        ' Recursive call - passes the same dictionary and FSO objects down
        GatherFiles fld.Path, d, fso
    Next fld
    On Error GoTo 0 ' Restore default error handling

    ' --- Cleanup objects for this level ---
    Set fil = Nothing
    Set fld = Nothing
End Sub

' --- Helper Functions for GatherFiles ---

'Utility: read entire file with chosen encoding, handles errors
Private Function ReadText(path As String, triFormat As Long) As String
    Dim fso As Object, ts As Object
    On Error Resume Next ' Handle FSO/FileOpen errors within this function
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(path, 1, False, triFormat) ' 1=ForReading
    If Err.Number = 0 And Not ts Is Nothing Then
        If Not ts.AtEndOfStream Then ReadText = ts.ReadAll
        ts.Close
    Else
        ReadText = "" ' Return empty string on error
        ' Error logged by caller (GatherFiles)
    End If
    Set ts = Nothing
    Set fso = Nothing
    On Error GoTo 0 ' Restore default error handling
End Function

'Utility: extract VB_Name (returns "" if not present or parse fails)
Private Function ParseVBName(fileContent As String) As String
    Dim p As Long, nameLine As String, extractedName As String
    ParseVBName = "" ' Default return value

    If Len(fileContent) = 0 Then Exit Function ' Skip empty content

    p = InStr(1, fileContent, "Attribute VB_Name", vbTextCompare)
    If p > 0 Then
        On Error Resume Next ' Handle potential errors splitting/parsing
        ' Extract the line containing the attribute
        nameLine = Mid$(fileContent, p) ' Get text from attribute onwards
        nameLine = Split(nameLine, vbCrLf)(0) ' Get only the first line
        nameLine = Split(nameLine, vbCr)(0)   ' Handle CR line endings too
        nameLine = Split(nameLine, vbLf)(0)   ' Handle LF line endings too

        ' Extract the name part after "="
        If InStr(nameLine, "=") > 0 Then
            extractedName = Split(nameLine, "=")(1)
            extractedName = Trim$(Replace(extractedName, """", "")) ' Clean quotes and spaces
            If Len(extractedName) > 0 Then ParseVBName = extractedName ' Success
        End If
        If Err.Number <> 0 Then ParseVBName = "": Err.Clear ' Reset on parse error
        On Error GoTo 0 ' Restore default error handling
    End If
End Function


' Copies validated exports to the source folder
Public Sub UpdateSourceFromExport()
    Dim fileName As String
    Dim fileCount As Integer
    Dim response As VbMsgBoxResult

    ' Confirm before overwriting source files
    response = MsgBox("This will copy all exported modules from:" & vbCrLf & EXPORT_FOLDER & vbCrLf & _
                     "to the source folder:" & vbCrLf & SOURCE_FOLDER & vbCrLf & vbCrLf & _
                     "Are you sure you want to update your source files?", _
                     vbQuestion + vbYesNo, "Update Source Files")

    If response = vbNo Then Exit Sub

    fileCount = 0

    ' Verify both folders exist
    Dim fso As Object ' Use FileSystemObject for better checks
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(EXPORT_FOLDER) Then
        MsgBox "Export folder doesn't exist: " & EXPORT_FOLDER, vbCritical
        Set fso = Nothing
        Exit Sub
    End If

    If Not fso.FolderExists(SOURCE_FOLDER) Then
        MsgBox "Source folder doesn't exist: " & SOURCE_FOLDER, vbCritical
        Set fso = Nothing
        Exit Sub
    End If

    ' Process all file types using FileSystemObject
    Dim exportFolderObj As Object
    Dim fileItem As Object
    Set exportFolderObj = fso.GetFolder(EXPORT_FOLDER)

    For Each fileItem In exportFolderObj.Files
        Dim fileExtension As String
        fileExtension = LCase$(fso.GetExtensionName(fileItem.Name))

        ' Only copy module files
        If fileExtension = "bas" Or fileExtension = "cls" Or fileExtension = "frm" Then
            On Error Resume Next
            fso.CopyFile fileItem.Path, SOURCE_FOLDER & fileItem.Name, True ' True = Overwrite
            If Err.Number <> 0 Then
                Debug.Print "ERROR copying " & fileItem.Name & " to " & SOURCE_FOLDER & ": " & Err.Description
                Err.Clear
            Else
                fileCount = fileCount + 1
                Debug.Print "Copied " & fileItem.Name & " to " & SOURCE_FOLDER
            End If
            On Error GoTo 0
        End If
    Next fileItem

    ' Cleanup
    Set fso = Nothing
    Set exportFolderObj = Nothing
    Set fileItem = Nothing

    MsgBox "Updated " & fileCount & " files in source folder: " & SOURCE_FOLDER, vbInformation
End Sub
