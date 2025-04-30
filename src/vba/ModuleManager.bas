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
' Incorporates Unicode reading and fallback to filename if VB_Name attribute is missing.
'-----------------------------------------------------------
Private Sub GatherFiles(startPath As String, ByRef moduleDict As Object, ByRef fileSystemObj As Object)
    ' Constants for file operations
    Const ForReading As Long = 1
    Const TriStateFalse As Long = 0      ' Use system default (ANSI/UTF-8) - Not used here
    Const TriStateTrue As Long = -1      ' Force Unicode (UTF-16)
    Const TriStateUseDefault As Long = -2 ' Use system default (ANSI/UTF-8) - Not used here
    Const ReadChunkSize As Long = 2048   ' Initial bytes to read for VB_Name attribute

    Dim folderName As String
    Dim folderObj As Object
    Dim fileObj As Object
    Dim subFolderObj As Object
    Dim fileText As String
    Dim moduleName As String
    Dim attributePos As Long
    Dim textStream As Object

    ' --- 1. Skip this folder if its own name starts with _legacy ---
    ' Use Trim$ to handle potential trailing spaces if startPath comes from user input/config
    folderName = LCase$(fileSystemObj.GetFileName(Trim$(startPath)))
    If folderName Like "_legacy*" Then
        Debug.Print "GatherFiles: SKIP folder (legacy name): " & startPath
        Exit Sub
    End If

    ' --- Check if folder exists before proceeding ---
    If Not fileSystemObj.FolderExists(startPath) Then
        Debug.Print "GatherFiles: ERROR - Folder not found: " & startPath
        Exit Sub
    End If

    Debug.Print "GatherFiles: Scanning folder: " & startPath

    ' --- Get the folder object once ---
    On Error Resume Next ' Handle potential permission errors getting folder object
    Set folderObj = fileSystemObj.GetFolder(startPath)
    If Err.Number <> 0 Then
        Debug.Print "GatherFiles: ERROR - Cannot access folder: " & startPath & " - " & Err.Description
        Err.Clear
        Exit Sub ' Cannot proceed with this folder
    End If
    On Error GoTo 0 ' Restore default error handling

    ' --- 2. Scan files in the current folder ---
    On Error Resume Next ' Handle potential errors enumerating files
    For Each fileObj In folderObj.Files
        If Err.Number <> 0 Then
            Debug.Print "GatherFiles: ERROR enumerating files in " & startPath & ": " & Err.Description
            Err.Clear
            Exit For ' Stop trying to process files in this folder if enumeration fails
        End If

        Select Case LCase$(fileSystemObj.GetExtensionName(fileObj.Name))
            Case "bas", "cls", "frm"
                moduleName = "" ' Reset for each file
                attributePos = 0
                Set textStream = Nothing ' Reset text stream object

                ' --- Attempt to open file as Unicode ---
                On Error Resume Next ' Handle file open errors (e.g., locked file)
                Set textStream = fileSystemObj.OpenTextFile(fileObj.Path, ForReading, False, TriStateTrue)
                If Err.Number <> 0 Or textStream Is Nothing Then
                    Debug.Print "GatherFiles: ERROR opening file (Unicode): " & fileObj.Path & " - " & Err.Description
                    Err.Clear
                    GoTo NextFile ' Skip this file
                End If
                On Error GoTo 0

                ' --- Read initial chunk to find VB_Name ---
                If Not textStream.AtEndOfStream Then
                    fileText = textStream.Read(ReadChunkSize)
                    attributePos = InStr(1, fileText, "Attribute VB_Name", vbTextCompare)
                Else
                    Debug.Print "GatherFiles: WARNING - File is empty: " & fileObj.Path
                    textStream.Close ' Close the empty file stream
                    GoTo NextFile ' Skip empty file
                End If

                ' --- If not found in first chunk, read the whole file (if not already at end) ---
                If attributePos = 0 And Not textStream.AtEndOfStream Then
                    textStream.Close ' Close the initial stream
                    Debug.Print "GatherFiles: VB_Name not in first " & ReadChunkSize & " bytes, reading full file: " & fileObj.Path
                    Set textStream = fileSystemObj.OpenTextFile(fileObj.Path, ForReading, False, TriStateTrue) ' Re-open
                    If Not textStream Is Nothing Then
                        If Not textStream.AtEndOfStream Then
                            fileText = textStream.ReadAll
                            attributePos = InStr(1, fileText, "Attribute VB_Name", vbTextCompare)
                        End If
                    Else
                         Debug.Print "GatherFiles: ERROR re-opening full file (Unicode): " & fileObj.Path
                         GoTo NextFile ' Skip if re-open fails
                    End If
                End If

                ' --- Close the stream ---
                If Not textStream Is Nothing Then
                    textStream.Close
                    Set textStream = Nothing
                End If

                ' --- Extract or Fallback ---
                If attributePos > 0 Then
                    ' Extract VB_Name from the attribute line
                    On Error Resume Next ' Handle potential errors splitting/parsing the line
                    moduleName = Split(Split(Mid$(fileText, attributePos), "=")(1), vbCrLf)(0)
                    moduleName = Trim$(Replace(moduleName, """", "")) ' Clean quotes and spaces
                    If Err.Number <> 0 Or Len(moduleName) = 0 Then
                        Debug.Print "GatherFiles: ERROR parsing VB_Name attribute in: " & fileObj.Path & " - Falling back to filename."
                        Err.Clear
                        moduleName = fileSystemObj.GetBaseName(fileObj.Name) ' Fallback on parse error
                    Else
                        Debug.Print "GatherFiles: Found VB_Name attribute: [" & moduleName & "] in " & fileObj.Path
                    End If
                    On Error GoTo 0
                Else
                    ' Attribute not found - Log error and skip this file (enforces best practice)
                    Debug.Print "GatherFiles: ERROR - Attribute VB_Name not found in " & fileObj.Path & ". Skipping file. Ensure it was exported correctly."
                    moduleName = "" ' Ensure moduleName is empty so it's not added
                End If

                ' --- Add to dictionary ONLY if a valid VB_Name was extracted ---
                If Len(moduleName) > 0 Then
                    If Not moduleDict.Exists(moduleName) Then
                        moduleDict.Add moduleName, fileObj.Path
                        Debug.Print "GatherFiles: ADDED [" & moduleName & "] -> " & fileObj.Path
                    Else
                        Debug.Print "GatherFiles: DUPLICATE VB_Name [" & moduleName & "] found. Ignoring " & fileObj.Path & " (Keeping existing: " & moduleDict(moduleName) & ")"
                    End If
                ' Else: If moduleName is empty (attribute missing or parse error), do nothing - file is skipped.
                End If
        End Select
NextFile:
        ' Loop cleanup (optional, good practice)
        Set textStream = Nothing
    Next fileObj
    On Error GoTo 0 ' Restore default error handling

    ' --- 3. Recurse into sub-folders ---
    On Error Resume Next ' Handle potential errors enumerating subfolders
    For Each subFolderObj In folderObj.SubFolders
         If Err.Number <> 0 Then
             Debug.Print "GatherFiles: ERROR enumerating subfolders in " & startPath & ": " & Err.Description
             Err.Clear
             Exit For ' Stop trying to process subfolders if enumeration fails
         End If
        ' Recursive call - passes the same dictionary and FSO objects down
        GatherFiles subFolderObj.Path, moduleDict, fileSystemObj
    Next subFolderObj
    On Error GoTo 0 ' Restore default error handling

    ' --- Cleanup objects for this level ---
    Set fileObj = Nothing
    Set subFolderObj = Nothing
    Set folderObj = Nothing
    Set textStream = Nothing ' Ensure cleanup even if errors occurred mid-loop
End Sub


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
