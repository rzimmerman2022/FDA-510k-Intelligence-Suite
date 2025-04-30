Option Explicit

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
' Recursively populate dict(VB_Name) = fullPath while
'  • skipping any path that contains "\_legacy\" or starts with _legacy
'  • reading internal VB_Name from file
'  • ignoring subsequent duplicates based on internal VB_Name
'-----------------------------------------------------------
Private Sub GatherFiles(ByVal startPath As String, ByRef d As Object, ByRef fso As Object)
    Dim fld As Object, fil As Object, vbName As String, txt As String, p As Long
    Dim currentFolder As Object ' Folder object

    ' 1) Skip everything under _legacy at any depth
    If InStr(1, LCase$(startPath), "\_legacy\", vbTextCompare) > 0 _
       Or LCase$(fso.GetFileName(startPath)) Like "_legacy*" Then
        Debug.Print "GatherFiles: Skipping folder path containing _legacy: " & startPath
        Exit Sub
    End If

    ' Check if folder exists before trying to access it
    If Not fso.FolderExists(startPath) Then
        Debug.Print "GatherFiles: Folder not found: " & startPath
        Exit Sub
    End If
    Set currentFolder = fso.GetFolder(startPath)

    ' 2) Scan files in the current folder
    On Error Resume Next ' Handle potential permission errors accessing files
    For Each fil In currentFolder.Files
        If Err.Number <> 0 Then
            Debug.Print "GatherFiles: Error accessing files in " & startPath & ": " & Err.Description
            Err.Clear
            Exit For ' Stop processing files in this folder on error
        End If

        Select Case LCase$(fso.GetExtensionName(fil.Name))
            Case "bas", "cls", "frm"
                ' Read first ~2 KB to find Attribute VB_Name
                Dim ts As Object ' TextStream
                vbName = "" ' Reset for each file
                Set ts = fso.OpenTextFile(fil.Path, 1) ' 1 = ForReading
                If Err.Number <> 0 Then
                    Debug.Print "GatherFiles: Error opening file " & fil.Path & ": " & Err.Description
                    Err.Clear
                Else
                    If Not ts.AtEndOfStream Then
                        txt = ts.Read(2048) ' Read chunk
                        ts.Close
                        p = InStr(1, txt, "Attribute VB_Name", vbTextCompare)
                        If p > 0 Then
                            ' Extract the name, handle quotes and whitespace
                            vbName = Split(Mid$(txt, p), "=")(1) ' Get part after '='
                            vbName = Split(vbName, vbCrLf)(0)    ' Get first line
                            vbName = Trim$(vbName)
                            If Left$(vbName, 1) = """" And Right$(vbName, 1) = """" Then
                                vbName = Mid$(vbName, 2, Len(vbName) - 2) ' Remove quotes
                            End If

                            ' Add to dictionary if the VB_Name is not already present
                            If Len(vbName) > 0 Then
                                If Not d.Exists(vbName) Then
                                    d.Add vbName, fil.Path
                                    Debug.Print "GatherFiles: Added [" & vbName & "] -> " & fil.Path
                                Else
                                    Debug.Print "GatherFiles: Duplicate VB_Name [" & vbName & "] found, ignoring " & fil.Path & " (Already have: " & d(vbName) & ")"
                                End If
                            Else
                                Debug.Print "GatherFiles: Could not extract valid VB_Name from " & fil.Path
                            End If
                        Else
                             Debug.Print "GatherFiles: Attribute VB_Name not found in first 2KB of " & fil.Path
                        End If
                    Else
                         ts.Close ' Close even if empty
                         Debug.Print "GatherFiles: File is empty " & fil.Path
                    End If
                    Set ts = Nothing
                End If
        End Select
    Next fil
    On Error GoTo 0 ' Restore error handling

    ' 3) Recurse into subfolders
    On Error Resume Next ' Handle potential permission errors accessing subfolders
    For Each fld In currentFolder.SubFolders
         If Err.Number <> 0 Then
            Debug.Print "GatherFiles: Error accessing subfolders in " & startPath & ": " & Err.Description
            Err.Clear
            Exit For ' Stop processing subfolders on error
        End If
        ' Recursive call - pass the subfolder's path
        GatherFiles fld.Path, d, fso ' Pass dictionary and fso
    Next fld
    On Error GoTo 0 ' Restore error handling

    ' Cleanup for this level
    Set currentFolder = Nothing
    Set fld = Nothing
    Set fil = Nothing
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
