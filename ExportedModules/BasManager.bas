Attribute VB_Name = "BasManager"
Option Explicit

' === Constants ===
Const EXT_BAS As String = ".bas"
Const EXT_CLS As String = ".cls"
Const EXT_FRM As String = ".frm"
Const MODULE_TYPE_STANDARD As Long = 1
Const MODULE_TYPE_CLASS As Long = 2
Const MODULE_TYPE_FORM As Long = 3

Sub DoExportAllModules()
    Dim path As String
    path = ThisWorkbook.path & Application.PathSeparator & "ExportedModules"
    Call ExportAllModules(path)
End Sub

Sub DoImportAllModules()
    Dim fd As FileDialog
    Dim path As String
    
    ' Create folder selection dialog
    Set fd = Application.fileDialog(msoFileDialogFolderPicker)
    
    path = ""
    With fd
        .Title = "Select folder to import from"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.path & Application.PathSeparator
        
        ' Show dialog
        If .Show = -1 Then ' If user clicked "OK"
            path = .SelectedItems(1)
            MsgBox "Selected folder: " & path
        Else
            MsgBox "Cancelled"
            Exit Sub
        End If
    End With
    
    Call ImportAllModules(path)
    
End Sub

' === Export all modules ===
Sub ExportAllModules(exportPath As String)
    On Error GoTo ErrHandler

    EnsureFolderExists exportPath

    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.vBComponents
        ExportModule vbComp, exportPath
    Next vbComp

    MsgBox "Successfully exported modules." & vbCrLf & exportPath, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error during export: " & Err.Description, vbCritical
End Sub


' === Import all modules (skip if exists) ===
Sub ImportAllModules(importPath As String)
    On Error GoTo ErrHandler

    If Dir(importPath, vbDirectory) = "" Then
        MsgBox "Specified folder does not exist: " & importPath, vbExclamation
        Exit Sub
    End If

    Dim existingModules As Object
    Set existingModules = CreateObject("Scripting.Dictionary")

    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.vBComponents
        existingModules(vbComp.Name) = True
    Next vbComp

    Dim fileName As String
    Dim fileExtension As String
    Dim fullPath As String
    
    ' Ensure importPath ends with a path separator
    If Right(importPath, 1) <> Application.PathSeparator Then
        importPath = importPath & Application.PathSeparator
    End If
    
    fileName = Dir(importPath & "*.*")

    Do While fileName <> ""
        fileExtension = LCase(Right(fileName, 4))
        
        ' Only process .bas, .cls, and .frm files
        If fileExtension = EXT_BAS Or fileExtension = EXT_CLS Or fileExtension = EXT_FRM Then
            fullPath = importPath & fileName
            Call ImportSingleModule(fullPath, existingModules)
        End If
        
        fileName = Dir
    Loop

    MsgBox "Successfully imported modules.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error during import: " & Err.Description, vbCritical
End Sub

' === Import single module file (check for existence) ===
Sub ImportSingleModule(fullPath As String, existingModules As Object)
    Dim baseName As String
    baseName = GetFileBaseName(Dir(fullPath))

    If Not existingModules.Exists(baseName) Then
        ThisWorkbook.VBProject.vBComponents.Import fullPath
    End If
End Sub

' === Export module (set extension by type) ===
Sub ExportModule(vbComp As Object, exportPath As String)
    Dim fileExt As String
    Select Case vbComp.Type
        Case MODULE_TYPE_STANDARD: fileExt = EXT_BAS
        Case MODULE_TYPE_CLASS:    fileExt = EXT_CLS
        Case MODULE_TYPE_FORM:     fileExt = EXT_FRM
        Case Else:                 Exit Sub
    End Select

    Dim fullPath As String
    fullPath = exportPath & Application.PathSeparator & vbComp.Name & fileExt
    vbComp.Export fullPath
End Sub

' === Create folder if it doesn't exist ===
Sub EnsureFolderExists(ByVal folderPath As String)
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
End Sub

' === Get base name from file name without extension ===
Function GetFileBaseName(fileName As String) As String
    GetFileBaseName = Left(fileName, InStrRev(fileName, ".") - 1)
End Function