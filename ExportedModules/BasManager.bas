Attribute VB_Name = "BasManager"
Option Explicit

' === �萔��` ===
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
    
    ' �t�H���_�I���_�C�A���O�̍쐬
    Set fd = Application.fileDialog(msoFileDialogFolderPicker)
    
    path = ""
    With fd
        .Title = "�C���|�[�g����t�H���_��I�����Ă�������"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.path & Application.PathSeparator
        
        ' �_�C�A���O�\��
        If .Show = -1 Then ' ���[�U�[���uOK�v���������ꍇ
            path = .SelectedItems(1)
            MsgBox "�I�����ꂽ�t�H���_: " & path
        Else
            MsgBox "�L�����Z������܂���"
            Exit Sub
        End If
    End With
    
    Call ImportAllModules(path)
    
End Sub

' === ���ׂẴ��W���[�����G�N�X�|�[�g ===
Sub ExportAllModules(exportPath As String)
    On Error GoTo ErrHandler

    EnsureFolderExists exportPath

    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.vBComponents
        ExportModule vbComp, exportPath
    Next vbComp

    MsgBox "���W���[���̃G�N�X�|�[�g���������܂����B" & vbCrLf & exportPath, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "�G�N�X�|�[�g���ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub


' === ���ׂẴ��W���[�����C���|�[�g�i���������̓X�L�b�v�j ===
Sub ImportAllModules(importPath As String)
    On Error GoTo ErrHandler

    If Dir(importPath, vbDirectory) = "" Then
        MsgBox "�w�肳�ꂽ�t�H���_�����݂��܂���: " & importPath, vbExclamation
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
    
    ' Look for files in the directory
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

    MsgBox "���W���[���̃C���|�[�g���������܂����B", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "�C���|�[�g���ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

' === �P��t�@�C���̃��W���[�����C���|�[�g�i�����ő��݃`�F�b�N�j ===
Sub ImportSingleModule(fullPath As String, existingModules As Object)
    Dim baseName As String
    baseName = GetFileBaseName(Dir(fullPath))

    If Not existingModules.Exists(baseName) Then
        ThisWorkbook.VBProject.vBComponents.Import fullPath
    End If
End Sub

' === ���W���[�����G�N�X�|�[�g�i��ނɉ����Ċg���q������j ===
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

' === �w��t�H���_�����݂��Ȃ��ꍇ�͍쐬 ===
Sub EnsureFolderExists(ByVal folderPath As String)
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
End Sub

' === �t�@�C��������g���q����������{�����擾 ===
Function GetFileBaseName(fileName As String) As String
    GetFileBaseName = Left(fileName, InStrRev(fileName, ".") - 1)
End Function

