Attribute VB_Name = "BasManager"
Option Explicit

' === 定数定義 ===
Const EXT_BAS As String = ".bas"
Const EXT_CLS As String = ".cls"
Const EXT_FRM As String = ".frm"
Const MODULE_TYPE_STANDARD As Long = 1
Const MODULE_TYPE_CLASS As Long = 2
Const MODULE_TYPE_FORM As Long = 3
Dim path As String

Sub DoExportAllModules()
    path = ThisWorkbook.path & Application.PathSeparator & "ExportedModules"
    Call ExportAllModules(path)
End Sub

Sub DoImportAllModules()
    Dim fd As fileDialog
    
    ' フォルダ選択ダイアログの作成
    Set fd = Application.fileDialog(msoFileDialogFolderPicker)
    
    path = ""
    With fd
        .Title = "インポートするフォルダを選択してください"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.path & Application.PathSeparator
        
        ' ダイアログ表示
        If .Show = -1 Then ' ユーザーが「OK」を押した場合
            path = .SelectedItems(1)
            MsgBox "選択されたフォルダ: " & path
        Else
            MsgBox "キャンセルされました"
            Exit Sub
        End If
    End With
    
    Call ImportAllModules(path)
    
End Sub

' === すべてのモジュールをエクスポート ===
Sub ExportAllModules(exportPath As String)
    On Error GoTo ErrHandler

    EnsureFolderExists exportPath

    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.vBComponents
        ExportModule vbComp, exportPath
    Next vbComp

    MsgBox "モジュールのエクスポートが完了しました。" & vbCrLf & exportPath, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "エクスポート中にエラーが発生しました: " & Err.Description, vbCritical
End Sub


' === すべてのモジュールをインポート（同名既存はスキップ） ===
Sub ImportAllModules(importPath As String)
    On Error GoTo ErrHandler

    If Dir(importPath, vbDirectory) = "" Then
        MsgBox "指定されたフォルダが存在しません: " & importPath, vbExclamation
        Exit Sub
    End If

    Dim existingModules As Object
    Set existingModules = CreateObject("Scripting.Dictionary")

    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.vBComponents
        existingModules(vbComp.Name) = True
    Next vbComp

    Dim fileName As String
    fileName = Dir(importPath & "*.*")

    Do While fileName <> ""
        Call ImportSingleModule(importPath & fileName, existingModules)
        fileName = Dir
    Loop

    MsgBox "モジュールのインポートが完了しました。", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "インポート中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' === 単一ファイルのモジュールをインポート（辞書で存在チェック） ===
Sub ImportSingleModule(fullPath As String, existingModules As Object)
    Dim baseName As String
    baseName = GetFileBaseName(Dir(fullPath))

    If Not existingModules.Exists(baseName) Then
        ThisWorkbook.VBProject.vBComponents.Import fullPath
    End If
End Sub

' === モジュールをエクスポート（種類に応じて拡張子を決定） ===
Sub ExportModule(vbComp As Object, exportPath As String)
    Dim fileExt As String
    Select Case vbComp.Type
        Case MODULE_TYPE_STANDARD: fileExt = EXT_BAS
        Case MODULE_TYPE_CLASS:    fileExt = EXT_CLS
        Case MODULE_TYPE_FORM:     fileExt = EXT_FRM
        Case Else:                 Exit Sub
    End Select

    Dim fullPath As String
    fullPath = exportPath & "\" & vbComp.Name & fileExt
    vbComp.Export fullPath
End Sub

' === 指定フォルダが存在しない場合は作成 ===
Sub EnsureFolderExists(ByVal folderPath As String)
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
End Sub

' === ファイル名から拡張子を除いた基本名を取得 ===
Function GetFileBaseName(fileName As String) As String
    GetFileBaseName = Left(fileName, InStrRev(fileName, ".") - 1)
End Function

