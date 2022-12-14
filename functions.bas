Attribute VB_Name = "functions"
 '// 関数を定義するモジュール
Option Explicit

'// ダイアログを表示してファイルを選択
Public Function selectFile(ByVal dialogTitle As String, ByVal initialFolder As String, ByVal filterDescription As String, ByVal filterExtension As String) As String

    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = initialFolder
        .AllowMultiSelect = False
        .title = dialogTitle
        
        .Filters.Clear
        .Filters.Add filterDescription, filterExtension

        If .Show Then
            selectFile = .SelectedItems(1)
        Else
            selectFile = ""
        End If
    End With
    
End Function

'// ダイアログを表示してフォルダを選択
Public Function selectFolder(ByVal dialogTitle As String, ByVal initialFolder As String) As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = initialFolder
        .AllowMultiSelect = False
        .title = dialogTitle
        
        If .Show Then
            selectFolder = .SelectedItems(1)
        Else
            selectFolder = ""
        End If
    End With
    
End Function

'// シートが無ければ作成
Public Sub createSheetIfNotExist(ByVal sheetName As String, ByVal targetBook As Workbook)

    Dim mySheet As Worksheet
    
    For Each mySheet In targetBook.Worksheets
        If mySheet.Name = sheetName Then
            Exit Sub
        End If
    Next
    
    With targetBook.Worksheets.Add(after:=targetBook.Worksheets(targetBook.Sheets.Count))
        .Name = sheetName
    End With
    
End Sub
