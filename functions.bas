Attribute VB_Name = "functions"
 '// �֐����`���郂�W���[��
Option Explicit

'// �_�C�A���O��\�����ăt�@�C����I��
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

'// �_�C�A���O��\�����ăt�H���_��I��
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

'// �V�[�g��������΍쐬
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
