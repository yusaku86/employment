Attribute VB_Name = "common"
Option Explicit

'// エラーデータ作成のフォーム起動
Public Sub openFormErrorData()

    frmErrorData.Show

End Sub

'// 未打刻者抽出のフォーム起動
Public Sub openFormUnpunched()

    frmUnpunched.Show
    
End Sub

'// PDF出力のフォーム起動
Public Sub openFormPDF()

    frmPDF.Show

End Sub

'// チャット送信のフォーム起動
Public Sub openFormChatwork()

    frmChatwork.Show

End Sub

'// 山岸運送のエラーデータ保存先変更
Public Sub changePathYamagishi()

    Call changePath("山岸運送㈱")

End Sub

'// YCLのエラーデータ保存先変更
Public Sub changePathYCL()

    Call changePath("㈱YCL")

End Sub


'// エラーデータの保存先変更
Private Sub changePath(ByVal company As String)

    Dim folderPath As String: folderPath = selectFolder("保存先変更:" & company, "G:\")
    
    If folderPath = "" Then
        Exit Sub
    End If
    
    If company = "山岸運送㈱" Then
        Sheets("設定").Cells(8, 13).Value = folderPath
    ElseIf company = "㈱YCL" Then
        Sheets("設定").Cells(8, 14).Value = folderPath
    End If
    
End Sub
