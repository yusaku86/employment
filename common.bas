Attribute VB_Name = "common"
Option Explicit

'// �G���[�f�[�^�쐬�̃t�H�[���N��
Public Sub openFormErrorData()

    frmErrorData.Show

End Sub

'// ���ō��Ғ��o�̃t�H�[���N��
Public Sub openFormUnpunched()

    frmUnpunched.Show
    
End Sub

'// PDF�o�͂̃t�H�[���N��
Public Sub openFormPDF()

    frmPDF.Show

End Sub

'// �`���b�g���M�̃t�H�[���N��
Public Sub openFormChatwork()

    frmChatwork.Show

End Sub

'// �R�݉^���̃G���[�f�[�^�ۑ���ύX
Public Sub changePathYamagishi()

    Call changePath("�R�݉^����")

End Sub

'// YCL�̃G���[�f�[�^�ۑ���ύX
Public Sub changePathYCL()

    Call changePath("��YCL")

End Sub


'// �G���[�f�[�^�̕ۑ���ύX
Private Sub changePath(ByVal company As String)

    Dim folderPath As String: folderPath = selectFolder("�ۑ���ύX:" & company, "G:\")
    
    If folderPath = "" Then
        Exit Sub
    End If
    
    If company = "�R�݉^����" Then
        Sheets("�ݒ�").Cells(8, 13).Value = folderPath
    ElseIf company = "��YCL" Then
        Sheets("�ݒ�").Cells(8, 14).Value = folderPath
    End If
    
End Sub
