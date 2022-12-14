VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChatwork 
   Caption         =   "���M���e����"
   ClientHeight    =   8535.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15840
   OleObjectBlob   =   "frmChatwork.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmChatwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// �`���b�g���M�̂��߂̃t�H�[��
Option Explicit

'/**
 '* ���C���v���O����(�`���b�g�ő�����e�ݒ�&�`���b�g���M)
'**/
Private Sub cmdEnter_Click()

    '// �o���f�[�V����
    If validate = False Then: Exit Sub
    
    '// ���M��O���[�vID
    Dim roomId As String: roomId = Split(cmbRoom.Value, ":")(0)
    
    '// API�g�[�N��
    Dim apiToken As String: apiToken = Sheets("�`���b�g�ݒ�").Cells(7, 4).Value
    
    '/**
     '* �`���b�g���M
    '**/
    Dim cc As New ChatWorkController
    
    '// �`���b�g���M�p�̕��͍쐬  createChatWorkText [�����V�������X�g], [�^�C�g��], [���M���b�Z�[�W]
    Dim message As String: message = cc.createChatWorkText(Split(Me.txtMentionList.Value, vbCrLf), "���ō��ꗗ:" & Me.cmbCompany.Value & " " & Me.cmbDepartment.Value, Me.txtMessage.Value)
    
    Dim result As Boolean
 
    '// ���b�Z�[�W�̂ݑ��M����ꍇ
    If Me.txtFile.Value = "" Then
        result = cc.sendMessage(message, roomId, apiToken)
    
    '// �t�@�C���ƃ��b�Z�[�W�𑗐M����ꍇ
    Else
        result = cc.sendMessageWithFile(message, Me.txtHiddenFileFullPath.Value, roomId, apiToken)
    End If
    
    If result = True Then
        MsgBox "���M���������܂����B", vbInformation, ThisWorkbook.Name
    Else
        MsgBox "���M�ł��܂���ł����B", vbExclamation, ThisWorkbook.Name
        Exit Sub
    End If
      
    Me.cmbDepartment.Value = ""
    Me.txtMentionList.Value = ""
    Me.txtHiddenFileFullPath.Value = ""
    
    With Me.txtFile
        .Locked = False
        .Value = ""
        .Locked = True
    End With
    
End Sub

'/**
 '* �o���f�[�V����
'**/
Private Function validate() As Boolean

    validate = False

    '// ��Ж������͂���Ă��邩
    If cmbCompany.Value = "" Then
        MsgBox "��Ж���I�����Ă��������B", vbQuestion, "�`���b�g���M"
        cmbCompany.SetFocus
        Exit Function
    End If
    
    '// ���������I������Ă��邩
    If cmbDepartment.Value = "" Then
        MsgBox "��������I�����Ă��������B", vbQuestion, "�`���b�g���M"
        cmbDepartment.SetFocus
        Exit Function
    End If
    
    '// ���M��O���[�v���I������Ă��邩
    If cmbRoom.Value = "" Then
        MsgBox "���M��O���[�v��I�����Ă��������B", vbQuestion, "�`���b�g���M"
        cmbRoom.SetFocus
        Exit Function
    End If
    
    '// ���悪�ݒ肳��Ă��邩
    If txtMentionList.Value = "" Then
        If MsgBox("���悪�ݒ肳��Ă��܂��񂪁A���M���Ă�낵���ł���?", vbQuestion + vbYesNo, "�`���b�g���M") = vbNo Then
            Exit Function
        End If
    End If
        
    '// ���b�Z�[�W�����͂���Ă��邩
    If txtMessage.Value = "" Then
        If MsgBox("���b�Z�[�W�����͂���Ă��܂��񂪁A���M���Ă�낵���ł���?", vbQuestion + vbYesNo, "�`���b�g���M") = vbNo Then
            Exit Function
        End If
    End If
    
    '// �Y�t�t�@�C�����I������Ă��邩
    If txtFile.Value = "" Then
        If MsgBox("�t�@�C�����Y�t����Ă��܂��񂪁A���M���Ă�낵���ł���?", vbQuestion + vbYesNo, "�`���b�g���M") = vbNo Then
            Exit Function
        End If
    End If
    
    validate = True
    
End Function

'/**
 '* ���[�U�[�t�H�[���N�����̐ݒ�
'**/
Private Sub UserForm_Initialize()
                
    '// �`���b�g�O���[�v���̑I�����ǉ�
    Dim i As Long
   
    For i = 7 To ThisWorkbook.Sheets("�`���b�g�ݒ�").Cells(Rows.Count, 5).End(xlUp).Row
        cmbRoom.AddItem ThisWorkbook.Sheets("�`���b�g�ݒ�").Cells(i, 5).Value
    Next
    
    '// ���M���胊�X�g�ǉ�
    For i = 7 To ThisWorkbook.Sheets("�`���b�g�ݒ�").Cells(Rows.Count, 6).End(xlUp).Row
        cmbMention.AddItem ThisWorkbook.Sheets("�`���b�g�ݒ�").Cells(i, 6).Value
    Next
    
    '// ��Ж��̑I�����ǉ�
    cmbCompany.AddItem "�R�݉^����"
    cmbCompany.AddItem "��YCL"
    
    '// �Ώی��R���{�{�b�N�X�̑I�����ǉ��Ə����l�ݒ�
    For i = 1 To 12
        Me.cmbMonth.AddItem i & "��"
    Next
    
    If Day(Now) < 16 Then
        Me.cmbMonth.Value = Month(Now) - 1 & "��"
    Else
        Me.cmbMonth.Value = Month(Now)
    End If
    
    '/**
     '* ���ԃR���{�{�b�N�X�̏����l�ݒ�
    '**/
    Dim endOfMonth As Long: endOfMonth = Day(DateSerial(Year(Now), Replace(Me.cmbMonth.Value, "��", "") + 1, 0))
    
    '// ��������1������15���̊Ԃ̏ꍇ
    If Day(Now) < 16 Then
        Me.cmbStartDay.Value = 26
        Me.cmbEndDay.Value = endOfMonth
        
    '// ��������16������25���̏ꍇ
    ElseIf 15 < Day(Now) And Day(Now) < 26 Then
        Me.cmbStartDay.Value = 1
        Me.cmbEndDay.Value = 15
        
    '// ��������26���ȍ~�̏ꍇ
    ElseIf 25 < Day(Now) Then
        Me.cmbStartDay.Value = 16
        Me.cmbEndDay.Value = 25
    End If
    
End Sub

'// �Ώی����ύX���ꂽ���̏���
Private Sub cmbMonth_Change()

    If Me.cmbMonth.Value = "" Then
        Exit Sub
    End If
    
    Dim previousStartDay As Variant: previousStartDay = Me.cmbStartDay.Value
    Dim previousEndDay As Variant: previousEndDay = Me.cmbEndDay.Value
    
    Me.cmbStartDay.Clear
    Me.cmbEndDay.Clear
    
    Dim endOfMonth As Long: endOfMonth = Day(DateSerial(Year(Now), Replace(Me.cmbMonth.Value, "��", "") + 1, 0))
    
    Dim i As Long
    
    '// ���ԃR���{�{�b�N�X�̑I���������ɍ��킹�ĕύX(���ɂ���ē������ς�邽��)
    For i = 1 To endOfMonth
        Me.cmbStartDay.AddItem i
        Me.cmbEndDay.AddItem i
    Next

    '// ���ԃR���{�{�b�N�X�̒l���Ώی��̌������傫�������̏ꍇ �� �Ώی��̌����������ԃR���{�{�b�N�X�̒l�Ƃ���
    If endOfMonth < previousStartDay Then
        previousStartDay = endOfMonth
    End If
    
    If endOfMonth < previousEndDay Then
        previousEndDay = endOfMonth
    End If

    Me.cmbStartDay.Value = previousStartDay
    Me.cmbEndDay.Value = previousEndDay
    
    Call setMessage

End Sub

'/**
 '* ���Ԃ��ύX���ꂽ�Ƃ��̏��� �� ���M���郁�b�Z�[�W��ύX
'**/
Private Sub cmbEndDay_Change()

    Call setMessage

End Sub

Private Sub cmbStartDay_Change()
    
    Call setMessage

End Sub

'/**
 '* ��Ж��̒l���ύX���ꂽ���̏���
'**/
Private Sub cmbCompany_Change()

    Me.cmbDepartment.Clear
    
    Dim targetColumn As Long
    
    If Me.cmbCompany.Value = "�R�݉^����" Then
        targetColumn = 8
    ElseIf Me.cmbCompany.Value = "��YCL" Then
        targetColumn = 9
    End If
    
    Dim i As Long
    
    '// �������R���{�{�b�N�X�̑I�����ǉ�
    For i = 8 To Sheets("�ݒ�").Cells(Rows.Count, targetColumn).End(xlUp).Row
        If Sheets("�ݒ�").Cells(i, targetColumn).Value = "" Then
            GoTo Continue
        End If
        
        Me.cmbDepartment.AddItem Sheets("�ݒ�").Cells(i, targetColumn).Value
Continue:
    Next
    
End Sub

'/**
 '* �`���b�g�ɑ��M���郁�b�Z�[�W�ݒ�
'**/
Private Sub setMessage()

    If (Me.cmbMonth.Value <> "" And Me.cmbStartDay.Value <> "" And Me.cmbEndDay.Value <> "") = False Then
        Exit Sub
    End If
    
    Me.txtMessage.Value = "�����l�ł��B" & _
                          vbLf & Me.cmbMonth.Value & Me.cmbStartDay.Value & "���`" & Me.cmbMonth.Value & Me.cmbEndDay.Value & "���܂ł̖��ō��ł��B" _
                        & vbLf & setDeadline() & "�܂łɕK���񓚂��肢���܂��B"

End Sub

'/**
 '* �񓚂̒��ߐ؂����ݒ�
'**/
Private Function setDeadline() As String

    Dim finished As Boolean: finished = False
    Dim i As Long: i = 2
    
    '// ��������2���オ��{�̒��ߐ؂肾���A�y���̏ꍇ�͕����ɂȂ�܂Œ��߂����L�΂�
    Do Until finished = True
        If Weekday(DateSerial(Year(Now), Month(Now), Day(Now) + i)) <> 0 And Weekday(DateSerial(Year(Now), Month(Now), Day(Now) + i)) <> 6 Then
            finished = True
            Exit Do
        End If
        
        i = i + 1
    Loop

    setDeadline = Format(DateSerial(Year(Now), Month(Now), Day(Now) + i), "m��d��(aaa)")

End Function

'/**
 '* �ǉ����������Ƃ��̏���
'**/
Private Sub cmdAdd_Click()
    
    If cmbMention.Value = "" Then
        Exit Sub
    End If
    
    '// ���Ɉ��悪�ǉ�����Ă����甲����
    If InStr(1, txtMentionList.Value, cmbMention.Value) > 0 Then
        cmbMention.Value = ""
        Exit Sub
    End If
    
        
    If txtMentionList.Value = "" Then
        txtMentionList.Value = cmbMention.Value
    Else
        txtMentionList.Value = txtMentionList.Value & vbLf & cmbMention.Value
    End If
    
    cmbMention.Value = ""

End Sub

'// ���Z�b�g���������Ƃ��̏���
Private Sub cmdReset_Click()

    txtMentionList.Value = ""

End Sub

'// �Q�Ƃ��������Ƃ��̏���
Private Sub cmdDialog_Click()

    Dim wsh As Object: Set wsh = CreateObject("Wscript.Shell")
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    Dim attachedFileName As String: attachedFileName = selectFile("�Y�t�t�@�C���I��", wsh.SpecialFolders(4) & "\", "Excel�t�@�C���EPDF", "*.xlsx;*.pdf;*.csv")
    
    If attachedFileName <> "" Then
        txtFile.Locked = False
    
        txtFile.Value = fso.GetFileName(attachedFileName)
        txtHiddenFileFullPath.Value = attachedFileName
    
        txtFile.Locked = True
    End If
    
    Set wsh = Nothing
    Set fso = Nothing
    
End Sub

'// �t�@�C�������Z�b�g�����������̏���
Private Sub cmdClearFile_Click()

    Me.txtFile.Locked = False
    
    Me.txtFile.Value = ""
    Me.txtHiddenFileFullPath.Value = ""
    
    Me.txtFile.Locked = True

End Sub

'// ������������Ƃ��̏���
Private Sub cmdCancel_Click()
    
    Unload Me

End Sub
