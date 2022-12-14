VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPDF 
   Caption         =   "PDF�o��"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6480
   OleObjectBlob   =   "frmPDF.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// PDF�o�͂��邽�߂̃t�H�[��
Option Explicit

'// ���s���������Ƃ��̏���(���C���v���V�[�W��)
Private Sub cmdEnter_Click()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '// �o���f�[�V����
    If validate = False Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(Me.txtHiddenFileFullPath)
    Dim targetSheet As Worksheet
    
    '// �w��̃t�@�C���̑S�V�[�g�̈����ݒ�
    For Each targetSheet In targetFile.Worksheets
        
        With targetSheet.PageSetup
            .Zoom = False
            .FitToPagesTall = 1
            .FitToPagesWide = 1
            .Orientation = xlLandscape
        End With
    Next
    
    Dim wsh As New WshShell
    
    targetFile.Worksheets.Select
    
    '// ExportAsFixedFormat [�t�@�C���^�C�v], [�t�@�C����]
    targetFile.ExportAsFixedFormat xlTypePDF, wsh.SpecialFolders(4) & "\" & Me.cmbCompany.Value & Me.cmbDepartment.Value & " " & Me.cmbMonth.Value & Me.cmbStartDay & "�`" & Me.cmbMonth.Value & Me.cmbEndDay.Value & "��.pdf"

    targetFile.Close False
    Set targetFile = Nothing
    
    Me.cmbDepartment.Value = ""
    Me.txtHiddenFileFullPath.Value = ""
    
    With Me.txtFileName
        .Locked = False
        .Value = ""
        .Locked = True
    End With
    
    MsgBox "PDF�o�͂��������܂����B", vbQuestion, "PDF�o��"

End Sub

'// ���͂��ꂽ�l�̃o���f�[�V����
Private Function validate() As Boolean
    
    validate = False
    
    '// ��Ж����I������Ă��邩
    If Me.cmbCompany.Value = "" Then
        MsgBox "��Ж���I�����Ă��������B", vbQuestion, "PDF�o��"
        Exit Function
    
    '// ���������I������Ă��邩
    ElseIf Me.cmbDepartment.Value = "" Then
        MsgBox "��������I�����Ă��������B"
        Exit Function
    
    '// �Ώی����I������Ă��邩
    ElseIf Me.cmbMonth.Value = "" Then
        MsgBox "�Ώی���I�����Ă��������B"
        Exit Function
    
    '// ���Ԃ��I������Ă��邩
    ElseIf Me.cmbStartDay.Value = "" Or Me.cmbEndDay.Value = "" Then
        MsgBox "���Ԃ�I�����Ă��������B"
        Exit Function

    '// ���Ԃ̏I�������J�n�����O�ɗ��Ă��Ȃ���
    ElseIf Me.cmbStartDay.Value > Me.cmbEndDay.Value Then
        MsgBox "���Ԃ̏I�����͊J�n������ɐݒ肵�Ă��������B"
        Exit Function
    End If
    
    validate = True
    
End Function

'// �t�H�[���N�����̏���
Private Sub UserForm_Initialize()

    Me.cmbCompany.AddItem "�R�݉^����"
    Me.cmbCompany.AddItem "��YCL"
    
    Dim i As Long
    
    '/**
     '* �Ώی��R���{�{�b�N�X�̑I�����ǉ��Ə����l�ݒ�
    '**/
    For i = 4 To 12
        Me.cmbMonth.AddItem i & "��"
    Next
    For i = 1 To 3
        Me.cmbMonth.AddItem i & "��"
    Next
    
    If Day(Now) < 16 Then
        Me.cmbMonth.Value = Month(Now) - 1 & "��"
    Else
        Me.cmbMonth.Value = Month(Now) & "��"
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

'// ��Ж����ύX���ꂽ�ꍇ
Private Sub cmbCompany_Change()

    If Me.cmbCompany.Value = "" Then
        Exit Sub
    End If
    
    Me.cmbDepartment.Clear
    
    Dim targetColumn As Long
    
    If Me.cmbCompany.Value = "�R�݉^����" Then
        targetColumn = 8
    ElseIf Me.cmbCompany.Value = "��YCL" Then
        targetColumn = 9
    End If
    
    Dim i As Long
    
    '// �����R���{�{�b�N�X�̑I��������Ђɍ��킹�ĕύX
    For i = 8 To Sheets("�ݒ�").Cells(Rows.Count, targetColumn).End(xlUp).Row
        If Sheets("�ݒ�").Cells(i, targetColumn).Value = "" Then
            GoTo Continue
        End If
        
        Me.cmbDepartment.AddItem Sheets("�ݒ�").Cells(i, targetColumn).Value
Continue:
    Next
    
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

End Sub

'// �Q�Ƃ��������Ƃ��̏���
Private Sub cmdDialog_Click()

    Dim wsh As New WshShell
    Dim fso As New FileSystemObject
    
    Dim targetFile As String: targetFile = selectFile("PDF�o�͂���t�@�C����I�����Ă��������B", wsh.SpecialFolders(4) & "\", "Excel�ECSV", "*.xlsx;*.csv")

    Me.txtFileName.Locked = False
    
    If targetFile <> "" Then
        Me.txtFileName.Value = fso.GetFileName(targetFile)
        Me.txtHiddenFileFullPath.Value = targetFile
    End If
    
    Me.txtFileName.Locked = True
    
    Set wsh = Nothing
    Set fso = Nothing

End Sub

'// ������������Ƃ��̏���
Private Sub cmdClose_Click()

    Unload Me

End Sub
