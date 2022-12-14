VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmErrorData 
   Caption         =   "�G���[�f�[�^�쐬"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   OleObjectBlob   =   "frmErrorData.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmErrorData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// �G���[�f�[�^���쐬����t�H�[��
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
    
    '// �I�����ꂽ�t�@�C�����K�؂Ȃ��̂��m�F
    If validateFile = False Then
        If MsgBox("�w�肵���t�@�C�����s�K�؂ȉ\��������܂��B" & vbLf & "�����𑱍s���܂���?", vbQuestion + vbYesNo, "�G���[�f�[�^�쐬") = vbNo Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End If
    
    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(Me.txtHiddenFileFullPath.Value)
    
    Dim i As Long
        
    '// �G���[�f�[�^�̓��s�v�Ȃ��̂��폜
    For i = 7 To ThisWorkbook.Sheets("�ݒ�").Cells(Rows.Count, 11).End(xlUp).Row
        If ThisWorkbook.Sheets("�ݒ�").Cells(i, 11).Value = "" Then
            GoTo Continue
        End If
        
        targetFile.Sheets(1).Cells(1, 1).AutoFilter 1, "*" & ThisWorkbook.Sheets("�ݒ�").Cells(i, 11).Value & "*"
        
        If targetFile.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row <> 1 Then
            targetFile.Sheets(1).Cells(1, 1).CurrentRegion.Offset(1).Delete
        End If

        targetFile.Sheets(1).Cells(1, 1).AutoFilter

Continue:
    Next

    '// A���<ObcUnacceptedReason></ObcUnacceptedReason>���폜
    For i = 1 To targetFile.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
        targetFile.Sheets(1).Cells(i, 1).Value = Replace(Replace(targetFile.Sheets(1).Cells(i, 1).Value, "<ObcUnacceptedReason>", ""), "</ObcUnacceptedReason>", "")
    Next

    '// �G���[�f�[�^���e�[�u���Ƃ��ĕۑ�
    Dim errTable As ListObject: Set errTable = targetFile.Sheets(1).ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, Cells(1, Columns.Count).End(xlToLeft).Column)), , xlYes)
    errTable.TableStyle = "TableStyleLight18"
    
    Set errTable = Nothing
    
    targetFile.Sheets(1).Cells.EntireColumn.AutoFit
    
    Dim folderPath As String: folderPath = createFolderPath
    
    '// �G���[�f�[�^��csv�Ȃ̂�Excel�t�@�C��(xlsx)�Ƃ��ĕۑ�
    targetFile.SaveAs folderPath & "\" & Me.cmbCompany.Value & " �ō��G���[�f�[�^.xlsx", xlOpenXMLWorkbook
    
    targetFile.Close False
    
    Set targetFile = Nothing
    
    Me.cmbCompany.Value = ""
    Me.txtFileName.Value = ""
    Me.txtHiddenFileFullPath.Value = ""
    
    MsgBox "�������������܂����B", vbInformation, "�G���[�f�[�^�쐬:" & Me.cmbCompany.Value

End Sub

'// ���͂��ꂽ�l�ƕۑ���̃o���f�[�V����
Private Function validate() As Boolean

    validate = False
    
    '// ��Ж����I������Ă��邩
    If Me.cmbCompany.Value = "" Then
        MsgBox "��Ж���I�����Ă��������B", vbQuestion, "�G���[�f�[�^�쐬"
        Exit Function

    '// �Ώی����I������Ă��邩
    ElseIf Me.cmbMonth.Value = "" Then
        MsgBox "�Ώی���I�����Ă��������B", vbQuestion, "�G���[�f�[�^�쐬"
        Exit Function
    
    '// ���Ԃ��I������Ă��邩
    ElseIf Me.cmbStartDay.Value = "" Or Me.cmbEndDay.Value = "" Then
        MsgBox "���Ԃ�I�����Ă��������B", vbQuestion, "�G���[�f�[�^�쐬"
        Exit Function

    '// ���Ԃ̏I�������J�n������ɂȂ��Ă��Ȃ���
    ElseIf Me.cmbEndDay.Value < Me.cmbStartDay.Value Then
        MsgBox "���Ԃ̏I�����͊J�n������ɐݒ肵�Ă��������B", vbQuestion, "�G���[�f�[�^�쐬"
        Exit Function
    End If
    
    '// �G���[�f�[�^�̕ۑ���t�H���_�����݂��邩
    Dim folderPath As String
    
    If Me.cmbCompany.Value = "�R�݉^����" Then
        folderPath = Sheets("�ݒ�").Cells(8, 13).Value
    ElseIf Me.cmbCompany.Value = "��YCL" Then
        folderPath = Sheets("�ݒ�").Cells(8, 14).Value
    End If
    
    Dim fso As New FileSystemObject
        
    If folderPath = "" Then
        MsgBox "�G���[�f�[�^�̕ۑ���t�H���_���ݒ肳��Ă��܂���B" & vbLf & "�V�[�g�u�ݒ�v�Őݒ肵�Ă��������B", vbQuestion, "�G���[�f�[�^�쐬"
        Exit Function
        
    ElseIf fso.FolderExists(folderPath) = False Then
        MsgBox "�G���[�f�[�^�̕ۑ���t�H���_�����݂��܂���B", vbQuestion, "�G���[�f�[�^�쐬"
        Exit Function
    End If
    
    validate = True

End Function

'// �I�����ꂽ�t�@�C�����K�؂Ȃ��̂��m�F
Private Function validateFile() As Boolean

    validateFile = False

    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(Me.txtHiddenFileFullPath.Value, ReadOnly:=True)
    
    If Cells(1, 1).Value = "<ObcUnacceptedReason> �G���[���e </ObcUnacceptedReason>" _
        And Cells(1, 2).Value = "���t" _
        And Cells(1, 3).Value = "�Ј��ԍ�" Then
            
        validateFile = True
    End If
    
    targetFile.Close False
    Set targetFile = Nothing
        
End Function

'// �G���[�f�[�^�̕ۑ���p�X�쐬
Private Function createFolderPath() As String

    Dim folderPath As String
    
    If Me.cmbCompany.Value = "�R�݉^����" Then
        folderPath = ThisWorkbook.Sheets("�ݒ�").Cells(8, 13).Value
    ElseIf Me.cmbCompany.Value = "��YCL" Then
        folderPath = ThisWorkbook.Sheets("�ݒ�").Cells(8, 14).Value
    End If
    
    '// �Ώ۔N
    Dim targetYear As Long: targetYear = Year(Now)
    
    '// ������(���̃t�@�C���𑀍삵�Ă��鎞�̌���1���ŁA�Ώی���12���̏ꍇ�͑Ώ۔N���������Ă���Ƃ��̔N����1�}�C�i�X�������ɂȂ�)
    If Month(Now) = 1 And Me.cmbMonth.Value = 12 Then
        targetYear = Year(Now) - 1
    End If
    
    Dim fso As New FileSystemObject
    
    '// �u�ݒ�V�[�g�ɓ��͂���Ă���t�H���_\�Ώ۔N.�Ώی��v�̃t�H���_��������΍쐬
    folderPath = folderPath & "\" & targetYear & "." & Me.cmbMonth.Value
    
    If fso.FolderExists(folderPath) = False Then
        fso.CreateFolder (folderPath)
    End If
    
    '// �ۑ���t�H���_:�u�ݒ�V�[�g�ɓ��͂���Ă���t�H���_\�Ώ۔N.�Ώی�\���ԊJ�n��-���ԏI�����v
    '// ��)�ݒ�V�[�g�ɓ��͂���Ă���t�H���_\2022.10��\1��-15��
    folderPath = folderPath & "\" & Me.cmbStartDay.Value & "��-" & Me.cmbEndDay.Value & "��"
    
    '// �ۑ��悪������΍쐬
    If fso.FolderExists(folderPath) = False Then
        fso.CreateFolder folderPath
    End If
    
    Set fso = Nothing
    
    createFolderPath = folderPath
    
End Function

'// �t�H�[���N�����̏���
Private Sub UserForm_Initialize()

    Me.cmbCompany.AddItem "�R�݉^����"
    Me.cmbCompany.AddItem "��YCL"
    
    '/**
     '* �Ώی��R���{�{�b�N�X�̑I�����ǉ��Ə����l�ݒ�
    '**/
    Dim i As Long
    
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
    
    Me.txtFileName.Locked = True

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
    
    Dim targetFile As String: targetFile = selectFile("���H����t�@�C����I�����Ă��������B", wsh.SpecialFolders(4) & "\", "Excel�ECSV", "*.xlsx;*.csv")

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
