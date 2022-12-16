VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUnpunched 
   Caption         =   "���ō��Ғ��o"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4200
   OleObjectBlob   =   "frmUnpunched.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmUnpunched"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ���ō��Ғ��o�p�t�H�[��
Option Explicit

'// ���s�������ꂽ���̏���(���C���v���V�[�W��)
Private Sub cmdEnter_Click()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '// �o���f�[�V����
    If validate = False Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    '// �w��̃t�@�C�����K�؂Ȃ��̂��m�F
    If validateFile = False Then
        If MsgBox("�w�肵���t�@�C�����s�K�؂ȉ\��������܂��B" & vbLf & "�����𑱍s���܂���?", vbYesNo + vbQuestion, "���ō��Ғ��o") = vbNo Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End If
    
    Me.Hide
    
    '// �p�[�g�Ζ��̌n�R�[�h�̘A�z�z��쐬
    Dim targetColumn As Long
    
    If Me.cmbCompany.Value = "�R�݉^����" Then
        targetColumn = 2
    ElseIf Me.cmbCompany.Value = "��YCL" Then
        targetColumn = 3
    End If
    
    Dim partTimerCodes As Dictionary: Set partTimerCodes = createDictionary(8, targetColumn)
    
    '// �x�����R�R�[�h�̘A�z�z��쐬
    If Me.cmbCompany.Value = "�R�݉^����" Then
        targetColumn = 5
    ElseIf Me.cmbCompany.Value = "��YCL" Then
        targetColumn = 6
    End If
    
    Dim holidayCodes As Dictionary: Set holidayCodes = createDictionary(8, targetColumn)
    
    '// ���ō��Ғ��o
    Call checkAllPunches(Me.txtHiddenFileFullPath.Value, partTimerCodes, holidayCodes)

    Set partTimerCodes = Nothing
    Set holidayCodes = Nothing

    MsgBox "�������������܂����B", vbInformation, ThisWorkbook.Name
    Unload Me

End Sub

'// ���͂��ꂽ�l�̃o���f�[�V����
Private Function validate() As Boolean

    validate = False

    '// ��Ж����I������Ă��邩
    If Me.cmbCompany.Value = "" Then
        MsgBox "��Ж���I�����Ă��������B", vbQuestion, "���ō��Ғ��o"
        Exit Function
    
    '// ���H�t�@�C���������͂���Ă��邩
    ElseIf Me.txtFileName.Value = "" Then
        MsgBox "���H����t�@�C����I�����Ă��������B", vbQuestion, "���ō��Ғ��o"
        Exit Function
    End If

    validate = True
    
End Function

'// �w�肳�ꂽ�t�@�C�����K�؂����f
Private Function validateFile() As Boolean

    validateFile = False

    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(Me.txtHiddenFileFullPath.Value, ReadOnly:=True)
    
    With targetFile.Sheets(1)
        If .Cells(6, 1).Value = "���t" _
            And .Cells(6, 2).Value = "�j" _
            And .Cells(6, 3).Value = "�Ζ��̌n" _
            And .Cells(6, 5).Value = "���R" _
            And .Cells(6, 7).Value = "�o�Ύ���" _
            And .Cells(6, 8).Value = "�ޏo����" _
            And .Cells(6, 9).Value = "�o�Ύ���" Then
            
            validateFile = True
        End If
    End With
    
    targetFile.Close
    
    Set targetFile = Nothing

End Function

'// �Z���̒l���L�[�Ƃ��Ċi�[�����A�z�z����쐬
Private Function createDictionary(ByVal startRow As Long, ByVal targetColumn As Long) As Dictionary

    Dim returnDic As New Dictionary
    
    Dim i As Long
    
    '// ����̓L�[��������Ηǂ��̂Œl�͋�(Exists���\�b�h���g�p���邽��Dictionary�^�̔z����쐬)
    For i = startRow To Sheets("�ݒ�").Cells(Rows.Count, targetColumn).End(xlUp).Row
        If Sheets("�ݒ�").Cells(i, targetColumn).Value = "" Then
            GoTo Continue
        End If
    
        returnDic.Add Sheets("�ݒ�").Cells(i, targetColumn).Value, ""
Continue:
    Next
    
    Set createDictionary = returnDic

    Set returnDic = Nothing
    
End Function

'/**
 '* ���ō��҂𒊏o
'**/
Private Sub checkAllPunches(ByVal targetFileName As String, ByVal partTimerCodes As Dictionary, ByVal holidayCodes As Dictionary)

    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(targetFileName)
    
    Dim targetSheet As Worksheet
    
    '/**
     '* ���ō��̊m�F
    '**/
    For Each targetSheet In targetFile.Worksheets
        targetSheet.Activate
        
        '// checkPunches [�p�[�g�Ζ��̌n�R�[�h], [�x�����R�R�[�h]
        Call checkPunches(partTimerCodes, holidayCodes)
    Next

    Dim i As Long
    
    '/**
     '* ���ō����Ȃ��l�̃V�[�g���폜
    '**/
    For i = targetFile.Worksheets.Count To 1 Step -1
        
        '// �V�[�g���̐擪��OK�łȂ��ꍇ(���ō�����̏ꍇ)
        If Left(targetFile.Sheets(i).Name, 2) <> "OK" Then
            GoTo Continue
        End If
        
        '// �V�[�g���̐擪��OK�̏ꍇ(���ō����Ȃ��ꍇ)
        If Left(targetFile.Sheets(i).Name, 2) = "OK" Then
            
            '// �V�[�g������1����葽���ꍇ �� �V�[�g�폜
            If targetFile.Sheets.Count > 1 Then
                targetFile.Sheets(i).Delete
            
            '// �V�[�g������1���̏ꍇ �� �V�[�g�Ɂu���ō��͂���܂���ł����B�v�ƕ\��
            Else
                targetFile.Sheets(i).Cells.Clear
                targetFile.Sheets(i).Name = "���ō��҂Ȃ�"
                targetFile.Sheets(i).Cells(5, 5).Value = "���ō��͂���܂���ł����B"
                targetFile.Sheets(i).Cells(5, 5).Font.Size = 26
            End If
        End If
        
Continue:
    Next
    
    targetFile.Sheets(1).Activate
    Cells(1, 1).Select
    
    Set targetFile = Nothing
    
End Sub

'/**
 '* ���ō����Ȃ������m�F(1�l��)
'**/
Private Sub checkPunches(ByVal partTimerCodes As Dictionary, ByVal holidayCodes As Dictionary)

    Dim lastRow As Long: lastRow = WorksheetFunction.Match("���v", Columns(1), 0) - 1
    
    Dim i As Long
    
    For i = 6 To lastRow
    
        '// ���ō��ł͂Ȃ��ꍇ
        If Cells(i, 7).Value <> "" And Cells(i, 8).Value <> "" Then
            GoTo Continue
    
        '// �ދ΂����ō��̏ꍇ
        ElseIf Cells(i, 7).Value <> "" And Cells(i, 8).Value = "" Then
        
            '// �T��̊��Ԃ̍ŏI���̏ꍇ �� ��荞�񂾎��Ԃ̊֌W�Ŗ��ō��̂悤�ɂȂ��Ă���ꍇ�����邽��King of Time���m�F���Ă��炤
            If i = lastRow Then
                Cells(i, 5).Value = "�v�m�F"
                Cells(i, 8).Interior.Color = vbBlue
            Else
                Cells(i, 5).Value = "���ō�����"
                Cells(i, 8).Interior.Color = vbYellow
            End If
        
        '// �o�΂����ō��̏ꍇ
        ElseIf Cells(i, 8).Value <> "" And Cells(i, 7).Value = "" Then
            Cells(i, 5).Value = "���ō�����"
            Cells(i, 7).Interior.Color = vbYellow
    
        '// �Ζ��̌n���󔒁A�������̓p�[�g���x���o�΂̏ꍇ
        ElseIf Cells(i, 3).Value = "" Or partTimerCodes.Exists(Val(Cells(i, 3).Value)) = True Then
            GoTo Continue
        
        '// ���R�R�[�h���x���̏ꍇ
        ElseIf holidayCodes.Exists(Val(Cells(i, 5).Value)) = True Then
            GoTo Continue
    
        '// �p�[�g�ł��x���ł��Ȃ��o�ދ΂Ƃ��ɖ��ō��̏ꍇ
        Else
            Cells(i, 5).Value = "���ō�����"
            Range(Cells(i, 7), Cells(i, 8)).Interior.Color = vbYellow
        End If
Continue:
    Next
    
    '// �s�v�ȗ�(�o�Ύ��ԗ���E)���폜
    Range(Columns(10), Columns(Cells(6, Columns.Count).End(xlToLeft).Column)).Delete
    Columns(5).AutoFit
    
    '// ���ō��ł͂Ȃ��ō��f�[�^���폜
    Range(Cells(6, 1), Cells(6, Columns.Count).End(xlToLeft)).AutoFilter 5, "<>���ō�����", xlAnd, "<>�v�m�F"
    Cells(1, 1).CurrentRegion.Offset(6).Delete
    
    Cells(6, 1).AutoFilter
    
    '// ���ō���������΃V�[�g���̐擪��OK�Ƃ���
    If WorksheetFunction.CountIf(Columns(5), "���ō�����") + WorksheetFunction.CountIf(Columns(5), "�v�m�F") = 0 Then
        ActiveSheet.Name = "OK" & Left(ActiveSheet.Name, 5)
    End If
    
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

'// �t�H�[���N�����̏���
Private Sub UserForm_Initialize()

    cmbCompany.AddItem "�R�݉^����"
    cmbCompany.AddItem "��YCL"
    
    txtFileName.Locked = True
    
End Sub

'// ������������Ƃ��̏���
Private Sub cmdClose_Click()
    
    Unload Me

End Sub
