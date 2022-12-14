VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPDF 
   Caption         =   "PDF出力"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6480
   OleObjectBlob   =   "frmPDF.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// PDF出力するためのフォーム
Option Explicit

'// 実行を押したときの処理(メインプロシージャ)
Private Sub cmdEnter_Click()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '// バリデーション
    If validate = False Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(Me.txtHiddenFileFullPath)
    Dim targetSheet As Worksheet
    
    '// 指定のファイルの全シートの印刷を設定
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
    
    '// ExportAsFixedFormat [ファイルタイプ], [ファイル名]
    targetFile.ExportAsFixedFormat xlTypePDF, wsh.SpecialFolders(4) & "\" & Me.cmbCompany.Value & Me.cmbDepartment.Value & " " & Me.cmbMonth.Value & Me.cmbStartDay & "～" & Me.cmbMonth.Value & Me.cmbEndDay.Value & "日.pdf"

    targetFile.Close False
    Set targetFile = Nothing
    
    Me.cmbDepartment.Value = ""
    Me.txtHiddenFileFullPath.Value = ""
    
    With Me.txtFileName
        .Locked = False
        .Value = ""
        .Locked = True
    End With
    
    MsgBox "PDF出力が完了しました。", vbQuestion, "PDF出力"

End Sub

'// 入力された値のバリデーション
Private Function validate() As Boolean
    
    validate = False
    
    '// 会社名が選択されているか
    If Me.cmbCompany.Value = "" Then
        MsgBox "会社名を選択してください。", vbQuestion, "PDF出力"
        Exit Function
    
    '// 部署名が選択されているか
    ElseIf Me.cmbDepartment.Value = "" Then
        MsgBox "部署名を選択してください。"
        Exit Function
    
    '// 対象月が選択されているか
    ElseIf Me.cmbMonth.Value = "" Then
        MsgBox "対象月を選択してください。"
        Exit Function
    
    '// 期間が選択されているか
    ElseIf Me.cmbStartDay.Value = "" Or Me.cmbEndDay.Value = "" Then
        MsgBox "期間を選択してください。"
        Exit Function

    '// 期間の終了日が開始日より前に来ていないか
    ElseIf Me.cmbStartDay.Value > Me.cmbEndDay.Value Then
        MsgBox "期間の終了日は開始日より後に設定してください。"
        Exit Function
    End If
    
    validate = True
    
End Function

'// フォーム起動時の処理
Private Sub UserForm_Initialize()

    Me.cmbCompany.AddItem "山岸運送㈱"
    Me.cmbCompany.AddItem "㈱YCL"
    
    Dim i As Long
    
    '/**
     '* 対象月コンボボックスの選択肢追加と初期値設定
    '**/
    For i = 4 To 12
        Me.cmbMonth.AddItem i & "月"
    Next
    For i = 1 To 3
        Me.cmbMonth.AddItem i & "月"
    Next
    
    If Day(Now) < 16 Then
        Me.cmbMonth.Value = Month(Now) - 1 & "月"
    Else
        Me.cmbMonth.Value = Month(Now) & "月"
    End If
    
    '/**
     '* 期間コンボボックスの初期値設定
    '**/
    Dim endOfMonth As Long: endOfMonth = Day(DateSerial(Year(Now), Replace(Me.cmbMonth.Value, "月", "") + 1, 0))
    
    '// 処理日が1日から15日の間の場合
    If Day(Now) < 16 Then
        Me.cmbStartDay.Value = 26
        Me.cmbEndDay.Value = endOfMonth
        
    '// 処理日が16日から25日の場合
    ElseIf 15 < Day(Now) And Day(Now) < 26 Then
        Me.cmbStartDay.Value = 1
        Me.cmbEndDay.Value = 15
        
    '// 処理日が26日以降の場合
    ElseIf 25 < Day(Now) Then
        Me.cmbStartDay.Value = 16
        Me.cmbEndDay.Value = 25
    End If
    
End Sub

'// 会社名が変更された場合
Private Sub cmbCompany_Change()

    If Me.cmbCompany.Value = "" Then
        Exit Sub
    End If
    
    Me.cmbDepartment.Clear
    
    Dim targetColumn As Long
    
    If Me.cmbCompany.Value = "山岸運送㈱" Then
        targetColumn = 8
    ElseIf Me.cmbCompany.Value = "㈱YCL" Then
        targetColumn = 9
    End If
    
    Dim i As Long
    
    '// 部署コンボボックスの選択肢を会社に合わせて変更
    For i = 8 To Sheets("設定").Cells(Rows.Count, targetColumn).End(xlUp).Row
        If Sheets("設定").Cells(i, targetColumn).Value = "" Then
            GoTo Continue
        End If
        
        Me.cmbDepartment.AddItem Sheets("設定").Cells(i, targetColumn).Value
Continue:
    Next
    
End Sub

'// 対象月が変更された時の処理
Private Sub cmbMonth_Change()

    If Me.cmbMonth.Value = "" Then
        Exit Sub
    End If
    
    Dim previousStartDay As Variant: previousStartDay = Me.cmbStartDay.Value
    Dim previousEndDay As Variant: previousEndDay = Me.cmbEndDay.Value
    
    Me.cmbStartDay.Clear
    Me.cmbEndDay.Clear
    
    Dim endOfMonth As Long: endOfMonth = Day(DateSerial(Year(Now), Replace(Me.cmbMonth.Value, "月", "") + 1, 0))
    
    Dim i As Long
    
    '// 期間コンボボックスの選択肢を月に合わせて変更(月によって日数が変わるため)
    For i = 1 To endOfMonth
        Me.cmbStartDay.AddItem i
        Me.cmbEndDay.AddItem i
    Next

    '// 期間コンボボックスの値が対象月の月末より大きい数字の場合 → 対象月の月末日を期間コンボボックスの値とする
    If endOfMonth < previousStartDay Then
        previousStartDay = endOfMonth
    End If
    
    If endOfMonth < previousEndDay Then
        previousEndDay = endOfMonth
    End If

    Me.cmbStartDay.Value = previousStartDay
    Me.cmbEndDay.Value = previousEndDay

End Sub

'// 参照を押したときの処理
Private Sub cmdDialog_Click()

    Dim wsh As New WshShell
    Dim fso As New FileSystemObject
    
    Dim targetFile As String: targetFile = selectFile("PDF出力するファイルを選択してください。", wsh.SpecialFolders(4) & "\", "Excel・CSV", "*.xlsx;*.csv")

    Me.txtFileName.Locked = False
    
    If targetFile <> "" Then
        Me.txtFileName.Value = fso.GetFileName(targetFile)
        Me.txtHiddenFileFullPath.Value = targetFile
    End If
    
    Me.txtFileName.Locked = True
    
    Set wsh = Nothing
    Set fso = Nothing

End Sub

'// 閉じるを押したときの処理
Private Sub cmdClose_Click()

    Unload Me

End Sub
