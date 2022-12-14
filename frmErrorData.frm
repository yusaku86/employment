VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmErrorData 
   Caption         =   "エラーデータ作成"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   OleObjectBlob   =   "frmErrorData.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmErrorData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// エラーデータを作成するフォーム
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
    
    '// 選択されたファイルが適切なものか確認
    If validateFile = False Then
        If MsgBox("指定したファイルが不適切な可能性があります。" & vbLf & "処理を続行しますか?", vbQuestion + vbYesNo, "エラーデータ作成") = vbNo Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End If
    
    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(Me.txtHiddenFileFullPath.Value)
    
    Dim i As Long
        
    '// エラーデータの内不要なものを削除
    For i = 7 To ThisWorkbook.Sheets("設定").Cells(Rows.Count, 11).End(xlUp).Row
        If ThisWorkbook.Sheets("設定").Cells(i, 11).Value = "" Then
            GoTo Continue
        End If
        
        targetFile.Sheets(1).Cells(1, 1).AutoFilter 1, "*" & ThisWorkbook.Sheets("設定").Cells(i, 11).Value & "*"
        
        If targetFile.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row <> 1 Then
            targetFile.Sheets(1).Cells(1, 1).CurrentRegion.Offset(1).Delete
        End If

        targetFile.Sheets(1).Cells(1, 1).AutoFilter

Continue:
    Next

    '// A列の<ObcUnacceptedReason></ObcUnacceptedReason>を削除
    For i = 1 To targetFile.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
        targetFile.Sheets(1).Cells(i, 1).Value = Replace(Replace(targetFile.Sheets(1).Cells(i, 1).Value, "<ObcUnacceptedReason>", ""), "</ObcUnacceptedReason>", "")
    Next

    '// エラーデータをテーブルとして保存
    Dim errTable As ListObject: Set errTable = targetFile.Sheets(1).ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, Cells(1, Columns.Count).End(xlToLeft).Column)), , xlYes)
    errTable.TableStyle = "TableStyleLight18"
    
    Set errTable = Nothing
    
    targetFile.Sheets(1).Cells.EntireColumn.AutoFit
    
    Dim folderPath As String: folderPath = createFolderPath
    
    '// エラーデータはcsvなのでExcelファイル(xlsx)として保存
    targetFile.SaveAs folderPath & "\" & Me.cmbCompany.Value & " 打刻エラーデータ.xlsx", xlOpenXMLWorkbook
    
    targetFile.Close False
    
    Set targetFile = Nothing
    
    Me.cmbCompany.Value = ""
    Me.txtFileName.Value = ""
    Me.txtHiddenFileFullPath.Value = ""
    
    MsgBox "処理が完了しました。", vbInformation, "エラーデータ作成:" & Me.cmbCompany.Value

End Sub

'// 入力された値と保存先のバリデーション
Private Function validate() As Boolean

    validate = False
    
    '// 会社名が選択されているか
    If Me.cmbCompany.Value = "" Then
        MsgBox "会社名を選択してください。", vbQuestion, "エラーデータ作成"
        Exit Function

    '// 対象月が選択されているか
    ElseIf Me.cmbMonth.Value = "" Then
        MsgBox "対象月を選択してください。", vbQuestion, "エラーデータ作成"
        Exit Function
    
    '// 期間が選択されているか
    ElseIf Me.cmbStartDay.Value = "" Or Me.cmbEndDay.Value = "" Then
        MsgBox "期間を選択してください。", vbQuestion, "エラーデータ作成"
        Exit Function

    '// 期間の終了日が開始日より先になっていないか
    ElseIf Me.cmbEndDay.Value < Me.cmbStartDay.Value Then
        MsgBox "期間の終了日は開始日より後に設定してください。", vbQuestion, "エラーデータ作成"
        Exit Function
    End If
    
    '// エラーデータの保存先フォルダが存在するか
    Dim folderPath As String
    
    If Me.cmbCompany.Value = "山岸運送㈱" Then
        folderPath = Sheets("設定").Cells(8, 13).Value
    ElseIf Me.cmbCompany.Value = "㈱YCL" Then
        folderPath = Sheets("設定").Cells(8, 14).Value
    End If
    
    Dim fso As New FileSystemObject
        
    If folderPath = "" Then
        MsgBox "エラーデータの保存先フォルダが設定されていません。" & vbLf & "シート「設定」で設定してください。", vbQuestion, "エラーデータ作成"
        Exit Function
        
    ElseIf fso.FolderExists(folderPath) = False Then
        MsgBox "エラーデータの保存先フォルダが存在しません。", vbQuestion, "エラーデータ作成"
        Exit Function
    End If
    
    validate = True

End Function

'// 選択されたファイルが適切なものか確認
Private Function validateFile() As Boolean

    validateFile = False

    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(Me.txtHiddenFileFullPath.Value, ReadOnly:=True)
    
    If Cells(1, 1).Value = "<ObcUnacceptedReason> エラー内容 </ObcUnacceptedReason>" _
        And Cells(1, 2).Value = "日付" _
        And Cells(1, 3).Value = "社員番号" Then
            
        validateFile = True
    End If
    
    targetFile.Close False
    Set targetFile = Nothing
        
End Function

'// エラーデータの保存先パス作成
Private Function createFolderPath() As String

    Dim folderPath As String
    
    If Me.cmbCompany.Value = "山岸運送㈱" Then
        folderPath = ThisWorkbook.Sheets("設定").Cells(8, 13).Value
    ElseIf Me.cmbCompany.Value = "㈱YCL" Then
        folderPath = ThisWorkbook.Sheets("設定").Cells(8, 14).Value
    End If
    
    '// 対象年
    Dim targetYear As Long: targetYear = Year(Now)
    
    '// 処理月(このファイルを操作している時の月が1月で、対象月が12月の場合は対象年が処理しているときの年から1マイナスした数になる)
    If Month(Now) = 1 And Me.cmbMonth.Value = 12 Then
        targetYear = Year(Now) - 1
    End If
    
    Dim fso As New FileSystemObject
    
    '// 「設定シートに入力されているフォルダ\対象年.対象月」のフォルダが無ければ作成
    folderPath = folderPath & "\" & targetYear & "." & Me.cmbMonth.Value
    
    If fso.FolderExists(folderPath) = False Then
        fso.CreateFolder (folderPath)
    End If
    
    '// 保存先フォルダ:「設定シートに入力されているフォルダ\対象年.対象月\期間開始日-期間終了日」
    '// 例)設定シートに入力されているフォルダ\2022.10月\1月-15日
    folderPath = folderPath & "\" & Me.cmbStartDay.Value & "日-" & Me.cmbEndDay.Value & "日"
    
    '// 保存先が無ければ作成
    If fso.FolderExists(folderPath) = False Then
        fso.CreateFolder folderPath
    End If
    
    Set fso = Nothing
    
    createFolderPath = folderPath
    
End Function

'// フォーム起動時の処理
Private Sub UserForm_Initialize()

    Me.cmbCompany.AddItem "山岸運送㈱"
    Me.cmbCompany.AddItem "㈱YCL"
    
    '/**
     '* 対象月コンボボックスの選択肢追加と初期値設定
    '**/
    Dim i As Long
    
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
    
    Me.txtFileName.Locked = True

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
    
    Dim targetFile As String: targetFile = selectFile("加工するファイルを選択してください。", wsh.SpecialFolders(4) & "\", "Excel・CSV", "*.xlsx;*.csv")

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
