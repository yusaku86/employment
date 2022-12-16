VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUnpunched 
   Caption         =   "未打刻者抽出"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4200
   OleObjectBlob   =   "frmUnpunched.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmUnpunched"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// 未打刻者抽出用フォーム
Option Explicit

'// 実行が押された時の処理(メインプロシージャ)
Private Sub cmdEnter_Click()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '// バリデーション
    If validate = False Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    '// 指定のファイルが適切なものか確認
    If validateFile = False Then
        If MsgBox("指定したファイルが不適切な可能性があります。" & vbLf & "処理を続行しますか?", vbYesNo + vbQuestion, "未打刻者抽出") = vbNo Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End If
    
    Me.Hide
    
    '// パート勤務体系コードの連想配列作成
    Dim targetColumn As Long
    
    If Me.cmbCompany.Value = "山岸運送㈱" Then
        targetColumn = 2
    ElseIf Me.cmbCompany.Value = "㈱YCL" Then
        targetColumn = 3
    End If
    
    Dim partTimerCodes As Dictionary: Set partTimerCodes = createDictionary(8, targetColumn)
    
    '// 休日事由コードの連想配列作成
    If Me.cmbCompany.Value = "山岸運送㈱" Then
        targetColumn = 5
    ElseIf Me.cmbCompany.Value = "㈱YCL" Then
        targetColumn = 6
    End If
    
    Dim holidayCodes As Dictionary: Set holidayCodes = createDictionary(8, targetColumn)
    
    '// 未打刻者抽出
    Call checkAllPunches(Me.txtHiddenFileFullPath.Value, partTimerCodes, holidayCodes)

    Set partTimerCodes = Nothing
    Set holidayCodes = Nothing

    MsgBox "処理が完了しました。", vbInformation, ThisWorkbook.Name
    Unload Me

End Sub

'// 入力された値のバリデーション
Private Function validate() As Boolean

    validate = False

    '// 会社名が選択されているか
    If Me.cmbCompany.Value = "" Then
        MsgBox "会社名を選択してください。", vbQuestion, "未打刻者抽出"
        Exit Function
    
    '// 加工ファイル名が入力されているか
    ElseIf Me.txtFileName.Value = "" Then
        MsgBox "加工するファイルを選択してください。", vbQuestion, "未打刻者抽出"
        Exit Function
    End If

    validate = True
    
End Function

'// 指定されたファイルが適切か判断
Private Function validateFile() As Boolean

    validateFile = False

    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(Me.txtHiddenFileFullPath.Value, ReadOnly:=True)
    
    With targetFile.Sheets(1)
        If .Cells(6, 1).Value = "日付" _
            And .Cells(6, 2).Value = "曜" _
            And .Cells(6, 3).Value = "勤務体系" _
            And .Cells(6, 5).Value = "事由" _
            And .Cells(6, 7).Value = "出勤時刻" _
            And .Cells(6, 8).Value = "退出時刻" _
            And .Cells(6, 9).Value = "出勤時間" Then
            
            validateFile = True
        End If
    End With
    
    targetFile.Close
    
    Set targetFile = Nothing

End Function

'// セルの値をキーとして格納した連想配列を作成
Private Function createDictionary(ByVal startRow As Long, ByVal targetColumn As Long) As Dictionary

    Dim returnDic As New Dictionary
    
    Dim i As Long
    
    '// 今回はキーだけあれば良いので値は空白(Existsメソッドを使用するためDictionary型の配列を作成)
    For i = startRow To Sheets("設定").Cells(Rows.Count, targetColumn).End(xlUp).Row
        If Sheets("設定").Cells(i, targetColumn).Value = "" Then
            GoTo Continue
        End If
    
        returnDic.Add Sheets("設定").Cells(i, targetColumn).Value, ""
Continue:
    Next
    
    Set createDictionary = returnDic

    Set returnDic = Nothing
    
End Function

'/**
 '* 未打刻者を抽出
'**/
Private Sub checkAllPunches(ByVal targetFileName As String, ByVal partTimerCodes As Dictionary, ByVal holidayCodes As Dictionary)

    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(targetFileName)
    
    Dim targetSheet As Worksheet
    
    '/**
     '* 未打刻の確認
    '**/
    For Each targetSheet In targetFile.Worksheets
        targetSheet.Activate
        
        '// checkPunches [パート勤務体系コード], [休日事由コード]
        Call checkPunches(partTimerCodes, holidayCodes)
    Next

    Dim i As Long
    
    '/**
     '* 未打刻がない人のシートを削除
    '**/
    For i = targetFile.Worksheets.Count To 1 Step -1
        
        '// シート名の先頭がOKでない場合(未打刻ありの場合)
        If Left(targetFile.Sheets(i).Name, 2) <> "OK" Then
            GoTo Continue
        End If
        
        '// シート名の先頭がOKの場合(未打刻がない場合)
        If Left(targetFile.Sheets(i).Name, 2) = "OK" Then
            
            '// シート枚数が1枚より多い場合 → シート削除
            If targetFile.Sheets.Count > 1 Then
                targetFile.Sheets(i).Delete
            
            '// シート枚数が1枚の場合 → シートに「未打刻はありませんでした。」と表示
            Else
                targetFile.Sheets(i).Cells.Clear
                targetFile.Sheets(i).Name = "未打刻者なし"
                targetFile.Sheets(i).Cells(5, 5).Value = "未打刻はありませんでした。"
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
 '* 未打刻がないかを確認(1人分)
'**/
Private Sub checkPunches(ByVal partTimerCodes As Dictionary, ByVal holidayCodes As Dictionary)

    Dim lastRow As Long: lastRow = WorksheetFunction.Match("合計", Columns(1), 0) - 1
    
    Dim i As Long
    
    For i = 6 To lastRow
    
        '// 未打刻ではない場合
        If Cells(i, 7).Value <> "" And Cells(i, 8).Value <> "" Then
            GoTo Continue
    
        '// 退勤が未打刻の場合
        ElseIf Cells(i, 7).Value <> "" And Cells(i, 8).Value = "" Then
        
            '// 週報の期間の最終日の場合 → 取り込んだ時間の関係で未打刻のようになっている場合もあるためKing of Timeを確認してもらう
            If i = lastRow Then
                Cells(i, 5).Value = "要確認"
                Cells(i, 8).Interior.Color = vbBlue
            Else
                Cells(i, 5).Value = "未打刻あり"
                Cells(i, 8).Interior.Color = vbYellow
            End If
        
        '// 出勤が未打刻の場合
        ElseIf Cells(i, 8).Value <> "" And Cells(i, 7).Value = "" Then
            Cells(i, 5).Value = "未打刻あり"
            Cells(i, 7).Interior.Color = vbYellow
    
        '// 勤務体系が空白、もしくはパートか休日出勤の場合
        ElseIf Cells(i, 3).Value = "" Or partTimerCodes.Exists(Val(Cells(i, 3).Value)) = True Then
            GoTo Continue
        
        '// 事由コードが休日の場合
        ElseIf holidayCodes.Exists(Val(Cells(i, 5).Value)) = True Then
            GoTo Continue
    
        '// パートでも休日でもなく出退勤ともに未打刻の場合
        Else
            Cells(i, 5).Value = "未打刻あり"
            Range(Cells(i, 7), Cells(i, 8)).Interior.Color = vbYellow
        End If
Continue:
    Next
    
    '// 不要な列(出勤時間列より右)を削除
    Range(Columns(10), Columns(Cells(6, Columns.Count).End(xlToLeft).Column)).Delete
    Columns(5).AutoFit
    
    '// 未打刻ではない打刻データを削除
    Range(Cells(6, 1), Cells(6, Columns.Count).End(xlToLeft)).AutoFilter 5, "<>未打刻あり", xlAnd, "<>要確認"
    Cells(1, 1).CurrentRegion.Offset(6).Delete
    
    Cells(6, 1).AutoFilter
    
    '// 未打刻が無ければシート名の先頭にOKとつける
    If WorksheetFunction.CountIf(Columns(5), "未打刻あり") + WorksheetFunction.CountIf(Columns(5), "要確認") = 0 Then
        ActiveSheet.Name = "OK" & Left(ActiveSheet.Name, 5)
    End If
    
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

'// フォーム起動時の処理
Private Sub UserForm_Initialize()

    cmbCompany.AddItem "山岸運送㈱"
    cmbCompany.AddItem "㈱YCL"
    
    txtFileName.Locked = True
    
End Sub

'// 閉じるを押したときの処理
Private Sub cmdClose_Click()
    
    Unload Me

End Sub
