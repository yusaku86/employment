VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChatwork 
   Caption         =   "送信内容入力"
   ClientHeight    =   8535.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15840
   OleObjectBlob   =   "frmChatwork.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmChatwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// チャット送信のためのフォーム
Option Explicit

'/**
 '* メインプログラム(チャットで送る内容設定&チャット送信)
'**/
Private Sub cmdEnter_Click()

    '// バリデーション
    If validate = False Then: Exit Sub
    
    '// 送信先グループID
    Dim roomId As String: roomId = Split(cmbRoom.Value, ":")(0)
    
    '// APIトークン
    Dim apiToken As String: apiToken = Sheets("チャット設定").Cells(7, 4).Value
    
    '/**
     '* チャット送信
    '**/
    Dim cc As New ChatWorkController
    
    '// チャット送信用の文章作成  createChatWorkText [メンションリスト], [タイトル], [送信メッセージ]
    Dim message As String: message = cc.createChatWorkText(Split(Me.txtMentionList.Value, vbCrLf), "未打刻一覧:" & Me.cmbCompany.Value & " " & Me.cmbDepartment.Value, Me.txtMessage.Value)
    
    Dim result As Boolean
 
    '// メッセージのみ送信する場合
    If Me.txtFile.Value = "" Then
        result = cc.sendMessage(message, roomId, apiToken)
    
    '// ファイルとメッセージを送信する場合
    Else
        result = cc.sendMessageWithFile(message, Me.txtHiddenFileFullPath.Value, roomId, apiToken)
    End If
    
    If result = True Then
        MsgBox "送信が完了しました。", vbInformation, ThisWorkbook.Name
    Else
        MsgBox "送信できませんでした。", vbExclamation, ThisWorkbook.Name
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
 '* バリデーション
'**/
Private Function validate() As Boolean

    validate = False

    '// 会社名が入力されているか
    If cmbCompany.Value = "" Then
        MsgBox "会社名を選択してください。", vbQuestion, "チャット送信"
        cmbCompany.SetFocus
        Exit Function
    End If
    
    '// 部署名が選択されているか
    If cmbDepartment.Value = "" Then
        MsgBox "部署名を選択してください。", vbQuestion, "チャット送信"
        cmbDepartment.SetFocus
        Exit Function
    End If
    
    '// 送信先グループが選択されているか
    If cmbRoom.Value = "" Then
        MsgBox "送信先グループを選択してください。", vbQuestion, "チャット送信"
        cmbRoom.SetFocus
        Exit Function
    End If
    
    '// 宛先が設定されているか
    If txtMentionList.Value = "" Then
        If MsgBox("宛先が設定されていませんが、送信してよろしいですか?", vbQuestion + vbYesNo, "チャット送信") = vbNo Then
            Exit Function
        End If
    End If
        
    '// メッセージが入力されているか
    If txtMessage.Value = "" Then
        If MsgBox("メッセージが入力されていませんが、送信してよろしいですか?", vbQuestion + vbYesNo, "チャット送信") = vbNo Then
            Exit Function
        End If
    End If
    
    '// 添付ファイルが選択されているか
    If txtFile.Value = "" Then
        If MsgBox("ファイルが添付されていませんが、送信してよろしいですか?", vbQuestion + vbYesNo, "チャット送信") = vbNo Then
            Exit Function
        End If
    End If
    
    validate = True
    
End Function

'/**
 '* ユーザーフォーム起動時の設定
'**/
Private Sub UserForm_Initialize()
                
    '// チャットグループ名の選択肢追加
    Dim i As Long
   
    For i = 7 To ThisWorkbook.Sheets("チャット設定").Cells(Rows.Count, 5).End(xlUp).Row
        cmbRoom.AddItem ThisWorkbook.Sheets("チャット設定").Cells(i, 5).Value
    Next
    
    '// 送信相手リスト追加
    For i = 7 To ThisWorkbook.Sheets("チャット設定").Cells(Rows.Count, 6).End(xlUp).Row
        cmbMention.AddItem ThisWorkbook.Sheets("チャット設定").Cells(i, 6).Value
    Next
    
    '// 会社名の選択肢追加
    cmbCompany.AddItem "山岸運送㈱"
    cmbCompany.AddItem "㈱YCL"
    
    '// 対象月コンボボックスの選択肢追加と初期値設定
    For i = 1 To 12
        Me.cmbMonth.AddItem i & "月"
    Next
    
    If Day(Now) < 16 Then
        Me.cmbMonth.Value = Month(Now) - 1 & "月"
    Else
        Me.cmbMonth.Value = Month(Now)
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
    
    Call setMessage

End Sub

'/**
 '* 期間が変更されたときの処理 → 送信するメッセージを変更
'**/
Private Sub cmbEndDay_Change()

    Call setMessage

End Sub

Private Sub cmbStartDay_Change()
    
    Call setMessage

End Sub

'/**
 '* 会社名の値が変更された時の処理
'**/
Private Sub cmbCompany_Change()

    Me.cmbDepartment.Clear
    
    Dim targetColumn As Long
    
    If Me.cmbCompany.Value = "山岸運送㈱" Then
        targetColumn = 8
    ElseIf Me.cmbCompany.Value = "㈱YCL" Then
        targetColumn = 9
    End If
    
    Dim i As Long
    
    '// 部署名コンボボックスの選択肢追加
    For i = 8 To Sheets("設定").Cells(Rows.Count, targetColumn).End(xlUp).Row
        If Sheets("設定").Cells(i, targetColumn).Value = "" Then
            GoTo Continue
        End If
        
        Me.cmbDepartment.AddItem Sheets("設定").Cells(i, targetColumn).Value
Continue:
    Next
    
End Sub

'/**
 '* チャットに送信するメッセージ設定
'**/
Private Sub setMessage()

    If (Me.cmbMonth.Value <> "" And Me.cmbStartDay.Value <> "" And Me.cmbEndDay.Value <> "") = False Then
        Exit Sub
    End If
    
    Me.txtMessage.Value = "お疲れ様です。" & _
                          vbLf & Me.cmbMonth.Value & Me.cmbStartDay.Value & "日～" & Me.cmbMonth.Value & Me.cmbEndDay.Value & "日までの未打刻です。" _
                        & vbLf & setDeadline() & "までに必ず回答お願いします。"

End Sub

'/**
 '* 回答の締め切り日を設定
'**/
Private Function setDeadline() As String

    Dim finished As Boolean: finished = False
    Dim i As Long: i = 2
    
    '// 処理日の2日後が基本の締め切りだが、土日の場合は平日になるまで締めきりを伸ばす
    Do Until finished = True
        If Weekday(DateSerial(Year(Now), Month(Now), Day(Now) + i)) <> 0 And Weekday(DateSerial(Year(Now), Month(Now), Day(Now) + i)) <> 6 Then
            finished = True
            Exit Do
        End If
        
        i = i + 1
    Loop

    setDeadline = Format(DateSerial(Year(Now), Month(Now), Day(Now) + i), "m月d日(aaa)")

End Function

'/**
 '* 追加を押したときの処理
'**/
Private Sub cmdAdd_Click()
    
    If cmbMention.Value = "" Then
        Exit Sub
    End If
    
    '// 既に宛先が追加されていたら抜ける
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

'// リセットを押したときの処理
Private Sub cmdReset_Click()

    txtMentionList.Value = ""

End Sub

'// 参照を押したときの処理
Private Sub cmdDialog_Click()

    Dim wsh As Object: Set wsh = CreateObject("Wscript.Shell")
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    Dim attachedFileName As String: attachedFileName = selectFile("添付ファイル選択", wsh.SpecialFolders(4) & "\", "Excelファイル・PDF", "*.xlsx;*.pdf;*.csv")
    
    If attachedFileName <> "" Then
        txtFile.Locked = False
    
        txtFile.Value = fso.GetFileName(attachedFileName)
        txtHiddenFileFullPath.Value = attachedFileName
    
        txtFile.Locked = True
    End If
    
    Set wsh = Nothing
    Set fso = Nothing
    
End Sub

'// ファイルをリセットを押した時の処理
Private Sub cmdClearFile_Click()

    Me.txtFile.Locked = False
    
    Me.txtFile.Value = ""
    Me.txtHiddenFileFullPath.Value = ""
    
    Me.txtFile.Locked = True

End Sub

'// 閉じるを押したときの処理
Private Sub cmdCancel_Click()
    
    Unload Me

End Sub
