VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChatWorkController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// チャットワークを操作するクラス
Option Explicit

'/**
 '* チャットにメッセージのみを送信
'**/
Public Function sendMessage(ByVal message As String, ByVal roomId, ByVal apiToken As String) As Boolean

    Dim roomUrl As String: roomUrl = "https://api.chatwork.com/v2/rooms/" & roomId & "/messages"

    Dim xmlHttp As New XMLHTTP60
    
    With xmlHttp
        .Open "POST", roomUrl, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .setRequestHeader "X-ChatWorkToken", apiToken
        .send "body=" & message
    End With
    
    '// 通知の実行結果を取得
    If InStr(xmlHttp.responseText, "message_id") > 0 Then
        sendMessage = True
    Else
        sendMessage = False
    End If
    
    Set xmlHttp = Nothing
    
End Function

'/**
 '* チャットに送信するメッセージを作成
'**/
Public Function createChatWorkText(ByVal mentions As Variant, ByVal title As String, ByVal message As String) As String

    Dim i As Long
    
    Dim mention As String
    
    '// 宛先作成
    For i = 0 To UBound(mentions)
        If mention = "" Then
            mention = "[To:" & Split(mentions(i), ":")(0) & "]" & Split(mentions(i), ":")(1) & "さん"
        Else
            mention = mention & vbLf & "[To:" & Split(mentions(i), ":")(0) & "]" & Split(mentions(i), ":")(1) & "さん"
        End If
    Next
    
    Dim text As String
    
    '// メッセージが空白の場合
    If message = "" Then
        text = mention
    Else
        text = mention & "[info][title]" & title & "[/title]" & message & "[/info]"
    End If
    
    createChatWorkText = text

End Function

'==============================================================================================================================
'/**
 '* チャットにファイルとメッセージを送信
'**/
Public Function sendMessageWithFile(ByVal message As String, ByVal filePath As String, ByVal roomId As String, ByVal apiToken As String) As Boolean
    
    '// HTTPリクエストで使用するデータの境界線(boundary)
    Dim httpBoundary As String: httpBoundary = createBoundary()
    
    Dim myStream As New ADODB.stream
    myStream.Open
    
    '// 添付ファイルをURLエンコード
    Dim fso As New FileSystemObject
    
    Dim fileUrl As String: fileUrl = WorksheetFunction.EncodeURL(fso.GetFileName(filePath))
    Set fso = Nothing
    
    '/**
     '* HTTPのリクエストボディ作成
    '**/
    
    '// 添付ファイルのリクエスト作成
    '// createHttpRequestOfFile [書き込むストリーム], [HTTP通信で使用する境界線], [添付ファイル名をURLエンコードしたもの], [添付ファイルパス]
    Call createHttpRequestOfFile(myStream, httpBoundary, fileUrl, filePath)
        
    '// メッセージ部分のリクエスト作成
    If message <> "" Then
        '// createHttpRequestOfMessage [書き込むストリーム], [HTTP通信で使用する境界線], [送信するメッセージ]
        Call createHttpRequestOfMessage(myStream, httpBoundary, message)
    End If
    
    '// HTTPリクエストの終了部分を作成
    '// createHttpFooter [書き込むストリーム], [Http通信で使用する境界線]
    Call createHttpFooter(myStream, httpBoundary)
    
    myStream.Position = 0
    myStream.Type = adTypeBinary
    
    '// HTTP通信をするurl
    Dim roomUrl As String: roomUrl = "https://api.chatwork.com/v2/rooms/" & roomId & "/files"
    
    '/**
     '* HTTPリクエスト実行
    '**/
    Dim xmlHttp As New XMLHTTP60

    With xmlHttp
        .Open "POST", roomUrl, False
        .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & httpBoundary
        .setRequestHeader "X-ChatWorkToken", apiToken
        .send myStream.Read()
    End With
    
    '// 通知の実行結果を取得
    If InStr(xmlHttp.responseText, "file_id") > 0 Then
        sendMessageWithFile = True
    Else
        sendMessageWithFile = False
    End If
    
    Set myStream = Nothing
    Set xmlHttp = Nothing

End Function

'// HTTPリクエストとして渡すためのデータの境界(boundary)を設定
Private Function createBoundary() As String
     
    '// HTTPリクエストで使用するデータの境界を作成
    Dim multipartChars As String: multipartChars = "-_1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim boundary As String: boundary = "--------------------"
 
    Dim i As Long
    Dim point As Long

    For i = 1 To 16
        Randomize
        point = Int(Len(multipartChars) * Rnd + 1)
        boundary = boundary + Mid(multipartChars, point, 1)
    Next

    createBoundary = boundary
 
End Function

'// 添付ファイルのHTTPリクエスト作成
Private Sub createHttpRequestOfFile(ByRef myStream As ADODB.stream, ByVal boundary As String, ByVal fileUrl As String, ByVal filePath)

    '// ストリームのキャラセットとタイプ変更
    Call changeCharsetAndType(myStream, adTypeText, "shift_jis")
    
    '// ヘッダー部分作成
    Dim httpRequest As String
    
    httpRequest = "--" & boundary & vbCrLf _
            & "Content-Disposition: form-data; name=""file""; filename*=utf-8''" & fileUrl & vbCrLf _
            & "Content-Type:application/octet-stream" & vbCrLf & vbCrLf
 
    myStream.WriteText httpRequest
 
    '/**
     '* 添付ファイルをバイナリデータ化
    '**/
    
    changeCharsetAndType myStream, adTypeBinary
 
    '// 新しいストリームに添付ファイルを読み込み、読み込んだ内容を元のストリームに追加する
    Dim secondStream As New ADODB.stream
    secondStream.Type = adTypeBinary
    secondStream.Open
    secondStream.LoadFromFile filePath

    myStream.Write secondStream.Read()
 
    secondStream.Close
    Set secondStream = Nothing
     
End Sub

'// メッセージのHTTPリクエスト作成
Private Sub createHttpRequestOfMessage(ByRef myStream As ADODB.stream, ByVal boundary As String, ByVal message As String)
      
    changeCharsetAndType myStream, adTypeText, "UTF-8"
    
    Dim httpRequest As String
    
    httpRequest = vbCrLf & "--" & boundary & vbCrLf _
                & "Content-Disposition: form-data; name=""message""" + vbCrLf + vbCrLf _
                & message + vbCrLf

    myStream.WriteText httpRequest

End Sub

'// HTTPリクエストの終了部分を作る
Private Function createHttpFooter(ByRef myStream As ADODB.stream, ByVal boundary As String) As Boolean
    
    changeCharsetAndType myStream, adTypeText, "shift_jis"
    myStream.WriteText vbCrLf & "--" & boundary & "--" & vbCrLf
 
End Function

'/**
 '* データの文字コードとタイプを変更
 '* @params stream データを書き込むストリーム
 '* @params adType targetStreamのタイプ(テキストかバイナリか)
 '* @params char   変更する文字コード
'**/
Private Sub changeCharsetAndType(ByRef targetStream As ADODB.stream, ByVal adType As Long, Optional ByVal char As String)
    
    '// ストリームのタイプや文字コードはポジションを0にしないと変更できないため
    '// 現在のポジションを変数に格納し、タイプと文字コードを変更後もとに戻す
    Dim currentPosition As Long: currentPosition = targetStream.Position
    
    targetStream.Position = 0
    
    targetStream.Type = adType
    
    If char <> "" Then
        targetStream.Charset = char
    End If
    
    targetStream.Position = currentPosition
 
End Sub
