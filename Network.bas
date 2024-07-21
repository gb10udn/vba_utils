Attribute VB_Name = "Network"
Option Explicit

Public Sub SendPost()  ' # TODO: 240721 名前をもっと一般的にして、vba utils に加えること。
  Dim http As Object
  Dim url As String
  Dim inputData As String
  Dim result As String

  ' MSXML2.XMLHTTPオブジェクトを作成
  Set http = CreateObject("MSXML2.XMLHTTP")
  url = "http://localhost:8000"  ' # TODO: 240721 8000 ポートがない場合のエラーハンドリング。(Msgbox で表示させること。)
  inputData = "{""a"": 3, ""b"": 5}"

  ' HTTP POSTリクエストを送信
  http.Open "POST", url, False
  http.setRequestHeader "Content-Type", "application/json"
  http.send inputData

  ' レスポンスを取得
  result = http.responseText  ' # TODO: 240721 辞書型 or 配列での取得を検討せよ。
  MsgBox result
  
  Set http = Nothing
End Sub


Public Function CheckPortAvailable(ByVal port As Long) As Boolean
  '
  ' 127.0.0.1 (localhost) のポートが使える可能性があるか (他のプロセスが使っていないか) どうか。
  ' 使える可能性がある場合は True を返す。
  '
  Dim stdout As String
  Dim cmd As String
  
  cmd = "cmd /c netstat -an | find ""127.0.0.1:" & port & """"
  stdout = RunSyncCommandAndCatchStdout(cmd)  ' # TODO: 240721 Utils.bas にまとめた場合に、うまく動かない可能性があるので修正せよ。
  
  If InStr(stdout, "127.0.0.1:" & port) > 0 Then
    CheckPortAvailable = False
  Else
    CheckPortAvailable = True
  End If

End Function


Public Sub TEST___CheckPortAvailable()
  MsgBox CheckPortAvailable(8000)
End Sub
