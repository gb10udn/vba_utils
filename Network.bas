Attribute VB_Name = "Network"
Option Explicit


Public Function SendPost(ByVal port As Long, Optional ByVal postData As String = "{}") As Object
  '
  ' Ex. postData = "{""a"": 1, ""b"": 2}" のように、ダブルクォーテーションが２ついる点に注意せよ。
  '
  Dim url As String
  Dim xmlhttp As Object
  Dim response As Variant
  
  Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
  url = "http://127.0.0.1:" & port & "/"
  xmlhttp.Open "POST", url, False
  xmlhttp.setRequestHeader "Content-Type", "application/json"
  xmlhttp.send postData

  response = xmlhttp.responseText
  Set SendPost = JsonConverter.ParseJson(response)
  Set xmlhttp = Nothing
End Function


Public Function WriteResponse(ByRef response As Object)
  '
  ' 受け取ったレスポンスから、エクセルシートに値を書き込む。
  '
  Dim res As Variant
  For Each res In response
    If res("sheet_name") = "" Then
      Cells(res("x"), res("y")).value = res("value")
    Else
      ThisWorkbook.Sheets(res("sheet_name")).Cells(res("x"), res("y")).value = res("value")
    End If
  Next res
  Set response = Nothing
End Function


Private Sub TEST___SendPost_and_WriteResponse()
  Dim res As Object
  Dim response As Object
  
  Set response = SendPost(8000)
    
  For Each res In response
    Debug.Print res("x")
    Debug.Print res("y")
    Debug.Print res("sheet_name")
    Debug.Print res("value")
  Next res
  
  WriteResponse response  ' INFO: 240721 引数はカッコ無しで渡すこと。(こうすることで、参照渡しが確定する。)
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


Private Sub TEST___CheckPortAvailable()
  MsgBox CheckPortAvailable(8001)
End Sub


Public Function ObtainMinAvailablePort(ByVal minPort As Long, Optional ByVal maxPort As Long = 65536) As Long
  '
  ' 使用できる可能性のある、最小のポート番号を返す。
  '
  Dim port As Long
  For port = minPort To maxPort
    If CheckPortAvailable(port) = True Then
      ObtainMinAvailablePort = port
      Exit For
    End If
  Next port
End Function


Private Sub TEST___ObtainMinAvailablePort()
  Debug.Print ObtainMinAvailablePort(8000)
End Sub


