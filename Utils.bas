Attribute VB_Name = "Utils"
Option Explicit


Public Function FilePath_ChangeExtension(ByVal FilePath As String, ByVal newExtension As String) As String
  '
  ' 拡張子を変更する関数。例えば、.csv --> .dat で使用する。
  ' 引数の、newExtension は、ピリオドから入力する。(Ex. ".dat")
  ' filePath が拡張子を持たない場合は、何もしない。
  '
  Dim fso As Object
  Dim hasExtension As Boolean
  Dim baseDir As String
  Dim fileNameWithoutExtension As String
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  hasExtension = fso.GetExtensionName(FilePath) <> ""
  baseDir = fso.GetParentFolderName(FilePath)
  fileNameWithoutExtension = fso.GetBaseName(FilePath)
  
  If hasExtension = True Then
    ChangeExtension = baseDir & "\" & fileNameWithoutExtension & newExtension
  Else
    ChangeExtension = baseDir & "\" & fileNameWithoutExtension
  End If
  
  Set fso = Nothing
  
End Function



Public Function FilePath_ObtainAbsPath(ByVal arg) As String
  '
  ' 絶対パスに変換して返す。なお、絶対パスを渡すとそのまま返す。(os.path.abspath)
  '
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
    
  With CreateObject("WScript.Shell")
    .CurrentDirectory = ThisWorkbook.Path  ' INFO: 231108 ファイル共有用。(ChDrive, ChDir では不可)
  End With
  
  ObtainAbsPath = fso.GetAbsolutePathName(arg)
  Set fso = Nothing
End Function




Public Function FilePath_ObtainFileName(ByVal arg As String) As String
  '
  ' ファイルパスの内、ファイル名 (拡張子を含む) を返す関数。(os.path.basename)
  '
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  ObtainFileName = fso.GetFileName(arg)
  Set fso = Nothing
End Function




Public Sub Network_SendPost()  ' # TODO: 240721 名前をもっと一般的にして、vba utils に加えること。
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


Public Function Network_CheckPortAvailable(ByVal port As Long) As Boolean
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









Public Function command_RunSyncCommandAndCatchStdout(ByVal cmd As String) As String  ' # TODO: 240127 timeout / isHidden を実装する。
  '
  ' 同期的にコマンドを実行し、その標準出力を受け取る関数。
  ' (外部ファイル (Ex. exe, cmd, bat etc...) の実行等を想定。)
  '
  ' Parameters
  ' ----------
  ' cmd : String
  '   実行するコマンド。
  '
  ' Return
  ' ------
  ' result : String
  '   取得した標準出力。
  '
  Dim wshShell As Object
  Dim wshShellExec As Object
  Dim wshShellStdout As Object
  Dim result As String
  
  Set wshShell = VBA.CreateObject("WScript.Shell")
  wshShell.CurrentDirectory = ThisWorkbook.Path  ' # FIXME: 240125 OneDrive 上で動かない懸念有り。
  Set wshShellExec = wshShell.Exec(cmd)          ' INFO: 240125 .Exec() は標準出力を受け取り可能。(https://www.bugbugnow.net/2018/06/wshrunexec.html)
  Set wshShellStdout = wshShellExec.stdout
  
  ' [START] run and catch stdout
  result = wshShellStdout.ReadAll
  Do While wshShellExec.Status = 0  ' # HACK: 240125 エラー時の対応を書く。
    VBA.DoEvents
  Loop
  ' [END] run and catch stdout
  
  ' [START] post process
  wshShellStdout.Close
  Set wshShell = Nothing
  Set wshShellExec = Nothing
  Set wshShellStdout = Nothing
  ' [END] post process
  
  RunSyncCommandAndCatchStdout = result
End Function




Public Sub command_RunAsyncCommandAndCatchStdout(ByVal cmd As String, Optional ByVal isHidden As Boolean = True)
  '
  ' 非同期的にコマンドを実行する関数。(標準出力を受け取らない。)
  ' (外部ファイル (Ex. exe, cmd, bat etc...) の実行等を想定。)
  '
  ' Parameters
  ' ----------
  ' cmd : String
  '   実行するコマンド。
  '
  Dim wshShell As Object

  Set wshShell = CreateObject("WScript.Shell")
  wshShell.CurrentDirectory = ThisWorkbook.Path  ' # FIXME: 240128 OneDrive 上で動かない懸念有り。
  
  If isHidden Then
    wshShell.Run cmd, vbHide, False   ' INFO: 第三引数 --> 同期する (True) or しない (False)。非同期処理のプロシージャなので、False とした。
  Else
    wshShell.Run cmd, vbNormalFocus, False
  End If
  Set wshShell = Nothing
End Sub



