Attribute VB_Name = "Command"
Option Explicit


Public Function RunSyncCommandAndCatchStdout(ByVal cmd As String) As String  ' TODO: 240127 timeout / isHidden を実装する。
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
  wshShell.CurrentDirectory = ThisWorkbook.Path  ' FIXME: 240125 OneDrive 上で動かない懸念有り。
  Set wshShellExec = wshShell.Exec(cmd)          ' INFO: 240125 .Exec() は標準出力を受け取り可能。(https://www.bugbugnow.net/2018/06/wshrunexec.html)
  Set wshShellStdout = wshShellExec.stdout
  
  ' [START] run and catch stdout
  result = wshShellStdout.ReadAll
  Do While wshShellExec.Status = 0  ' HACK: 240125 エラー時の対応を書く。
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


Private Sub TEST___RunExeAndObtainStdout()
  '
  ' python の実行できる windows であることを前提とする。
  '
  Dim result As String
  result = RunSyncCommandAndCatchStdout("python .\py\run_print.py")
  Debug.Print result
  
  result = RunSyncCommandAndCatchStdout("ipconfig")
  Debug.Print result
End Sub


Public Sub RunAsyncCommandAndCatchStdout(ByVal cmd As String, Optional ByVal isHidden As Boolean = True)
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
  wshShell.CurrentDirectory = ThisWorkbook.Path  ' FIXME: 240128 OneDrive 上で動かない懸念有り。
  
  If isHidden Then
    wshShell.Run cmd, vbHide, False   ' INFO: 第三引数 --> 同期する (True) or しない (False)。非同期処理のプロシージャなので、False とした。
  Else
    wshShell.Run cmd, vbNormalFocus, False
  End If
  Set wshShell = Nothing
End Sub


Private Sub TEST_RunAsyncCommandAndCatchStdout()
  Call RunAsyncCommandAndCatchStdout("python .\py\run_print.py", True)
  Debug.Print "isHidden = True で実行しました。"
  
  Call RunAsyncCommandAndCatchStdout("python .\py\run_print.py", False)
  Debug.Print "isHidden = False で実行しました。一瞬、プロンプトのウィンドウが出ていれば OK"
End Sub
