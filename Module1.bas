Attribute VB_Name = "Module1"
Option Explicit

' HACK: 240125 vba の出力を自動実行する。
' HACK: 240125 複数ある場合、utils.bas などにまとめる。

Public Function RunSyncCommandAndCatchStdout(command As String, Optional isHidden As Boolean = False) As String
  '
  ' 同期的にコマンドを実行し、その標準出力を受け取る関数。
  ' (外部ファイル (Ex. exe, cmd, bat etc...) の実行等を想定。)
  '
  ' Parameters
  ' ----------
  ' command : String
  '   実行するコマンド。
  '
  ' isHidden : Boolean (default : False)
  '   コマンドウィンドウを開くかどうか。  ' TODO: 240125 機能追加する (少し面倒そうだった。)
  '
  ' Return
  ' ------
  ' result : String
  '   取得した標準出力。
  '
  Dim wshShell As Object
  Dim wshShellExec As Object
  Dim wshShellstdout As Object
  Dim result As String
  
  Set wshShell = VBA.CreateObject("WScript.Shell")
  wshShell.currentDirectory = ThisWorkbook.Path  ' FIXME: 240125 OneDrive 上で動かない懸念有り。
  Set wshShellExec = wshShell.Exec(command)      ' INFO: 240125 .Exec() は標準出力を受け取り可能。(https://www.bugbugnow.net/2018/06/wshrunexec.html)
  Set wshShellstdout = wshShellExec.stdout
  
  ' [START] run and catch stdout
  result = wshShellstdout.ReadAll
  Do While wshShellExec.Status = 0  ' HACK: 240125 エラー時の対応を書く。
    VBA.DoEvents
  Loop
  ' [END] run and catch stdout
  
  ' [START] post process
  wshShellstdout.Close
  Set wshShell = Nothing
  Set wshShellExec = Nothing
  Set wshShellstdout = Nothing
  ' [END] post process
  
  RunSyncCommandAndCatchStdout = result
End Function


Private Sub TEST___RunExeAndObtainStdout()
 '
 ' python の実行できる windows であることを前提とする。
 '
 Dim result As String
 result = RunSyncCommandAndCatchStdout("python .\py\run_print.py", True)
 Debug.Print result
End Sub
