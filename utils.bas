Attribute VB_Name = "utils"
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


Public Function ChangeExtension(ByVal FilePath As String, ByVal newExtension As String) As String
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


Private Sub TEST___ChangeExtension()
  Dim testPath As String
  
  testPath = "hoge\fuga.csv"  ' INFO: 231106 想定するケース
  Debug.Print ChangeExtension(testPath, ".dat")
  
  testPath = "hog.e\fuga.csv"  ' INFO 231123 フォルダ名に拡張子があっても問題なし。
  Debug.Print ChangeExtension(testPath, ".dat")
  
  testPath = "hoge\fuga"  ' INFO 231123 拡張子がない場合は、何もしない。
  Debug.Print ChangeExtension(testPath, ".dat")

  testPath = "hoge\fuga.a.a.a.csv"  ' INFO 231123 ファイル名にピリオドが複数あっても、誤動作しない。
  Debug.Print ChangeExtension(testPath, ".dat")

End Sub

Public Function ObtainAbsPath(ByVal arg) As String
  '
  ' 絶対パスに変換して返す。なお、絶対パスを渡すとそのまま返す。(os.path.abspath)
  '
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
    
  With CreateObject("WScript.Shell")
    .CurrentDirectory = ThisWorkbook.Path  ' INFO: 231108 ファイル共有 (\\sv1401) 用。(ChDrive, ChDir では不可)
  End With
  
  ObtainAbsPath = fso.GetAbsolutePathName(arg)
  Set fso = Nothing
End Function


Private Sub TEST___ObtainAbsPath()
  Dim testPath As String
  
  ' INFO: 231106 フォルダ名から開始
  testPath = "misc\hoge.cmd"
  Debug.Print ObtainAbsPath(testPath)
  
  ' INFO: 231106 .\ から書き始め
  testPath = ".\misc\hoge.cmd"
  Debug.Print ObtainAbsPath(testPath)
  
  ' INFO: 231106 ./ から書き始め
  testPath = "./misc\hoge.cmd"
  Debug.Print ObtainAbsPath(testPath)
  
  ' INFO: 231106 絶対パスで書いた
  testPath = "C:\hoge.cmd"
  Debug.Print ObtainAbsPath(testPath)
  
  ' INFO: 231108 簡易の相対パス
  testPath = "misc"
  Debug.Print ObtainAbsPath(testPath)
  
  ' INFO: 231108 ファイル共有箇所
  testPath = "\\ShareServer\dev"
  Debug.Print ObtainAbsPath(testPath)
End Sub


Public Function ObtainFileName(ByVal arg As String) As String
  '
  ' ファイルパスの内、ファイル名 (拡張子を含む) を返す関数。(os.path.basename)
  '
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  ObtainFileName = fso.GetFileName(arg)
  Set fso = Nothing
End Function


Private Sub TEST___ObtainFileName()
  Dim testPath As String
  
  testPath = "C:\misc\hoge.cmd"
  Debug.Print ObtainFileName(testPath)  ' INFO: 231123 標準的なユースケース。
  
  testPath = "C:\misc\hoge"
  Debug.Print ObtainFileName(testPath)  ' INFO: 231123 拡張子無しのファイルでも動く。
  
End Sub





