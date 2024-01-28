Attribute VB_Name = "FilePath"
Option Explicit

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
