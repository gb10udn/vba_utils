Attribute VB_Name = "Module1"
Option Explicit

' HACK: 240125 vba �̏o�͂��������s����B
' HACK: 240125 ��������ꍇ�Autils.bas �Ȃǂɂ܂Ƃ߂�B

Public Function RunSyncCommandAndCatchStdout(command As String, Optional isHidden As Boolean = False) As String
  '
  ' �����I�ɃR�}���h�����s���A���̕W���o�͂��󂯎��֐��B
  ' (�O���t�@�C�� (Ex. exe, cmd, bat etc...) �̎��s����z��B)
  '
  ' Parameters
  ' ----------
  ' command : String
  '   ���s����R�}���h�B
  '
  ' isHidden : Boolean (default : False)
  '   �R�}���h�E�B���h�E���J�����ǂ����B  ' TODO: 240125 �@�\�ǉ����� (�����ʓ|�����������B)
  '
  ' Return
  ' ------
  ' result : String
  '   �擾�����W���o�́B
  '
  Dim wshShell As Object
  Dim wshShellExec As Object
  Dim wshShellstdout As Object
  Dim result As String
  
  Set wshShell = VBA.CreateObject("WScript.Shell")
  wshShell.currentDirectory = ThisWorkbook.Path  ' FIXME: 240125 OneDrive ��œ����Ȃ����O�L��B
  Set wshShellExec = wshShell.Exec(command)      ' INFO: 240125 .Exec() �͕W���o�͂��󂯎��\�B(https://www.bugbugnow.net/2018/06/wshrunexec.html)
  Set wshShellstdout = wshShellExec.stdout
  
  ' [START] run and catch stdout
  result = wshShellstdout.ReadAll
  Do While wshShellExec.Status = 0  ' HACK: 240125 �G���[���̑Ή��������B
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
 ' python �̎��s�ł��� windows �ł��邱�Ƃ�O��Ƃ���B
 '
 Dim result As String
 result = RunSyncCommandAndCatchStdout("python .\py\run_print.py", True)
 Debug.Print result
End Sub
