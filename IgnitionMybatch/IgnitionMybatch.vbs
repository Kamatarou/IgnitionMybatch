Option Explicit
 
Const vbHide = 0             '�E�B���h�E���\��
Const vbNormalFocus = 1      '�ʏ�̃E�B���h�E�A���őO�ʂ̃E�B���h�E
 
Dim objWShell
Set objWShell = CreateObject("WScript.Shell")
objWShell.Run "cmd /c youWantRunSomething.bat", vbHide, False
 
Set objWShell = Nothing