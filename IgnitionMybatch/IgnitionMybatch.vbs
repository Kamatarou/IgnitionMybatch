Option Explicit
 
Const vbHide = 0             '�E�B���h�E���\��
Const vbNormalFocus = 1      '�ʏ�̃E�B���h�E�A���őO�ʂ̃E�B���h�E
 
Dim objWShell
Set objWShell = CreateObject("WScript.Shell")
'"youWantRunSomething.bat"�̕��������Ȃ����p�ӂ����o�b�`�t�@�C���̖��O�ɒu�������Ă��������B
' Please move a part of "youWantRunSomething.bat" to the name of the batch file you prepared.
objWShell.Run "cmd /c ..\batchFolder\youWantRunSomething.bat", vbNormalFocus
'objWShell.Run "cmd /c ..\batchFolder\youWantRunSomething.bat", vbHide , False
 
Set objWShell = Nothing