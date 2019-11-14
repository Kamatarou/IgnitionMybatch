Option Explicit
 
Const vbHide = 0             'ウィンドウを非表示
Const vbNormalFocus = 1      '通常のウィンドウ、かつ最前面のウィンドウ
 
Dim objWShell
Set objWShell = CreateObject("WScript.Shell")
objWShell.Run "cmd /c youWantRunSomething.bat", vbHide, False
 
Set objWShell = Nothing