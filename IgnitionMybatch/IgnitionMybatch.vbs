Option Explicit
 
Const vbHide = 0             'ウィンドウを非表示
Const vbNormalFocus = 1      '通常のウィンドウ、かつ最前面のウィンドウ
 
Dim objWShell
Set objWShell = CreateObject("WScript.Shell")
'"youWantRunSomething.bat"の部分をあなたが用意したバッチファイルの名前に置き換えてください。
' Please move a part of "youWantRunSomething.bat" to the name of the batch file you prepared.
objWShell.Run "cmd /c ..\batchFolder\youWantRunSomething.bat", vbNormalFocus
'objWShell.Run "cmd /c ..\batchFolder\youWantRunSomething.bat", vbHide , False
 
Set objWShell = Nothing