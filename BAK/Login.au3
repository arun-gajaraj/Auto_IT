Func _Au3RecordSetupLogin()
Opt('WinWaitDelay',100)
Opt('WinDetectHiddenText',1)
Opt('MouseCoordMode',0)
Local $aResult = DllCall('User32.dll', 'int', 'GetKeyboardLayoutNameW', 'wstr', '')
If $aResult[1] <> '00000409' Then
  MsgBox(64, 'Warning', 'Recording has been done under a different Keyboard layout' & @CRLF & '(00000409->' & $aResult[1] & ')')
EndIf

EndFunc

Func _WinWaitActivateLogin($title,$text,$timeout=0)
	WinWait($title,$text,$timeout)
	If Not WinActive($title,$text) Then WinActivate($title,$text)
	WinWaitActive($title,$text,$timeout)
EndFunc

_AU3RecordSetupLogin()
#endregion --- Internal functions Au3Recorder End ---

Run('D:\DINESH\PROJECT\WO TRAFFIC\v8.1.0\20160711a_534611 DEV\RunImage\WOTraffic\WOTraffic.exe')
_WinWaitActivateLogin("WideOrbit Login","")
MouseClick("left",376,319,1)
sleep(7000)