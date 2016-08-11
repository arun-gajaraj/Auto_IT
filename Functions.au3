#include <Date.au3>
#include <MsgBoxConstants.au3>


Func _OpenTask($appTitle,$taskToOpen)

	ControlSetText($appTitle, "", "[CLASS:Edit; INSTANCE:1]", $taskToOpen)
	ControlFocus($appTitle,"","[CLASS:Edit; INSTANCE:1]")
	Send("{ENTER}")

EndFunc

Func _SearchWithTitle($WndHandle,$searchText)
	Local $id = "[CLASS:TEdit; INSTANCE:2]"
		ControlSetText($WndHandle, "",$id, $searchText)
		ControlFocus($WndHandle,"",$id)
		Send("{ENTER}")
		Sleep(1000)
	EndFunc

Func _GetDateTime()

;~ 	Return (_NowDate() &" "& _NowTime())
	Return (@MDAY & "-" & @MON& "-" &@YEAR& "--" &@HOUR& "-" &@MIN& "-" & @SEC)
EndFunc

Func _OpenFirstItem($WndHandle,$GridID)
	Sleep(1000)
	Local $iPos = ControlGetPos($WndHandle, "", $GridID)
	ControlFocus($WndHandle,"",$GridID)
	Sleep(100)
	ControlClick($WndHandle,"",$GridID,"left",1, ($iPos[2] / 2),($iPos[3] / 2))
	Sleep(100)
	Send("{HOME}")
	Sleep(100)
	Send("{ALTDOWN}e{ALTUP}")

EndFunc

Func _RunReport($runWnd,$resGrid,$timerStart,$log1)

	Local $lPos = ControlGetPos($runWnd,"",$resGrid)
	WinActivate($runWnd,"")
	MouseMove($lPos[2]/2,$lPos[3]/2)
	$timerStart = TimerInit()		; Starts the Timer when Report is Run
	Send("{ALTDOWN}r{ALTUP}")

	Local $iCursor = MouseGetCursor()
	  Do
		 Sleep(500)
		 $iCursor = MouseGetCursor()

	  Until $iCursor = 2 Or $iCursor = 0  ;~~~~~~~~~~~~~~~~~ 2-ARROW, 15-BUSY
	  	 $timerDurationRun = TimerDiff($timerStart)
		 Local $roundedTime = Round($timerDurationRun/1000,1)
		 FileWrite($log1,(@CRLF & "Report Ran for "& $roundedTime & " Seconds"))
		 Sleep(500)
;~ 		 Local $timesec=($timerDurationRun)/1000)
;~ 	  	 MsgBox($MB_SYSTEMMODAL, '', "Done Loading after"&$timesec)

EndFunc

Func _SaveToFile($runWnd,$fileName)
	Local $gPos = ControlGetPos($runWnd,"","[CLASS:TwoDBGrid; INSTANCE:1]")
	MouseClick("right",$gPos[2]/2,$gPos[3]/2)
	Send("{UP}")
	Sleep(200)
	Send("{UP}")
	Sleep(200)
	Send("{RIGHT}")
	Sleep(200)
	Send("{ENTER}")

	Sleep(200)
	Send($fileName)
	Send("{ENTER}")

	Local $WOMsg = WinWait("WideOrbit Message","",0)
	Sleep(500)
;~ 	ControlClick($WOMsg,"","[CLASS:TButton; INSTANCE:2]","left",1)		;---- Unstable
	Send("{ALTDOWN}n{ALTUP}")



EndFunc

Func _SaveToFile2($runWnd,$fileName)
	Local $gPos = ControlGetPos($runWnd,"","[CLASS:TdxTreeList; INSTANCE:1]")
	MouseClick("right",(($gPos[2]/2)+288), (($gPos[3]/2)+163))   ;~~~ 288,163 is Position Offset

	Send("{UP}")
	Sleep(200)
;~ 	Send("{UP}")
	Sleep(200)
	Send("{RIGHT}")
	Sleep(200)
	Send("{ENTER}")

	Sleep(200)
	Send($fileName)
	Send("{ENTER}")

	Local $WOMsg = WinWait("WideOrbit Message","",0)
	Sleep(500)
;~ 	ControlClick($WOMsg,"","[CLASS:TButton; INSTANCE:2]","left",1)	;---- Unstable
	Send("{ALTDOWN}n{ALTUP}")


EndFunc


Func _RunReport2($runWnd,$resGrid,$timerStart1,$log1)

	Local $lPos = ControlGetPos($runWnd,"",$resGrid)
	WinActivate($runWnd,"")
	MouseMove($lPos[2]/2,$lPos[3]/2)
	$timerStart1 = TimerInit()		; Starts the Timer when Report is Run
	Send("{ALTDOWN}r{ALTUP}")

	Local $iCursor1 = MouseGetCursor()
	  Do
		 Sleep(250)
		 $iCursor1 = MouseGetCursor()
;~ 		 FileWrite($log,$iCursor1 )
	  Until $iCursor1 = 2 Or $iCursor1 = 0 ;~~~~~~~~~~~~~~~~~ 2-ARROW, 15-BUSY, 0- Resize (somewhere in grid)
	  $timerDurationRun = TimerDiff($timerStart1)
		 Local $roundedTime = Round($timerDurationRun/1000,1)
		 FileWrite($log1,(@CRLF & "Report Ran for "& $roundedTime & " Seconds"))
;~ 	  MsgBox($MB_SYSTEMMODAL, '', "Done Loading")

;~ Sleep(15000)

EndFunc