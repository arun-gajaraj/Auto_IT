#region --- Internal functions Au3Recorder Start ---
#include <MsgBoxConstants.au3>

Func _Au3RecordSetupRevRep()
Opt('WinWaitDelay',100)
Opt('WinDetectHiddenText',1)
Opt('MouseCoordMode',0)
Local $aResult = DllCall('User32.dll', 'int', 'GetKeyboardLayoutNameW', 'wstr', '')
If $aResult[1] <> '00000409' Then
  MsgBox(64, 'Warning', 'Recording has been done under a different Keyboard layout' & @CRLF & '(00000409->' & $aResult[1] & ')')
EndIf

EndFunc

Func _WinWaitActivateRevRep($title,$text,$timeout=0)
	WinWait($title,$text,$timeout)
	If Not WinActive($title,$text) Then WinActivate($title,$text)
	WinWaitActive($title,$text,$timeout)
EndFunc

_AU3RecordSetupRevRep()
#endregion --- Internal functions Au3Recorder End ---


_WinWaitActivateRevRep("WO Traffic 8.1.0","")

;---------- OPEN REVENUE REPORT TASK ------------
;Local $hwnd = WinWait("[CLASS:TfrmRevenueReports4]", "", 10)
;ControlClick($hWnd,"","[CLASS:tbRun]")
ControlClick("WO Traffic 8.1.0", "", "[CLASS:Edit; INSTANCE:1]")
Send("revenue{SPACE}report{ENTER}")
_WinWaitActivateRevRep("Saved Reports - Revenue Report","")
WinSetState("Saved Reports - Revenue Report","", @SW_MAXIMIZE)

;----------SEARCH FOR REVENUE REPORT - CUSTOM-------------
MouseClick("left",450,125,1)
Send("custom1{ENTER}")

;While $PgxBarExists = 0
;	MsgBox($MB_SYSTEMMODAL, "", "Progress Bar is Visible")
 ;   Sleep(3000)
;WEnd
;MsgBox($MB_SYSTEMMODAL, "", "Progress Bar is NOT Visible")
;--------- OPEN REV REPORT-------------
MouseClick("left",690,211,1)
MouseClick("left",690,211,2)
_WinWaitActivateRevRep("Revenue Report  [2016, , As Booked, AE, Cash, Revenue , Adjustments , Historical Revenue ] [CUSTOM1]","")
Local $hwnd = WinWait("[CLASS:TfrmRevenueReports4]", "", 10)

WinSetState("Revenue Report  [2016, , As Booked, AE, Cash, Revenue , Adjustments , Historical Revenue ] [CUSTOM1]","", @SW_MAXIMIZE)
;-----------Click on Run Button---------------
MouseClick("left",261,61,1)
;----------- Wait for Revenue Report Title Change ---------------
_WinWaitActivateRevRep("Revenue Report  [2014, Jul, Aug, Sep, As Booked, AE, Cash, Revenue , Adjustments , Historical Revenue , Includes preempted spots]","")
;Local $PgxBarExists = ControlCommand($hWnd, "", "[CLASS:TcxProgressBar]", "IsVisible", "")
;$PgxBarExists
Sleep(3000)
While ControlCommand($hWnd, "", "[CLASS:TcxProgressBar]", "IsVisible", "") = 1
	;MsgBox($MB_SYSTEMMODAL, "", "Progress Bar is Visible")
	;Local $PgxBarExists = ControlCommand($hWnd, "", "[CLASS:TcxProgressBar]", "IsVisible", "")
    Sleep(3000)
WEnd
;Do
;	MsgBox($MB_SYSTEMMODAL, "", "Progress Bar is Visible")
;	Sleep(3000)
;Until $PgxBarExists = 1
Sleep(2000)
;MsgBox($MB_SYSTEMMODAL, "", "Progress Bar is Not Visible1")

;----------- Save to File ---------------
MouseClick("right",348,458,1)
MouseClick("left",483,744,1)
MouseClick("left",582,742,1)

;--------------wait for save as Dialog ----------------
_WinWaitActivateRevRep("WideOrbit: Save to file","")

;------------- click on File name Textbox ----------------
MouseClick("left",441,514,1)
;Send("{SHIFTDOWN}c{SHIFTUP}{SHIFTDOWN};{SHIFTUP}\{SHIFTDOWN}r{SHIFTUP}eports{SHIFTDOWN}-{SHIFTUP}{SHIFTDOWN}i{SHIFTUP}{BACKSPACE}{BACKSPACE}\{SHIFTDOWN}i{SHIFTUP}nve{BACKSPACE}{BACKSPACE}{BACKSPACE}{BACKSPACE}{SHIFTDOWN}r{SHIFTUP}ev{SHIFTDOWN}r{SHIFTUP}ep{SHIFTDOWN}-{SHIFTUP}{SHIFTDOWN}c{SHIFTUP}ustom")
Send("C:\Reports\Custom_RevReport")

;---------------- Select XLS option -------------------
MouseClick("left",382,538,1)
MouseClick("left",371,559,1)

;--------------- Click on Ok Button -------------------
MouseClick("left",793,587,1)

;--------------- Wait for 'Do you want to open Exported File?' message Dialog -----------------
_WinWaitActivateRevRep("WideOrbit Message","")

;-------------- Click on No Button in message Dialog -------------------
MouseClick("left",350,181,1)
_WinWaitActivateRevRep("frmFormManager","")
MouseClick("right",502,407,1)
_WinWaitActivateRevRep("Revenue Report  [2014, Jul, Aug, Sep, As Booked, AE, Cash, Revenue , Adjustments , Historical Revenue , Includes preempted spots]","")
MouseClick("right",560,465,1)
MouseClick("left",641,749,1)
MouseClick("left",791,755,1)
_WinWaitActivateRevRep("WideOrbit: Save to file","")
MouseClick("left",442,520,1)
Send("C:\Reports\Custom_RevReport")
;Send("{SHIFTDOWN}re{SHIFTUP}v{SHIFTDOWN}re{SHIFTUP}{BACKSPACE}ep{SHIFTDOWN}-{SHIFTUP}{SHIFTDOWN}c{SHIFTUP}ustom")
MouseClick("left",322,585,1)
MouseClick("left",324,541,1)
MouseClick("left",315,569,1)
MouseClick("left",768,587,1)
_WinWaitActivateRevRep("WideOrbit Message","")

MouseClick("left",350,181,1)

Sleep(3000)

WinClose("Revenue Report  [2014, Jul, Aug, Sep, As Booked, AE, Cash, Revenue , Adjustments , Historical Revenue , Includes preempted spots]","")
WinClose("Saved Reports - Revenue Report","")

Exit