#include <MsgBoxConstants.au3>
#include <Functions.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <WinAPIFiles.au3>

Global $filepath = "D:\AutoIT_Reports\"
Global $filename = "Report from "& @MDAY & "-" & @MON& "-" &@YEAR& "--" &@HOUR& "-" &@MIN& "-" & @SEC&".txt"
Global $fullPath = ($filepath & "" &$filename)
Global $file = _FileCreate ( $fullPath )
Global $log = FileOpen($fullPath , $FO_APPEND)
Global $SaveReportAsName = ""  ;--- Filename Pre-Fix for both .xls and .xlsx formats

Global $timerRun
Global $timerDurationRun

	Local $hWnd = WinWait("WO Traffic 8.1.0", "", 0)

;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Reports to Run   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

;~ OpenRevReportScreen()
;~ OpenInventoryReportScreen()
;~ OpenPacingReportScreen()
;~ OpenLogSummaryReportScreen()
;~ OpenMissingInsReportScreen()

;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Report Run Methods  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Func OpenRevReportScreen($savedReportName,$reportScreen)
	;~ From Open TaskBar

	FileWrite($log,("" &@CRLF & "Revenue Report : "&$savedReportName&" Started its run at" & @HOUR & "-" & @MIN & "-" & @SEC) )

	_OpenTask($hWnd, "Revenue Report") ;~ Window "Saved Reports - Revenue Report" is opened

	Local $rWnd = WinWait("Saved Reports - Revenue Report", "", 10)
	_SearchWithTitle($rWnd,$savedReportName)
	_OpenFirstItem($rWnd,"[CLASS:TcxGridSite; INSTANCE:1]")
	Local $runWnd = WinWait($reportScreen, "", 10)
	WinSetState($runWnd,"",@SW_MAXIMIZE)
	_RunReport($runWnd,"[CLASS:TwoPageControl; INSTANCE:1]",$timerRun,$log)
 	_SaveToFile($runWnd,$SaveReportAsName &$savedReportName&"-Revenue Report"& ".xlsx")  ;~~~~ Saves to File in specified extension ~~~~~~~~~~~~~
 	_SaveToFile($runWnd,$SaveReportAsName &$savedReportName&"-Revenue Report"& ".xls")

	;--------- Operation done! Closing windows!! -----------------------------
;~ 	MsgBox($MB_SYSTEMMODAL, '', "Done. Closing!")
	Sleep(2000)
	;-------------------Writing to Log ---------------------------------------
	FileWrite($log,("" &@CRLF & "Revenue Report : "&$savedReportName&" finished its run at" & @HOUR & "-" & @MIN & "-" & @SEC) )
	WinClose($runWnd)
	WinClose($rWnd)

EndFunc   ;==>OpenRevReportScreen

Func OpenInventoryReportScreen($savedReportName,$reportScreen) ;~ From Open TaskBar

	FileWrite($log,("" &@CRLF & "Inventory Report : "&$savedReportName&" Started its run at" & @HOUR & "-" & @MIN & "-" & @SEC) )

	_OpenTask($hWnd, "Inventory Report") ;~ Window "Saved Reports - Inventory Report" is opened

	Local $rWnd = WinWait("Saved Reports - Inventory Report", "", 10)

	_SearchWithTitle($rWnd,$savedReportName)
	_OpenFirstItem($rWnd,"[CLASS:TcxGridSite; INSTANCE:1]")

	Local $runWnd = WinWait($reportScreen, "", 0)

	WinSetState($runWnd,"",@SW_MAXIMIZE)

	_RunReport2($runWnd,"[CLASS:TwoPageControl; INSTANCE:1]",$timerRun,$log)

 	_SaveToFile2($runWnd,$SaveReportAsName &$savedReportName&"-Inventory Report"& ".xlsx")  ;~~~~ Saves to File in specified extension ~~~~~~~~~~~~~
 	_SaveToFile2($runWnd,$SaveReportAsName &$savedReportName&"-Inventory Report"& ".xls")

	;--------- Operation done! Closing windows!! -----------------------------
;~ 	MsgBox($MB_SYSTEMMODAL, '', "Done. Closing!")
	Sleep(2000)
	;-------------------Writing to Log ---------------------------------------
	FileWrite($log,("" &@CRLF & "Inventory Report finished its run at" & @HOUR & "-" & @MIN & "-" & @SEC) )
	WinClose($runWnd)
	WinClose($rWnd)

EndFunc

Func OpenPacingReportScreen() ;~ From Open TaskBar


	_OpenTask($hWnd, "Pacing Report") ;~ Window "Saved Reports - Pacing Report" is opened

	Local $rWnd = WinWait("Saved Reports - Pacing Report", "", 10)

	_SearchWithTitle($rWnd,"CUSTOM")
	_OpenFirstItem($rWnd,"[CLASS:TcxGridSite; INSTANCE:1]")

	Local $runWnd = WinWait("Pacing Report [CUSTOM_Pacing_Report]", "", 0)

	WinSetState($runWnd,"",@SW_MAXIMIZE)

	_RunReport($runWnd,"[CLASS:TwoDBGrid; INSTANCE:1]",$timerRun,$log)

 	_SaveToFile($runWnd,$SaveReportAsName &"Pacing Report"& ".xlsx")  ;~~~~ Saves to File in specified extension ~~~~~~~~~~~~~
 	_SaveToFile($runWnd,$SaveReportAsName &"Pacing Report"& ".xls")

	;--------- Operation done! Closing windows!! -----------------------------
	MsgBox($MB_SYSTEMMODAL, '', "Done. Closing!")
	Sleep(2000)
	;-------------------Writing to Log ---------------------------------------
	FileWrite($log,("" &@CRLF & "Pacing Report finished its run at" & @HOUR & "-" & @MIN & "-" & @SEC) )
	WinClose($runWnd)
	WinClose($rWnd)

EndFunc

Func OpenLogSummaryReportScreen() ;~ From Open TaskBar


	_OpenTask($hWnd, "Log Summary Report") ;~ Window "Saved Reports - Pacing Report" is opened

	Local $rWnd = WinWait("Saved Reports - Log Summary Report", "", 10)

	_SearchWithTitle($rWnd,"CUSTOM")
	_OpenFirstItem($rWnd,"[CLASS:TcxGridSite; INSTANCE:1]")

	Local $runWnd = WinWait("Log Summary [CUSTOM]", "", 0)

	WinSetState($runWnd,"",@SW_MAXIMIZE)

	_RunReport($runWnd,"[CLASS:TwoDBGrid; INSTANCE:1]",$timerRun,$log)

 	_SaveToFile($runWnd,$SaveReportAsName &"-Log Summary Report"& ".xlsx")  ;~~~~ Saves to File in specified extension ~~~~~~~~~~~~~
 	_SaveToFile($runWnd,$SaveReportAsName &"-Log Summary Report"& ".xls")

	;--------- Operation done! Closing windows!! -----------------------------
	MsgBox($MB_SYSTEMMODAL, '', "Done. Closing!")
	Sleep(2000)
	;-------------------Writing to Log ---------------------------------------
	FileWrite($log,("" &@CRLF & "Log Summary Report finished its run at" & @HOUR & "-" & @MIN & "-" & @SEC) )
	WinClose($runWnd)
	WinClose($rWnd)

EndFunc

Func OpenMissingInsReportScreen() ;~ From Open TaskBar


	_OpenTask($hWnd, "Missing Instructions Report") ;~ Window "Saved Reports - Pacing Report" is opened

	Local $rWnd = WinWait("Saved Reports - Missing Instructions Report", "", 10)

	_SearchWithTitle($rWnd,"CUSTOM")
	_OpenFirstItem($rWnd,"[CLASS:TcxGridSite; INSTANCE:1]")

	Local $runWnd = WinWait("Missing Instructions [CUSTOM]", "", 0)

	WinSetState($runWnd,"",@SW_MAXIMIZE)

	_RunReport($runWnd,"[CLASS:TdxDBGrid; INSTANCE:1]",$timerRun,$log)

 	_SaveToFile($runWnd,$SaveReportAsName &"-Missing Instructions Report"& ".xlsx")  ;~~~~ Saves to File in specified extension ~~~~~~~~~~~~~ CHANGE THE ID
 	_SaveToFile($runWnd,$SaveReportAsName &"-Missing Instructions Report"& ".xls")

	;--------- Operation done! Closing windows!! -----------------------------
	MsgBox($MB_SYSTEMMODAL, '', "Done. Closing!")
	Sleep(2000)
	;-------------------Writing to Log ---------------------------------------
	FileWrite($log,("" &@CRLF & "Missing Instructions Report finished its run at" & @HOUR & "-" & @MIN & "-" & @SEC) )
	WinClose($runWnd)
	WinClose($rWnd)

EndFunc

;~ FileClose($log)

;~ MsgBox($MB_SYSTEMMODAL, '', "Done!")

