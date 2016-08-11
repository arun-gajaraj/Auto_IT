#include <MsgBoxConstants.au3>
#include <Functions.au3>


OpenRevReportScreen()

Func OpenRevReportScreen()  ;~ From Open TaskBar

Local $hWnd = WinWait("WO Traffic 8.1.0","",0)
_OpenTask($hWnd,"Revenue Report") ;~ Window "Saved Reports - Revenue Report" is opened

Local $rWnd = WinWait("Saved Reports - Revenue Report","",10)


	EndFunc