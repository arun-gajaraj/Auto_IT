#include <Array.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <RevReport.au3>



; Create application object and open an example workbook

Local $oExcel = _Excel_Open()
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
Local $oWorkbook = _Excel_BookOpen($oExcel, @ScriptDir & "\Test.xls")
If @error Then
    MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example", "Error opening workbook '" & @ScriptDir & "\Test.xls'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
    _Excel_Close($oExcel)
    Exit
EndIf

; *****************************************************************************
; Read data from a single cell on the active sheet of the specified workbook
; *****************************************************************************


; ----------------    Check and Run Revenue Report  OLD  ------------------------------

Local $sResultRev = _Excel_RangeRead($oWorkbook,"ReportsToRun", "A11")
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	;;;;MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Data successfully read." & @CRLF & "Value of cell A1: " & $sResult)
If $sResultRev = "Y" Then

		;; Calling All Revenue Reports Listed in Sheet
		_RunAllRevenue()

EndIf

Local $sResultInv = _Excel_RangeRead($oWorkbook,"ReportsToRun", "A6")
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	;;;;MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Data successfully read." & @CRLF & "Value of cell A1: " & $sResultInv)
If $sResultInv = "Y" Then
		;; Calling All Inventory Reports Listed in Sheet
		_RunAllInventory()

EndIf


;~~~~~~~~~~~ Read Data To Run Revenue Report   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Func _RunAllRevenue()

 Local $aResult = _Excel_RangeRead($oWorkbook, "Revenue_Report", "A1:A3")		;~~~~ Change Range when adding/removing Reports
 If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 3", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

 Local $arrayBoundReports = UBound($aResult,$UBOUND_ROWS)

;~  MsgBox($MB_SYSTEMMODAL, '',$arrayBoundReports )

  For $i = 1 To $arrayBoundReports-1 Step +1
     OpenRevReportScreen($aResult[$i],"Revenue Report  [2016, , As Booked, AE, Cash, Revenue , Adjustments , Historical Revenue ] ["&$aResult[$i]&"]")
	 Sleep(1000)
 Next
EndFunc


;~~~~~~~~~~~ Read Data To Run Inventory Report   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Func _RunAllInventory()

Local $iResult = _Excel_RangeRead($oWorkbook, "Inventory_Report", "A1:A3")		;~~~~ Change Range when adding/removing Reports
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 3", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

Local $arrayBoundReports1 = UBound($iResult,$UBOUND_ROWS)

For $i = 1 To $arrayBoundReports1-1 Step +1
    OpenInventoryReportScreen($iResult[$i],"Inventory Report ["&$iResult[$i]&"]")
	Sleep(1000)
	MsgBox($MB_SYSTEMMODAL, '',$iResult[$i] )
Next
EndFunc


;#CS 	Local $iRows = UBound($aArray, $UBOUND_ROWS) ; Total number of rows. In this example it will be 10.
;     	Local $iCols = UBound($aArray, $UBOUND_COLUMNS) ; Total number of columns. In this example it will be 20.
;     	Local $iDimension = UBound($aArray, $UBOUND_DIMENSIONS) ; The dimension of the array e.g. 1/2/3 dimensional. #CE


;~ OpenRevReportScreen("CUSTOM","Revenue Report  [2016, , As Booked, AE, Cash, Revenue , Adjustments , Historical Revenue ] [CUSTOM]") ;
;~ OpenRevReportScreen("Kathy's Revenue Report","Revenue Report  [2016, , As Booked, AE, Cash, Revenue , Adjustments , Historical Revenue ] [Kathy's Revenue Report]")

;~ 	OpenInventoryReportScreen("CUSTOM","Inventory Report [Custom]")
;~ 	OpenInventoryReportScreen("ShareBuilders Inventory Report Summary WD  - KWTV","Inventory Report [ShareBuilders Inventory Report Summary WD  - KWTV]")




