#include <Array.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>

; Create application object and open an example workbook
Local $oExcel = _Excel_Open()
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
Local $oWorkbook = _Excel_BookOpen($oExcel, @ScriptDir & "\Test.xls",False)
If @error Then
    MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example", "Error opening workbook '" & @ScriptDir & "Test.xls'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
    _Excel_Close($oExcel)
    Exit
EndIf

; *****************************************************************************
; Read the formulas of a cell range (all used cells in column A)
; *****************************************************************************

Local $aResult = _Excel_RangeRead($oWorkbook, "Revenue_Report", $oWorkbook.ActiveSheet.Usedrange.Columns("A:A"))
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 3", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
;~ MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 3", "Data successfully read." & @CRLF & "Please click 'OK' to display all formulas in column B.")
;~ _ArrayDisplay($aResult, "Excel UDF: _Excel_RangeRead Example 3 - Formulas in column B")

Local $arrayBoundReports = UBound($aResult,$UBOUND_ROWS)

Local $SavedReportName
Local $SavedTitle = "Revenue Report  [2016, , As Booked, AE, Cash, Revenue , Adjustments , Historical Revenue ] ["&$SavedReportName&"]"

MsgBox($MB_SYSTEMMODAL, "", $arrayBoundReports)

For $i = 1 To $arrayBoundReports-1 Step -1
    OpenRevReportScreen($SavedReportName,$SavedTitle)
Next





;#CS 	Local $iRows = UBound($aArray, $UBOUND_ROWS) ; Total number of rows. In this example it will be 10.
;     	Local $iCols = UBound($aArray, $UBOUND_COLUMNS) ; Total number of columns. In this example it will be 20.
;     	Local $iDimension = UBound($aArray, $UBOUND_DIMENSIONS) ; The dimension of the array e.g. 1/2/3 dimensional. #CE





