#include<Excel.au3>



Func getContentFromExcel($sheetname)
	Local $pathExcel = "E:\Work\Outlook Plugin\test calculation\TestScenarioOutlookPlugIn\InputData.xlsx"
	Local $oExcel_1 = _Excel_Open()
	Local $oWorkBook = _Excel_BookOpen($oExcel_1,$pathExcel)

	Local $Reicpt_SendTo = _Excel_RangeRead($oWorkBook,$sheetname,"B1")
	Local $Reicpt_SendCc = _Excel_RangeRead($oWorkBook,$sheetname,"B2")
	Local $Reicpt_SendBcc = _Excel_RangeRead($oWorkBook,$sheetname,"B3")
	Local $titleEmail = _Excel_RangeRead($oWorkBook,$sheetname,"B4")
	Local $BodyMess = _Excel_RangeRead($oWorkBook,$sheetname,"B5",1)

	$BodyMess = StringReplace(StringReplace($BodyMess,"{ALT}"," "),"{ENTER}", "  ")
;Local $aResult = _Excel_RangeRead($oWorkbook, 2, "B5", 3)

	Dim $aExcelContent[5]
	$aExcelContent[0] = $Reicpt_SendTo
	$aExcelContent[1] = $Reicpt_SendCc
	$aExcelContent[2] = $Reicpt_SendBcc
	$aExcelContent[3] = $titleEmail
	$aExcelContent[4] = $BodyMess


	_Excel_BookClose($oWorkBook, False)
	_Excel_Close($oExcel_1, False)

	Return $aExcelContent

EndFunc

;MsgBox(0,"Read",getContentFromExcel("sheet1")[0] & @CRLF & getContentFromExcel("sheet1")[1])
;MsgBox(0,"Read",$Reicpt_SendTo & @CRLF & $Reicpt_SendCc & @CRLF & $titleEmail & @CRLF &       $BodyMess)




