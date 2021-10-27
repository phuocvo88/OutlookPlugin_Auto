#include <Excel.au3>
Global $sheetname = "Sheet2"

Local $oExcel = _Excel_Open()
Local $oWorkbook = _Excel_BookOpen($oExcel, @ScriptDir & "\InputData.xlsx")

;~ Local $oRange = $oWorkbook.ActiveSheet.Range("B5").Select ;;=> to copy a text then we paste later
     $oExcel.CopyObjectsWithCells = True
     $oExcel.Selection.Copy

Local $sBodyHeader = _Excel_RangeRead($oWorkBook,$sheetname,"B5",1) & @CRLF & @CRLF
Local $BodyMess = _Excel_RangeRead($oWorkBook,$sheetname,"B6",1) & @CRLF & @CRLF
Local $sBodyFooter = _Excel_RangeRead($oWorkBook,$sheetname,"B7",1)  ; @CRLF & @CRLF & "This is footer of body message"  & @CRLF & @CRLF

Local $Reicpt_SendTo = _Excel_RangeRead($oWorkBook,$sheetname,"B1")
	Local $Reicpt_SendCc = _Excel_RangeRead($oWorkBook,$sheetname,"B2")
	Local $Reicpt_SendBcc = _Excel_RangeRead($oWorkBook,$sheetname,"B3")
	Local $titleEmail = _Excel_RangeRead($oWorkBook,$sheetname,"B4")



Local $oOutlook = ObjCreate("Outlook.Application")
Local $oMail = $oOutlook.CreateItem(0)
    $oMail.Display
    $oMail.To = $Reicpt_SendTo ;"sample@example.com"
    $oMail.Subject = $titleEmail ;"Sample Subject"





Local $oWordEditor = $oOutlook.ActiveInspector.wordEditor
    $oWordEditor.Range(0, 0).Select
    $oWordEditor.Application.Selection.TypeText($sBodyHeader)
    ;$oWordEditor.Application.Selection.Paste   ;;=> paste the text we copied previously
	$oWordEditor.Application.Selection.TypeText($BodyMess)
    $oWordEditor.Application.Selection.TypeText($sBodyFooter)

$oMail.Display

Sleep(3000)
	_Excel_BookClose($oWorkBook, False)
	_Excel_Close($oExcel, False)