#include <Excel.au3>

$oExcel = _Excel_Open()
$oWorkbook = _Excel_BookOpen($oExcel, @ScriptDir & "\InputData.xlsx")
$oRange = $oWorkbook.ActiveSheet.Range("B5").Select
$oExcel.CopyObjectsWithCells = True
$oExcel.Selection.Copy

$oOutlook = ObjCreate("Outlook.Application")
$oMail = $oOutlook.CreateItem(0)

$oMail.Display
$oMail.To = "sample@example.com"
$oMail.Subject = "Sample Subject"
$oWordEditor = $oOutlook.ActiveInspector.wordEditor
$oMail.Body = "Hello" & @CRLF & "Please find charts above." & @CRLF & @CRLF & "Regards Subz"
$oWordEditor.Range(0, 0).Select
$oWordEditor.Application.Selection.Paste
$oMail.Display