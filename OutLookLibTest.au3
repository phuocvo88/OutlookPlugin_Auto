#AutoIt3Wrapper_UseX64=Y
#include <Testing.au3>
#include <FunctionLib.au3>

Global $result
Global $attachmentPath, $screenshotPath
Global $test1 = "test Red icon"
Global $test2 = "test Yellow icon"
Global $test3 = "test Green icon"
Global $test4 = "test attachments"



HotKeySet("{ESC}","Quit") ;Press ESC key to quit




; Function with the name "TestSetup" and "TestTearDown" are executed
; before and after every test. Test scripts should include these
; functions to prevent non-fatal AU3check errors.
Func TestSetup()
	StartApp()
	;Call("StartApp")
EndFunc

Func TestTearDown()
	;Call("CloseApp")
	CloseApp()
EndFunc



#Region Test("test sheet 1"): test red icon
Test($test1)
	InputDataFromExcel("Sheet1")
	AddAttachmentToEmail( @ScriptDir & "\mailAttachments\The_Outlook.jpg")
	Sleep(2000)
	Local $resultStatus = IsIconRed()
	ConsoleWrite("icon status red: " & $resultStatus & @CRLF)

	$screenshotPath = TakeScreenShot($test1)
	ConsoleWrite("attachment path in Test: " & $screenshotPath & @CRLF)

	$result= AssertTrue($resultStatus,$screenshotPath)
	_RefreshSystemTray()
	Sleep(1000)

#EndRegion

#Region Test("test sheet 2"): test yellow icon
Test($test2)
	InputDataFromExcel("Sheet2")
	AddAttachmentToEmail( @ScriptDir & "\mailAttachments\The_Outlook.jpg")
	Sleep(2000)
	Local $resultStatus = IsIconYellow()
	ConsoleWrite("icon status yellow: " & $resultStatus & @CRLF)

	$screenshotPath = TakeScreenShot($test2)
	ConsoleWrite("attachment path in Test: " & $screenshotPath & @CRLF)
	$result= AssertTrue($resultStatus,$screenshotPath)
	_RefreshSystemTray()
	Sleep(1000)
#EndRegion

#Region Test("test sheet 3"): test green icon
Test($test3)
	InputDataFromExcel("Sheet3")
	AddAttachmentToEmail( @ScriptDir & "\mailAttachments\The_Outlook.jpg")
	Sleep(2000)
	Local $resultStatus = IsIconGreen()
	ConsoleWrite("icon status Green: " & $resultStatus & @CRLF)

	;Local $resultStatus = IsIconRed()
 	;ConsoleWrite("icon status red: " & $resultStatus & @CRLF)
 	;$resultStatus = IsIconGreen()
	;ConsoleWrite("icon status green: " & $resultStatus & @CRLF)
 	;Sleep(1000)


	$screenshotPath = TakeScreenShot($test3)
	ConsoleWrite("attachment path in Test: " & $screenshotPath & @CRLF)
	$result= AssertTrue($resultStatus,$screenshotPath)
	_RefreshSystemTray()
	Sleep(1000)
#EndRegion


#Region Test("test sheet 4"): test attach all kinds of attachments
Test($test4)
	InputDataFromExcel("Sheet4")
	AddAttachmentToEmail( @ScriptDir & "\mailAttachments\The_Outlook.jpg")
	AddAttachmentToEmail( @ScriptDir & "\mailAttachments\sampleDOC.docx")
	AddAttachmentToEmail( @ScriptDir & "\mailAttachments\samplePDF.pdf")
	AddAttachmentToEmail( @ScriptDir & "\mailAttachments\sampleXLS.xlsx")
	Sleep(1000)
	Local $resultStatus = IsIconGreen()
	ConsoleWrite("icon status Green: " & $resultStatus & @CRLF)

	$screenshotPath = TakeScreenShot($test4)
	ConsoleWrite("attachment path in Test: " & $screenshotPath & @CRLF)
	$result= AssertTrue($resultStatus,$screenshotPath)
	_RefreshSystemTray()
	Sleep(1000)
#EndRegion


;FlushTestResults()
;Exit

;~ While True      ;Here start main loop
;~     Sleep(20)
;~ WEnd

Func Quit()
    Exit
EndFunc