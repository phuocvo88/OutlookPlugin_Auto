#AutoIt3Wrapper_UseX64=Y
#include <Testing.au3>
#include <FunctionLib.au3>

Global $result
Global $attachmentPath, $screenshotPath
Global $test1 = "test Red icon"
Global $test2 = "test Yellow icon"
Global $test3 = "test Green icon"
Global $test4 = "test attachments"
Global $test5 = "test message pop-up"

Global $preQA_TC1 = "CheckGreenIcon"
Global $preQA_TC2 = "CheckRedIcon"

Global $PreQA_TestBuild = "CheckRedIcon"

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
	_RefreshSystemTray()
EndFunc


#cs
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


#ce

#cs
#Region Test("test sheet 5"): test mesage on popup
Test($test5)
	InputDataFromExcel("Sheet5")
	AddAttachmentToEmail( @ScriptDir & "\mailAttachments\The_Outlook.jpg")
	;AddAttachmentToEmail( @ScriptDir & "\mailAttachments\sampleDOC.docx")
	;AddAttachmentToEmail( @ScriptDir & "\mailAttachments\samplePDF.pdf")
	;AddAttachmentToEmail( @ScriptDir & "\mailAttachments\sampleXLS.xlsx")
	Sleep(1000)
	;Local $resultStatus = IsIconGreen()
	;ConsoleWrite("icon status Green: " & $resultStatus & @CRLF)
ClickSendMail2()
	$screenshotPath = TakeScreenShot($test5)
	ConsoleWrite("attachment path in Test: " & $screenshotPath & @CRLF)
GetMessFromPopup()
ClickCancelOnPopup()

	;$result= AssertTrue($resultStatus,$screenshotPath)
	_RefreshSystemTray()
	Sleep(6000)
#EndRegion
#ce



#Region test cases for PreQA build

Test("test green")
	InputDataFromExcel("Sheet1")
	Sleep(1000)

	Local $winErrMsgBox =  WinWaitActive("AutoIt Error","",5)
	If NOT $winErrMsgBox = 0 then
		GetAllWindowsControls(WinGetHandle("AutoIt Error"))

;~ 	ControlClick("[CLASS:Button; INSTANCE:1]","OK",
;~ 	ControlClick(WinActivate("AutoIt Error"),"OK",

	EndIf

	Local $resultStatus = IsIconGreen()
	$screenshotPath = TakeScreenShot("test green")
	ClickSendMail2()

	ConsoleWrite("attachment path in Test: " & $screenshotPath & @CRLF)
	AssertTrue($resultStatus,$screenshotPath)
	_RefreshSystemTray()
	Sleep(3000)


Test($PreQA_TestBuild)
	InputDataFromExcel("Sheet2")
	Sleep(1000)
	Local $resultStatus = IsIconRed()
	ClickSendMail2()
	$screenshotPath = TakeScreenShot($PreQA_TestBuild)
	ConsoleWrite("attachment path in Test: " & $screenshotPath & @CRLF)
	Local $popupTxt = GetMessFromPopup()
	Local $compareTxt = "Email contains claim '1000081' for 26E5"
	ClickCancelOnPopup()
	Local $resultCompareMess = IsMessContainsTxt($popupTxt,$compareTxt)

	If $resultStatus = True AND $resultCompareMess = True Then
	Local $result = True
	EndIf

	AssertTrue($result,$screenshotPath)

	_RefreshSystemTray()
	Sleep(3000)

#EndRegion



;FlushTestResults()
;Exit

;~ While True      ;Here start main loop
;~     Sleep(20)
;~ WEnd

Func Quit()
    Exit
EndFunc