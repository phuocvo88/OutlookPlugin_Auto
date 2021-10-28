#AutoIt3Wrapper_UseX64=Y
#include <Testing.au3>
#include <FunctionLib.au3>

Global $result
Global $attachmentPath, $screenshotPath

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


Test("test sheet 1")
	InputDataFromExcel("Sheet1")
	;Sleep(2000)
	;ClickInsightIcon()
	AddAttachmentToEmail( @ScriptDir & "\The_Outlook.jpg")
	Sleep(1000)
	$screenshotPath = TakeScreenShot("test sheet 1")
ConsoleWrite("attachment path in Test: " & $screenshotPath & @CRLF)
$result = AssertEqual(3.4,3.5, $screenshotPath)
_RefreshSystemTray()
Sleep(1000)


Test("test sheet 2")
	InputDataFromExcel("Sheet2")
	;Sleep(2000)
	;ClickInsightIcon()
	AddAttachmentToEmail( @ScriptDir & "\The_Outlook.jpg")
	Sleep(1000)
	$screenshotPath = TakeScreenShot("test sheet 2")
ConsoleWrite("attachment path in Test: " & $screenshotPath & @CRLF)
$result = AssertEqual(3.4,3.5, $screenshotPath)
_RefreshSystemTray()
Sleep(1000)

Test("test sheet 3")
	InputDataFromExcel("Sheet3")
	;Sleep(2000)
	;ClickInsightIcon()
	AddAttachmentToEmail( @ScriptDir & "\The_Outlook.jpg")
	Sleep(1000)
	$screenshotPath = TakeScreenShot("test sheet 3")
ConsoleWrite("attachment path in Test: " & $screenshotPath & @CRLF)
$result = AssertEqual(3.5,3.5, $screenshotPath)
Sleep(2000)










;FlushTestResults()
;Exit

;~ While True      ;Here start main loop
;~     Sleep(20)
;~ WEnd

Func Quit()
    Exit
EndFunc