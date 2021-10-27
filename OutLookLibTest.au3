#AutoIt3Wrapper_UseX64=Y
#include <Testing.au3>
#include <FunctionLib.au3>

Global $result
Global $attachmentPath

HotKeySet("{ESC}","Quit") ;Press ESC key to quit

; Function with the name "TestSetup" and "TestTearDown" are executed
; before and after every test. Test scripts should include these
; functions to prevent non-fatal AU3check errors.
Func TestSetup()
	;StartApp()
	Call("StartApp")
EndFunc

Func TestTearDown()
	Call("CloseApp")
	;CloseApp()
EndFunc


Test( "Test Yellow" )

	ClickNewEmail()
	;SetDataForEmail("Sheet2")
	SetDataForEmail_2("Sheet2")
;~ 	sleep(3000)
;~ 	ClickInsightIcon()
;~ 	$attachmentPath = TakeScreenShot("Test Yellow")
;~ 	ConsoleWrite("attachment path in Test: " & $attachmentPath & @CRLF)

;~ 	$result = AssertEqual(3.4,3.5, $attachmentPath &  @CRLF)



;~ Test( "Test Red" )
;~ 	ClickNewEmail()
;~ 	SetDataForEmail("Sheet1")
;~ 	sleep(3000)
;~ 	ClickInsightIcon()
;~ 	$attachmentPath = TakeScreenShot("Test Red")
;~ 	ConsoleWrite("attachment path in Test: " & $attachmentPath & @CRLF)
;~ 	$result = AssertEqual(3.4,3.4, $attachmentPath &  @CRLF)



;~ Test( "TestScenario3" )
;~ 	$attachmentPath = TakeScreenShot("TestScenario3")
;~ 	$result = AssertFalse(True, $attachmentPath &  @CRLF)

;~ 	Test( "TestScenario4" )
;~ 	$attachmentPath = TakeScreenShot("TestScenario4")
;~ 	$result = AssertFalse(False, $attachmentPath &  @CRLF)




While True      ;Here start main loop
    Sleep(20)
WEnd

Func Quit()
    Exit
EndFunc