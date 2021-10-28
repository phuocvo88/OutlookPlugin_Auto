#AutoIt3Wrapper_UseX64=Y
#include <Array.au3>


OnAutoItExitRegister("FlushTestResults")

Global $iTestCount = 0
Global $iTotalAssertions =0
Global $iTotalPass = 0
Global $iTotalFail = 0
Global $aResults[1]
Global $iTimeBegin = TimerInit()
Global $bUseSetup = false
Global $sCurrentTestName = ""
Global $result = False
Global $attachmentPath

Func Test( $testName )
	; Reset Global Variables
	$iTestCount += 1

	$sCurrentTestName = $testName

	If $iTestCount > 1 Then
		Call ("TestTearDown" )
	EndIf

Sleep(3000)
	Call( "TestSetup" )
EndFunc

Func AssertTrue( $expression, $attachmentPath )
	$iTotalAssertions += 1
	If Not $expression then
		DoTestFail( $attachmentPath )
		$result = False
	Else
		DoTestPass($attachmentPath)
		$result = True
	EndIf
	Return $result
EndFunc

Func AssertFalse( $expression, $attachmentPath )
	$iTotalAssertions += 1
	If $expression then
		DoTestFail( $attachmentPath )
		$result = False
	Else
		DoTestPass($attachmentPath)
		$result = True
	EndIf
	Return $result
EndFunc

;testing here
Func AssertEqual($a, $b, $attachmentPath)
	$iTotalAssertions += 1
	If $a <> $b Then
		DoTestFail($attachmentPath)
		$result = False
	Else
		DoTestPass($attachmentPath)
		$result = True
	EndIf
	Return $result
EndFunc


Func FlushTestResults()
	Call ("TestTearDown" )

	If FileExists( @ScriptFullPath & ".log" ) then
		FileDelete( @ScriptFullPath & ".log" )
	EndIf

	$TimeEnd = Round(TimerDiff($iTimeBegin) / 1000, 2)

	$FinalResults = StringFormat( "(%d) Tests (%d) Total Assertions (%d) Pass (%d) Fail", _
		$iTestCount, $iTotalAssertions, $iTotalPass, $iTotalFail )
	ReportTestInfo( "" )
	ReportTestInfo( $FinalResults )
	ReportTestInfo( "Total Time: " & $TimeEnd & " seconds" )

	$aResultsMessages = 	_ArrayToString( $aResults, @CRLF, 1 )
	If StringLen( $aResultsMessages ) <> 0 then
		ReportTestInfo( "" )
		ReportTestInfo( "ASSERTION" & @TAB & "MESSAGE" )
		ReportTestInfo( $aResultsMessages )
	EndIf

EndFunc

Func DoTestPass($path)
	$iTotalPass += 1
	Local $message = StringFormat( "%-15s %s", _
		$iTotalAssertions, "Test is PASSED: " & $sCurrentTestName & " -> " & "See attachment for more details: " & $path  )
	_ArrayAdd( $aResults, $message )
	ConsoleWrite( "test case " & $sCurrentTestName & " is Passed" & @CRLF)
EndFunc

Func DoTestFail( $path )
	$iTotalFail += 1
	Local $message = StringFormat( "%-15s %s", _
		$iTotalAssertions, "Error in " & $sCurrentTestName & " -> " & "See attachment for more details: " & $path  )
	_ArrayAdd( $aResults, $message )
	ConsoleWrite( "test case " & $sCurrentTestName & " is Failed"  & @CRLF)
EndFunc

Func ReportTestInfo( $data )
	ConsoleWrite( $data & @CRLF )
	FileWriteLine( @ScriptFullPath & ".log", $data )
EndFunc