; Basic Functionality of the AUT
#AutoIt3Wrapper_UseX64=Y
#include <ImageSearch.au3>
#include <MsgBoxConstants.au3>
#include <Utility_Excel.au3>
#include <GuiEdit.au3>
#include <WinAPIProc.au3>
#include <Date.au3>
#include <WinAPIFiles.au3>

#include <ScreenCapture.au3>

AutoItSetOption('MouseCoordMode', 0)
Global $iPID

Global Enum $iScientific, $iStandard = 1

Global $iWinTransitionTime = 250	; Time to wait for a window transition to take place (Scientific -> Standard)


Global $dateTimeFormated
Global $date, $time, $dateTime
Global $testcaseName


Func StartApp()


	;this part will be put in set up
	$iPID = Run("C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE")
	ConsoleWrite("PID value = " & $iPID & @CRLF)
	AutoItSetOption('MouseCoordMode', 0)
	Local $outLook = WinWait("Inbox - P.Minh@aswhiteglobal.com - Outlook")
	WinSetState("Inbox - P.Minh@aswhiteglobal.com - Outlook", "", @SW_MAXIMIZE )
	WinGetHandle("Inbox - P.Minh@aswhiteglobal.com - Outlook")
	ConsoleWrite("launch ok"  & @CRLF)
	Sleep(3000)


EndFunc

Func CloseApp()
	ProcessClose( $iPID )
	_WinAPI_TerminateProcess($iPID,0)
	If ProcessWaitClose( $iPID, 10) = 1 Then

		Return True
	Else
		Return False
	EndIf
EndFunc

Func SetDataForEmail($sheetname)
	$inputData = getContentFromExcel($sheetname)
	$receipient_SendTo = $inputData[0]
	$receipient_SendCc = $inputData[1]
	$receipient_SendBcc = $inputData[2]
	$mail_Title = $inputData[3]
	$mail_BodyMess = $inputData[4]


	ControlSetText("Untitled - Message (HTML) ", "", "[Class:RichEdit20WPT; Instance:2]",$receipient_SendTo,1)
	Sleep(1000)
	Send( "{TAB} ")

	ControlSetText("Untitled - Message (HTML) ", "", "[Class:RichEdit20WPT; Instance:3]",$receipient_SendCc ,1)
	Sleep(1000)
	Send( "{TAB} ")

	ControlSetText("Untitled - Message (HTML) ", "", "[Class:RichEdit20WPT; Instance:5]", $mail_Title,1)
	Sleep(1000)

	Send( "{TAB} ")
	Send( "{TAB} ")
	Send( $mail_BodyMess)


Sleep(3000)

EndFunc

Func SetDataForEmail_2($sheetname)
	$inputData = getContentFromExcel($sheetname)
	$receipient_SendTo = $inputData[0]
	$receipient_SendCc = $inputData[1]
	$receipient_SendBcc = $inputData[2]
	$mail_Title = $inputData[3]
	$mail_BodyMess = $inputData[4]


;~ 	ControlSetText("Untitled - Message (HTML) ", "", "[Class:RichEdit20WPT; Instance:2]",$receipient_SendTo,1)
;~ 	Sleep(1000)
;~ 	Send( "{TAB} ")

;~ 	ControlSetText("Untitled - Message (HTML) ", "", "[Class:RichEdit20WPT; Instance:3]",$receipient_SendCc ,1)
;~ 	Sleep(1000)
	Send( "{TAB} ")

	ControlSetText("Untitled - Message (HTML) ", "", "[Class:RichEdit20WPT; Instance:5]", $mail_Title,1)
	Sleep(1000)

	Send( "{TAB} ")
	Send( "{TAB} ")
	;$mail_BodyMess =   StringReplace(StringReplace($mail_BodyMess,"{ALT}"," "),"{ENTER}", "  ")
	Send( $mail_BodyMess)


Sleep(3000)

EndFunc

Func ClickNewEmail()
	Local $x=0,$y=0
	Local $search

	MouseMove(@DesktopWidth/2 , @DesktopHeight/2)
	Sleep(2000)
	$search = _ImageSearch(@ScriptDir&'\newEmailBtn.bmp',1, $x, $y, 0)

	ConsoleWrite("step 1"  & @CRLF)
	;MsgBox(0, "Location", "x = " & $x & " & y= " & $y)
	if $search = 1 then
		ConsoleWrite("step 2"  & @CRLF)
		MouseMove ( $x , $y )
		;MouseClick($MOUSE_CLICK_LEFT)
		MouseClick($MOUSE_CLICK_LEFT, $x, $y ,1)
	EndIf

	Sleep(500)
	WinWait("Untitled - Message (HTML) ", "", 2)
	WinSetState("Untitled - Message (HTML) ",  "", @SW_MAXIMIZE)


	;WinWaitActive("[CLASS:DirectUIHWND]")
	Sleep(3000)

EndFunc

Func ClickInsightIcon()
;Move mouse over insight icon so it has status hover (darker background)
	WinGetHandle("Untitled - Message (HTML) ")
	Local $x1=0, $y1=0
	Local $x2=0, $y2=0
	Local $search


	MouseMove(@DesktopWidth/2 , @DesktopHeight/2)
	Sleep(2000)
	$search = _ImageSearch(@ScriptDir&'\Insights_icon.bmp',1, $x1, $y1, 0)

	;convert position to zoom level 125%
	$x1 = ($x1 * 100) / 125
	$y1 = ($y1 * 100) / 125
	If $search = 1 then
		ConsoleWrite("found insight icon "  & @CRLF)
		ConsoleWrite("position x1 and y1 = " & $x1 & " | "  & $y1 & @CRLF)
		MouseMove ( $x1 , $y1 )
		MouseClick($MOUSE_CLICK_LEFT, $x1, $y1 ,1)
		Sleep(2000)
	EndIf
#cs
		$search2 = _ImageSearch(@ScriptDir&'\Insight_Selected_icon.bmp',1, $x2, $y2, 0)
		;convert position to zoom level 125%
	$x2 = ($x2 * 100) / 125
	$y2 = ($y2 * 100) / 125
		ConsoleWrite("position x2 and y2 = " & $x2 & " | "  & $y2 & @CRLF)
		If $search2 = 1 Then
			$isHover = True
			MouseClick($MOUSE_CLICK_LEFT, $x2, $y2 ,1)
		EndIf
	EndIf
#ce

EndFunc



Func ShowDateTime()
	ConsoleWrite("date = " & _NowDate() & " time = " & _NowTime(5)	)
EndFunc



Func GetCurrentDateTimeForSaveFile()
	ConsoleWrite("date = " & _NowDate() & @CRLF )
	 $date = ChangeDateFormatForSaveFile(_NowDate())
	ConsoleWrite("date formated = " & $date & @CRLF )

	ConsoleWrite("time = " & _NowTime(5) & @CRLF )
	 $time = ChangeTimeFormatForSaveFile(_NowTime(5))
	ConsoleWrite("time formated = " & $time & @CRLF )

	;ConsoleWrite("Date time formated = " & $date & "_" & $time & @CRLF)

	$dateTimeFormated = $date & "_" & $time
	ConsoleWrite("date and time = " & $dateTimeFormated & @CRLF)
	return $dateTimeFormated
EndFunc

; dd.mm.yyyy --> yyyy-mm-dd
Func ChangeDateFormatForSaveFile($date)
    If $date == '' Then Return ''
    $ret = StringMid($date,7,4) & '-' & StringMid($date,4,2) & '-' & StringLeft($date,2)
;~  If StringLen($date) > 10 Then $ret &= StringMid($date,11) ; optionally including time
    Return $ret
EndFunc


Func ChangeTimeFormatForSaveFile($time)
    If $time == '' Then Return ''
    $ret = StringReplace($time,":","")
	$ret2 = StringMid($ret,1,2) & "-" & StringMid($ret,3,2)  & "-" & StringMid($ret,5,2)

    Return $ret2
EndFunc



Func TakeScreenShot($testcaseName)
	 $dateTimeScr = GetCurrentDateTimeForSaveFile()
	 Local $folder = @ScriptDir & "\Report_" & @YEAR &@MON &@MDAY

	 If DirCreate($folder) Then
	 Local $filePath = $folder & "\" & $testcaseName & "_" & $dateTimeScr & ".jpg"
	_ScreenCapture_Capture($filePath, True)
	 EndIf

;~ 	 	IF FileExists($folder) <> 1 Then
;~ 			Local $filePath = $folder & "\" & $testcaseName & "_" & $dateTimeScr & ".jpg"
;~ 			_ScreenCapture_Capture($filePath, True)
;~ 		Else
;~ 			$folder = DirCreate(@ScriptDir & "\Report_" &  @YEAR &@MON &@MDAY)
;~ 			Local $filePath = $folder & "\" & $testcaseName & "_" & $dateTimeScr & ".jpg"
;~ 			_ScreenCapture_Capture($filePath, True)
;~ 		EndIf




;~ 	IF FileExists($folder) <> 1 Then
;~ 	$folder = DirCreate(@ScriptDir & "\Report_" &  @YEAR &@MON &@MDAY)
;~ 	EndIf
;~ 	Local $filePath = $folder & "\" & $testcaseName & "_" & $dateTimeScr & ".jpg"
;~ 	_ScreenCapture_Capture($filePath, True)




	;ShellExecute(@MyDocumentsDir & "\GDIPlus_Image1.jpg")

	return $filePath
EndFunc















Func GetViewMode()
	IF ( ControlCommand( "Calculator", "", "Sta", "IsVisible", "" ) ) = 1 Then
		Return $iScientific
	Else
		Return $iStandard
	EndIf
EndFunc

Func ClickOn( $controltext )
	If StringLen( $controltext ) = 0 then return False

	If ( ControlClick( "Calculator", "", $controltext ) ) = 1 Then
		Return True
	Else
		Return False
	EndIf
EndFunc

Func ClickString( $text )
	If StringLen( $text ) = 0 then return False

	If StringLen( $text ) = 1 then
		If Not ClickOn( $text ) then Return False
	Else
		$arrText = StringSplit( $text, "" )
		For $i = 1 to $arrText[0]
			If Not ClickOn( $arrText[$i] ) then Return False
		Next
	EndIf

	Return True
EndFunc

Func Read()
	$test = ControlGetText( "Calculator", "", "Edit1" )
	$test = StringReplace( $test, ". ", "" )
	$test = StringStripWS( $test, 3 )

	Return $test
EndFunc