; Basic Functionality of the AUT
#AutoIt3Wrapper_UseX64=Y
#include <ImageSearch.au3>
#include <MsgBoxConstants.au3>
#include <Utility_Excel.au3>
#include <GuiEdit.au3>
#include <WinAPIProc.au3>
#include <Date.au3>
#include <WinAPIFiles.au3>
#include <Excel.au3>
#include <ScreenCapture.au3>

AutoItSetOption('MouseCoordMode', 0)
Global $iPID, $oOutlook, $oMail

Global Enum $iScientific, $iStandard = 1

Global $iWinTransitionTime = 250	; Time to wait for a window transition to take place (Scientific -> Standard)


Global $dateTimeFormated
Global $date, $time, $dateTime
Global $testcaseName



Func StartApp()
	 $oOutlook = ObjCreate("Outlook.Application")
	 Sleep(1000)
	 $iPID = ProcessExists("OUTLOOK.EXE")
	 ConsoleWrite("PID in start app = " & $iPID & @CRLF)
	 $oMail = $oOutlook.CreateItem(0)
	$oMail.Display

Sleep(1000)
EndFunc


Func CloseApp()
	$iPID = ProcessExists("OUTLOOK.EXE")
	ConsoleWrite("PID gain in close app = " & $iPID & @CRLF)
;~ 	;ConsoleWrite("outlook process value gain in close app = " & $oOutlook & @CRLF)
	ProcessClose( $iPID )
	;_WinAPI_TerminateProcess($iPID,0)
;~ 	ProcessClose( $oOutlook )
;~ 	_WinAPI_TerminateProcess($oOutlook,0)
	If ProcessWaitClose( $iPID, 10) = 1 Then

		Return True
	Else
		Return False
	EndIf
EndFunc


Func InputDataFromExcel($sheetname)

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



;~ Local $oOutlook = ObjCreate("Outlook.Application")
;~ Local $oMail = $oOutlook.CreateItem(0)
;~     $oMail.Display
    $oMail.To = $Reicpt_SendTo ;"sample@example.com"
    $oMail.Subject = $titleEmail ;"Sample Subject"





Local $oWordEditor = $oOutlook.ActiveInspector.wordEditor
    $oWordEditor.Range(0, 0).Select
    $oWordEditor.Application.Selection.TypeText($sBodyHeader)
    ;$oWordEditor.Application.Selection.Paste   ;;=> paste the text we copied previously
	$oWordEditor.Application.Selection.TypeText($BodyMess)
    $oWordEditor.Application.Selection.TypeText($sBodyFooter)

$oMail.Display

Sleep(2000)
	_Excel_BookClose($oWorkBook, False)
	_Excel_Close($oExcel, False)
Sleep(2000)
EndFunc

Func AddAttachmentToEmail($attachmentFullPath)
	$oMail.attachments.add ($attachmentFullPath)

EndFunc

#Region Phuoc workaround to send data into Outlook mail composer
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

Func ClickInsightIcon()
;Move mouse over insight icon so it has status hover (darker background)
	;WinGetHandle("Untitled - Message (HTML) ")
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
		;ConsoleWrite("found insight icon "  & @CRLF)
		;ConsoleWrite("position x1 and y1 = " & $x1 & " | "  & $y1 & @CRLF)
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

#EndRegion



#Region This part is to convert date time to name screenshot
Func ShowDateTime()
	;ConsoleWrite("date = " & _NowDate() & " time = " & _NowTime(5)	)
EndFunc

Func GetCurrentDateTimeForSaveFile()
	;ConsoleWrite("date = " & _NowDate() & @CRLF )
	 $date = ChangeDateFormatForSaveFile(_NowDate())
	;ConsoleWrite("date formated = " & $date & @CRLF )

	;ConsoleWrite("time = " & _NowTime(5) & @CRLF )
	 $time = ChangeTimeFormatForSaveFile(_NowTime(5))
	;ConsoleWrite("time formated = " & $time & @CRLF )

	;ConsoleWrite("Date time formated = " & $date & "_" & $time & @CRLF)

	$dateTimeFormated = $date & "_" & $time
	ConsoleWrite("date and time = " & $dateTimeFormated & @CRLF)
	return $dateTimeFormated
EndFunc

; dd.mm.yyyy --> yyyy-mm-dd
Func ChangeDateFormatForSaveFile($date)
    If $date == '' Then Return ''
    $ret = StringMid($date,7,4) & '-' & StringLeft($date,2) & '-'  & StringMid($date,4,2)
;~  If StringLen($date) > 10 Then $ret &= StringMid($date,11) ; optionally including time
    Return $ret
EndFunc


Func ChangeTimeFormatForSaveFile($time)
    If $time == '' Then Return ''
    $ret = StringReplace($time,":","")
	$ret2 = StringMid($ret,1,2) & "-" & StringMid($ret,3,2)  & "-" & StringMid($ret,5,2)

    Return $ret2
EndFunc

#EndRegion


Func TakeScreenShot($testcaseName)
	 $dateTimeScr = GetCurrentDateTimeForSaveFile()
	 Local $folder = @ScriptDir & "\Report_" & @YEAR &@MON &@MDAY

	 If DirCreate($folder) Then
	 Local $filePath = $folder & "\" & $testcaseName & "_" & $dateTimeScr & ".jpg"
	_ScreenCapture_Capture($filePath, True)
	 EndIf

	;ShellExecute(@MyDocumentsDir & "\GDIPlus_Image1.jpg")
Sleep(2000)
	return $filePath
EndFunc


#Region
; ===================================================================
; _RefreshSystemTray($nDealy = 1000)
;
; Removes any dead icons from the notification area.
; Parameters:
;   $nDelay - IN/OPTIONAL - The delay to wait for the notification area to expand with Windows XP's
;       "Hide Inactive Icons" feature (In milliseconds).
; Returns:
;   Sets @error on failure:
;       1 - Tray couldn't be found.
;       2 - DllCall error.
; ===================================================================
Func _RefreshSystemTray($nDelay = 1000)
; Save Opt settings
    Local $oldMatchMode = Opt("WinTitleMatchMode", 4)
    Local $oldChildMode = Opt("WinSearchChildren", 1)
    Local $error = 0
    Do; Pseudo loop
        Local $hWnd = WinGetHandle("classname=TrayNotifyWnd")
        If @error Then
            $error = 1
            ExitLoop
        EndIf

        Local $hControl = ControlGetHandle($hWnd, "", "Button1")

    ; We're on XP and the Hide Inactive Icons button is there, so expand it
        If $hControl <> "" And ControlCommand($hWnd, "", $hControl, "IsVisible") Then
            ControlClick($hWnd, "", $hControl)
            Sleep($nDelay)
        EndIf

        Local $posStart = MouseGetPos()
        Local $posWin = WinGetPos($hWnd)

        Local $y = $posWin[1]
        While $y < $posWin[3] + $posWin[1]
            Local $x = $posWin[0]
            While $x < $posWin[2] + $posWin[0]
                DllCall("user32.dll", "int", "SetCursorPos", "int", $x, "int", $y)
                If @error Then
                    $error = 2
                    ExitLoop 3; Jump out of While/While/Do
                EndIf
                $x = $x + 8
            WEnd
            $y = $y + 8
        WEnd
        DllCall("user32.dll", "int", "SetCursorPos", "int", $posStart[0], "int", $posStart[1])
    ; We're on XP so we need to hide the inactive icons again.
        If $hControl <> "" And ControlCommand($hWnd, "", $hControl, "IsVisible") Then
            ControlClick($hWnd, "", $hControl)
        EndIf
    Until 1

; Restore Opt settings
    Opt("WinTitleMatchMode", $oldMatchMode)
    Opt("WinSearchChildren", $oldChildMode)
    SetError($error)
EndFunc; _RefreshSystemTray()
#EndRegion









;;;;Below code is from original source for testing calculator application
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