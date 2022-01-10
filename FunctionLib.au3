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
#include "_ImageSearch_UDF.au3"
#include <Array.au3>
#include <WinAPI.au3>

Opt("WinWaitDelay", 1000)
AutoItSetOption('MouseCoordMode', 0)
Global $iPID, $oOutlook, $oMail

Global Enum $iScientific, $iStandard = 1

Global $iWinTransitionTime = 250	; Time to wait for a window transition to take place (Scientific -> Standard)


Global $dateTimeFormated
Global $date, $time, $dateTime
Global $testcaseName
Global $oOutlook , $oMail
Global $popupTxt, $txtToVerify

Func StartApp()
	 $oOutlook = ObjCreate("Outlook.Application")
	 Sleep(5000)
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
Local $oWorkbook = _Excel_BookOpen($oExcel, @ScriptDir & "\InputData_PreQA.xlsx")

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


Sleep(2000)
Local $winErrMsgBox =  WinWaitActive("AutoIt Error","",5)
	If NOT $winErrMsgBox = 0 then
		GetAllWindowsControls(WinGetHandle("AutoIt Error"))

;~ 	ControlClick("[CLASS:Button; INSTANCE:1]","OK",
;~ 	ControlClick(WinActivate("AutoIt Error"),"OK",

	EndIf


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
	MouseMove(@DesktopWidth/3 , @DesktopHeight/3)
	Sleep(1000)
EndFunc

Func FindPlugInStatus()


	Local $x3=0,$y3=0
	Local $searchIconRed

	MouseMove(@DesktopWidth/2 , @DesktopHeight/2)
	Sleep(3000)
	ConsoleWrite("full path icon= "&  @ScriptDir&'\statusIcon\status_Red.bmp' & @CRLF)
	$searchIconRed = _ImageSearch(@ScriptDir&'\statusIcon\status_Red.bmp',1, $x3, $y3, 0)
	;convert position to zoom level 125%
	$x3 = ($x3 * 100) / 125
	$y3 = ($y3 * 100) / 125
	ConsoleWrite("searchIconred value = " & $searchIconRed & @CRLF)
	ConsoleWrite("x value = " & $x3 & "y value = " & $y3 & @CRLF)

Sleep(5000)
	if $searchIconRed = 1 then
		ConsoleWrite("icon red is found"  & @CRLF)
		MsgBox(0, "icon red is found at: ", "x = " & $x3 & " & y= " & $y3)
		MouseMove ( $x3 , $y3 )
		;MouseClick($MOUSE_CLICK_LEFT)
		MouseClick($MOUSE_CLICK_LEFT, $x3, $y3 ,1)
	Else
		MsgBox(0, "icon red is not found at: ", "x = " & $x3 & " & y= " & $y3)
	EndIf


EndFunc


Func findicon3()
	Local $_Image_1 = @ScriptDir & "\statusIcon\noMatch.bmp"
	Local $search1 = _ImageSearchByPath($_Image_1)
Local $x= ($search1[1] * 100) / 125, $y= ($search1[2] * 100) / 125

	If $search1[0] = 1 Then
		MsgBox(0, 'Ex 1 - Success', 'Image found:' & " X=" & $x & " Y=" & $y & @CRLF & $_Image_1)
	Else
		MsgBox(48, 'Ex 1 - Failed', 'Image not found')
	EndIf

	ConsoleWrite("step 1"  & @CRLF)
	;MsgBox(0, "Location", "x = " & $x & " & y= " & $y)
	if $search1[0] = 1  then
		ConsoleWrite("step 2"  & @CRLF)
		MouseMove ( $x , $y )
		;MouseClick($MOUSE_CLICK_LEFT)
		MouseClick($MOUSE_CLICK_LEFT, $x, $y ,1)
		Sleep(5000)
	EndIf

EndFunc

Func IsIconRed()
	Local $_Image_1 = @ScriptDir & "\statusIcon\red.bmp"
	Local $return = _ImageSearchByPath($_Image_1)
	Local $x= ($return[1] * 100) / 125, $y= ($return[2] * 100) / 125

	If $return[0] = 1 Then
		MouseMove ( $x , $y )
		Return True
	Else
		Return False
	EndIf
EndFunc

Func IsIconGreen()
	Local $_Image_1 = @ScriptDir & "\statusIcon\green.bmp"
	Local $return = _ImageSearchByPath($_Image_1)
	Local $x= ($return[1] * 100) / 125, $y= ($return[2] * 100) / 125

	If $return[0] = 1 Then
		MouseMove ( $x , $y )
		Return True
	Else
		Return False
	EndIf
EndFunc

Func IsIconYellow()
	Local $_Image_1 = @ScriptDir & "\statusIcon\noMatch.bmp"
	Local $return = _ImageSearchByPath($_Image_1)
	Local $x= ($return[1] * 100) / 125, $y= ($return[2] * 100) / 125

	If $return[0] = 1 Then
		MouseMove ( $x , $y )
		Return True
	Else
		Return False
	EndIf
EndFunc

Func GetMessFromPopup()
	Sleep(1000)

Local $winHwd = WinGetHandle("Mutual_DataBreachPrevention");
ConsoleWriteError("Win Handle is " & $winHwd & @CRLF) ; first check. You should seen a handle

Local  $sText2 = ControlGetText($winHwd, "", "[CLASS:Static; INSTANCE:2]")
Local $sText3 = ControlGetText("Mutual_DataBreachPrevention","","[CLASS:Static; INSTANCE:2]")

ConsoleWrite("text 2 = " & $sText2 & @CRLF)
ConsoleWrite("text 3 = " & $sText3 & @CRLF)
;~ Local $ctrlHwd = ControlGetHandle($winHwd, "", "[CLASS:Button; INSTANCE:2]") ;~ or Local $ctrlHwd = ControlGetHandle($winHwd, "", "4427")
;~ ConsoleWriteError("Control Handle is " & $ctrlHwd & @CRLF) ; second  check. You should seen a handle
;~ ControlSend($ctrlHwd, "", "", "{ENTER}")

ConsoleWrite("PID getmessage = " & $iPID & @CRLF)

return $sText2


EndFunc

Func IsMessContainsTxt($popupTxt, $txtToVerify)
	Local $iPosition  = StringInStr($popupTxt,$txtToVerify)
	If 	$iPosition  <> 0 then
		ConsoleWrite("The search string first appears at position: " & $iPosition & @CRLF)
	Else
		ConsoleWrite("search string not found" & @CRLF)
	EndIf
EndFunc



Func ClickSendMail()
	;$oMail.Send()
	MouseClick("left",39,225)
Sleep(1000)
MouseMove(@DesktopWidth/4 , @DesktopHeight/4)
ConsoleWrite("PID in SendMAil part = " & $iPID & @CRLF)
EndFunc

Func ClickSendMail2()

	Local $winHwd = WinGetHandle("[ACTIVE]")



	$iOriginal = Opt("MouseCoordMode",2)             ;Get the current MouseCoordMode    ;Change the MouseCoordMode to relative coords

$aPos = ControlGetPos($winHwd,"","[CLASS:Button; INSTANCE:1]")          ;Get the position of the given control
ConsoleWrite("position is " & "x=" & $aPos[0] & " y= " & $aPos[1]  & @CRLF)
MouseClick("left",$aPos[0],$aPos[1])
Opt("MouseCoordMode",$iOriginal)               ;Change the MouseCoordMode back to the original




	;_ControlMouseClick("Untitled - Message (HTML) ","","[CLASS:Button; INSTANCE:1]","left",1,10)
	Sleep(500)
	MouseMove(@DesktopWidth/4 , @DesktopHeight/4)
EndFunc

Func ClickCancelOnPopup()

;~ If (WinWaitActive("Mutual_DataBreachPrevention") = 0) Then
;~     MsgBox(0, "Timeout", "Window not seen")
;~     Exit
;~ EndIf
Sleep(1000)
Local $winHwd = WinGetHandle("Mutual_DataBreachPrevention");
ConsoleWriteError("Win Handle is " & $winHwd & @CRLF) ; first check. You should seen a handle
Local $ctrlHwd = ControlGetHandle($winHwd, "", "[CLASS:Button; INSTANCE:2]") ;~ or Local $ctrlHwd = ControlGetHandle($winHwd, "", "4427")
ConsoleWriteError("Control Handle is " & $ctrlHwd & @CRLF) ; second  check. You should seen a handle
ControlSend($ctrlHwd, "", "", "{ENTER}")

	EndFunc







#Region This part is to convert date time to name screenshot
Func ShowDateTime()
	;ConsoleWrite("date = " & _NowDate() & " time = " & _NowTime(5)	)
EndFunc

Func GetCurrentDateTimeForSaveFile()
	ConsoleWrite("date = " & _NowCalcDate() & @CRLF )
	 $date = ChangeDateFormatForSaveFile(_NowCalcDate())
	ConsoleWrite("date formated = " & $date & @CRLF )

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
    ;If $date == '' Then Return ''
    ;$ret = StringMid($date,7,4) & '-' & StringLeft($date,2) & '-'  & StringMid($date,4,2)
;~  If StringLen($date) > 10 Then $ret &= StringMid($date,11) ; optionally including time
	;Return $ret

	If $date == '' Then Return ''
    $ret = StringLeft($date,4) & '-' & StringMid($date,6,2) & '-'  & StringMid($date,9,2)
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
ConsoleWrite("PID  in screenshot part = " & $iPID & @CRLF)
	return $filePath
EndFunc






#Region  _RefreshSystemTray($nDealy = 1000)
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




#Region GetAllWindowsControls with filter: https://www.autoitscript.com/forum/topic/164226-get-all-windows-controls/


;~ ConsoleWrite("Make your window active!" & @CRLF)
;~ Sleep(5000)

;~ GetAllWindowsControls(WinGetHandle("[ACTIVE]"))

Func GetAllWindowsControls($hCallersWindow, $bOnlyVisible=Default, $sStringIncludes=Default, $sClass=Default)
    If Not IsHWnd($hCallersWindow) Then
        ConsoleWrite("$hCallersWindow must be a handle...provided=[" & $hCallersWindow & "]" & @CRLF)
        Return False
    EndIf
    ; Get all list of controls
    If $bOnlyVisible = Default Then $bOnlyVisible = False
    If $sStringIncludes = Default Then $sStringIncludes = ""
    If $sClass = Default Then $sClass = ""
    $sClassList = WinGetClassList($hCallersWindow)

    ; Create array
    $aClassList = StringSplit($sClassList, @CRLF, 2)

    ; Sort array
    _ArraySort($aClassList)
    _ArrayDelete($aClassList, 0)

    ; Loop
    $iCurrentClass = ""
    $iCurrentCount = 1
    $iTotalCounter = 1

    If StringLen($sClass)>0 Then
        For $i = UBound($aClassList)-1 To 0 Step - 1
            If $aClassList[$i]<>$sClass Then
                _ArrayDelete($aClassList,$i)
            EndIf
        Next
    EndIf

    For $i = 0 To UBound($aClassList) - 1
        If $aClassList[$i] = $iCurrentClass Then
            $iCurrentCount += 1
        Else
            $iCurrentClass = $aClassList[$i]
            $iCurrentCount = 1
        EndIf

        $hControl = ControlGetHandle($hCallersWindow, "", "[CLASSNN:" & $iCurrentClass & $iCurrentCount & "]")
        $text = StringRegExpReplace(ControlGetText($hCallersWindow, "", $hControl), "[\n\r]", "{@CRLF}")
        $aPos = ControlGetPos($hCallersWindow, "", $hControl)
        $sControlID = _WinAPI_GetDlgCtrlID($hControl)
        $bIsVisible = ControlCommand($hCallersWindow, "", $hControl, "IsVisible")
        If $bOnlyVisible And Not $bIsVisible Then
            $iTotalCounter += 1
            ContinueLoop
        EndIf

        If StringLen($sStringIncludes) > 0 Then
            If Not StringInStr($text, $sStringIncludes) Then
                $iTotalCounter += 1
                ContinueLoop
            EndIf
        EndIf





        If IsArray($aPos) Then
            ;ConsoleWrite("Func=[GetAllWindowsControls]: ControlCounter=[" & StringFormat("%3s", $iTotalCounter) & "] ControlID=[" & StringFormat("%5s", $sControlID) & "] Handle=[" & StringFormat("%10s", $hControl) & "] ClassNN=[" & StringFormat("%19s", $iCurrentClass & $iCurrentCount) & "] XPos=[" & StringFormat("%4s", $aPos[0]) & "] YPos=[" & StringFormat("%4s", $aPos[1]) & "] Width=[" & StringFormat("%4s", $aPos[2]) & "] Height=[" & StringFormat("%4s", $aPos[3]) & "] IsVisible=[" & $bIsVisible & "] Text=[" & $text & "]." & @CRLF)
			FileWriteLine( @ScriptFullPath & "_listOfWinCtrl", "Func=[GetAllWindowsControls]: ControlCounter=[" & StringFormat("%3s", $iTotalCounter) & "] ControlID=[" & StringFormat("%5s", $sControlID) & "] Handle=[" & StringFormat("%10s", $hControl) & "] ClassNN=[" & StringFormat("%19s", $iCurrentClass & $iCurrentCount) & "] XPos=[" & StringFormat("%4s", $aPos[0]) & "] YPos=[" & StringFormat("%4s", $aPos[1]) & "] Width=[" & StringFormat("%4s", $aPos[2]) & "] Height=[" & StringFormat("%4s", $aPos[3]) & "] IsVisible=[" & $bIsVisible & "] Text=[" & $text & "]." & @CRLF )
			FileWriteLine( @ScriptFullPath & "_listOfWinCtrl", "------------------------END of log-------------------------- " & @CRLF)
		Else
			FileWriteLine( @ScriptFullPath & "_listOfWinCtrl", "Func=[GetAllWindowsControls]: ControlCounter=[" & StringFormat("%3s", $iTotalCounter) & "] ControlID=[" & StringFormat("%5s", $sControlID) & "] Handle=[" & StringFormat("%10s", $hControl) & "] ClassNN=[" & StringFormat("%19s", $iCurrentClass & $iCurrentCount) & "] XPos=[winclosed] YPos=[winclosed] Width=[winclosed] Height=[winclosed] Text=[" & $text & "]." & @CRLF)
			FileWriteLine( @ScriptFullPath & "_listOfWinCtrl", "------------------------END of log-------------------------- " & @CRLF)
			;ConsoleWrite("Func=[GetAllWindowsControls]: ControlCounter=[" & StringFormat("%3s", $iTotalCounter) & "] ControlID=[" & StringFormat("%5s", $sControlID) & "] Handle=[" & StringFormat("%10s", $hControl) & "] ClassNN=[" & StringFormat("%19s", $iCurrentClass & $iCurrentCount) & "] XPos=[winclosed] YPos=[winclosed] Width=[winclosed] Height=[winclosed] Text=[" & $text & "]." & @CRLF)

		EndIf

        If Not WinExists($hCallersWindow) Then ExitLoop
        $iTotalCounter += 1
    Next
EndFunc   ;==>GetAllWindowsControls
#EndRegion


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