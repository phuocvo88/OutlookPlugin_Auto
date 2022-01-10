#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <FunctionLib.au3>
#Include <Array.au3>


Example()


;GetAllWindowsControls(WinGetHandle("[ACTIVE]"))


Func Example()
   $data = GetAllWindowsControls(WinGetHandle("Untitled - Message (HTML) "))
	If FileExists( @ScriptFullPath & ".testTemp" ) then
		FileDelete( @ScriptFullPath & ".testTemp" )
	EndIf

FileWriteLine( @ScriptFullPath & "_testTemp", $data )

EndFunc   ;==>Example






#cs
#Include <WinAPI.au3>



; Example ######################################################################################################
;#Include <Array.au3>

Local $list = _WinGetControlList(WinGetHandle("[ACTIVE]") )
If $list = 0 Then Exit MsgBox(0, "Controls list" , "No control found for " & WinGetTitle("[ACTIVE]") )

_ArrayDisplay($list, "Controls list", Default, 0, "|", "CLASS|ClassnameNN|Advanced Mode|Handle|Text|ID|Position in window|Size|Position in screen|Visible")
; ##############################################################################################################




Func _WinGetControlList($sTitle, $sText = "")
    Local $n = 0, $iCount
    Local $aControlPos, $iInstance, $aScreenPos

    Local $sClassList = WinGetClassList($sTitle, $sText)
    If $sClassList = "" Then Return SetError(1, 0, 0)

    Local $aResult = StringRegExp($sClassList, "(\N+)", 3)
    If NOT IsArray($aResult) Then Return SetError(1, 0, 0)
    Redim $aResult[UBound($aResult)][10]

    Local $aClasses = StringRegExp($sClassList, "(?s)(?:\A|\R)(\N+)(?=\R)(?!(?:\R\N+)*\R\1\R)", 3)
    For $i = 0 To UBound($aClasses) - 1
        StringRegExpReplace($sClassList, "\Q" & $aClasses[$i] & "\E\R", "")
        $iCount = @extended

        For $iInstance = 1 To $iCount
            $aResult[$n][0] = $aClasses[$i]                                                  ; Class
            $aResult[$n][1] = $aClasses[$i] & $iInstance                                     ; ClassnameNN
            $aResult[$n][2] = "[CLASS:" & $aClasses[$i] & "; INSTANCE:" & $iInstance& "]"    ; Advanced mode
            $aResult[$n][3] = ControlGetHandle($sTitle, $sText, $aResult[$n][2])             ; Handle
            $aResult[$n][4] = ControlGetText($sTitle, $sText, $aResult[$n][3] )              ; Text
            $aResult[$n][5] = _WinAPI_GetDlgCtrlID($aResult[$n][3])                          ; ID

            $aControlPos = ControlGetPos($sTitle, $sText, $aResult[$n][3])
            If IsArray($aControlPos) Then
                $aResult[$n][6] = "X=" & $aControlPos[0] & " ; Y=" & $aControlPos[1]          ; Position X,Y (in the Window)
                $aResult[$n][7] = "W=" & $aControlPos[2] & " ; H=" & $aControlPos[3]          ; Width and Height
            EndIf

            $aScreenPos = WinGetpos($aResult[$n][2])
            If IsArray($aScreenPos) Then
                $aResult[$n][8] = "X=" & $aScreenPos[0] & " ;  Y=" & $aScreenPos[1]           ; Position X,Y (relative to screen)
            EndIf

            $aResult[$n][9] = ControlCommand($sTitle, $sText, $aResult[$n][3], "IsVisible")   ; Visible

            $n += 1
        Next
    Next

    Return $aResult
EndFunc
#ce