#include <Date.au3>
#include <FunctionLib.au3>

Local $iMsg = "Test record"
ConsoleWrite("temp dir = " &   @TempDir          & @CRLF )
;FileWriteLine(@TempDir & "\Pgm.log", _NowCalcDate() & " :" & $iMsg)
ConsoleWrite("nowcalcDate = " & _NowCalcDate() & @CRLF )



Local $testDate2 =   convertDate(_NowCalcDate())
ConsoleWrite("convert date = " & $testDate2 & @CRLF )




Func convertDate($date)
    If $date == '' Then Return ''
    $ret = StringLeft($date,4) & '-' & StringMid($date,6,2) & '-'  & StringMid($date,9,2)
;~  If StringLen($date) > 10 Then $ret &= StringMid($date,11) ; optionally including time
    Return $ret
EndFunc