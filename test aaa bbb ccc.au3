#include <Date.au3>

$nowDate = _NowDate()
ConsoleWrite("now date = " & $nowDate & @CRLF)

$date = ChangeDateFormatForSaveFile(_NowDate())
ConsoleWrite("converted date = " & $date & @CRLF)
; dd.mm.yyyy --> yyyy-mm-dd
Func ChangeDateFormatForSaveFile($date)
    If $date == '' Then Return ''
    $ret = StringMid($date,7,4) & '-' & StringLeft($date,2) & '-'  & StringMid($date,4,2)
;~  If StringLen($date) > 10 Then $ret &= StringMid($date,11) ; optionally including time
    Return $ret
EndFunc