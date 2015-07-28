#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3> ; Required for _ArrayDisplay.

$sText = FileRead('C:\coding\meta\Backup_Mo_Sa_22_Uhr_alle VMs.vbm')
global $RegExp = StringRegExp($sText, '(?s)<string>(.*?)</string>', 3)
 Local $iRows = UBound($RegExp, $UBOUND_ROWS)
 ConsoleWrite($iRows&@CR)
 _ArrayDisplay($RegExp)