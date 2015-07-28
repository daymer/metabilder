#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Comment=beta
#AutoIt3Wrapper_Res_Fileversion=0.2.0.0
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <Date.au3>
#include <MsgBoxConstants.au3>
#include <File.au3>
#include <Array.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <TabConstants.au3>
#include <WindowsConstants.au3>
#include <GUIListBox.au3>
#include <GuiStatusBar.au3>
#include <GuiToolbar.au3>
#include <ToolbarConstants.au3>

If Not IsAdmin() Then
    MsgBox($MB_SYSTEMMODAL, "Warning","Admin rights are not detected."& @CR& 'You need to be admin to run this tool')
	Exit
EndIf
#Region main paths
$ProgrammDir = 'C:\VeeamMetaBilder\'
$temp_dir = @TempDir&'\VeeamMetaBilder\'
If Not FileExists($ProgrammDir) Then DirCreate($ProgrammDir)
If Not FileExists($temp_dir) Then DirCreate($temp_dir)
#EndRegion
#Region global vars
#EndRegion
$Debug_TB = False
$storages_reparse_test = False


Global $hToolbar, $iMemo
Global $iItem ; Командный идентификатор кнопки связанный с уведомлением.
Global Enum $idNew = 1000, $idOpen, $idSave

_Main()

Func _Main()
    Local $hGUI, $aSize

    ; Создаёт GUI
    $hGUI = GUICreate(StringTrimRight(@ScriptName, 4), 600, 400)
    $hToolbar = _GUICtrlToolbar_Create($hGUI)
    $aSize = _GUICtrlToolbar_GetMaxSize($hToolbar)

    ;$iMemo = GUICtrlCreateEdit("", 2, $aSize[1] + 20, 596, 396 - ($aSize[1] + 20), $WS_VSCROLL)
    GUICtrlSetFont($iMemo, 9, 400, 0, "Courier New")
    GUISetState()
    GUIRegisterMsg($WM_NOTIFY, "_WM_NOTIFY")

    ; Добавляет стандартный системный bitmaps
    _GUICtrlToolbar_AddBitmap($hToolbar, 1, -1, $IDB_STD_LARGE_COLOR)

    ; Добавляет кнопки
    _GUICtrlToolbar_AddButton($hToolbar, $idNew, $STD_FILENEW)
    _GUICtrlToolbar_AddButton($hToolbar, $idOpen, $STD_FILEOPEN)
    _GUICtrlToolbar_AddButton($hToolbar, $idSave, $STD_FILESAVE)
    _GUICtrlToolbar_AddButtonSep($hToolbar)
    _GUICtrlToolbar_AddButton($hToolbar, $idHelp, $STD_HELP)

    ; Цикл выполняется, пока окно не будет закрыто
    Do
    Until GUIGetMsg() = $GUI_EVENT_CLOSE

EndFunc   ;==>_Main

; Записывает строку в элемент для заметок
Func MemoWrite($sMessage = "")
    GUICtrlSetData($iMemo, $sMessage & @CRLF, 1)
EndFunc   ;==>MemoWrite

; Обработчик уведомлений - WM_NOTIFY
Func _WM_NOTIFY($hWndGUI, $MsgID, $wParam, $lParam)
    #forceref $hWndGUI, $MsgID, $wParam
    Local $tNMHDR, $hwndFrom, $code, $i_idNew, $dwFlags, $i_idOld ; , $idFrom
    Local $tNMTBHOTITEM
    $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
    $hwndFrom = DllStructGetData($tNMHDR, "hWndFrom")
    ; $idFrom = DllStructGetData($tNMHDR, "IDFrom")
    $code = DllStructGetData($tNMHDR, "Code")
    Switch $hwndFrom
        Case $hToolbar
            Switch $code
                Case $NM_LDOWN
                    ;----------------------------------------------------------------------------------------------
                   ; MemoWrite("$NM_LDOWN: Кликнут элемент: " & $iItem & " с индексом: " & _GUICtrlToolbar_CommandToIndex($hToolbar, $iItem))
					if $iItem = $idOpen Then
						$sFileOpenDialog_vbm=Import_some_vbm()

					#Region REPARSING META

if Import_some_vbm <> 0 Then
   ConsoleWrite("$NM_LDOWN: Import_some_vbm " & $sFileOpenDialog_vbm &@CR)
$sText = FileRead($sFileOpenDialog_vbm)
local $RegExp_string = StringRegExp($sText, '(?s)<string>(.*?)</string>', 3)
local $numm_string = UBound($RegExp_string, $UBOUND_ROWS)
local $RegExp_Vbm_Id = StringRegExp($sText, '(?s)&lt;Vbm Id="(.*?)" JobId="', 3)
local $numm_Vbm_Id = UBound($RegExp_Vbm_Id, $UBOUND_ROWS)
if $numm_Vbm_Id > 1 Then MsgBox($MB_SYSTEMMODAL, "Error", "More than 1 Vbm_Id.")
local $RegExp_JobId = StringRegExp($sText, '(?s)" JobId="(.*?)" JobName="', 3)

local $RegExp_JobName = StringRegExp($sText, '(?s)JobName="(.*?)" JobType="', 3)
local $RegExp_JobType = StringRegExp($sText, '(?s)JobType="(.*?)" JobSourceType="', 3)
local $RegExp_JobSourceType = StringRegExp($sText, '(?s)JobSourceType="(.*?)" Platform="', 3)
local $RegExp_Platform = StringRegExp($sText, '(?s)Platform="(.*?)" CreationTimeUtc="', 3)


$RegExp_CreationTimeUtc = StringRegExp($sText, '(?s)CreationTimeUtc="(.*?)" FormatVersion=', 3)
local $numm_CreationTimeUtc = UBound($RegExp_CreationTimeUtc, $UBOUND_ROWS)
if $numm_CreationTimeUtc  = 0 Then local $RegExp_CreationTimeUtc = StringRegExp($sText, '(?s) CreationTimeUtc="(.*?)" BackupCreationTime="', 3)

local $RegExp_FormatVersion = StringRegExp($sText, '(?s)FormatVersion="(.*?)" BackupCreationTime=', 3)
local $numm_FormatVersion = UBound($RegExp_FormatVersion, $UBOUND_ROWS)
local $RegExp_BackupCreationTime = StringRegExp($sText, '(?s) BackupCreationTime="(.*?)"&gt;&lt;Hosts', 3)
local $RegExp_Hosts = StringRegExp($sText, '(?s)&lt;Hosts&gt;(.*?)&lt;/Hosts&gt;', 3)
;main arrays
local $RegExp_Storages = StringRegExp($sText, '(?s)&lt;Storages&gt;(.*?)&lt;/Storages&gt;', 3)
local $RegExp_Storages_str = $RegExp_Storages[0]
local $RegExp_JobPoints = StringRegExp($sText, '(?s)&lt;JobPoints&gt;(.*?)&lt;/JobPoints&gt;', 3)
local $RegExp_Vms = StringRegExp($sText, '(?s)&lt;Vms&gt;(.*?)&lt;/Vms&gt;', 3)
local $RegExp_VmPoints = StringRegExp($sText, '(?s)&lt;VmPoints&gt;(.*?)&lt;/VmPoints&gt;', 3)
;/main arrays
_Storages_reparse($RegExp_Storages_str)




#Region repars vars
global $numm_Vbm_Id = $RegExp_Vbm_Id[0]
global $VM_ids_in_meta = $numm_string
global $numm_JobId = $RegExp_JobId[0]
global $numm_JobName = $RegExp_JobName[0]
global $numm_JobType = $RegExp_JobType[0]
global $numm_SourceType = $RegExp_JobSourceType[0]
global $numm_Platform = $RegExp_Platform[0]
global $numm_CreationTimeUtc = $RegExp_CreationTimeUtc[0]
if $numm_FormatVersion = 0 Then
  global $numm_FormatVersion = 'not_set'
   Else
global $numm_FormatVersion = $RegExp_FormatVersion [0]
EndIf
global $numm_BackupCreationTime = $RegExp_BackupCreationTime[0]
global $array_Hosts = $RegExp_Hosts
global $array_VM_ids = $RegExp_string
global $array_Storages = $RegExp_Storages
global $array_JobPoints = $RegExp_JobPoints
global $array_Vms = $RegExp_Vms
global $array_VmPoints = $RegExp_VmPoints
#EndRegion



EndIf
#EndRegion
					EndIf
                    ;----------------------------------------------------------------------------------------------
                Case $TBN_HOTITEMCHANGE
                  $tNMTBHOTITEM = DllStructCreate($tagNMTBHOTITEM, $lParam)
                   $i_idOld = DllStructGetData($tNMTBHOTITEM, "idOld")
                    $i_idNew = DllStructGetData($tNMTBHOTITEM, "idNew")
                    $iItem = $i_idNew
                    $dwFlags = DllStructGetData($tNMTBHOTITEM, "dwFlags")
                    If BitAND($dwFlags, $HICF_LEAVING) = $HICF_LEAVING Then
                        ;MemoWrite("$HICF_LEAVING: " & $i_idOld)
                    Else
                        Switch $i_idNew
                           Case $idNew
                                ;----------------------------------------------------------------------------------------------
                    ;            MemoWrite("$TBN_HOTITEMCHANGE: $idNew")
                                ;----------------------------------------------------------------------------------------------
                            Case $idOpen
                                ;----------------------------------------------------------------------------------------------
                    ;            MemoWrite("$TBN_HOTITEMCHANGE: $idOpen")
                                ;----------------------------------------------------------------------------------------------
                            Case $idSave
                                ;----------------------------------------------------------------------------------------------
                    ;            MemoWrite("$TBN_HOTITEMCHANGE: $idSave")
                                ;----------------------------------------------------------------------------------------------
                            Case $idHelp
                                ;----------------------------------------------------------------------------------------------
                    ;            MemoWrite("$TBN_HOTITEMCHANGE: $idHelp")
                                ;----------------------------------------------------------------------------------------------
                        EndSwitch
                    EndIf
            EndSwitch
    EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc   ;==>_WM_NOTIFY



#Region test vars
global $job_name ='test_job_name'
#EndRegion

#Region vbm_read fuctions
#Region main arrays reparse Func
Func _Storages_reparse($RegExp_Storages_str)
local $Storages_array_rep[1][13]
$Storages_array_rep[0][0] = 'Storage Id'
$Storages_array_rep[0][1] = 'FileName'
$Storages_array_rep[0][2] = 'CreationTimeUtc'
$Storages_array_rep[0][3] = 'BlockSize'
$Storages_array_rep[0][4] = 'State'
$Storages_array_rep[0][5] = 'LinkId'
$Storages_array_rep[0][6] = 'PartialIncrement'
$Storages_array_rep[0][7] = 'GfsPeriod'
$Storages_array_rep[0][8] = 'CreationMode'
$Storages_array_rep[0][9] = 'BackupSize'
$Storages_array_rep[0][10] = 'DataSize'
$Storages_array_rep[0][11] = 'DedupRatio'
$Storages_array_rep[0][12] = 'CompressRatio'

local $RegExp_Storage_Id = StringRegExp($RegExp_Storages_str, '(?s)&lt;Storage Id="(.*?)" FileName="', 3)
local $numm_Storage_Id = UBound($RegExp_Storage_Id, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Storage_Id step +1
   local $Storages_array_rep_temp[1][13]
   $Storages_array_rep_temp[0][0] =  $RegExp_Storage_Id[$i]
   _ArrayConcatenate($Storages_array_rep,$Storages_array_rep_temp)
   Next

local $RegExp_Storage_FileName = StringRegExp($RegExp_Storages_str, '(?s)" FileName="(.*?)" CreationTimeUtc=', 3)
local $numm_Storage_FileName = UBound($RegExp_Storage_FileName, $UBOUND_ROWS) -1
For $i = 0 To $numm_Storage_FileName step +1
   $Storages_array_rep[$i+1][1] =  $RegExp_Storage_FileName[$i]
Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)" CreationTimeUtc="(.*?)" BlockSize=', 3)
local $numm_Storage_temp = UBound($RegExp_Storage_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][2] =  $RegExp_Storage_temp[$i]
Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)" BlockSize="(.*?)" State=', 3)
local $numm_Storage_temp = UBound($RegExp_Storage_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][3] =  $RegExp_Storage_temp[$i]
Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)" State="(.*?)" PartialIncrement=', 3)
local $numm_Storage_temp = UBound($RegExp_Storage_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][4] =  $RegExp_Storage_temp[$i]
Next

local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)" State="(.*?)" PartialIncrement=', 3)
local $numm_Storage_temp = UBound($RegExp_Storage_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][5] =  $RegExp_Storage_temp[$i]


Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)" PartialIncrement="(.*?)" GfsPeriod=', 3)
local $numm_Storage_temp = UBound($RegExp_Storage_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][6] =  $RegExp_Storage_temp[$i]
Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)" GfsPeriod="(.*?)" CreationMode=', 3)
local $numm_Storage_temp = UBound($RegExp_Storage_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][7] =  $RegExp_Storage_temp[$i]
Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)" CreationMode="(.*?)"&gt;&lt;Stats&gt;', 3)
local $numm_Storage_temp = UBound($RegExp_Storage_temp, $UBOUND_ROWS) -1
   For $i = 0 To $numm_Storage_temp step +1
   $text = $RegExp_Storage_temp[$i]
   $trim = StringInStr($text, '" LinkId="')
   if not $trim=0 Then
  $LinkId_temp =  StringTrimLeft($text, $trim+9)
    $Storages_array_rep[$i+1][5] =  $LinkId_temp
	  $Storages_array_rep[$i+1][8] =  StringLeft($text, $trim-1)
Else
   $LinkId_temp = '0000000-0000-0000-0000-000000000000'
   $Storages_array_rep[$i+1][5] =  $LinkId_temp
   $Storages_array_rep[$i+1][8] =  $RegExp_Storage_temp[$i]
EndIf
Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)&lt;BackupSize&gt;(.*?)&lt;/BackupSize&gt;', 3)
local $numm_Storage_temp = UBound($RegExp_Storage_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][9] =  $RegExp_Storage_temp[$i]
Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)&lt;DataSize&gt;(.*?)&lt;/DataSize&gt;', 3)
local $numm_Storage_temp = UBound($RegExp_Storage_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][10] =  $RegExp_Storage_temp[$i]
Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)&lt;DedupRatio&gt;(.*?)&lt;/DedupRatio&gt;', 3)
local $numm_Storage_temp = UBound($RegExp_Storage_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][11] =  $RegExp_Storage_temp[$i]
Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)&lt;CompressRatio&gt;(.*?)&lt;/CompressRatio&gt;', 3)
local $numm_Storage_temp = UBound($RegExp_Storage_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][12] =  $RegExp_Storage_temp[$i]
Next
if $storages_reparse_test = True  Then
FileDelete('Storages_reparse_test.txt')
_FileWriteFromArray('Storages_array_rep.txt', $Storages_array_rep)
EndIf
ConsoleWrite('$Storages_array_rep is done'&@cr)
Return $Storages_array_rep
EndFunc
#EndRegion



#EndRegion
#Region vbm_create fuctions

Func _create_blank_vbm($job_name)
	local $vbm_file=$job_name&".vbm"
	If Not FileExists($vbm_file) Then
		_FileCreate($vbm_file)
		Else
		MsgBox($MB_SYSTEMMODAL, "Error", " Error Creating/Resetting vbm file: " & @error)
		 $stop = 'true'
	EndIf
	if not IsDeclared('stop') Then
#Region vbm content
	FileWrite($vbm_file,'<?xml version="1.0" encoding="UTF-8"?>')
	FileWrite($vbm_file,'<BackupMeta>')
	#Region header
			FileWrite($vbm_file,'<VmObjects>')
				FileWrite($vbm_file,'<VmId>')
				FileWrite($vbm_file,'</VmId>')
			FileWrite($vbm_file,'</VmObjects>')
	#EndRegion
	#Region KeySets
	FileWrite($vbm_file,'<KeySets />')
	FileWrite($vbm_file,'<Vbm KeySetId="">')
	#EndRegion
	EndIf
#EndRegion
EndFunc
#EndRegion



Func Import_some_vbm()
    ; Create a constant variable in Local scope of the message to display in FileOpenDialog.
    Local Const $sMessage = "Import VBM"

    ; Display an open dialog to select a list of file(s).
    Local $sFileOpenDialog = FileOpenDialog($sMessage, @WindowsDir & "\", "Meta File (*.vbm;)", $FD_FILEMUSTEXIST + $FD_MULTISELECT)
    If @error Then
        ; Display the error message.
        MsgBox($MB_SYSTEMMODAL, "", "No file(s) were selected.")
;Exit
        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)
    Else
        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)

        ; Replace instances of "|" with @CRLF in the string returned by FileOpenDialog.
        $sFileOpenDialog_vbm = StringReplace($sFileOpenDialog, "|", @CRLF)

        ; Display the list of selected files.
        ;MsgBox($MB_SYSTEMMODAL, "", "You chose the following files:" & @CRLF & $sFileOpenDialog)
		Return($sFileOpenDialog_vbm)
    EndIf
EndFunc   ;==>Example
