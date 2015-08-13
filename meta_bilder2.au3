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
#include "Table.au3"
#include <StaticConstants.au3>
#include <GuiButton.au3>
#include <ColorConstants.au3>
$version_exe = '0.2.0.0'
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
#Region test vars
global $job_name ='test_job_name'
global $Debug_TB = False
global $storages_reparse_test = False
global $VMPoints_reparse_test = False
global $points_reparse_test = False
global $hosts_reparse_test = False
global $vms_reparse_test = False
#EndRegion

_Main()
Func _Main()
    Local $hGUI, $aSize
    If @AutoItX64 Then $sWow64 = "\Wow6432Node"
    Local $sPath = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE" & $sWow64 & "\AutoIt v3\AutoIt", "InstallDir") & "\Examples\GUI\Advanced\Images"
    ; —оздаЄт GUI
    $hGUI = GUICreate(StringTrimRight("Veeam Meta Parser "&$version_exe, 2), 1000, 750)
GUISetBkColor(0xF0F4F9)
$tools = GUICtrlCreateGroup("", -1, -10, 1000, 55)
$Group2 = GUICtrlCreateGroup("", 16, 48, 465, 670)
$Group3 = GUICtrlCreateGroup("", 485, 48, 490, 670)
  $open = GUICtrlCreateIcon("shell32.dll", 5, 5, 5, 32, 32)
    $save= GUICtrlCreateIcon("shell32.dll", 7, 37, 5, 32, 32)
	GUICtrlCreateLabel('|', 74, 7, 5, 30, $SS_ETCHEDVERT)
	  $help = GUICtrlCreateIcon("shell32.dll", 24, 79, 5, 32, 32)
$tools = GUICtrlCreateGroup("", 0, 715, 1000, 40)
Global $idProgressbar1 = GUICtrlCreateProgress(745, 725, 250, 20)
GUICtrlSetState($idProgressbar1, $GUI_HIDE )
    GUICtrlSetColor(-1, 32250); not working with Windows XP Style
GUISetState(@SW_SHOW)




     While 1
        $GUIMsg = GUIGetMsg()

        Switch $GUIMsg
            Case $GUI_EVENT_CLOSE
            			ExitLoop

			Case $open
				GUICtrlSetState($idProgressbar1, $GUI_show )
				_Parse()
				GUICtrlSetData($idProgressbar1, 100)
				Sleep(1000)
				GUICtrlSetState($idProgressbar1, $GUI_HIDE )
$Group2 = GUICtrlCreateGroup("", 16, 48, 465, 670)
$Checkbox1 = GUICtrlCreateCheckbox("", 33, 70, 15, 15)
$Checkbox2 = GUICtrlCreateCheckbox("", 78, 100, 15, 15)
$Checkbox3 = GUICtrlCreateCheckbox("", 78, 130, 15, 15)
$Checkbox4 = GUICtrlCreateCheckbox("", 78, 160, 15, 15)
$Checkbox5 = GUICtrlCreateCheckbox("", 78, 190, 15, 15)
$Checkbox6 = GUICtrlCreateCheckbox("", 33, 220, 15, 15)
$Button1 = GUICtrlCreateLabel("Eugene Test NetApp OnHost2015-07-14T135429.vbk", 50, 65, 410, 25,  BitOR($SS_Left, $SS_CENTERIMAGE))
$Button2 = GUICtrlCreateLabel("Eugene Test NetApp OnHost2015-07-14T191524.vib", 95, 95, 363, 25,  BitOR($SS_Left, $SS_CENTERIMAGE))
$Button3 = GUICtrlCreateLabel("Eugene Test NetApp OnHost2015-07-14T191524.vib", 95, 125, 363, 25, BitOR($SS_Left, $SS_CENTERIMAGE))
$Button4 = GUICtrlCreateLabel("Eugene Test NetApp OnHost2015-07-14T191524.vib", 95, 155, 363, 25, BitOR($SS_Left, $SS_CENTERIMAGE))
$Button5 = GUICtrlCreateLabel("Eugene Test NetApp OnHost2015-07-14T191524.vib", 95, 185, 363, 25, BitOR($SS_Left, $SS_CENTERIMAGE))
$Button6 = GUICtrlCreateLabel("Eugene Test NetApp OnHost2015-07-14T135429.vbk", 51, 215, 410, 25, BitOR($SS_Left, $SS_CENTERIMAGE))
GUICtrlCreateGroup("", -99, -99, 1, 1)


$Group3 = GUICtrlCreateGroup("", 485, 48, 490, 670)
$Button7 = GUICtrlCreateLabel("VM: aff77f0e-a090-46ad-8463-2846ca8b7c57 Point: 0856ea98-a473-4072-b256-341572aa4a4d", 499, 65, 470, 25, BitOR($SS_Left, $SS_CENTERIMAGE))
$Button7 = GUICtrlCreateLabel("VM: aff77f0e-a090-46ad-8463-2846ca8b7c57 Point: 0856ea98-a473-4072-b256-341572aa4a4d", 499, 85, 470, 25, BitOR($SS_Left, $SS_CENTERIMAGE))
$Button7 = GUICtrlCreateLabel("VM: aff77f0e-a090-46ad-8463-2846ca8b7c57 Point: 0856ea98-a473-4072-b256-341572aa4a4d", 499, 105, 470, 25, BitOR($SS_Left, $SS_CENTERIMAGE))
$Button7 = GUICtrlCreateLabel("VM: aff77f0e-a090-46ad-8463-2846ca8b7c57 Point: 0856ea98-a473-4072-b256-341572aa4a4d", 499, 125, 470, 25, BitOR($SS_Left, $SS_CENTERIMAGE))
$Button7 = GUICtrlCreateLabel("VM: aff77f0e-a090-46ad-8463-2846ca8b7c57 Point: 0856ea98-a473-4072-b256-341572aa4a4d", 499, 145, 470, 25, BitOR($SS_Left, $SS_CENTERIMAGE))
$Button7 = GUICtrlCreateLabel("VM: aff77f0e-a090-46ad-8463-2846ca8b7c57 Point: 0856ea98-a473-4072-b256-341572aa4a4d", 499, 165, 470, 25, BitOR($SS_Left, $SS_CENTERIMAGE))
GUICtrlCreateGroup("", -99, -99, 1, 1)


			EndSwitch
			WEnd

EndFunc   ;==>_Main


Func _Parse()

						$sFileOpenDialog_vbm=Import_some_vbm()
						 ConsoleWrite("$NM_LDOWN: Import_some_vbm " & $sFileOpenDialog_vbm &@CR)
						if Not $sFileOpenDialog_vbm = 0 Then
						  	  $sText = FileRead($sFileOpenDialog_vbm)

;;;;
					#Region REPARSING META
;;;;


#Region $BackupMeta version
global $BackupMeta_version = StringRegExp($sText, '(?s)<BackupMeta Version="(.*?)">', 3)
$BackupMeta_version = $BackupMeta_version[0]
ConsoleWrite('$BackupMeta_version = '&$BackupMeta_version&@cr)
if $BackupMeta_version = '0' then
	MsgBox($MB_SYSTEMMODAL, "Error", "VBM version isn't found")
Else
#EndRegion
#Region pre-header
local $RegExp_string = StringRegExp($sText, '(?s)<string>(.*?)</string>', 3)
local $numm_string = UBound($RegExp_string, $UBOUND_ROWS)
#EndRegion

#Region KeySets
local $RegExp_Vbm_Id = StringRegExp($sText, '(?s)&lt;Vbm Id="(.*?)" JobId="', 3)
local $numm_Vbm_Id = UBound($RegExp_Vbm_Id, $UBOUND_ROWS)
if $numm_Vbm_Id > 1 Then MsgBox($MB_SYSTEMMODAL, "Error", "More than 1 Vbm_Id.")
local $RegExp_JobId = StringRegExp($sText, '(?s)" JobId="(.*?)" JobName="', 3)
local $RegExp_JobName = StringRegExp($sText, '(?s)JobName="(.*?)" JobType="', 3)
local $RegExp_JobType = StringRegExp($sText, '(?s)JobType="(.*?)" JobSourceType="', 3)
local $RegExp_JobSourceType = StringRegExp($sText, '(?s)JobSourceType="(.*?)" Platform="', 3)
local $RegExp_Platform = StringRegExp($sText, '(?s)Platform="(.*?)" CreationTimeUtc="', 3)
local $RegExp_CreationTimeUtc = StringRegExp($sText, '(?s)CreationTimeUtc="(.*?)"', 3)
local $RegExp_FormatVersion = StringRegExp($sText, '(?s)FormatVersion="(.*?)" BackupCreationTime=', 3)
local $numm_FormatVersion = UBound($RegExp_FormatVersion, $UBOUND_ROWS)
local $RegExp_BackupCreationTime = StringRegExp($sText, '(?s) BackupCreationTime="(.*?)"&gt;&lt;Hosts', 3)
#EndRegion
#Region Meta global fields

local $RegExp_Hosts = StringRegExp($sText, '(?s)&lt;Hosts&gt;(.*?)&lt;/Hosts&gt;', 3)
local $RegExp_Storages = StringRegExp($sText, '(?s)&lt;Storages&gt;(.*?)&lt;/Storages&gt;', 3)
local $RegExp_JobPoints = StringRegExp($sText, '(?s)&lt;JobPoints&gt;(.*?)&lt;/JobPoints&gt;', 3) ;<---- verisions!
local $RegExp_Vms = StringRegExp($sText, '(?s)&lt;Vms&gt;(.*?)&lt;/Vms&gt;', 3)
local $RegExp_VmPoints = StringRegExp($sText, '(?s)&lt;VmPoints&gt;(.*?)&lt;/VmPoints&gt;', 3)
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
global $array_VM_IDs = $RegExp_string
#EndRegion
;;
;Hosts
;Storages
;JobPoints
;Vms
;VmPoints
;;
global $array_Hosts = _Hosts_reparse($RegExp_Hosts)
local $array_Storages = _Storages_reparse($RegExp_Storages)
global $array_Storages_sorted = _Array_Storages_sorting($array_Storages)
global $array_JobPoints = _JobPoints_reparse($RegExp_JobPoints)
global $array_Vms = _VMs_reparse($RegExp_Vms)
global $array_VmPoints = _VMPoints_reparse($RegExp_VmPoints)
EndIf
#Region consistensy check
$targetarray_Storages= UBound($array_Storages, $UBOUND_ROWS)
$targetarray_Storages_sorted = UBound($array_Storages_sorted, $UBOUND_ROWS)
if $targetarray_Storages <> $targetarray_Storages_sorted Then MsgBox($MB_SYSTEMMODAL, "Fatal Error", "Orphaned increment storage!"&@CR&"See debug log for more info")
#EndRegion
#EndRegion
						EndIf
						EndFunc

						#Region vbm_read fuctions
Func _Array_Storages_sorting($array_Storages)
#Region sorting $array_Storages
local $array_Storages_sorted[1][14]
$array_Storages_sorted[0][0] = 'Storage Id'
$array_Storages_sorted[0][1] = 'FileName'
$array_Storages_sorted[0][2] = 'CreationTimeUtc'
$array_Storages_sorted[0][3] = 'BlockSize'
$array_Storages_sorted[0][4] = 'State'
$array_Storages_sorted[0][5] = 'LinkId'
$array_Storages_sorted[0][6] = 'Partial Increment'
$array_Storages_sorted[0][7] = 'GfsPeriod'
$array_Storages_sorted[0][8] = 'CreationMode'
$array_Storages_sorted[0][9] = 'BackupSize'
$array_Storages_sorted[0][10] = 'DataSize'
$array_Storages_sorted[0][11] = 'DedupRatio'
$array_Storages_sorted[0][12] = 'CompressRatio'
$array_Storages_sorted[0][13] = 'All'
local $numm_array_Storages = UBound($array_Storages, $UBOUND_ROWS) - 1
For $i = 1 To $numm_array_Storages step +1
Local $temp = $array_Storages[$i][5]
if $temp = '0000000-0000-0000-0000-000000000000' Then
_ArrayAdd($array_Storages_sorted, $array_Storages[$i][0], 0)
$target_index = UBound($array_Storages_sorted, $UBOUND_ROWS) -1
$array_Storages_sorted[$target_index][1] = $array_Storages[$i][1]
$array_Storages_sorted[$target_index][2] = $array_Storages[$i][2]
$array_Storages_sorted[$target_index][3] = $array_Storages[$i][3]
$array_Storages_sorted[$target_index][4] = $array_Storages[$i][4]
$array_Storages_sorted[$target_index][5] = $array_Storages[$i][5]
$array_Storages_sorted[$target_index][6] = $array_Storages[$i][6]
$array_Storages_sorted[$target_index][7] = $array_Storages[$i][7]
$array_Storages_sorted[$target_index][8] = $array_Storages[$i][8]
$array_Storages_sorted[$target_index][9] = $array_Storages[$i][9]
$array_Storages_sorted[$target_index][10] = $array_Storages[$i][10]
$array_Storages_sorted[$target_index][11] = $array_Storages[$i][11]
$array_Storages_sorted[$target_index][12] = $array_Storages[$i][12]
$array_Storages_sorted[$target_index][13] = $array_Storages[$i][13]
EndIf
Next
$target_source= UBound($array_Storages, $UBOUND_ROWS)-1
$target_index = UBound($array_Storages_sorted, $UBOUND_ROWS)-1
;sorting by date
if $target_index <> 0 Then
	;;need to sort vbk by time
	EndIf




While $target_index < $target_source
$target_index = UBound($array_Storages_sorted, $UBOUND_ROWS)-1
For $i = 1 To $numm_array_Storages step +1
if $array_Storages[$i][5] <> '0000000-0000-0000-0000-000000000000' Then
		$iIndex_exist_in_vbk = _ArraySearch($array_Storages_sorted, $array_Storages[$i][5], 0, 0, 0, 1, 1, 5) ;есть ли ты в $array_Storages_sorted
		if $iIndex_exist_in_vbk = -1 Then
			$iIndex_link = _ArraySearch($array_Storages, $array_Storages[$i][5], 0, 0, 0, 1, 1, 0) ;на кого ты ссылаешьс€
				$iIndex_exist_in_vbk_link =_ArraySearch($array_Storages_sorted, $array_Storages[$iIndex_link][0], 0, 0, 0, 1, 1, 0);есть ли твой перент в $array_Storages_sorted
				if $iIndex_exist_in_vbk_link <> -1 Then ;если перент есть
					;ConsoleWrite($iIndex_exist_in_vbk_link&@cr&$array_Storages[$i][0]&@cr)
					FileDelete('$array_Storages_sorted.txt')
					$target_index = UBound($array_Storages_sorted, $UBOUND_ROWS)-1
					if $target_index = $iIndex_exist_in_vbk_link Then
							_ArrayAdd($array_Storages_sorted, $array_Storages[$i][0],0)
							Else
					_ArrayInsert($array_Storages_sorted, $iIndex_exist_in_vbk_link+1,$array_Storages[$i][0])
					EndIf
					$array_Storages_sorted[$iIndex_exist_in_vbk_link+1][1] = $array_Storages[$i][1]
					$array_Storages_sorted[$iIndex_exist_in_vbk_link+1][2] = $array_Storages[$i][2]
					$array_Storages_sorted[$iIndex_exist_in_vbk_link+1][3] = $array_Storages[$i][3]
					$array_Storages_sorted[$iIndex_exist_in_vbk_link+1][4] = $array_Storages[$i][4]
					$array_Storages_sorted[$iIndex_exist_in_vbk_link+1][5] = $array_Storages[$i][5]
					$array_Storages_sorted[$iIndex_exist_in_vbk_link+1][6] = $array_Storages[$i][6]
					$array_Storages_sorted[$iIndex_exist_in_vbk_link+1][7] = $array_Storages[$i][7]
					$array_Storages_sorted[$iIndex_exist_in_vbk_link+1][8] = $array_Storages[$i][8]
					$array_Storages_sorted[$iIndex_exist_in_vbk_link+1][9] = $array_Storages[$i][9]
					$array_Storages_sorted[$iIndex_exist_in_vbk_link+1][10] = $array_Storages[$i][10]
					$array_Storages_sorted[$iIndex_exist_in_vbk_link+1][11] = $array_Storages[$i][11]
					$array_Storages_sorted[$iIndex_exist_in_vbk_link+1][12] = $array_Storages[$i][12]
					$array_Storages_sorted[$iIndex_exist_in_vbk_link+1][13] = $array_Storages[$i][13]
					;ConsoleWrite($target_index &' '& $target_source&@CR)
					EndIf
		EndIf
		EndIf
Next
WEnd
#EndRegion
Return($array_Storages_sorted)
EndFunc
#Region main arrays reparse Func
Func _Storages_reparse($RegExp_Storages)
local $Storages_array_rep[1][14]
$Storages_array_rep[0][0] = 'Storage Id'
$Storages_array_rep[0][1] = 'FileName'
$Storages_array_rep[0][2] = 'CreationTimeUtc'
$Storages_array_rep[0][3] = 'BlockSize'
$Storages_array_rep[0][4] = 'State'
$Storages_array_rep[0][5] = 'LinkId'
$Storages_array_rep[0][6] = 'Partial Increment'
$Storages_array_rep[0][7] = 'GfsPeriod'
$Storages_array_rep[0][8] = 'CreationMode'
$Storages_array_rep[0][9] = 'BackupSize'
$Storages_array_rep[0][10] = 'DataSize'
$Storages_array_rep[0][11] = 'DedupRatio'
$Storages_array_rep[0][12] = 'CompressRatio'
$Storages_array_rep[0][13] = 'All'

local $RegExp_Storages_str = $RegExp_Storages[0]
local $RegExp_Storage_Id = StringRegExp($RegExp_Storages_str, '(?s)&lt;Storage Id="(.*?)" FileName="', 3)
local $numm_Storage_Id = UBound($RegExp_Storage_Id, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Storage_Id step +1
   local $Storages_array_rep_temp[1][14]
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
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][3] =  $RegExp_Storage_temp[$i]
Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)" State="(.*?)" PartialIncrement=', 3)
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][4] =  $RegExp_Storage_temp[$i]
Next

local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)" State="(.*?)" PartialIncrement=', 3)
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][5] =  $RegExp_Storage_temp[$i]
Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)" PartialIncrement="(.*?)" GfsPeriod=', 3)
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][6] =  $RegExp_Storage_temp[$i]
Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)" GfsPeriod="(.*?)" CreationMode=', 3)
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][7] =  $RegExp_Storage_temp[$i]
Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)" CreationMode="(.*?)"&gt;&lt;Stats&gt;', 3)
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
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][9] =  $RegExp_Storage_temp[$i]
Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)&lt;DataSize&gt;(.*?)&lt;/DataSize&gt;', 3)
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][10] =  $RegExp_Storage_temp[$i]
Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)&lt;DedupRatio&gt;(.*?)&lt;/DedupRatio&gt;', 3)
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][11] =  $RegExp_Storage_temp[$i]
Next
local $RegExp_Storage_temp = StringRegExp($RegExp_Storages_str, '(?s)&lt;CompressRatio&gt;(.*?)&lt;/CompressRatio&gt;', 3)
For $i = 0 To $numm_Storage_temp step +1
   $Storages_array_rep[$i+1][12] =  $RegExp_Storage_temp[$i]
Next
local $RegExp_Storage_FileName = StringRegExp($RegExp_Storages_str, '(?s)&lt;Storage Id="(.*?)&lt;/Storage&gt;', 3)
local $numm_Storage_FileName = UBound($RegExp_Storage_FileName, $UBOUND_ROWS) -1
For $i = 0 To $numm_Storage_FileName step +1
   $Storages_array_rep[$i+1][13] =  $RegExp_Storage_FileName[$i]
Next
if $storages_reparse_test = True  Then
FileDelete('Storages_reparse_test.txt')
_FileWriteFromArray('Storages_array_rep.txt', $Storages_array_rep)
EndIf
ConsoleWrite('$Storages_array_rep is done'&@cr)
GUICtrlSetData($idProgressbar1, 40)


Return $Storages_array_rep
EndFunc
Func _VMPoints_reparse($RegExp_VMPoints)
local $VMPoints_reparse_array_rep[1][10]
$VMPoints_reparse_array_rep[0][0] = 'Point Id'
$VMPoints_reparse_array_rep[0][1] = 'LinkId'
$VMPoints_reparse_array_rep[0][2] = 'GroupId'
$VMPoints_reparse_array_rep[0][3] = 'CreationTime'
$VMPoints_reparse_array_rep[0][4] = 'Type'
$VMPoints_reparse_array_rep[0][5] = 'BackupId'
$VMPoints_reparse_array_rep[0][6] = 'InsideDir'
$VMPoints_reparse_array_rep[0][7] = 'StgId'
$VMPoints_reparse_array_rep[0][8] = 'ParentId'
$VMPoints_reparse_array_rep[0][9] = 'Point all'
$RegExp_VMPoints = $RegExp_VMPoints[0]

local $RegExp_Point_Id = StringRegExp($RegExp_VMPoints, '(?s)&lt;Pnt Id="(.*?)"', 3)
local $numm_Point_Id = UBound($RegExp_Point_Id, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Point_Id step +1
   local $VMPoints_array_rep_temp[1][10]
   $VMPoints_array_rep_temp[0][0] =  $RegExp_Point_Id[$i]
   _ArrayConcatenate($VMPoints_reparse_array_rep,$VMPoints_array_rep_temp)
   Next
local $RegExp_Points_temp = StringRegExp($RegExp_VMPoints, '(?s)" LinkId="(.*?)"', 3)
local $numm_Points_temp = UBound($RegExp_Points_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Points_temp step +1
   $VMPoints_reparse_array_rep[$i+1][1] =  $RegExp_Points_temp[$i]
Next

local $RegExp_Points_temp = StringRegExp($RegExp_VMPoints, '(?s)" GroupId="(.*?)"', 3)
local $numm_Points_temp = UBound($RegExp_Points_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Points_temp step +1
   $VMPoints_reparse_array_rep[$i+1][2] =  $RegExp_Points_temp[$i]
Next
local $RegExp_Points_temp = StringRegExp($RegExp_VMPoints, '(?s)" CreationTime="(.*?)"', 3)
local $numm_Points_temp = UBound($RegExp_Points_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Points_temp step +1
   $VMPoints_reparse_array_rep[$i+1][3] =  $RegExp_Points_temp[$i]
Next
local $RegExp_Points_temp = StringRegExp($RegExp_VMPoints, '(?s)" Type="(.*?)"', 3)
local $numm_Points_temp = UBound($RegExp_Points_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Points_temp step +1
   $VMPoints_reparse_array_rep[$i+1][4] =  $RegExp_Points_temp[$i]
Next
local $RegExp_Points_temp = StringRegExp($RegExp_VMPoints, '(?s)" BackupId="(.*?)" /&gt;', 3)
local $numm_Points_temp = UBound($RegExp_Points_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Points_temp step +1
   $VMPoints_reparse_array_rep[$i+1][5] =  $RegExp_Points_temp[$i]
Next
local $RegExp_Points_temp = StringRegExp($RegExp_VMPoints, '(?s)" InsideDir="(.*?)"', 3)
local $numm_Points_temp = UBound($RegExp_Points_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Points_temp step +1
   $VMPoints_reparse_array_rep[$i+1][6] =  $RegExp_Points_temp[$i]
Next
local $RegExp_Points_temp = StringRegExp($RegExp_VMPoints, '(?s)" StgId="(.*?)"', 3)
local $numm_Points_temp = UBound($RegExp_Points_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Points_temp step +1
   $VMPoints_reparse_array_rep[$i+1][7] =  $RegExp_Points_temp[$i]
Next
local $RegExp_Points_temp = StringRegExp($RegExp_VMPoints, '(?s)" ParentId="(.*?)"', 3)
local $numm_Points_temp = UBound($RegExp_Points_temp, $UBOUND_ROWS) -1
For $i = 0 To $numm_Points_temp step +1
   $VMPoints_reparse_array_rep[$i+1][8] =  $RegExp_Points_temp[$i]
Next
local $RegExp_Points_temp = StringRegExp($RegExp_VMPoints, '(?s)&lt;Pnt Id="(.*?)&lt;/Pnt&gt;', 3)
local $numm_Point_Id = UBound($RegExp_Points_temp, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Points_temp step +1
   $VMPoints_reparse_array_rep[$i+1][9] =  $RegExp_Points_temp[$i]
Next
GUICtrlSetData($idProgressbar1, 80)
if $VMPoints_reparse_test = True  Then
FileDelete('VMPoints_reparse_test.txt')
_FileWriteFromArray('VMPoints_array_rep.txt', $VMPoints_reparse_array_rep)
EndIf
ConsoleWrite('$VMPoints_array_rep is done'&@cr)

Return $VMPoints_reparse_array_rep
EndFunc
Func _JobPoints_reparse($RegExp_JobPoints)
local $RegExp_JobPoints_str =$RegExp_JobPoints[0]
local $JobPoints_reparse_array_rep[1][8]
$JobPoints_reparse_array_rep[0][0] = 'Point Id'
$JobPoints_reparse_array_rep[0][1] = 'LinkId'
$JobPoints_reparse_array_rep[0][2] = 'Num'
$JobPoints_reparse_array_rep[0][3] = 'GroupId'
$JobPoints_reparse_array_rep[0][4] = 'CreationTime'
$JobPoints_reparse_array_rep[0][5] = 'Algoritm'
$JobPoints_reparse_array_rep[0][6] = 'BackupId'
$JobPoints_reparse_array_rep[0][7] = 'JobPoints all'

local $RegExp_Point_Id = StringRegExp($RegExp_JobPoints_str, '&lt;Point Id="(.*?)"', 3)
local $numm_Point_Id = UBound($RegExp_Point_Id, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Point_Id step +1
   local $Points_array_rep_temp[1][8]
   $Points_array_rep_temp[0][0] =  $RegExp_Point_Id[$i]
   _ArrayConcatenate($JobPoints_reparse_array_rep,$Points_array_rep_temp)
   Next
local $RegExp_Point_temp = StringRegExp($RegExp_JobPoints_str, ' LinkId="(.*?)"', 3)
local $numm_Points_temp = UBound($RegExp_Point_temp, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Points_temp step +1
   $JobPoints_reparse_array_rep[$i+1][1] =  $RegExp_Point_temp[$i]
Next
local $RegExp_Point_temp = StringRegExp($RegExp_JobPoints_str, ' Num="(.*?)"', 3)
local $numm_Points_temp = UBound($RegExp_Point_temp, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Points_temp step +1
   $JobPoints_reparse_array_rep[$i+1][2] =  $RegExp_Point_temp[$i]
Next
local $RegExp_Point_temp = StringRegExp($RegExp_JobPoints_str, ' GroupId="(.*?)"', 3)
local $numm_Points_temp = UBound($RegExp_Point_temp, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Points_temp step +1
   $JobPoints_reparse_array_rep[$i+1][3] =  $RegExp_Point_temp[$i]
Next
local $RegExp_Point_temp = StringRegExp($RegExp_JobPoints_str, ' CreationTime"(.*?)"', 3)
local $numm_Points_temp = UBound($RegExp_Point_temp, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Points_temp step +1
   $JobPoints_reparse_array_rep[$i+1][4] =  $RegExp_Point_temp[$i]
Next
local $RegExp_Point_temp = StringRegExp($RegExp_JobPoints_str, ' Algoritm="(.*?)"', 3)
local $numm_Points_temp = UBound($RegExp_Point_temp, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Points_temp step +1
   $JobPoints_reparse_array_rep[$i+1][5] =  $RegExp_Point_temp[$i]
Next
local $RegExp_Point_temp = StringRegExp($RegExp_JobPoints_str, ' BackupId="(.*?)"', 3)
local $numm_Points_temp = UBound($RegExp_Point_temp, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Points_temp step +1
   $JobPoints_reparse_array_rep[$i+1][6] =  $RegExp_Point_temp[$i]
Next
local $RegExp_Point_temp = StringRegExp($RegExp_JobPoints_str, '&lt;Point Id="(.*?) /&gt;', 3)
local $numm_Points_temp = UBound($RegExp_Point_temp, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Points_temp step +1
   $JobPoints_reparse_array_rep[$i+1][7] =  $RegExp_Point_temp[$i]
Next
if $points_reparse_test = True  Then
FileDelete('points_reparse_test.txt')
_FileWriteFromArray('points_array_rep.txt', $JobPoints_reparse_array_rep)
EndIf
ConsoleWrite('$Points_array_rep is done'&@cr)
Return $JobPoints_reparse_array_rep
GUICtrlSetData($idProgressbar1, 60)
EndFunc
Func _Hosts_reparse($RegExp_Hosts)
local $RegExp_Hosts_str =$RegExp_Hosts[0]
local $Hosts_reparse_array_rep[1][5]
$Hosts_reparse_array_rep[0][0] = 'Id'
$Hosts_reparse_array_rep[0][1] = 'MoRef'
$Hosts_reparse_array_rep[0][2] = 'Name'
$Hosts_reparse_array_rep[0][3] = 'Type'
$Hosts_reparse_array_rep[0][4] = 'Host all'

local $RegExp_Host_Id = StringRegExp($RegExp_Hosts_str, 'Id="(.*?)"', 3)
local $numm_Host_Id = UBound($RegExp_Host_Id, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Host_Id step +1
   local $Hosts_array_rep_temp[1][5]
   $Hosts_array_rep_temp[0][0] =  $RegExp_Host_Id[$i]
   _ArrayConcatenate($Hosts_reparse_array_rep,$Hosts_array_rep_temp)
   Next
local $RegExp_Host_temp = StringRegExp($RegExp_Hosts_str, ' MoRef="(.*?)"', 3)
local $numm_Host_Id_temp = UBound($RegExp_Host_temp, $UBOUND_ROWS) - 1
For $i = 0 To $RegExp_Host_temp step +1
   $Hosts_reparse_array_rep[$i+1][1] =  $RegExp_Host_temp[$i]
Next

local $RegExp_Host_temp = StringRegExp($RegExp_Hosts_str, ' Name="(.*?)"', 3)
local $numm_Host_Id_temp = UBound($RegExp_Host_temp, $UBOUND_ROWS) - 1
For $i = 0 To $RegExp_Host_temp step +1
   $Hosts_reparse_array_rep[$i+1][2] =  $RegExp_Host_temp[$i]
Next
local $RegExp_Host_temp = StringRegExp($RegExp_Hosts_str, ' Type="(.*?)"', 3)
local $numm_Host_Id_temp = UBound($RegExp_Host_temp, $UBOUND_ROWS) - 1
For $i = 0 To $RegExp_Host_temp step +1
   $Hosts_reparse_array_rep[$i+1][3] =  $RegExp_Host_temp[$i]
Next
local $RegExp_Host_temp = StringRegExp($RegExp_Hosts_str, '&lt;Host Id="(.*?)" /&gt;', 3)
local $numm_Host_Id_temp = UBound($RegExp_Host_temp, $UBOUND_ROWS) - 1
For $i = 0 To $RegExp_Host_temp step +1
   $Hosts_reparse_array_rep[$i+1][4] =  $RegExp_Host_temp[$i]
Next

if $hosts_reparse_test = True  Then
FileDelete('Hosts_array_rep.txt')
_FileWriteFromArray('Hosts_array_rep.txt', $Hosts_reparse_array_rep)
EndIf
ConsoleWrite('$Hosts_array_rep is done'&@cr)
Return $Hosts_reparse_array_rep
EndFunc
Func _VMs_reparse($RegExp_Vms)
local $RegExp_Vms_str =$RegExp_Vms[0]
local $Vms_reparse_array_rep[1][6]
$Vms_reparse_array_rep[0][0] = 'Id'
$Vms_reparse_array_rep[0][1] = 'ObjectId'
$Vms_reparse_array_rep[0][2] = 'VmRef'
$Vms_reparse_array_rep[0][3] = 'DisplayName'
$Vms_reparse_array_rep[0][4] = 'IsCorrupted'
$Vms_reparse_array_rep[0][5] = 'VM_all'

local $RegExp_Vms_Id = StringRegExp($RegExp_Vms_str, '&lt;Vm Id="(.*?)"', 3)
local $numm_Vms_Id = UBound($RegExp_Vms_Id, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Vms_Id step +1
   local $Vms_ID_array_rep_temp[1][6]
   $Vms_ID_array_rep_temp[0][0] =  $RegExp_Vms_Id[$i]
   _ArrayConcatenate($Vms_reparse_array_rep,$Vms_ID_array_rep_temp)
   Next
local $RegExp_Vms_temp = StringRegExp($RegExp_Vms_str, ' ObjectId="(.*?)"', 3)
local $numm_Vms_temp = UBound($RegExp_Vms_temp, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Vms_temp step +1
   $Vms_reparse_array_rep[$i+1][1] =  $RegExp_Vms_temp[$i]
Next
local $RegExp_Vms_temp = StringRegExp($RegExp_Vms_str, ' VmRef="(.*?)"', 3)
local $numm_Vms_temp = UBound($RegExp_Vms_temp, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Vms_temp step +1
   $Vms_reparse_array_rep[$i+1][2] =  $RegExp_Vms_temp[$i]
Next
local $RegExp_Vms_temp = StringRegExp($RegExp_Vms_str, ' DisplayName="(.*?)"', 3)
local $numm_Vms_temp = UBound($RegExp_Vms_temp, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Vms_temp step +1
   $Vms_reparse_array_rep[$i+1][3] =  $RegExp_Vms_temp[$i]
Next
local $RegExp_Vms_temp = StringRegExp($RegExp_Vms_str, ' IsCorrupted="(.*?)"', 3)
local $numm_Vms_temp = UBound($RegExp_Vms_temp, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Vms_temp step +1
   $Vms_reparse_array_rep[$i+1][4] =  $RegExp_Vms_temp[$i]
Next
local $RegExp_Vms_temp = StringRegExp($RegExp_Vms_str, '&lt;Vm Id="(.*?)&lt;/Vm&gt;', 3)
local $numm_Vms_temp = UBound($RegExp_Vms_temp, $UBOUND_ROWS) - 1
For $i = 0 To $numm_Vms_temp step +1
   $Vms_reparse_array_rep[$i+1][5] = '&lt;Vm Id="'& $RegExp_Vms_temp[$i]&'&lt;/Vm&gt;'
Next
if $VMs_reparse_test = True  Then
FileDelete('VMs_array_rep.txt')
_FileWriteFromArray('VMs_array_rep.txt', $Vms_reparse_array_rep)
EndIf
ConsoleWrite('$VMs_array_rep is done'&@cr)
GUICtrlSetData($idProgressbar1, 80)
Return $Vms_reparse_array_rep
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