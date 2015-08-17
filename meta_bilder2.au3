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
;AutoItSetOption("MustDeclareVars", 1)
global $version_exe = '0.2.0.0'
If Not IsAdmin() Then
    MsgBox($MB_SYSTEMMODAL, "Warning","Admin rights are not detected."& @CR& 'You need to be admin to run this tool')
	Exit
EndIf
#Region main paths
global $ProgrammDir = 'C:\VeeamMetaBilder\'
global $temp_dir = @TempDir&'\VeeamMetaBilder\'
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
global $storages_sorted_test = False
#EndRegion
#Region GUI Test
global $GUI_Test = False
#EndRegion
_Main()
Func _Main()
    Local $hGUI, $aSize
    If @AutoItX64 Then local $sWow64 = "\Wow6432Node"
    Local $sPath = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE" & $sWow64 & "\AutoIt v3\AutoIt", "InstallDir") & "\Examples\GUI\Advanced\Images"
    ; Создаёт GUI
    $hGUI = GUICreate(StringTrimRight("Veeam Meta Parser "&$version_exe, 2), 1000, 750)
	;GUICtrlSetResizing ( $hGUI, resizing )
GUISetBkColor(0xF0F4F9)
global $tools = GUICtrlCreateGroup("", -1, -10, 1000, 55)
global $Group2 = GUICtrlCreateGroup("", 16, 48, 465, 670)
global $Group3 = GUICtrlCreateGroup("", 485, 48, 490, 670)
  global $open = GUICtrlCreateIcon("shell32.dll", 5, 5, 5, 32, 32)
   global  $save= GUICtrlCreateIcon("shell32.dll", 7, 37, 5, 32, 32)
	GUICtrlCreateLabel('|', 74, 7, 5, 30, $SS_ETCHEDVERT)
	global  $help = GUICtrlCreateIcon("shell32.dll", 24, 79, 5, 32, 32)
$tools = GUICtrlCreateGroup("", 0, 715, 1000, 40)
Global $idProgressbar1 = GUICtrlCreateProgress(745, 725, 250, 20)
GUICtrlSetState($idProgressbar1, $GUI_HIDE )
    GUICtrlSetColor(-1, 32250); not working with Windows XP Style
#Region грязный, отвратительный костыль, которому не место в нашем мире.
Global $Button_storage1, $Button_storage2, $Button_storage3, $Button_storage4, $Button_storage5, $Button_storage6, $Button_storage7, $Button_storage8, $Button_storage9, $Button_storage10, $Button_storage11, $Button_storage12, $Button_storage13, $Button_storage14, $Button_storage15, $Button_storage16, $Button_storage17, $Button_storage18, $Button_storage19, $Button_storage20, $Button_storage21, $Button_storage22, $Button_storage23, $Button_storage24, $Button_storage25, $Button_storage26, $Button_storage27, $Button_storage28, $Button_storage29, $Button_storage30, $Button_storage31, $Button_storage32, $Button_storage33, $Button_storage34, $Button_storage35, $Button_storage36, $Button_storage37, $Button_storage38, $Button_storage39, $Button_storage40, $Button_storage41, $Button_storage42, $Button_storage43, $Button_storage44, $Button_storage45, $Button_storage46, $Button_storage47, $Button_storage48, $Button_storage49, $Button_storage50, $Button_storage51, $Button_storage52, $Button_storage53, $Button_storage54, $Button_storage55, $Button_storage56, $Button_storage57, $Button_storage58, $Button_storage59, $Button_storage60, $Button_storage61, $Button_storage62, $Button_storage63, $Button_storage64, $Button_storage65, $Button_storage66, $Button_storage67, $Button_storage68, $Button_storage69, $Button_storage70, $Button_storage71, $Button_storage72, $Button_storage73, $Button_storage74, $Button_storage75, $Button_storage76, $Button_storage77, $Button_storage78, $Button_storage79, $Button_storage80, $Button_storage81, $Button_storage82, $Button_storage83, $Button_storage84, $Button_storage85, $Button_storage86, $Button_storage87, $Button_storage88, $Button_storage89, $Button_storage90, $Button_storage91, $Button_storage92, $Button_storage93, $Button_storage94, $Button_storage95, $Button_storage96, $Button_storage97, $Button_storage98, $Button_storage99, $Button_storage100, $Button_storage101, $Button_storage102, $Button_storage103, $Button_storage104, $Button_storage105, $Button_storage106, $Button_storage107, $Button_storage108, $Button_storage109, $Button_storage110, $Button_storage111, $Button_storage112, $Button_storage113, $Button_storage114, $Button_storage115, $Button_storage116, $Button_storage117, $Button_storage118, $Button_storage119, $Button_storage120, $Button_storage121, $Button_storage122, $Button_storage123, $Button_storage124, $Button_storage125, $Button_storage126, $Button_storage127, $Button_storage128, $Button_storage129, $Button_storage130, $Button_storage131, $Button_storage132, $Button_storage133, $Button_storage134, $Button_storage135, $Button_storage136, $Button_storage137, $Button_storage138, $Button_storage139, $Button_storage140, $Button_storage141, $Button_storage142, $Button_storage143, $Button_storage144, $Button_storage145, $Button_storage146, $Button_storage147, $Button_storage148, $Button_storage149, $Button_storage150, $Button_storage151, $Button_storage152, $Button_storage153, $Button_storage154, $Button_storage155, $Button_storage156, $Button_storage157, $Button_storage158, $Button_storage159, $Button_storage160, $Button_storage161, $Button_storage162, $Button_storage163, $Button_storage164, $Button_storage165, $Button_storage166, $Button_storage167, $Button_storage168, $Button_storage169, $Button_storage170, $Button_storage171, $Button_storage172, $Button_storage173, $Button_storage174, $Button_storage175, $Button_storage176, $Button_storage177, $Button_storage178, $Button_storage179, $Button_storage180, $Button_storage181, $Button_storage182, $Button_storage183, $Button_storage184, $Button_storage185, $Button_storage186, $Button_storage187, $Button_storage188, $Button_storage189, $Button_storage190, $Button_storage191, $Button_storage192, $Button_storage193, $Button_storage194, $Button_storage195, $Button_storage196, $Button_storage197, $Button_storage198, $Button_storage199, $Button_storage200 ;грязный, отвратительный костыль, которому не место в нашем мире.
Global $Backup_size_storage1, $Backup_size_storage2, $Backup_size_storage3, $Backup_size_storage4, $Backup_size_storage5, $Backup_size_storage6, $Backup_size_storage7, $Backup_size_storage8, $Backup_size_storage9, $Backup_size_storage10, $Backup_size_storage11, $Backup_size_storage12, $Backup_size_storage13, $Backup_size_storage14, $Backup_size_storage15, $Backup_size_storage16, $Backup_size_storage17, $Backup_size_storage18, $Backup_size_storage19, $Backup_size_storage20, $Backup_size_storage21, $Backup_size_storage22, $Backup_size_storage23, $Backup_size_storage24, $Backup_size_storage25, $Backup_size_storage26, $Backup_size_storage27, $Backup_size_storage28, $Backup_size_storage29, $Backup_size_storage30, $Backup_size_storage31, $Backup_size_storage32, $Backup_size_storage33, $Backup_size_storage34, $Backup_size_storage35, $Backup_size_storage36, $Backup_size_storage37, $Backup_size_storage38, $Backup_size_storage39, $Backup_size_storage40, $Backup_size_storage41, $Backup_size_storage42, $Backup_size_storage43, $Backup_size_storage44, $Backup_size_storage45, $Backup_size_storage46, $Backup_size_storage47, $Backup_size_storage48, $Backup_size_storage49, $Backup_size_storage50, $Backup_size_storage51, $Backup_size_storage52, $Backup_size_storage53, $Backup_size_storage54, $Backup_size_storage55, $Backup_size_storage56, $Backup_size_storage57, $Backup_size_storage58, $Backup_size_storage59, $Backup_size_storage60, $Backup_size_storage61, $Backup_size_storage62, $Backup_size_storage63, $Backup_size_storage64, $Backup_size_storage65, $Backup_size_storage66, $Backup_size_storage67, $Backup_size_storage68, $Backup_size_storage69, $Backup_size_storage70, $Backup_size_storage71, $Backup_size_storage72, $Backup_size_storage73, $Backup_size_storage74, $Backup_size_storage75, $Backup_size_storage76, $Backup_size_storage77, $Backup_size_storage78, $Backup_size_storage79, $Backup_size_storage80, $Backup_size_storage81, $Backup_size_storage82, $Backup_size_storage83, $Backup_size_storage84, $Backup_size_storage85, $Backup_size_storage86, $Backup_size_storage87, $Backup_size_storage88, $Backup_size_storage89, $Backup_size_storage90, $Backup_size_storage91, $Backup_size_storage92, $Backup_size_storage93, $Backup_size_storage94, $Backup_size_storage95, $Backup_size_storage96, $Backup_size_storage97, $Backup_size_storage98, $Backup_size_storage99, $Backup_size_storage100, $Backup_size_storage101, $Backup_size_storage102, $Backup_size_storage103, $Backup_size_storage104, $Backup_size_storage105, $Backup_size_storage106, $Backup_size_storage107, $Backup_size_storage108, $Backup_size_storage109, $Backup_size_storage110, $Backup_size_storage111, $Backup_size_storage112, $Backup_size_storage113, $Backup_size_storage114, $Backup_size_storage115, $Backup_size_storage116, $Backup_size_storage117, $Backup_size_storage118, $Backup_size_storage119, $Backup_size_storage120, $Backup_size_storage121, $Backup_size_storage122, $Backup_size_storage123, $Backup_size_storage124, $Backup_size_storage125, $Backup_size_storage126, $Backup_size_storage127, $Backup_size_storage128, $Backup_size_storage129, $Backup_size_storage130, $Backup_size_storage131, $Backup_size_storage132, $Backup_size_storage133, $Backup_size_storage134, $Backup_size_storage135, $Backup_size_storage136, $Backup_size_storage137, $Backup_size_storage138, $Backup_size_storage139, $Backup_size_storage140, $Backup_size_storage141, $Backup_size_storage142, $Backup_size_storage143, $Backup_size_storage144, $Backup_size_storage145, $Backup_size_storage146, $Backup_size_storage147, $Backup_size_storage148, $Backup_size_storage149, $Backup_size_storage150, $Backup_size_storage151, $Backup_size_storage152, $Backup_size_storage153, $Backup_size_storage154, $Backup_size_storage155, $Backup_size_storage156, $Backup_size_storage157, $Backup_size_storage158, $Backup_size_storage159, $Backup_size_storage160, $Backup_size_storage161, $Backup_size_storage162, $Backup_size_storage163, $Backup_size_storage164, $Backup_size_storage165, $Backup_size_storage166, $Backup_size_storage167, $Backup_size_storage168, $Backup_size_storage169, $Backup_size_storage170, $Backup_size_storage171, $Backup_size_storage172, $Backup_size_storage173, $Backup_size_storage174, $Backup_size_storage175, $Backup_size_storage176, $Backup_size_storage177, $Backup_size_storage178, $Backup_size_storage179, $Backup_size_storage180, $Backup_size_storage181, $Backup_size_storage182, $Backup_size_storage183, $Backup_size_storage184, $Backup_size_storage185, $Backup_size_storage186, $Backup_size_storage187, $Backup_size_storage188, $Backup_size_storage189, $Backup_size_storage190, $Backup_size_storage191, $Backup_size_storage192, $Backup_size_storage193, $Backup_size_storage194, $Backup_size_storage195, $Backup_size_storage196, $Backup_size_storage197, $Backup_size_storage198, $Backup_size_storage199, $Backup_size_storage200
Global $Button_vms_in_stg0, $Button_vms_in_stg1, $Button_vms_in_stg2, $Button_vms_in_stg3, $Button_vms_in_stg4, $Button_vms_in_stg5, $Button_vms_in_stg6, $Button_vms_in_stg7, $Button_vms_in_stg8, $Button_vms_in_stg9, $Button_vms_in_stg10, $Button_vms_in_stg11, $Button_vms_in_stg12, $Button_vms_in_stg13, $Button_vms_in_stg14, $Button_vms_in_stg15, $Button_vms_in_stg16, $Button_vms_in_stg17, $Button_vms_in_stg18, $Button_vms_in_stg19, $Button_vms_in_stg20, $Button_vms_in_stg21, $Button_vms_in_stg22, $Button_vms_in_stg23, $Button_vms_in_stg24, $Button_vms_in_stg25, $Button_vms_in_stg26, $Button_vms_in_stg27, $Button_vms_in_stg28, $Button_vms_in_stg29, $Button_vms_in_stg30, $Button_vms_in_stg31, $Button_vms_in_stg32, $Button_vms_in_stg33, $Button_vms_in_stg34, $Button_vms_in_stg35, $Button_vms_in_stg36, $Button_vms_in_stg37, $Button_vms_in_stg38, $Button_vms_in_stg39, $Button_vms_in_stg40, $Button_vms_in_stg41, $Button_vms_in_stg42, $Button_vms_in_stg43, $Button_vms_in_stg44, $Button_vms_in_stg45, $Button_vms_in_stg46, $Button_vms_in_stg47, $Button_vms_in_stg48, $Button_vms_in_stg49, $Button_vms_in_stg50, $Button_vms_in_stg51, $Button_vms_in_stg52, $Button_vms_in_stg53, $Button_vms_in_stg54, $Button_vms_in_stg55, $Button_vms_in_stg56, $Button_vms_in_stg57, $Button_vms_in_stg58, $Button_vms_in_stg59, $Button_vms_in_stg60, $Button_vms_in_stg61, $Button_vms_in_stg62, $Button_vms_in_stg63, $Button_vms_in_stg64, $Button_vms_in_stg65, $Button_vms_in_stg66, $Button_vms_in_stg67, $Button_vms_in_stg68, $Button_vms_in_stg69, $Button_vms_in_stg70, $Button_vms_in_stg71, $Button_vms_in_stg72, $Button_vms_in_stg73, $Button_vms_in_stg74, $Button_vms_in_stg75, $Button_vms_in_stg76, $Button_vms_in_stg77, $Button_vms_in_stg78, $Button_vms_in_stg79, $Button_vms_in_stg80, $Button_vms_in_stg81, $Button_vms_in_stg82, $Button_vms_in_stg83, $Button_vms_in_stg84, $Button_vms_in_stg85, $Button_vms_in_stg86, $Button_vms_in_stg87, $Button_vms_in_stg88, $Button_vms_in_stg89, $Button_vms_in_stg90, $Button_vms_in_stg91, $Button_vms_in_stg92, $Button_vms_in_stg93, $Button_vms_in_stg94, $Button_vms_in_stg95, $Button_vms_in_stg96, $Button_vms_in_stg97, $Button_vms_in_stg98, $Button_vms_in_stg99

#EndRegion
For $i = 1 To 200 step +1 ;delete all Checkboxs and Buttons

	Local $sEvalButton = GUICtrlCreateLabel('', -1, -1, 0, 0,  BitOR($SS_Left, $SS_CENTERIMAGE))
		Assign('Button_storage'&$i, $sEvalButton)
												Local $sEvalString = Eval('Button_storage'&$i)
												GUICtrlSetState($sEvalString, $GUI_HIDE )
												GUICtrlDelete($sEvalString)
	Local $sEvalButton = GUICtrlCreateLabel('', -1, -1, 0, 0,  BitOR($SS_Left, $SS_CENTERIMAGE))
		Assign('Checkbox_storage'&$i, $sEvalButton)
												Local $sEvalString = Eval('Checkbox_storage'&$i)
												GUICtrlSetState($sEvalString, $GUI_HIDE )
												GUICtrlDelete($sEvalString)
												Local $sEvalButton = GUICtrlCreateLabel('', -1, -1, 0, 0,  BitOR($SS_Left, $SS_CENTERIMAGE))
		Assign('Backup_size_storage'&$i, $sEvalButton)
												Local $sEvalString = Eval('Backup_size_storage'&$i)
												GUICtrlSetState($sEvalString, $GUI_HIDE )
												GUICtrlDelete($sEvalString)
Next
For $i = 0 To 99 step +1 ;delete all Checkboxs and Buttons


	Local $sEvalButton = GUICtrlCreateLabel('temp', -1, -1, 0, 0,  BitOR($SS_Left, $SS_CENTERIMAGE))
		Assign('Button_vms_in_stg'&$i, $sEvalButton)
												Local $sEvalString = Eval('Button_vms_in_stg'&$i)
												GUICtrlDelete($sEvalString)
												Local $sEvalvms_in_stg_IsCorrupted = Eval('Button_vms_in_stg_IsCorrupted'&$i)
												GUICtrlDelete($sEvalvms_in_stg_IsCorrupted)
Next


												GUISetState(@SW_SHOW)
#Region dummy buttons


     While 1
       local $GUIMsg = GUIGetMsg()

        Switch $GUIMsg
            Case $GUI_EVENT_CLOSE
            			ExitLoop

			Case $open


				GUICtrlSetState($idProgressbar1, $GUI_show )
				local $parce_func_array = 'no'
				$parce_func_array =_Parse()
				;_ArrayDisplay($parce_func_array,'$parce_func_array')
				if $parce_func_array <> 'no' Then
					For $i = 1 To 200 step +1 ;delete all Checkboxs and Buttons

		Local $sEvalString = Eval('Checkbox_storage'&$i)
		GUICtrlDelete($sEvalString)
		Local $sEvalString = Eval('Button_storage'&$i)
		GUICtrlDelete($sEvalString)
		Local $sEvalString = Eval('Backup_size_storage'&$i)
		GUICtrlDelete($sEvalString)

					Next
	For $i = 0 To 99 step +1 ;delete all Checkboxs and Buttons
		Local $sEvalString = Eval('Button_vms_in_stg'&$i)
		GUICtrlDelete($sEvalString)
		Local $sEvalvms_in_stg_IsCorrupted = Eval('Button_vms_in_stg_IsCorrupted'&$i)
		GUICtrlDelete($sEvalvms_in_stg_IsCorrupted)
		Next
				Global $array_Hosts = $parce_func_array[0][1]
				Global $array_Storages_sorted = $parce_func_array[2][1]
				Global $array_JobPoints  = $parce_func_array[3][1]
				Global $array_Vms = $parce_func_array[4][1]
				Global $array_VmPoints = $parce_func_array[5][1]
				Global $array_VM_IDs = $parce_func_array[16][1]
				GUICtrlSetData($idProgressbar1, 100)
						;_ArrayDisplay($array_Hosts,'$array_Hosts')
						;_ArrayDisplay($array_Storages_sorted,'$array_Storages_sorted')
						;_ArrayDisplay($array_JobPoints,'$array_JobPoints')
						;_ArrayDisplay($array_Vms,'$array_Vms')
						;_ArrayDisplay($array_VmPoints,'$array_VmPoints')
						;_ArrayDisplay($array_VM_IDs,'$array_VM_IDs')
												Sleep(1000)
				GUICtrlSetState($idProgressbar1, $GUI_HIDE )
ConsoleWrite('############################################################################################meta_bilder_starts_working############################################################################################'&@CR)
#Region list creating
$Group2 = GUICtrlCreateGroup("", 16, 48, 465, 670)
local $numm_array_Storages_sorted = UBound($array_Storages_sorted, $UBOUND_ROWS) - 1
	if $GUI_Test = True Then ConsoleWrite('############################################################################################GUI CREATE############################################################################################'&@CR)
For $i = 1 To $numm_array_Storages_sorted step +1
	local $EXTENTION = StringRight($array_Storages_sorted[$i][1],3)
	local $x=$i-1
		IF $EXTENTION = 'vbk' Then
			Local $sEvalCheckbox = GUICtrlCreateCheckbox("", 33, 70+18*$x, 15, 15)
		Assign('Checkbox_storage'&$i, $sEvalCheckbox)
												Local $sEvalString1 = Eval('Checkbox_storage'&$i)
												GUICtrlSetState($sEvalString1, $GUI_show )

			Local $sEvalButton = GUICtrlCreateLabel($array_Storages_sorted[$i][1], 50, 65 + 18*$x, 300, 25,  BitOR($SS_Left, $SS_CENTERIMAGE))
		Assign('Button_storage'&$i, $sEvalButton)
												Local $sEvalString2 = Eval('Button_storage'&$i)
												GUICtrlSetState($sEvalString2, $GUI_show )
			Local $bytes = _ByteSuffix($array_Storages_sorted[$i][9])
			Local $sEvalBackup_size = GUICtrlCreateLabel($bytes, 425, 65+18*$x, 50, 25,  BitOR($SS_right, $SS_CENTERIMAGE))
			Assign('Backup_size_storage'&$i, $sEvalBackup_size)
												Local $sEvalString3 = Eval('Backup_size_storage'&$i)
												GUICtrlSetState($sEvalString3, $GUI_show )

												$sEvalvms_stg_id = $array_Storages_sorted[$i][1]
												Assign('Button_stg_id'&$i, $sEvalvms_stg_id,2);holds stg id from button
if $GUI_Test = True Then ConsoleWrite('Group '&$i&' '&$sEvalString1&' '&$sEvalString2&$sEvalString3&@CR)

		Else

			Local $sEvalCheckbox = GUICtrlCreateCheckbox("", 72, 70 + 18*$x, 15, 15)
		Assign('Checkbox_storage'&$i, $sEvalCheckbox)
												Local $sEvalString1 = Eval('Checkbox_storage'&$i)
												GUICtrlSetState($sEvalString1, $GUI_show )

			Local $sEvalButton = GUICtrlCreateLabel($array_Storages_sorted[$i][1], 90, 65 + 18*$x, 300, 25,  BitOR($SS_Left, $SS_CENTERIMAGE))
			Assign('Button_storage'&$i, $sEvalButton)
												Local $sEvalString2 = Eval('Button_storage'&$i)
												GUICtrlSetState($sEvalString2, $GUI_show )
			Local $bytes = _ByteSuffix($array_Storages_sorted[$i][9])
			Local $sEvalBackup_size = GUICtrlCreateLabel($bytes, 425, 65+18*$x, 50, 25,  BitOR($SS_right, $SS_CENTERIMAGE))
			Assign('Backup_size_storage'&$i, $sEvalBackup_size)
												Local $sEvalString3 = Eval('Backup_size_storage'&$i)
												GUICtrlSetState($sEvalString3, $GUI_show )
												$sEvalvms_stg_id = $array_Storages_sorted[$i][1]
												Assign('Button_stg_id'&$i, $sEvalvms_stg_id,2);holds stg id from button
if $GUI_Test = True Then ConsoleWrite('Group '&$i&' '&$sEvalString1&' '&$sEvalString2&' '&$sEvalString3&@CR)

						EndIf
Next
GUICtrlCreateGroup("", -99, -99, 1, 1)
	if $GUI_Test = True Then ConsoleWrite('############################################################################################GUI CREATE END############################################################################################'&@CR)



EndIf

#Region CASE $Button_storagex
#Region
Case $Button_storage1
$button_numm=1
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage2
$button_numm=2
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage3
$button_numm=3
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage4
$button_numm=4
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage5
$button_numm=5
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage6
$button_numm=6
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage7
$button_numm=7
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage8
$button_numm=8
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage9
$button_numm=9
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage10
$button_numm=10
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage11
$button_numm=11
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage12
$button_numm=12
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage13
$button_numm=13
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage14
$button_numm=14
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage15
$button_numm=15
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage16
$button_numm=16
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage17
$button_numm=17
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage18
$button_numm=18
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage19
$button_numm=19
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage20
$button_numm=20
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage21
$button_numm=21
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage22
$button_numm=22
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage23
$button_numm=23
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage24
$button_numm=24
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage25
$button_numm=25
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage26
$button_numm=26
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage27
$button_numm=27
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage28
$button_numm=28
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage29
$button_numm=29
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage30
$button_numm=30
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage31
$button_numm=31
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage32
$button_numm=32
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage33
$button_numm=33
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage34
$button_numm=34
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage35
$button_numm=35
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage36
$button_numm=36
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage37
$button_numm=37
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage38
$button_numm=38
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage39
$button_numm=39
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage40
$button_numm=40
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage41
$button_numm=41
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage42
$button_numm=42
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage43
$button_numm=43
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage44
$button_numm=44
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage45
$button_numm=45
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage46
$button_numm=46
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage47
$button_numm=47
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage48
$button_numm=48
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage49
$button_numm=49
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage50
$button_numm=50
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage51
$button_numm=51
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage52
$button_numm=52
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage53
$button_numm=53
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage54
$button_numm=54
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage55
$button_numm=55
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage56
$button_numm=56
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage57
$button_numm=57
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage58
$button_numm=58
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage59
$button_numm=59
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage60
$button_numm=60
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage61
$button_numm=61
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage62
$button_numm=62
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage63
$button_numm=63
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage64
$button_numm=64
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage65
$button_numm=65
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage66
$button_numm=66
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage67
$button_numm=67
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage68
$button_numm=68
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage69
$button_numm=69
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage70
$button_numm=70
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage71
$button_numm=71
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage72
$button_numm=72
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage73
$button_numm=73
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage74
$button_numm=74
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage75
$button_numm=75
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage76
$button_numm=76
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage77
$button_numm=77
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage78
$button_numm=78
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage79
$button_numm=79
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage80
$button_numm=80
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage81
$button_numm=81
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage82
$button_numm=82
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage83
$button_numm=83
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage84
$button_numm=84
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage85
$button_numm=85
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage86
$button_numm=86
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage87
$button_numm=87
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage88
$button_numm=88
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage89
$button_numm=89
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage90
$button_numm=90
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage91
$button_numm=91
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage92
$button_numm=92
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage93
$button_numm=93
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage94
$button_numm=94
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage95
$button_numm=95
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage96
$button_numm=96
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage97
$button_numm=97
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage98
$button_numm=98
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage99
$button_numm=99
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage100
$button_numm=100
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage101
$button_numm=101
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage102
$button_numm=102
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage103
$button_numm=103
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage104
$button_numm=104
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage105
$button_numm=105
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage106
$button_numm=106
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage107
$button_numm=107
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage108
$button_numm=108
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage109
$button_numm=109
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage110
$button_numm=110
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage111
$button_numm=111
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage112
$button_numm=112
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage113
$button_numm=113
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage114
$button_numm=114
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage115
$button_numm=115
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage116
$button_numm=116
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage117
$button_numm=117
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage118
$button_numm=118
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage119
$button_numm=119
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage120
$button_numm=120
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage121
$button_numm=121
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage122
$button_numm=122
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage123
$button_numm=123
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage124
$button_numm=124
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage125
$button_numm=125
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage126
$button_numm=126
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage127
$button_numm=127
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage128
$button_numm=128
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage129
$button_numm=129
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage130
$button_numm=130
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage131
$button_numm=131
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage132
$button_numm=132
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage133
$button_numm=133
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage134
$button_numm=134
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage135
$button_numm=135
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage136
$button_numm=136
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage137
$button_numm=137
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage138
$button_numm=138
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage139
$button_numm=139
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage140
$button_numm=140
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage141
$button_numm=141
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage142
$button_numm=142
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage143
$button_numm=143
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage144
$button_numm=144
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage145
$button_numm=145
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage146
$button_numm=146
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage147
$button_numm=147
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage148
$button_numm=148
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage149
$button_numm=149
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage150
$button_numm=150
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage151
$button_numm=151
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage152
$button_numm=152
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage153
$button_numm=153
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage154
$button_numm=154
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage155
$button_numm=155
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage156
$button_numm=156
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage157
$button_numm=157
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage158
$button_numm=158
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage159
$button_numm=159
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage160
$button_numm=160
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage161
$button_numm=161
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage162
$button_numm=162
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage163
$button_numm=163
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage164
$button_numm=164
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage165
$button_numm=165
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage166
$button_numm=166
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage167
$button_numm=167
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage168
$button_numm=168
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage169
$button_numm=169
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage170
$button_numm=170
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage171
$button_numm=171
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage172
$button_numm=172
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage173
$button_numm=173
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage174
$button_numm=174
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage175
$button_numm=175
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage176
$button_numm=176
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage177
$button_numm=177
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage178
$button_numm=178
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage179
$button_numm=179
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage180
$button_numm=180
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage181
$button_numm=181
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage182
$button_numm=182
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage183
$button_numm=183
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage184
$button_numm=184
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage185
$button_numm=185
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage186
$button_numm=186
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage187
$button_numm=187
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage188
$button_numm=188
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage189
$button_numm=189
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage190
$button_numm=190
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage191
$button_numm=191
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage192
$button_numm=192
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage193
$button_numm=193
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage194
$button_numm=194
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage195
$button_numm=195
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage196
$button_numm=196
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage197
$button_numm=197
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage198
$button_numm=198
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage199
$button_numm=199
_button_stg($button_numm, $array_VmPoints)
Case $Button_storage200
$button_numm=200
_button_stg($button_numm, $array_VmPoints)
#EndRegion
#EndRegion
#Region CASE Backup_size_storage
#Region
Case $Backup_size_storage1
$button_size_storage_numm=1
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage2
$button_size_storage_numm=2
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage3
$button_size_storage_numm=3
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage4
$button_size_storage_numm=4
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage5
$button_size_storage_numm=5
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage6
$button_size_storage_numm=6
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage7
$button_size_storage_numm=7
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage8
$button_size_storage_numm=8
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage9
$button_size_storage_numm=9
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage10
$button_size_storage_numm=10
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage11
$button_size_storage_numm=11
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage12
$button_size_storage_numm=12
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage13
$button_size_storage_numm=13
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage14
$button_size_storage_numm=14
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage15
$button_size_storage_numm=15
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage16
$button_size_storage_numm=16
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage17
$button_size_storage_numm=17
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage18
$button_size_storage_numm=18
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage19
$button_size_storage_numm=19
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage20
$button_size_storage_numm=20
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage21
$button_size_storage_numm=21
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage22
$button_size_storage_numm=22
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage23
$button_size_storage_numm=23
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage24
$button_size_storage_numm=24
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage25
$button_size_storage_numm=25
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage26
$button_size_storage_numm=26
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage27
$button_size_storage_numm=27
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage28
$button_size_storage_numm=28
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage29
$button_size_storage_numm=29
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage30
$button_size_storage_numm=30
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage31
$button_size_storage_numm=31
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage32
$button_size_storage_numm=32
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage33
$button_size_storage_numm=33
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage34
$button_size_storage_numm=34
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage35
$button_size_storage_numm=35
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage36
$button_size_storage_numm=36
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage37
$button_size_storage_numm=37
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage38
$button_size_storage_numm=38
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage39
$button_size_storage_numm=39
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage40
$button_size_storage_numm=40
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage41
$button_size_storage_numm=41
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage42
$button_size_storage_numm=42
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage43
$button_size_storage_numm=43
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage44
$button_size_storage_numm=44
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage45
$button_size_storage_numm=45
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage46
$button_size_storage_numm=46
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage47
$button_size_storage_numm=47
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage48
$button_size_storage_numm=48
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage49
$button_size_storage_numm=49
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage50
$button_size_storage_numm=50
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage51
$button_size_storage_numm=51
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage52
$button_size_storage_numm=52
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage53
$button_size_storage_numm=53
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage54
$button_size_storage_numm=54
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage55
$button_size_storage_numm=55
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage56
$button_size_storage_numm=56
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage57
$button_size_storage_numm=57
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage58
$button_size_storage_numm=58
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage59
$button_size_storage_numm=59
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage60
$button_size_storage_numm=60
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage61
$button_size_storage_numm=61
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage62
$button_size_storage_numm=62
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage63
$button_size_storage_numm=63
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage64
$button_size_storage_numm=64
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage65
$button_size_storage_numm=65
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage66
$button_size_storage_numm=66
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage67
$button_size_storage_numm=67
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage68
$button_size_storage_numm=68
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage69
$button_size_storage_numm=69
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage70
$button_size_storage_numm=70
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage71
$button_size_storage_numm=71
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage72
$button_size_storage_numm=72
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage73
$button_size_storage_numm=73
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage74
$button_size_storage_numm=74
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage75
$button_size_storage_numm=75
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage76
$button_size_storage_numm=76
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage77
$button_size_storage_numm=77
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage78
$button_size_storage_numm=78
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage79
$button_size_storage_numm=79
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage80
$button_size_storage_numm=80
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage81
$button_size_storage_numm=81
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage82
$button_size_storage_numm=82
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage83
$button_size_storage_numm=83
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage84
$button_size_storage_numm=84
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage85
$button_size_storage_numm=85
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage86
$button_size_storage_numm=86
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage87
$button_size_storage_numm=87
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage88
$button_size_storage_numm=88
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage89
$button_size_storage_numm=89
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage90
$button_size_storage_numm=90
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage91
$button_size_storage_numm=91
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage92
$button_size_storage_numm=92
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage93
$button_size_storage_numm=93
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage94
$button_size_storage_numm=94
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage95
$button_size_storage_numm=95
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage96
$button_size_storage_numm=96
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage97
$button_size_storage_numm=97
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage98
$button_size_storage_numm=98
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage99
$button_size_storage_numm=99
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage100
$button_size_storage_numm=100
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage101
$button_size_storage_numm=101
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage102
$button_size_storage_numm=102
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage103
$button_size_storage_numm=103
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage104
$button_size_storage_numm=104
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage105
$button_size_storage_numm=105
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage106
$button_size_storage_numm=106
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage107
$button_size_storage_numm=107
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage108
$button_size_storage_numm=108
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage109
$button_size_storage_numm=109
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage110
$button_size_storage_numm=110
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage111
$button_size_storage_numm=111
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage112
$button_size_storage_numm=112
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage113
$button_size_storage_numm=113
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage114
$button_size_storage_numm=114
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage115
$button_size_storage_numm=115
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage116
$button_size_storage_numm=116
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage117
$button_size_storage_numm=117
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage118
$button_size_storage_numm=118
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage119
$button_size_storage_numm=119
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage120
$button_size_storage_numm=120
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage121
$button_size_storage_numm=121
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage122
$button_size_storage_numm=122
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage123
$button_size_storage_numm=123
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage124
$button_size_storage_numm=124
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage125
$button_size_storage_numm=125
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage126
$button_size_storage_numm=126
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage127
$button_size_storage_numm=127
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage128
$button_size_storage_numm=128
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage129
$button_size_storage_numm=129
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage130
$button_size_storage_numm=130
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage131
$button_size_storage_numm=131
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage132
$button_size_storage_numm=132
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage133
$button_size_storage_numm=133
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage134
$button_size_storage_numm=134
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage135
$button_size_storage_numm=135
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage136
$button_size_storage_numm=136
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage137
$button_size_storage_numm=137
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage138
$button_size_storage_numm=138
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage139
$button_size_storage_numm=139
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage140
$button_size_storage_numm=140
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage141
$button_size_storage_numm=141
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage142
$button_size_storage_numm=142
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage143
$button_size_storage_numm=143
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage144
$button_size_storage_numm=144
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage145
$button_size_storage_numm=145
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage146
$button_size_storage_numm=146
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage147
$button_size_storage_numm=147
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage148
$button_size_storage_numm=148
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage149
$button_size_storage_numm=149
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage150
$button_size_storage_numm=150
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage151
$button_size_storage_numm=151
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage152
$button_size_storage_numm=152
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage153
$button_size_storage_numm=153
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage154
$button_size_storage_numm=154
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage155
$button_size_storage_numm=155
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage156
$button_size_storage_numm=156
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage157
$button_size_storage_numm=157
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage158
$button_size_storage_numm=158
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage159
$button_size_storage_numm=159
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage160
$button_size_storage_numm=160
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage161
$button_size_storage_numm=161
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage162
$button_size_storage_numm=162
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage163
$button_size_storage_numm=163
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage164
$button_size_storage_numm=164
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage165
$button_size_storage_numm=165
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage166
$button_size_storage_numm=166
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage167
$button_size_storage_numm=167
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage168
$button_size_storage_numm=168
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage169
$button_size_storage_numm=169
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage170
$button_size_storage_numm=170
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage171
$button_size_storage_numm=171
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage172
$button_size_storage_numm=172
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage173
$button_size_storage_numm=173
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage174
$button_size_storage_numm=174
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage175
$button_size_storage_numm=175
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage176
$button_size_storage_numm=176
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage177
$button_size_storage_numm=177
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage178
$button_size_storage_numm=178
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage179
$button_size_storage_numm=179
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage180
$button_size_storage_numm=180
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage181
$button_size_storage_numm=181
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage182
$button_size_storage_numm=182
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage183
$button_size_storage_numm=183
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage184
$button_size_storage_numm=184
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage185
$button_size_storage_numm=185
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage186
$button_size_storage_numm=186
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage187
$button_size_storage_numm=187
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage188
$button_size_storage_numm=188
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage189
$button_size_storage_numm=189
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage190
$button_size_storage_numm=190
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage191
$button_size_storage_numm=191
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage192
$button_size_storage_numm=192
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage193
$button_size_storage_numm=193
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage194
$button_size_storage_numm=194
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage195
$button_size_storage_numm=195
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage196
$button_size_storage_numm=196
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage197
$button_size_storage_numm=197
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage198
$button_size_storage_numm=198
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage199
$button_size_storage_numm=199
_button_stg_size($button_size_storage_numm)
Case $Backup_size_storage200
$button_size_storage_numm=200
_button_stg_size($button_size_storage_numm)
	#EndRegion
	#EndRegion
#Region Case Button_vms_in_stg
#Region
Case $Button_vms_in_stg0
$button_vms_in_stg_numm=0
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg1
$button_vms_in_stg_numm=1
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg2
$button_vms_in_stg_numm=2
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg3
$button_vms_in_stg_numm=3
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg4
$button_vms_in_stg_numm=4
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg5
$button_vms_in_stg_numm=5
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg6
$button_vms_in_stg_numm=6
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg7
$button_vms_in_stg_numm=7
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg8
$button_vms_in_stg_numm=8
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg9
$button_vms_in_stg_numm=9
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg10
$button_vms_in_stg_numm=10
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg11
$button_vms_in_stg_numm=11
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg12
$button_vms_in_stg_numm=12
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg13
$button_vms_in_stg_numm=13
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg14
$button_vms_in_stg_numm=14
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg15
$button_vms_in_stg_numm=15
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg16
$button_vms_in_stg_numm=16
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg17
$button_vms_in_stg_numm=17
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg18
$button_vms_in_stg_numm=18
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg19
$button_vms_in_stg_numm=19
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg20
$button_vms_in_stg_numm=20
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg21
$button_vms_in_stg_numm=21
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg22
$button_vms_in_stg_numm=22
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg23
$button_vms_in_stg_numm=23
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg24
$button_vms_in_stg_numm=24
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg25
$button_vms_in_stg_numm=25
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg26
$button_vms_in_stg_numm=26
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg27
$button_vms_in_stg_numm=27
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg28
$button_vms_in_stg_numm=28
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg29
$button_vms_in_stg_numm=29
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg30
$button_vms_in_stg_numm=30
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg31
$button_vms_in_stg_numm=31
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg32
$button_vms_in_stg_numm=32
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg33
$button_vms_in_stg_numm=33
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg34
$button_vms_in_stg_numm=34
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg35
$button_vms_in_stg_numm=35
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg36
$button_vms_in_stg_numm=36
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg37
$button_vms_in_stg_numm=37
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg38
$button_vms_in_stg_numm=38
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg39
$button_vms_in_stg_numm=39
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg40
$button_vms_in_stg_numm=40
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg41
$button_vms_in_stg_numm=41
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg42
$button_vms_in_stg_numm=42
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg43
$button_vms_in_stg_numm=43
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg44
$button_vms_in_stg_numm=44
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg45
$button_vms_in_stg_numm=45
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg46
$button_vms_in_stg_numm=46
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg47
$button_vms_in_stg_numm=47
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg48
$button_vms_in_stg_numm=48
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg49
$button_vms_in_stg_numm=49
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg50
$button_vms_in_stg_numm=50
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg51
$button_vms_in_stg_numm=51
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg52
$button_vms_in_stg_numm=52
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg53
$button_vms_in_stg_numm=53
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg54
$button_vms_in_stg_numm=54
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg55
$button_vms_in_stg_numm=55
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg56
$button_vms_in_stg_numm=56
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg57
$button_vms_in_stg_numm=57
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg58
$button_vms_in_stg_numm=58
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg59
$button_vms_in_stg_numm=59
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg60
$button_vms_in_stg_numm=60
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg61
$button_vms_in_stg_numm=61
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg62
$button_vms_in_stg_numm=62
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg63
$button_vms_in_stg_numm=63
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg64
$button_vms_in_stg_numm=64
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg65
$button_vms_in_stg_numm=65
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg66
$button_vms_in_stg_numm=66
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg67
$button_vms_in_stg_numm=67
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg68
$button_vms_in_stg_numm=68
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg69
$button_vms_in_stg_numm=69
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg70
$button_vms_in_stg_numm=70
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg71
$button_vms_in_stg_numm=71
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg72
$button_vms_in_stg_numm=72
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg73
$button_vms_in_stg_numm=73
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg74
$button_vms_in_stg_numm=74
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg75
$button_vms_in_stg_numm=75
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg76
$button_vms_in_stg_numm=76
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg77
$button_vms_in_stg_numm=77
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg78
$button_vms_in_stg_numm=78
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg79
$button_vms_in_stg_numm=79
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg80
$button_vms_in_stg_numm=80
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg81
$button_vms_in_stg_numm=81
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg82
$button_vms_in_stg_numm=82
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg83
$button_vms_in_stg_numm=83
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg84
$button_vms_in_stg_numm=84
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg85
$button_vms_in_stg_numm=85
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg86
$button_vms_in_stg_numm=86
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg87
$button_vms_in_stg_numm=87
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg88
$button_vms_in_stg_numm=88
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg89
$button_vms_in_stg_numm=89
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg90
$button_vms_in_stg_numm=90
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg91
$button_vms_in_stg_numm=91
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg92
$button_vms_in_stg_numm=92
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg93
$button_vms_in_stg_numm=93
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg94
$button_vms_in_stg_numm=94
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg95
$button_vms_in_stg_numm=95
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg96
$button_vms_in_stg_numm=96
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg97
$button_vms_in_stg_numm=97
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg98
$button_vms_in_stg_numm=98
_button_vms_in_stg($button_vms_in_stg_numm)
Case $Button_vms_in_stg99
$button_vms_in_stg_numm=99
_button_vms_in_stg($button_vms_in_stg_numm)
	#EndRegion
	#EndRegion

			EndSwitch
			WEnd

EndFunc   ;==>_Main

#Region _Parse
Func _Parse()

						local $sFileOpenDialog_vbm=Import_some_vbm()
						 ConsoleWrite("$NM_LDOWN: Import_some_vbm " & $sFileOpenDialog_vbm &@CR)
						if Not $sFileOpenDialog_vbm = 0 Then
						  	 local  $sText = FileRead($sFileOpenDialog_vbm)

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
local $numm_Vbm_Id = $RegExp_Vbm_Id[0]
local $VM_ids_in_meta = $numm_string
local $numm_JobId = $RegExp_JobId[0]
local $numm_JobName = $RegExp_JobName[0]
local $numm_JobType = $RegExp_JobType[0]
local $numm_SourceType = $RegExp_JobSourceType[0]
local $numm_Platform = $RegExp_Platform[0]
local $numm_CreationTimeUtc = $RegExp_CreationTimeUtc[0]
if $numm_FormatVersion = 0 Then
  local $numm_FormatVersion = 'not_set'
   Else
local $numm_FormatVersion = $RegExp_FormatVersion [0]
EndIf
local $numm_BackupCreationTime = $RegExp_BackupCreationTime[0]
local $array_VM_IDs = $RegExp_string
#EndRegion
;;
;Hosts
;Storages
;JobPoints
;Vms
;VmPoints
;;
local $array_Hosts = _Hosts_reparse($RegExp_Hosts)
local $array_Storages = _Storages_reparse($RegExp_Storages)
local $array_Storages_sorted = _Array_Storages_sorting($array_Storages)
local $array_JobPoints = _JobPoints_reparse($RegExp_JobPoints)
local $array_Vms = _VMs_reparse($RegExp_Vms)
local $array_VmPoints = _VMPoints_reparse($RegExp_VmPoints)
EndIf
#Region consistensy check
local $targetarray_Storages= UBound($array_Storages, $UBOUND_ROWS)
local $targetarray_Storages_sorted = UBound($array_Storages_sorted, $UBOUND_ROWS)
if $targetarray_Storages <> $targetarray_Storages_sorted Then MsgBox($MB_SYSTEMMODAL, "Fatal Error", "Orphaned increment storage!"&@CR&"See debug log for more info")
#EndRegion
Local $parce_func_array[17][2]
$parce_func_array[0][0] = '$array_Hosts'
$parce_func_array[0][1] = $array_Hosts
$parce_func_array[1][0] = '$array_Storages'
$parce_func_array[1][1] = $array_Storages
$parce_func_array[2][0] = '$array_Storages_sorted'
$parce_func_array[2][1] = $array_Storages_sorted
$parce_func_array[3][0] = '$array_JobPoints'
$parce_func_array[3][1] = $array_JobPoints
$parce_func_array[4][0] = '$array_Vms'
$parce_func_array[4][1] = $array_Vms
$parce_func_array[5][0] = '$array_VmPoints'
$parce_func_array[5][1] = $array_VmPoints
$parce_func_array[6][0] = '$numm_Vbm_Id'
$parce_func_array[6][1] = $numm_Vbm_Id
$parce_func_array[7][0] = '$VM_ids_in_meta'
$parce_func_array[7][1] = $VM_ids_in_meta
$parce_func_array[8][0] = '$numm_JobId'
$parce_func_array[8][1] = $numm_JobId
$parce_func_array[9][0] = '$numm_JobName'
$parce_func_array[9][1] = $numm_JobName
$parce_func_array[10][0] = '$numm_JobType'
$parce_func_array[10][1] = $numm_JobType
$parce_func_array[11][0] = '$numm_SourceType'
$parce_func_array[11][1] = $numm_SourceType
$parce_func_array[12][0] = '$numm_Platform'
$parce_func_array[12][1] = $numm_Platform
$parce_func_array[13][0] = '$numm_CreationTimeUtc'
$parce_func_array[13][1] = $numm_CreationTimeUtc
$parce_func_array[14][0] = '$numm_FormatVersion'
$parce_func_array[14][1] = $numm_FormatVersion
$parce_func_array[15][0] = '$numm_BackupCreationTime'
$parce_func_array[15][1] = $numm_BackupCreationTime
$parce_func_array[16][0] = '$array_VM_IDs'
$parce_func_array[16][1] = $array_VM_IDs
Return($parce_func_array)

#EndRegion
						EndIf
						EndFunc
#EndRegion
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
local $target_index = UBound($array_Storages_sorted, $UBOUND_ROWS) -1
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
local $target_source= UBound($array_Storages, $UBOUND_ROWS)-1
local $target_index = UBound($array_Storages_sorted, $UBOUND_ROWS)-1
;sorting by date
if $target_index <> 0 Then
	;;need to sort vbk by time
	EndIf




While $target_index < $target_source
$target_index = UBound($array_Storages_sorted, $UBOUND_ROWS)-1
For $i = 1 To $numm_array_Storages step +1
if $array_Storages[$i][5] <> '0000000-0000-0000-0000-000000000000' Then
		local $iIndex_exist_in_vbk = _ArraySearch($array_Storages_sorted, $array_Storages[$i][5], 0, 0, 0, 1, 1, 5) ;есть ли ты в $array_Storages_sorted
		if $iIndex_exist_in_vbk = -1 Then
			local $iIndex_link = _ArraySearch($array_Storages, $array_Storages[$i][5], 0, 0, 0, 1, 1, 0) ;на кого ты ссылаешься
				local $iIndex_exist_in_vbk_link =_ArraySearch($array_Storages_sorted, $array_Storages[$iIndex_link][0], 0, 0, 0, 1, 1, 0);есть ли твой перент в $array_Storages_sorted
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
if $storages_sorted_test = True  Then
FileDelete('Storages_sorted.txt')
_FileWriteFromArray('Storages_sorted.txt', $array_Storages_sorted)
EndIf
ConsoleWrite('$Storages_sorted is done'&@cr)
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
  local  $text = $RegExp_Storage_temp[$i]
  local  $trim = StringInStr($text, '" LinkId="')
   if not $trim=0 Then
 local  $LinkId_temp =  StringTrimLeft($text, $trim+9)
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
;_ArrayDisplay($Storages_array_rep)
EndIf
ConsoleWrite('$Storages_array_rep is done'&@cr)
GUICtrlSetData($idProgressbar1, 40)


Return $Storages_array_rep
EndFunc
Func _VMPoints_reparse($RegExp_VMPoints)
local $VMPoints_reparse_array_rep[1][10]
$VMPoints_reparse_array_rep[0][0] = 'Point Id'
$VMPoints_reparse_array_rep[0][1] = 'LinkId'
$VMPoints_reparse_array_rep[0][2] = 'IsCorrupted'
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

local $RegExp_Points_temp = StringRegExp($RegExp_VMPoints, '(?s)" IsCorrupted="(.*?)"', 3)
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
		local  $stop = 'true'
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
       local  $sFileOpenDialog_vbm = StringReplace($sFileOpenDialog, "|", @CRLF)

        ; Display the list of selected files.
        ;MsgBox($MB_SYSTEMMODAL, "", "You chose the following files:" & @CRLF & $sFileOpenDialog)
		Return($sFileOpenDialog_vbm)
    EndIf
EndFunc   ;==>Example
Func _ByteSuffix($iBytes, $iRound = 2) ; By Spiff59
    Local $A, $aArray[9] = [" B", " KB", " MB", " GB", " TB", " PB", "EB", "ZB", "YB"]
    While $iBytes > 1023
        $A += 1
        $iBytes /= 1024
    WEnd
    Return Round($iBytes, $iRound) & $aArray[$A]
EndFunc   ;==>_ByteSuffix
Func _button_stg($button_numm, $array_VmPoints)
	For $i = 0 To 99 step +1 ;delete all

												Global $sEvalvms_in_stg = Eval('Button_vms_in_stg'&$i)
												GUICtrlDelete($sEvalvms_in_stg)
												Global $sEvalvms_in_stg_IsCorrupted = Eval('Button_vms_in_stg_IsCorrupted'&$i)
												GUICtrlDelete($sEvalvms_in_stg_IsCorrupted)
Next
	;ConsoleWrite($button_numm&@CR)
	local $avArray = $array_VmPoints
	local $stg_id_to_search = $array_Storages_sorted[$button_numm][0]
	local $iIndex = _ArrayFindAll($avArray, $stg_id_to_search, 0, 0, 0, 0, 7)
	local $numm_vms_in_stg = UBound($iIndex, $UBOUND_ROWS)
	ConsoleWrite('$stg_id_to_search = '&$stg_id_to_search&' found = '&$numm_vms_in_stg&@CR)
if $numm_vms_in_stg <> 0 Then
local $Group3 = GUICtrlCreateGroup("", 485, 48, 490, 670)
For $i = 0 To $numm_vms_in_stg-1 step +1
	;$vm_to_show = StringRegExp($array_VmPoints[$iIndex[$i]][6],'\s\((.*?)\)', 3); isn't working...
	$vm_to_show = StringTrimLeft($array_VmPoints[$iIndex[$i]][6],38)
	$vm_to_show = StringTrimRight($vm_to_show,1)
		;ConsoleWrite($vm_to_show&@CR)
				Global $sEvalvms_in_stg = GUICtrlCreateLabel('VM ID: '&$vm_to_show&' ('&$array_VmPoints[$iIndex[$i]][3]&')', 499, 65+18*$i, 400, 25, BitOR($SS_Left, $SS_CENTERIMAGE))
							Assign('Button_vms_in_stg'&$i, $sEvalvms_in_stg,2)
								$sEvalvms_in_stg_id = $array_VmPoints[$iIndex[$i]][0]
												Assign('Button_vms_in_stg_id'&$i, $sEvalvms_in_stg_id,2);holds vm id from button


																Global $sEvalvms_in_stg_IsCorrupted = GUICtrlCreateLabel('IsCorrupted: '&$array_VmPoints[$iIndex[$i]][2], 870, 65+18*$i, 100, 25, BitOR($SS_right, $SS_CENTERIMAGE))
																	Assign('Button_vms_in_stg_IsCorrupted'&$i, $sEvalvms_in_stg_IsCorrupted,2)

Next
GUICtrlCreateGroup("", -99, -99, 1, 1)
Else
	;nothing to display
	For $i = 0 To 0 step +1
	local $Group3 = GUICtrlCreateGroup("", 485, 48, 490, 670)
					Local $sEvalvms_in_stg = GUICtrlCreateLabel('Nothing to display, this storage is empty', 499, 65+18*$i, 400, 25, BitOR($SS_Left, $SS_CENTERIMAGE))
			Assign('Button_vms_in_stg'&$i, $sEvalvms_in_stg,2)
												Local $sEvalString = Eval('Button_vms_in_stg'&$i)
													GUICtrlCreateGroup("", -99, -99, 1, 1)
	Next
EndIf
EndFunc
Func _button_stg_size($button_size_storage_numm)
	$temp_id = Eval('Button_stg_id'&$button_size_storage_numm)
	ConsoleWrite($temp_id&@CR)
	local $iIndex = _ArraySearch($array_Storages_sorted, $temp_id, 0, 0, 0, 0, 1, 1)
			MsgBox($MB_SYSTEMMODAL, "Not parsed info", $array_Storages_sorted[$iIndex][13])
EndFunc
Func _button_vms_in_stg($button_vms_in_stg_numm)
	$temp_id = Eval('Button_vms_in_stg_id'&$button_vms_in_stg_numm)
	local $iIndex = _ArraySearch($array_VmPoints, $temp_id, 0, 0, 0, 0, 1, 0)
			MsgBox($MB_SYSTEMMODAL, "Not parsed info", $array_VmPoints[$iIndex][9])
EndFunc