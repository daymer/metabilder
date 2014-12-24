#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Comment=beta
#AutoIt3Wrapper_Res_Fileversion=0.1.0.3
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
$sFileOpenDialog_VBK = ''
#EndRegion

While 1
#Region ### START Koda GUI section ### Form=c:\users\dremsama\desktop\koda\templates\tabbed pages.kxf
$dlgTabbed = GUICreate("Veeam Meta Bilder ver 0.1 beta", 810, 550, @DesktopWidth / 2 - 405, @DesktopHeight / 2 - 225 )
GUISetIcon("", -1)


$PageControl1 = GUICtrlCreateTab(8, 8, 602, 425)
GUICtrlSetFont(-1, 12, 400, 0, "Arial")
#Region Initialization
$TabSheet1 = GUICtrlCreateTabItem("Initialization")

$Button_Browse_vbk = GUICtrlCreateButton("Browse", 512, 79, 75, 25)
GUICtrlSetFont(-1, 10, 400, 0, "Arial")
$Label1 = GUICtrlCreateLabel("Step #1: Import a VBK", 32, 50, 164, 20)
GUICtrlSetFont(-1, 10, 400, 0, "Arial")
$Input_vbk = GUICtrlCreateInput($sFileOpenDialog_VBK, 32, 80, 473, 24)
GUICtrlSetFont(-1, 10, 400, 0, "Arial")



#EndRegion
GUICtrlCreateTabItem("")
$Button1 = GUICtrlCreateButton("OK", 358, 456, 75, 25)
$Button2 = GUICtrlCreateButton("Cancel", 446, 456, 75, 25)
$Button3 = GUICtrlCreateButton("Help", 528, 456, 75, 25)

GUISetState(@SW_SHOW)



#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
			Case $Button_Browse_vbk
$sFileOpenDialog_VBK = Import_some_vbk()
GUIDelete($dlgTabbed)
ContinueLoop(2)
;MsgBox($MB_SYSTEMMODAL, "", "You chose the following files:" & @CRLF & $sFileOpenDialog_VBK)
Case $Button3
$sFileOpenDialog_VBK =  GUICtrlRead($Input_vbk)
if $sFileOpenDialog_VBK = '' Then
   MsgBox($MB_SYSTEMMODAL, "Warning","Import VBK first")
Else
$Dispatch_port_server = _start_server_agent()
_start_client_agent($Dispatch_port_server, $sFileOpenDialog_VBK)
	EndIf
	EndSwitch



WEnd
WEnd







Func Import_some_vbk()
    ; Create a constant variable in Local scope of the message to display in FileOpenDialog.
    Local Const $sMessage = "Import VBK."

    ; Display an open dialog to select a list of file(s).
    Local $sFileOpenDialog = FileOpenDialog($sMessage, @WindowsDir & "\", "Full Backups (*.vbk;)", $FD_FILEMUSTEXIST + $FD_MULTISELECT)
    If @error Then
        ; Display the error message.
        ;MsgBox($MB_SYSTEMMODAL, "", "No file(s) were selected.")
Exit
        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)
    Else
        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)

        ; Replace instances of "|" with @CRLF in the string returned by FileOpenDialog.
        $sFileOpenDialog_VBK = StringReplace($sFileOpenDialog, "|", @CRLF)

        ; Display the list of selected files.
        ;MsgBox($MB_SYSTEMMODAL, "", "You chose the following files:" & @CRLF & $sFileOpenDialog)
		Return($sFileOpenDialog_VBK)
    EndIf
EndFunc   ;==>Example



Func _start_server_agent()
$iPID = Run(@ComSpec & " /c " & '"'&_ProgramFilesDir()&'\Veeam\Backup Transport\VeeamAgent.exe"', "", @SW_HIDE, $STDIN_CHILD + $STDOUT_CHILD)
Sleep(500)
$Message = StdoutRead($iPID, True)
;ConsoleWrite( $Message)
$aLines = StringSplit($Message, @CRLF, 1)
Global $client[1][2]
 $client[0][0] = 'line'
  $client[0][1] = 'value'
;_ArrayDisplay($aLines)
For $i = 1 To $aLines[0] Step +1
   If $aLines[$i] = '' Then
	  ContinueLoop
   EndIf
  $aLines_temp = StringSplit($aLines[$i], ':', 1)
  ;_ArrayDisplay($aLines_temp)
  if $aLines_temp[1] ='>' Then ContinueLoop
Local $Temp_writer[1][2]
$Temp_writer[0][0] = $aLines_temp[1]
$Temp_writer[0][1] = $aLines_temp[2]
_ArrayConcatenate($client, $Temp_writer)
Next
;_ArrayDisplay($client)

$client_ub = UBound($client, 1) -1
For $i = 1 To $client_ub Step +1
   if $client[$i][0] = 'Dispatch port' Then $client_server_Dispatch_port = $client[$i][1]
   Next
StdinWrite($iPID, "startServer" & @CRLF & "us" & @CRLF & "ps" & @CRLF)
Sleep(500)
$Message = StdoutRead($iPID, True)
;ConsoleWrite( $Message)
$aLines = StringSplit($Message, @CRLF, 1)
Global $client[1][2]
 $client[0][0] = 'line'
  $client[0][1] = 'value'
;_ArrayDisplay($aLines)
For $i = 1 To $aLines[0] Step +1
   If $aLines[$i] = '' Then
	  ContinueLoop
   EndIf
  $aLines_temp = StringSplit($aLines[$i], ':', 1)
  ;_ArrayDisplay($aLines_temp)
  if $aLines_temp[1] ='>' Then ContinueLoop
Local $Temp_writer[1][2]
$Temp_writer[0][0] = $aLines_temp[1]
if IsDeclared("$aLines_temp[2]") Then
   $Temp_writer[0][1] = $aLines_temp[2]
Else
   $Dispatch_port_server = $Temp_writer[0][0]
   EndIf
_ArrayConcatenate($client, $Temp_writer)
Next
;_ArrayDisplay($client)
#Region ConsoleWrite
ConsoleWrite('$Dispatch_port '&$client_server_Dispatch_port&@CR)
ConsoleWrite('$Dispatch_port server '&$Dispatch_port_server&@CR)
#EndRegion
Return($Dispatch_port_server)
EndFunc


Func _start_client_agent($Dispatching_port_server, $sFileOpenDialog_VBK)
$iPID2 = Run(@ComSpec & " /c " & '"'&_ProgramFilesDir()&'\Veeam\Backup Transport\VeeamAgent.exe"', "", @SW_HIDE, $STDIN_CHILD + $STDOUT_CHILD)
Sleep(500)
$Dispatching_port_server = Number($Dispatching_port_server)
StdinWrite($iPID2, "connectByIPs" & @CRLF & "127.0.0.1" & @CRLF & "." & @CRLF& $Dispatching_port_server& @CRLF & @CRLF& "us" & @CRLF& "ps" & @CRLF&"dir" & @CRLF & $sFileOpenDialog_VBK & @CRLF)
Sleep(1000)

;StdinWrite($iPID2, "dir" & @CRLF & $sFileOpenDialog_VBK & @CRLF)
$Message = StdoutRead($iPID2, True)
;ConsoleWrite( $Message)
;ConsoleWrite( $sFileOpenDialog_VBK)
$aLines = StringSplit($Message, @CRLF, 1)
Global $client[1][2]
 $client[0][0] = 'line'
  $client[0][1] = 'value'
;_ArrayDisplay($aLines)
For $i = 1 To $aLines[0] Step +1
   If $aLines[$i] = '' Then
	  ContinueLoop
   EndIf
  $aLines_temp = StringSplit($aLines[$i], '/')
  ;_ArrayDisplay($aLines_temp)
  Local $Temp_writer[1][2]
$Temp_writer[0][0] = $aLines_temp[1]
if  $aLines_temp[0] = 1 Then
   $Temp_writer[0][1] = 'not declared'
Elseif $aLines_temp[0] = 2 Then
   $Temp_writer[0][1] = $aLines_temp[2]
   EndIf
_ArrayConcatenate($client, $Temp_writer)
Next
;_ArrayDisplay($client)
$iIndex = _ArraySearch($client, 'summary.xml', 0, 0, 0, 1)
$premeta_num = StringInStr ($client[$iIndex][0], ' ')
$premeta = StringTrimLeft($client[$iIndex][0],$premeta_num)
$source_meta = 'veeamfs:0:'&$premeta&'/summary.xml@'&$sFileOpenDialog_VBK
;ConsoleWrite( $source_meta&@cr)
StdinWrite($iPID2, "restore" & @CRLF & $temp_dir&'/summary.xml' & @CRLF & $source_meta & @CRLF&'.'& @CRLF)
$Message = StdoutRead($iPID2, True)
;ConsoleWrite( $Message)
ProcessClose ($iPID2)
#Region ConsoleWrite
ConsoleWrite('restore of '&$premeta& ' is done' &@cr)
#EndRegion
EndFunc


Func _ProgramFilesDir()
    Local $ProgramFileDir
    Switch @OSArch
        Case "X32"
            $ProgramFileDir = "Program Files"
        Case "X64"
            $ProgramFileDir = "Program Files (x86)"
    EndSwitch
    Return @HomeDrive & "\" & $ProgramFileDir
EndFunc   ;==>_ProgramFilesDirh

