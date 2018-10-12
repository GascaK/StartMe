; Script Name:  StartMeGUI.au3
; Desc:     A work related script that
;			controls various job related
;			duties.
; Author:       Gasca, K
; Created:      2 / 1 / 2017
; Version:      2.0

#include <MsgBoxConstants.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <WinAPIFiles.au3>
#include <File.au3>
#include <Crypt.au3>

; UDF's ;
#include "SnitchMe2.au3"

Opt("WinTitleMatchMode", 2)
Opt("WinWaitDelay", 1000)
Opt("MouseCoordMode", 0)
Opt("SendKeyDelay", 15)
Global $StartMe = GUICreate("StartMe GUI v 1.02", 315, 150, -1, -1)

; Cryptography ;
_Crypt_Startup();
Local $hkey = _Crypt_DeriveKey(StringToBinary("LiamisBadass"), $CALG_AES_128) ; Change "LiamisBadass" for machine specific encryption.
;---------------;

; Menu Selection ;
$idFileMenu = GUICtrlCreateMenu("File")
$idLMSPass = GUICtrlCreateMenuItem("LMS Password", $idFileMenu)
$idHOTPass = GUICtrlCreateMenuItem("HotSOS Password", $idFileMenu)
$idSnitchMe = GUICtrlCreateMenuitem("SnitchMe", $idFileMenu)
;---------------;

; Create Main Controller Window ;
$startText = GUICtrlCreateLabel("Click to Launch and Login:", 32, 16, 252, 28, $SS_CENTER)
GUICtrlSetFont(-1, 14, 800, 0, "MS Sans Serif")
GUICtrlSetColor(-1, 0x000000)
$hotButton = GUICtrlCreateButton("HotSOS", 40, 56, 91, 33)
$lmsButton = GUICtrlCreateButton("LMS", 168, 56, 91, 33)
Global $hotBox = GUICtrlCreateCheckbox("Prevent Auto Log ", 32, 96, 113, 17)
GUISetState(@SW_SHOW, $StartMe)
;---------------;

; Password Retrieval ;
Global $saveIndexHOT = "hsIndex.smg"
Global $saveIndexLMS = "lmsIndex.smg"
Local $sHsUser = retUser("sHsUser")
Global $sHotPass = retPass($saveIndexHOT)
Local $sLmsUser = retUser("sLmsUser")
Global $sLmsPass = retPass($saveIndexLMS)
;---------------;
Local $counter = 0

While 1
   $counter += 1
   $nMsg = GUIGetMsg($GUI_EVENT_ARRAY)
   Switch $nMsg[0]
	  Case $GUI_EVENT_CLOSE
		 Exit
	  ; Change LMS Password ;
	  Case $idLMSPass
		 GUISetState(@SW_DISABLE, $StartMe)
		 changeLMS()
		 retPass($saveIndexLMS)
		 GUISetState(@SW_ENABLE, $StartMe)
	  ; Change Hotsos Password ;
	  Case $idHOTPass
		 GUISetState(@SW_DISABLE, $StartMe)
		 changeHOT()
		 retPass($saveIndexHOT)
		 GUISetState(@SW_ENABLE, $StartMe)
	  ; Run Snitch report ;
	  Case $idSnitchMe
		 SnitchMe2()
	  ; Run Hotsos ;
	  Case $hotButton
		 GUISetState(@SW_DISABLE, $StartMe)
		 loginHOT(retUser("sHsUser"), retPass($saveIndexHOT))
		 GUISetState(@SW_ENABLE, $StartMe)
	  ; Run LMS ;
	  Case $lmsButton
		 GUISetState(@SW_DISABLE, $StartMe)
		 loginLMS(retUser("sLmsUser"), retPass($saveIndexLMS))
		 GUISetState(@SW_ENABLE, $StartMe)
	  ; nothing.. ;
	  Case $hotBox

   EndSwitch

   If _IsChecked($hotBox) And ($counter >= 30000) Then
	  If WinActivate("HotSOS") Then
		 Sleep(1000)
		 Send("{F5}")
		 WinSetState("HotSOS", "", @SW_MINIMIZE)
	  EndIf
	  $counter = 0
   EndIf
WEnd

; Open HOTSOS in .exe ;
Func loginHOT ($userName, $passWord)
   Global $hsID = Run("C:\Program Files (x86)\MTech\hotsos\client_na2\HotSOS.exe")
   WinWaitActive("Login")
   Send($userName & "{TAB}" & $passWord & "{ENTER}")
   WinWaitActive("HotSOS")
   MouseClick("primary", 89, 285)
   Sleep(2000)
   Send("{TAB}" & "fac - ")
   Send("{ENTER 2}")
EndFunc

; Open LMS in .ws ;
Func loginLMS($userName, $passWord)

   ShellExecute("C:\Users\Public\Desktop\LMS.ws")
   Sleep(2000)
   if WinExists("PC5250") Then
	  Send("{ENTER}")
	  Send("{TAB}" & $userName & "{ENTER}")
   EndIf
	  ;Local $hWND = WinGetHandle("IBM")
   Send($passWord & "{TAB 3}" & $userName & "{ENTER}")
	  ;Send($userName & "{TAB}" & $passWord & "{ENTER}")

   WinWaitActive("HOTEL(LMS)")
   Send($userName & "{TAB}" & $passWord & "{ENTER}")
   Send("1" & "{ENTER 2}" & "1" & "{ENTER}")

EndFunc

; Change Password Settings LMS;
Func changeLMS()

   ; Create popUP for LMS ;
   Local $gChangeLMS = GUICreate("", 300, 60)
   Local $labelPass = GUICtrlCreateLabel("LMS Password: ", 16, 16, 81, 17)
   Local $lmsInput = GUICtrlCreateInput("New Password", 112, 16, 121)
   Local $lmsButton = GUICtrlCreateButton("Change", 250, 16)
   GUICtrlSetLimit($lmsInput, 10)
   GUISetState(@SW_SHOW, $gChangeLMS)

   While 1
	  $idMsg = GUIGetMsg($GUI_EVENT_ARRAY)
	  if $idMsg[0] = $lmsButton Then
		 $sLmsPass = GUICtrlRead($lmsInput) ; Edits the GLOBAL LMS password
		 savePass($sLmsPass, $saveIndexLMS) ; Save Pass into saveIndexLMS
		 ExitLoop
	  ElseIf $idMsg[0] = $GUI_EVENT_CLOSE Then
		 ExitLoop
	  EndIf
   WEnd
   GUIDelete($gChangeLMS)

EndFunc

; Change Password Settings HOT ;
Func changeHOT()

   ; Create popUP for HOTsos ;
   Local $gChangeHOT = GUICreate("", 300, 60)
   Local $labelPass = GUICtrlCreateLabel("HOT Password: ", 16, 16, 81, 17)
   Local $hotInput = GUICtrlCreateInput("New Password", 112, 16, 121)
   Local $hotButton = GUICtrlCreateButton("Change", 250, 16)
   GUISetState(@SW_SHOW, $gChangeHOT)

   While 1
	  $idMsg = GUIGetMsg($GUI_EVENT_ARRAY)
	  if $idMsg[0] = $hotButton Then
		 $sHotPass = GUICtrlRead($hotInput) ; Edits the GLOBAL HOT password
		 savePass($sHotPass, $saveIndexHOT) ; Save Pass into saveIndexHOT
		 ExitLoop
	  ElseIf $idMsg[0] = $GUI_EVENT_CLOSE Then
		 ExitLoop
	  EndIf
   WEnd
   GUIDelete($gChangeHOT)

EndFunc

; Save Password ;
Func savePass($passInput, $fileSaved)

   Local $sFilePath = @WorkingDir & "\Res\" & $fileSaved
   Local $hFile = FileOpen($sFilePath, $FO_OVERWRITE)
   Local $dEncrypt = _Crypt_EncryptData($passInput, $hkey, $CALG_USERKEY)

   if FileExists($sFilePath) NOT Then
	  MsgBox($MB_SYSTEMMODAL, "StartMe GUI", "File does not exist. Shut Down and restart with Admin Priveleges?")
   Else
	  FileWrite($hFile, $dEncrypt)
   EndIf

   FileClose($hFile)
EndFunc

; Return Password ;
Func retPass($sPassFileName)

   Local $sFilePath = @WorkingDir & "\Res\" & $sPassFileName

   If FileExists($sFilePath) NOT Then
	  Local $hFile = FileOpen($sFilePath, $FO_READ + $FO_OVERWRITE + $FO_CREATEPATH)
	  FileWrite($hFile, "tempPass")
	  MsgBox($MB_SYSTEMMODAL, "StartMe GUI", "Change Password before using!")
	  Return "tempPass"

   Else
	  Local $hFile = FileOpen($sFilePath, $FO_READ)
   EndIf

   $sFileRead = FileRead($hFile)
   Local $dDecrypted = _Crypt_DecryptData($sFileRead, $hKey, $CALG_USERKEY)

   FileClose($hFile)
   Return BinaryToString($dDecrypted)

EndFunc

; Return Usernames ;
Func retUser($sToReturn)

   Local $sFilePath = @WorkingDir & "\Res\startMe.ini"

   if FileExists($sFilePath) NOT Then
	  MsgBox($MB_OK, "StartMe GUI", "File Does not exist. Creating .ini . . .")
	  DirCreate($sFilePath)
	  IniWrite($sFilePath , "StartMe - UserNames", "HotSOS", "Change Here")
	  IniWrite($sFilePath , "StartMe - UserNames", "LMS", "Change Here")
	  ShellExecute($sFilePath )

   ElseIf $sToReturn = "sHsUser" Then
	  return IniRead($sFilePath & "startMe.ini", "StartMe - UserNames", "HotSOS", "Default")

   ElseIf $sToReturn = "sLmsUser" Then
	  return IniRead($sFilePath & "startMe.ini", "StartMe - UserNames", "LMS", "Default")
   EndIf

EndFunc

; Check for Control Checked ;
Func _IsChecked($idControlID)
    Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc   ;==>_IsChecked