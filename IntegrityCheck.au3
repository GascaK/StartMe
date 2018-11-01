;=========================
; Function Name: _VerifyIntegritySMG()
;
; Checks for .ini's for the StartMeGUI
; application.
;=========================
Func __VerifyIniIntegritySMG($sFilePath)

	;Checking configuration files;
	If FileExists($sFilePath) NOT Then
		MsgBox($MB_OK, "StartMeGUI", "File does not exist. Creating .ini . . .")
		FileOpen($sFilePath, $FO_READ + $FO_OVERWRITE + $FO_CREATEPATH)
		If @error Then
			MsgBox("", "StartMeGui - Integrity", "Unable to create .ini's")
			Return False
		EndIf

		IniWrite($sFilePath , "StartMe - UserNames", "HotSOS", "Change Here")
		IniWrite($sFilePath , "StartMe - UserNames", "LMS", "Change Here")
		IniWrite($sFilePath, "SnitchMe - Columns", "Engineer", "A")
		IniWrite($sFilePath, "SnitchMe - Columns", "Status", "B")
		IniWrite($sFilePath, "SnitchMe - Columns", "Trade", "C")
		IniWrite($sFilePath, "SnitchMe - Columns", "Room", "D")
		IniWrite($sFilePath, "SnitchMe - Columns", "Issue", "E")
		IniWrite($sFilePath, "SnitchMe - Columns", "Order", "F")
		ShellExecute($sFilePath)
		Return True

	Else
		Return True
	EndIf

EndFunc

;=========================
; Function Name: __VerifyPassIntegritySMG()
;
; Checks for .smg files in specified file path.
;=========================
Func __VerifyPassIntegritySMG($sFilePath)

	; Checking .smg files ;
	If FileExists($sFilePath) NOT Then
		MsgBox($MB_OK, "StartMeGUI", "File does not exist. Creating .smg. . .")
		FileOpen($sFilePath, $FO_READ + $FO_OVERWRITE + $FO_CREATEPATH)
		If @error Then
			MsgBox("", "StartMeGUI - Integrity", "Unable to create .smg's")
			Return False
		EndIf
		FileWrite($sFilePath, "tempPass")
		Return True
	Else
		Return True
	EndIf

EndFunc