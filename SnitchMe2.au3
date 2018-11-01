;===============================
;Script Name:  SnitchMe.au3
; Desc: A work related script that
; creates a detailed list of
; work related search queries
; Author:       Gasca, K
; Created:      7 / 1 / 2017
; Version:      2.0
;===============================

#include-once
#include <Excel.au3>
#include <File.au3>
#include <Array.au3>
#include <EditConstants.au3>
#include <AutoItConstants.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <WindowsConstants.au3>


Func SnitchMe2()

    Opt("ExpandVarStrings", 1) ; For cleanliness.
	Local $oLocalCOMErrorHandler = ObjEvent("AutoIt.Error", "_ErrFunc")

	; Create the Report Window ;
	Local $hGUIhan = GUICreate("StartMeGUI - SnitchMe", 500, 700, -1, $WS_EX_ACCEPTFILES)
	Local $listView = GUICtrlCreateListView("Row # | HotSOS | Engineer | Room | Issue ", -1, -1, 500, 550)
	$inputEdit = GUICtrlCreateInput("", 150, 600, 335, 25)
	GUICtrlSetState(-1, $GUI_DROPACCEPTED)
	$dndLabel = GUICtrlCreateLabel("Drag and Drop Row Here", 8, 600, 125, 17)
	$clipButton = GUICtrlCreateButton("Copy to Clipboard", 152, 650, 163, 25)
	;---------------;

	; Create the Progress Bar Window ;
	ProgressOn("SnitchMe2 - Progress..", "Warming Up...", "0%", -1, -1, $DLG_MOVEABLE)
	;---------------;

	; Open Excel (.CSV) object ;
	Local $oExcel = _Excel_Open(False)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "SnitchMe", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	Local $sWork = @ScriptDir & "\orders\hotelOrders.csv"
	Local $oWork = _Excel_BookOpen($oExcel, $sWork)
	ProgressSet(0, 0 & "%", "Open: " & $sWork)
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "SnitchMe", "Error opening '" & $sWork & "'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Excel_Close($oExcel)
		Exit
	EndIf
	;----------------;

	; Assign cell variables and Delete unqualified cells ;
	Local $aHotelRoom[0][4]
	Local $i = 1
	Local $sValue = Null
	;----------------;

	; Assisgn Columns ;
	Local $sEngRow = GetSCol("Engineer")
	Local $sStatusRow = GetSCol("Status")
	Local $sTradeRow = GetSCol("Trade")
	Local $sRoomRow = GetSCol("Room")
	Local $sIssueRow = GetSCol("Issue")
	Local $sOrderRow = GetSCol("Order")
	;----------------;

	Do

		$iRows = $oExcel.ActiveSheet.UsedRange.Rows.Count ; Continously update Row Count

		$Engineer = _Excel_RangeRead($oWork, Default, $sEngRow & $i)

		$iCalcProg = ($i / $iRows) * 100 ; Perform Progess Calc only once per loop (Efficiency?)
		ProgressSet($iCalcProg, "", "Checking Row: $i$ of $iRows$") ; Update Progress

	    If $Engineer = "" Then
			ProgressSet($iCalcProg, "Deleting Row: $i$ @@ $sEngRow$ $i$ ")
			_Excel_RangeDelete($oWork.ActiveSheet, $oWork.ActiveSheet.rows($i), $xlShiftUp, 1)
			If @error Then Exit MsgBox($MB_SYSTEMMODAL, "SnitchMe", "Error deleting Row @ I" & $i & @CRLF & "@error = " & @error & ", @extended = " & @extended)

		ElseIf _Excel_RangeRead($oWork, Default, "$sStatusRow$" & "$i$") <> "Closed" Then
			ProgressSet($iCalcProg, "Deleting Row: $i$ @@ $sStatusRow$ $i$ ")
			_Excel_RangeDelete($oWork.ActiveSheet, $oWork.ActiveSheet.rows($i), $xlShiftUp, 1)
			If @error Then Exit MsgBox($MB_SYSTEMMODAL, "SnitchMe", "Error deleting Row @ C" & $i & @CRLF & "@error = " & @error & ", @extended = " & @extended)

		ElseIf _Excel_RangeRead($oWork, Default, "$sTradeRow$" & "$i$") <> "FAC - Hotel Shop" Then
			ProgressSet($iCalcProg, "Deleting Row: $i$ @@ $sTradeRow$ $i$")
			_Excel_RangeDelete($oWork.ActiveSheet, $oWork.ActiveSheet.rows($i), $xlShiftUp, 1)
			If @error Then Exit MsgBox($MB_SYSTEMMODAL, "SnitchMe", "Error deleting Row @ E" & $i & @CRLF & "@error = " & @error & ", @extended = " & @extended)

		Else
			$aGValue = StringSplit(_Excel_RangeRead($oWork, Default, "$sRoomRow$" & "$i$"), " - ", $STR_ENTIRESPLIT)
			; Only Add to Array if verified number in String $aGValue (ROOM Cell)
			For $b = 1 To $aGValue[0]

				If Number($aGValue[$b]) <> 0 Then
				    $Issue = _Excel_RangeRead($oWork, Default, "$sIssueRow$" & "$i$")
					$HotSOS = _Excel_RangeRead($oWork, Default, "$sOrderRow$" & "$i$")

 					_ArrayAdd($aHotelRoom, $HotSOS & '|' & $Engineer & '|' & $aGValue[$b] & '|' & $Issue, 0)
					$i += 1 ; Increment row counter
					ExitLoop
				ElseIf $b = $aGValue[0] Then ; (Room Cell) has no valid room number
					ProgressSet($iCalcProg, "Deleting Row: $i$ @@ $sRoomRow$ $i$")
					_Excel_RangeDelete($oWork.ActiveSheet, $oWork.ActiveSheet.rows($i), $xlShiftUp, 1)
					If @error Then Exit MsgBox($MB_SYSTEMMODAL, "SnitchMe", "Error deleting cells. @ G" & $i & @CRLF & "@error = " & @error & ", @extended = " & @extended)
				EndIf
			Next

		EndIf

	Until $i = $iRows + 1

	; Array Sorting ;
	ProgressSet(100, "", "Sorting Array..")
	_ArraySort($aHotelRoom, 0,0,0,2)
	$holdValue = 0
	$temp = 0

	; Verify total room # count ;
	For $e = 0 To UBound($aHotelRoom, $UBOUND_ROWS) - 1
		GUICtrlCreateListViewItem($e + 1 & "|" & $aHotelRoom[$e][0] & "|" & $aHotelRoom[$e][1] & "|" & $aHotelRoom[$e][2] & "|" & $aHotelRoom[$e][3], $listView)

		If $temp <> $aHotelRoom[$e][2] Then
			$holdValue += 1
		EndIf
	$temp = $aHotelRoom[$e][2]
	Next

	ProgressOff()
	$roomLabel = GUICtrlCreateLabel("Total of: $holdValue$ different rooms." , 200, 575, 150, 25)
	GUISetState(@SW_SHOW, $hGUIhan)
	_Excel_Close($oExcel)

	; Main Loop controller ;
	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
			    GUIDelete($hGUIhan)
				Opt("ExpandVarStrings", 0)
				ExitLoop
			Case $clipButton
				Local $aEdit = StringSplit(GUICtrlRead($inputEdit, $GUI_READ_DEFAULT), "|")
				ClipPut($aEdit[3] & " went to room " & $aEdit[4] & " and completed order #" & $aEdit[2] & " (" & $aEdit[5] & "), but did not complete ")
		EndSwitch
	WEnd
EndFunc

Func ChangeSnitchColumn()

	Local $sFilePath = @WorkingDir & "\Res\startMe.ini"

	If __VerifyIniIntegritySMG($sFilePath) Not Then Exit MsgBox("", "StartMeGUI", "Integrity Check Failed")

	$sEng = GetSCol("Engineer")
	$sStat = GetSCol("Status")
	$sTrade = GetSCol("Trade")
	$sRoom = GetSCol("Room")
	$sIssue = GetSCol("Issue")
	$sOrder = GetSCol("Order")

	; Create PopUp for Column Change ;
	$gSnitchPop = GUICreate("SnitchMe", 182, 259, 416, 272)
	$Label1 = GUICtrlCreateLabel("Engineer Col: ", 16, 16, 70, 17)
	$Label2 = GUICtrlCreateLabel("Status Col:", 16, 48, 55, 17)
	$Label3 = GUICtrlCreateLabel("Trade Col:", 16, 80, 53, 17)
	$Label4 = GUICtrlCreateLabel("Room Col", 16, 112, 50, 17)
	$Label5 = GUICtrlCreateLabel("Issue Col:", 16, 144, 50, 17)
	$Label6 = GUICtrlCreateLabel("Order Col:", 16, 176, 51, 17)
	$iEng = GUICtrlCreateInput($sEng, 104, 16, 49, 21, $ES_UPPERCASE)
	GUICtrlSetLimit(-1, 1, 1)
	$iStat = GUICtrlCreateInput($sStat, 104, 48, 49, 21, $ES_UPPERCASE)
	GUICtrlSetLimit(-1, 1, 1)
	$iTrade = GUICtrlCreateInput($sTrade, 104, 80, 49, 21, $ES_UPPERCASE)
	GUICtrlSetLimit(-1, 1, 1)
	$iRoom = GUICtrlCreateInput($sRoom, 104, 112, 49, 21, $ES_UPPERCASE)
	GUICtrlSetLimit(-1, 1, 1)
	$iIssue = GUICtrlCreateInput($sIssue, 104, 144, 49, 21, $ES_UPPERCASE)
	GUICtrlSetLimit(-1, 1, 1)
	$iOrder = GUICtrlCreateInput($sOrder, 104, 176, 49, 21, $ES_UPPERCASE)
	GUICtrlSetLimit(-1, 1, 1)
	$bChangePop = GUICtrlCreateButton("Change", 48, 216, 75, 25, $ES_UPPERCASE)
	GUISetState(@SW_SHOW, $gSnitchPop)
	;---------------;

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				GUIDelete($gSnitchPop)
				ExitLoop
			Case $bChangePop
				IniWrite($sFilePath, "SnitchMe - Columns", "Engineer", GUICtrlRead($iEng))
				IniWrite($sFilePath, "SnitchMe - Columns", "Status", GUICtrlRead($iStat))
				IniWrite($sFilePath, "SnitchMe - Columns", "Trade", GUICtrlRead($iTrade))
				IniWrite($sFilePath, "SnitchMe - Columns", "Room", GUICtrlRead($iRoom))
				IniWrite($sFilePath, "SnitchMe - Columns", "Issue", GUICtrlRead($iIssue))
				IniWrite($sFilePath, "SnitchMe - Columns", "Order", GUICtrlRead($iOrder))
				GUIDelete($gSnitchPop)
				ExitLoop
		EndSwitch
	WEnd

EndFunc

Func GetSCol($sColumn)

	Local $sFilePath = @WorkingDir &  "\Res\startMe.ini"
	If __VerifyIniIntegritySMG($sFilePath) Not Then Exit MsgBox("", "StartMeGUI", "Integrity Check Failed")

	Return IniRead($sFilePath, "SnitchMe - Columns", $sColumn, "A")

EndFunc

Func _ErrFunc($oError)
    MsgBox("", "",@ScriptName & " (" & $oError.scriptline & ") : ==> Local COM error handler - COM Error intercepted !" & @CRLF & _
            @TAB & "err.number is: " & @TAB & @TAB & "0x" & Hex($oError.number) & @CRLF & _
            @TAB & "err.windescription:" & @TAB & $oError.windescription & @CRLF & _
            @TAB & "err.description is: " & @TAB & $oError.description & @CRLF & _
            @TAB & "err.source is: " & @TAB & @TAB & $oError.source & @CRLF & _
            @TAB & "err.helpfile is: " & @TAB & $oError.helpfile & @CRLF & _
            @TAB & "err.helpcontext is: " & @TAB & $oError.helpcontext & @CRLF & _
            @TAB & "err.lastdllerror is: " & @TAB & $oError.lastdllerror & @CRLF & _
            @TAB & "err.scriptline is: " & @TAB & $oError.scriptline & @CRLF & _
            @TAB & "err.retcode is: " & @TAB & "0x" & Hex($oError.retcode) & @CRLF & @CRLF)
EndFunc   ;==>_ErrFunc