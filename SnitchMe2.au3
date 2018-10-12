; Script Name:  SnitchMe.au3
; Desc:     A work related script that
;			creates a detailed list of
;			work related search queries
; Author:       Gasca, K
; Created:      7 / 1 / 2017
; Version:      2.0

#include-once
#include <Excel.au3>
#include <File.au3>
#include <Array.au3>
#include <AutoItConstants.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <WindowsConstants.au3>

Func SnitchMe2()

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
	Local $oExcel = _Excel_Open()
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "SnitchMe", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	Local $sWork = @ScriptDir & "\orders\hotelOrders.csv"
	Local $oWork = _Excel_BookOpen($oExcel, $sWork)
	ProgressSet(0, 0 & "%", "Open: " & $sWork)
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "Snitch on ME", "Error opening '" & $sWork & "'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Excel_Close($oExcel)
		Exit
	EndIf
	;----------------;

	; Assign cell variables and Delete unqualified cells ;
	Local $aHotelRoom[0][4]
	Local $i = 1
	Local $sValue = Null

	Do
	
		$iRows = $oExcel.ActiveSheet.UsedRange.Rows.Count ; Continously update Row Count
		$HotSOS = _Excel_RangeRead($oWork, Default, "B" & $i)
		$Engineer = _Excel_RangeRead($oWork, Default, "I" & $i)
		$Issue = _Excel_RangeRead($oWork, Default, "F" & $i)
	
		$iCalcProg = ($i / $iRows) * 100 - 10 ; Perform Progess Calc only once per loop (Efficiency?)
		ProgressSet($iCalcProg, "Row: " & $i, "Checking Excel Rows") ; Update Progress
	
		If _Excel_RangeRead($oWork, Default, "C" & $i) <> "Closed" Then
			ProgressSet($iCalcProg, "Deleting Row: " & $i & " @ C" & $i)
			_Excel_RangeDelete($oWork.ActiveSheet, $oWork.ActiveSheet.rows($i), $xlShiftUp, 1)
		
		ElseIf _Excel_RangeRead($oWork, Default, "E" & $i) <> "FAC - Hotel Shop" Then
			ProgressSet($iCalcProg, "Deleting Row: " & $i & " @ E" & $i)
			_Excel_RangeDelete($oWork.ActiveSheet, $oWork.ActiveSheet.rows($i), $xlShiftUp, 1)
	
		ElseIf _Excel_RangeRead($oWork, Default, "I" & $i) = "" Then
			ProgressSet($iCalcProg, "Deleting Row: " & $i & " @ I" & $i)
			_Excel_RangeDelete($oWork.ActiveSheet, $oWork.ActiveSheet.rows($i), $xlShiftUp, 1)

		Else
			$aGValue = StringSplit(_Excel_RangeRead($oWork, Default, "G" & $i), " - ", $STR_ENTIRESPLIT)
			; Only Add to Array if verified number in String $aGValue (ROOM Cell)
			For $b = 1 To $aGValue[0]
			
				If Number($aGValue[$b]) <> 0 Then
					_ArrayAdd($aHotelRoom, $HotSOS & '|' & $Engineer & '|' & $aGValue[$b] & '|' & $Issue, 0)
					$i += 1 ; Increment row counter
					ExitLoop
				ElseIf $b = $aGValue[0] Then ; (Room Cell) has no valid room number
					ProgressSet($iCalcProg, "Deleting Row: " & $i & " @ G" & $i)
					_Excel_RangeDelete($oWork.ActiveSheet, $oWork.ActiveSheet.rows($i), $xlShiftUp, 1)
				EndIf
			Next
		
		EndIf

	Until $i = $iRows + 1

	; Array Sorting ;
	ProgressSet(90, "90%", "Sorting Array..")
	_ArraySort($aHotelRoom, 0,0,0,2)
	$holdValue = 0
	$temp = 0

	; Verify total room # count ;
	For $e = 0 To UBound($aHotelRoom, $UBOUND_ROWS) - 1
		ProgressSet(90 + ($e / UBound($aHotelRoom,$UBOUND_ROWS)), "Adding row " & $e & " to ListView", "Compiling List")
		GUICtrlCreateListViewItem($e + 1 & "|" & $aHotelRoom[$e][0] & "|" & $aHotelRoom[$e][1] & "|" & $aHotelRoom[$e][2] & "|" & $aHotelRoom[$e][3], $listView)
		
		If $temp <> $aHotelRoom[$e][2] Then
			$holdValue += 1
		EndIf
	$temp = $aHotelRoom[$e][2]
	Next

	ProgressOff()
	$roomLabel = GUICtrlCreateLabel("Total of: " & $holdValue & " different rooms." , 200, 575, 150, 25)
	GUISetState(@SW_SHOW, $hGUIhan)

	; Main Loop controller ;
	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				ExitLoop
			Case $clipButton
				Local $aEdit = StringSplit(GUICtrlRead($inputEdit, $GUI_READ_DEFAULT), "|")
				ClipPut($aEdit[3] & " went to room " & $aEdit[4] & " and completed order #" & $aEdit[2] & " (" & $aEdit[5] & "), but did not complete ")
		EndSwitch
	WEnd
EndFunc

Func Debug($input)
   MsgBox($MB_SYSTEMMODAL, "StartMe GUI: DEBUG MODE", $input)
EndFunc