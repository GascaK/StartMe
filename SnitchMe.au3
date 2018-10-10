; Snitch Report Automation V 0.00 ;
; Unreleased. NO Brandon cant have this ;

#include <Excel.au3>
#include <File.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <WindowsConstants.au3>


; Create the Report Window ;
GUICreate("StartMeGUI - SnitchMe", 500, 700, -1, $WS_EX_ACCEPTFILES)
Local $listView = GUICtrlCreateListView("Row # | HotSOS | Engineer | Room | Issue ", -1, -1, 500, 550)
$inputEdit = GUICtrlCreateInput("", 150, 600, 335, 25)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
$dndLabel = GUICtrlCreateLabel("Drag and Drop Row Here", 8, 600, 125, 17)
$clipButton = GUICtrlCreateButton("Copy to Clipboard", 152, 650, 163, 25)
;--------------;

Local $oExcel = _Excel_Open()
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookOpen Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

Local $sWork = @ScriptDir & "\orders\hotelOrders.csv"
Local $oWork = _Excel_BookOpen($oExcel, $sWork)

If @error Then
   MsgBox($MB_SYSTEMMODAL, "Snitch on ME", "Error opening '" & $sWork & "'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
   _Excel_Close($oExcel)
   Exit
EndIf

Local $iRows = $oExcel.ActiveSheet.UsedRange.Rows.Count
Local $i = 1
Local $sValue = Null

; TODO: Change to For loop.. why did i do this? ;
Do

   $aGValue = StringSplit(_Excel_RangeRead($oWork, Default, "G" & $i), " - ", $STR_ENTIRESPLIT)
   Local $bDel = True

   For $b = 1 To $aGValue[0]
	  If Number($aGValue[$b]) <> 0 Then
		 $bDel = False
	  EndIf
   Next

   If _Excel_RangeRead($oWork, Default, "C" & $i) <> "Closed" Then
	  _Excel_RangeDelete($oWork.ActiveSheet, $oWork.ActiveSheet.rows($i), $xlShiftUp, 1)

   ElseIf _Excel_RangeRead($oWork, Default, "E" & $i) <> "FAC - Hotel Shop" Then
	  _Excel_RangeDelete($oWork.ActiveSheet, $oWork.ActiveSheet.rows($i), $xlShiftUp, 1)

   ElseIf $bDel Then
	  _Excel_RangeDelete($oWork.ActiveSheet, $oWork.ActiveSheet.rows($i), $xlShiftUp, 1)

   ElseIf _Excel_RangeRead($oWork, Default, "I" & $i) = "" Then
	  _Excel_RangeDelete($oWork.ActiveSheet, $oWork.ActiveSheet.rows($i), $xlShiftUp, 1)

   Else
	  $i += 1
   EndIf

   $iRows = $oExcel.ActiveSheet.UsedRange.Rows.Count

Until $i = $iRows + 1

Local $aHotelRoom[0][4]

For $c = 1 To $oExcel.ActiveSheet.UsedRange.Rows.Count

   $aNewRowValue = StringSplit(_Excel_RangeRead($oWork, Default, "G" & $c), " - ", $STR_ENTIRESPLIT)
   $HotSOS = _Excel_RangeRead($oWork, Default, "B" & $c)
   $Engineer = _Excel_RangeRead($oWork, Default, "I" & $c)
   $Issue = _Excel_RangeRead($oWork, Default, "F" & $c)

   For $d = 1 To $aNewRowValue[0]
	  If Number($aNewRowValue[$d]) <> 0 Then
		 _ArrayAdd($aHotelRoom, $HotSOS & '|' & $Engineer & '|' & $aNewRowValue[$d] & '|' & $Issue, 0)
		 ExitLoop
	  EndIf
   Next

Next

Local $iCol = UBound($aHotelRoom, $UBOUND_ROWS)

_ArraySort($aHotelRoom, 0,0,0,2)

$holdValue = 0
$temp = 0

For $e = 0 To UBound($aHotelRoom, $UBOUND_ROWS) - 1

   GUICtrlCreateListViewItem($e + 1 & "|" & $aHotelRoom[$e][0] & "|" & $aHotelRoom[$e][1] & "|" & $aHotelRoom[$e][2] & "|" & $aHotelRoom[$e][3], $listView)
   If $temp <> $aHotelRoom[$e][2] Then
	  $holdValue += 1
   EndIf
   $temp = $aHotelRoom[$e][2]
Next

$roomLabel = GUICtrlCreateLabel("Total of: " & $holdValue & " different rooms." , 200, 575, 100, 25)
GUISetState(@SW_SHOW)

While 1
   Switch GUIGetMsg()
	  Case $GUI_EVENT_CLOSE
		 ExitLoop
	  Case $clipButton
		 Local $aEdit = StringSplit(GUICtrlRead($inputEdit, $GUI_READ_DEFAULT), "|")
		 ClipPut($aEdit[3] & " went to room " & $aEdit[4] & " and completed order #" & $aEdit[2] & " (" & $aEdit[5] & "), but did not complete ")
   EndSwitch
WEnd
