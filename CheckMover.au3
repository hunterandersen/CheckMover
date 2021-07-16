;----------------------------------------Press F5 to RUN---------------------------
#include <File.au3>
#include <Misc.au3>
#include <Constants.au3>
#include <MsgBoxConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>

;-----------------------The Current FDrive Drive Letter-------------------
$FDrive = "F:\DIETARY\Dietary Financials\Vendor Check Copies\"
;-----------------------The Current FDrive Drive Letter-------------------

$boxNum = 0
$valid = 0
$perfect = 1
$renameMe = 0
$vendorName = ""
$pdfReader = "[CLASS:classFoxitPhantom]" ; "[CLASS:AcrobatSDIWindow]"

MsgBox(0, "Reminder", "Make sure Vendor Check Copies and Separated PDFs Good is open")
If Not(WinExists("Vendor Check Copies")) Then
	$iPid = RunWait("C:\Windows\Explorer.exe /n,/e," & $FDrive)
	WinWait("Vendor Check Copies", "", 10)
	WinMove("Vendor Check Copies", "", 700, 200)
	Send("#{Right}")
EndIf

If Not(WinExists("Vendor Check Copies") Or WinExists("Vendor Checks") Or WinExists("Renamed")) Then
	MsgBox(0, "Doh", "One of the folders isn't open")
	Exit
EndIf

$folderPath = InputBox("File Location", "Enter the location to pull pdfs from", "C:\Users\anderseh1\Desktop\Separated PDFs Good\")
$error = @error
If $folderPath == "" Or $folderPath == " " Or Not (IsString($folderPath)) Then
	$folderPath = "C:\Users\anderseh1\Desktop\Separated PDFs Good\"
EndIf
If $error <> 0 Then
	_BoxClosing()
EndIf

$fileArray = _FileListToArray($folderPath, "*.pdf", 1)
If Not (IsArray($fileArray)) Or $fileArray[0] == 0 Then
	MsgBox(0, "Done", "No pdfs found in that folder location - Program will exit")
	Exit
 EndIf

For $i = 1 To $fileArray[0] Step 1
	$title = $fileArray[$i]

	Local $folderName = ""
	$aSlashes = StringSplit($folderPath, "\")
	If $aSlashes[$aSlashes[0]] <> "" Then
		$folderName = $aSlashes[$aSlashes[0]]
	Else
		$folderName = $aSlashes[$aSlashes[0] - 1]
	EndIf

	If WinExists($pdfReader) Then WinClose($pdfReader)
	While WinExists($folderName) <> True And $folderName <> "Desktop"
		MsgBox(0, "Folder not open", "Please open the folder named: " & $folderName)
	WEnd
	;============== Open the pdf file ===========================
	$hWin = WinActivate($folderName)
	$sFileToSelect = ''
	$oSHFolderView = _ObjectSHFolderViewFromWin($hWin)
	If Not @error Then
		$sFileToSelect = $title
		WinActivate($hWin)
		; Select ONLY the item
		$oSHFolderView.SelectItem($oSHFolderView.Folder.ParseName($sFileToSelect), 4 + 1 + 8 + 16)
	Else
		MsgBox(0, "FolderView Error", "Dat's no bueno")
		_BoxClosing()
	EndIf
	Send("{Enter}")
	WinWaitActive($pdfReader)
	Send("#{Left}")
	Sleep(300)
	Send("{Escape}")
	Sleep(100)
	;============== END OPEN PDF ============================
	;//Start Renaming Process
	$newName = "Generic Title"
	$suggestedVendor = "I have no idea"


;And StringInStr($title, "File"))
	If ((StringIsAlNum(StringLeft($title, 16)) Or StringInStr($title, "_")) Or $renameMe = 1) Then
		$needsName = 1
		While ($needsName = 1)
			WinActivate($pdfReader)
			$guiH = GUICreate("New Title", 280, 170, 1000, 600, $GUI_SS_DEFAULT_GUI, $WS_EX_TOPMOST)
			$guiEdit = GUICtrlCreateLabel("Enter new name for file", 10, 15)
			$guiInput = GUICtrlCreateInput("", 10, 110, 255, 20)
			$guiOK = GUICtrlCreateButton("OK", 37, 135, 80, 23, $BS_DEFPUSHBUTTON)
			GUISetState(@SW_SHOW, $guiH)
			Sleep(100)
			GUICtrlSetState($guiInput, $GUI_FOCUS)
			While 1
				Switch GUIGetMsg()
					Case $GUI_EVENT_CLOSE
						$perfect = 0
						GUIDelete($guiH)
						_BoxClosing()
					Case $guiOK
						$renameMe = 0
						$needsName = 0
						$newName = GUICtrlRead($guiInput)
						If StringInStr($newName, "0") = 1 Then
							$pressedYes = MsgBox($MB_YESNO, "Trim Leading Zeros?", $newName & "-->" & Number($newName))
							If $pressedYes = $IDYES Then
								$newName = Number($newName)
							EndIf
						Else
							If StringInStr($newName, "ACH") = 1 Then

								If Not(_NumDashes($newName) = 4) Then
									MsgBox(0, "New ACH RegExp", "There aren't 4 dashes. Double check the file name. Program will continue WITHOUT FIXING IT.")
								EndIf

								;$pressedYes = MsgBox($MB_YESNO, "Delete space after ACH?", $newName)
								;If $pressedYes = $IDYES Then
								;	$newName = StringReplace($newName, " ", "")
								;EndIf
							EndIf
						EndIf
						GUIDelete($guiH)
						ExitLoop
				EndSwitch
			WEnd
			If $newName == "" Then
				$error = MsgBox(0, "Empty Title", "You gotta give me somethin")
				$needsName = 1
			EndIf
			If StringInStr($newName, "\") Or StringInStr($newName, "/") Or StringInStr($newName, ".") Or StringInStr($newName, ":") Then
				$needsName = 1
				$error = MsgBox(0, "Special Characters", "Friends don't let friends use \ or / or . or : in their file names")
			EndIf
		WEnd
	EndIf
	;//End Renaming Process
	$valid = 0
	$startLoc = $folderPath & $title
	$endLocation = ""
	$secondVendor = 0
	While $valid = 0 Or $valid = 2 Or $valid = 3
		WinActivate($pdfReader)
		$guiH = GUICreate("Company Name", 280, 170, -1, -1, $GUI_SS_DEFAULT_GUI, $WS_EX_TOPMOST)
		$guiEdit = GUICtrlCreateLabel("Type the name of the folder to place the check in", 10, 15)
		If $secondVendor = 2 Then
			GUICtrlSetData($guiEdit, "Type the second folder to put the check in")
		EndIf
		If $valid = 3 Then
			If _VendorGuess($vendorName, "Carlisle Food Services") > 0 Then
				$suggestedVendor = "Dinex"
			ElseIf (_VendorGuess($vendorName, "HC Brill") > 0) Or (_VendorGuess($vendorName, "h. c. brill") > 0) Then
				$suggestedVendor = "CSM Bakeries"
			ElseIf _VendorGuess($vendorName, "Aryzta") > 0 Then
				$suggestedVendor = "Otis Spunkmeyer"
			ElseIf _VendorGuess($vendorName, "Gordon") > 0 Then
				$suggestedVendor = "GFS, Gordon Food Service"
			ElseIf _VendorGuess($vendorName, "Cafe Puree") > 0 Then
				$suggestedVendor = "Medtrition"
			ElseIf _VendorGuess($vendorName, "Muller Pinehurst") > 0 Then
				$suggestedVendor = "Prairie Farms"
			ElseIf _VendorGuess($vendorName, "Sysco Hawaii") > 0 Then
				$suggestedVendor = "HFM"
			ElseIf _VendorGuess($vendorName, "Ralcorp") > 0 Then
				$suggestedVendor = "Conagra Carriage House"
			ElseIf _VendorGuess($vendorName, "Ralston") > 0 Then
				$suggestedVendor = "Conagra Ralston"
			ElseIf _VendorGuess($vendorName, "Burry") > 0 Then
				$suggestedVendor = "Merisant Burry"
			ElseIf _VendorGuess($vendorName, "Ottenberg") > 0 Or (_VendorGuess($vendorName, "quality baker") > 0) Then
				$suggestedVendor = "H&S Bakery Ottenberg"
			ElseIf _VendorGuess($vendorName, "Square H") > 0 Then
				$suggestedVendor = "Hoffy?"
			ElseIf _VendorGuess($vendorName, "Pinnacle") > 0 Then
				$suggestedVendor = "Conagra Pinnacle"
			ElseIf _VendorGuess($vendorName, "ifp") > 0 Then
				$suggestedVendor = "Leahy IFP"
			ElseIf _VendorGuess($vendorName, "Oneida") > 0 Then
				$suggestedVendor = "Everywhere Global"
			ElseIf _VendorGuess($vendorName, "Continental Mills") > 0 Then
				$suggestedVendor = "Krusteaz Professional"
			ElseIf _VendorGuess($vendorName, "Dart") > 0 Then
				$suggestedVendor = "looking for Solo Cup on the check"
			ElseIf _VendorGuess($vendorName, "American Foods Group") > 0 Then
				$suggestedVendor = "looking for King's Command on the check"
			ElseIf _VendorGuess($vendorName, "Heartland") > 0 Then
				$suggestedVendor = "Splenda"
			ElseIf _VendorGuess($vendorName, "Johnson & Johnson") > 0 Then
				$suggestedVendor = "looking for Splenda on the check"
			EndIf
			GUICtrlSetData($guiEdit, "Need help? Try " & $suggestedVendor)
		EndIf
		$guiInput = GUICtrlCreateInput("", 10, 110, 264, 20)
		$guiOK = GUICtrlCreateButton("OK", 10, 135, 80, 23, $BS_DEFPUSHBUTTON)
		$guiSkip = GUICtrlCreateButton("Skip", 103, 135, 80, 23)
		$guiRename = GUICtrlCreateButton("Force Rename", 196, 135, 80, 23)
		GUISetState(@SW_SHOW, $guiH)
		GUICtrlSetState($guiInput, $GUI_FOCUS)
		While 1
			Switch GUIGetMsg()
				Case $GUI_EVENT_CLOSE
					$perfect = 0
					GUIDelete($guiH)
					_BoxClosing()
				Case $guiSkip
					$perfect = 0
					GUIDelete($guiH)
					If WinExists($pdfReader) Then WinClose($pdfReader)
					Sleep(100)
					ContinueLoop(3)
				Case $guiOK
					$vendorName = GUICtrlRead($guiInput)
					GUIDelete($guiH)
					If Not(WinExists("Vendor Check Copies")) Then
						$wPID = Run("explorer.exe " & $FDrive)
						Sleep(1000)
					EndIf
					$hWin = WinActivate("Vendor Check Copies")
					$oSHFolderView = _ObjectSHFolderViewFromWin($hWin)
					Send($vendorName)
					Sleep(700)
					If Not($secondVendor = 2) Then
						$vendorFile = $oSHFolderView.SelectedItems.Item(0).Path
					Else
						$vendorFile2 = $oSHFolderView.SelectedItems.Item(0).Path
					EndIf
					$vendorFolder = $oSHFolderView.SelectedItems.Item(0).Name
					ExitLoop
				Case $guiRename
					$renameMe = 1
					$i = $i - 1
					If WinExists($pdfReader) Then WinClose($pdfReader)
					GUIDelete($guiH)
					ContinueLoop(3)
			EndSwitch
		WEnd
		$buttonWords = ""
		If $secondVendor = 2 Then
			$buttonWords = "1. Move Files"
		Else
			$buttonWords = "1. Move File"
		EndIf
		$guiH = GUICreate("Move File", 300, 170, -1, -1, $GUI_SS_DEFAULT_GUI, $WS_EX_TOPMOST)
		$guiEdit = GUICtrlCreateLabel("Move to this folder?" & $vendorFolder, 10, 15)
		$guiOK = GUICtrlCreateButton($buttonWords, 20, 105, 100, 23, $BS_DEFPUSHBUTTON)
		$guiDiff = GUICtrlCreateButton("2. Different Vendor", 60, 135, 100, 23)
		$guiSkip = GUICtrlCreateButton("3. Skip File", 150, 105, 100, 23)
		If Not($secondVendor = 2) Then
			$guiDouble = GUICtrlCreateButton("4. Double Vendor", 190, 135, 100, 23)
		EndIf
		GUISetState(@SW_SHOW, $guiH)
		While 1
			If Not(WinExists($guiH)) Then
				ExitLoop
			EndIf
			If _IsPressed(31) Or _IsPressed(61) Then ;1 on keyboard
				GUICtrlSetState($guiOK, $GUI_FOCUS)
			ElseIf _IsPressed(32) Or _IsPressed(62) Then ;2 on keyboard
				GUICtrlSetState($guiDiff, $GUI_FOCUS)
			ElseIf _IsPressed(33) Or _IsPressed(63) Then ;3 on keyboard
				GUICtrlSetState($guiSkip, $GUI_FOCUS)
			ElseIf _IsPressed(34) Or _IsPressed(64) Then ;4 on keyboard
				GUICtrlSetState($guiDouble, $GUI_FOCUS)
			EndIf
			Switch GUIGetMsg()
				Case $GUI_EVENT_CLOSE
					$perfect = 0
					GUIDelete($guiH)
					_BoxClosing()
				Case $guiSkip
					$perfect = 0
					GUIDelete($guiH)
					If WinExists($pdfReader) Then WinClose($pdfReader)
					Sleep(100)
					ContinueLoop(3)
				Case $guiDiff
					GUIDelete($guiH)
					$valid = 3
				Case $guiDouble
					;Have it repeat
					$secondVendor = 2
					$valid = 2
					GUIDelete($guiH)
					ExitLoop
				Case $guiOK
					$valid = 1
					GUIDelete($guiH)
					ExitLoop
			EndSwitch
		WEnd
	WEnd
	; --------------------MOVE FILE TO CORRECT VENDOR FOLDER------------------------------

	If Not($newName = "Generic Title") Then
		$title = $newName
	EndIf
	$endLocation = $vendorFile & "\" & $title & ".pdf"

	MsgBox(0, "start, end" , $startLoc & "    |    " & $endLocation)


	If FileExists($endLocation) Then
		MsgBox(0, "Two Files, One Name", "That file already exists in the folder. Double check the " & $vendorFile & " files")
	Else ;Only move the file if it doesn't already exist at the end location.
		While FileExists($endLocation) = 0
			$fileMoved = FileMove($startLoc, $endLocation, $FC_NOOVERWRITE); Move the file WITHOUT overwriting existing
			Sleep(200)
		WEnd
		If $secondVendor = 2 Then
			While FileExists($vendorFile2 & "\" & $title & ".pdf") = 0
				$fileMoved = FileCopy($endLocation, $vendorFile2, $FC_NOOVERWRITE); Copy the file WITHOUT overwriting existing
				Sleep(200)
			WEnd
		EndIf

		While WinExists($pdfReader)
			WinClose($pdfReader)
			Sleep(300)
		WEnd

		$finalDest = "C:\Users\anderseh1\Desktop\Separated PDFs Good\Finished Products\Vendor Checks\" & $title & ".pdf"
		$fileMoved = 0
		While $fileMoved = 0
			$fileMoved = FileMove($startLoc, $finalDest, 1); Move the file and overwrite existing
		WEnd
	EndIf

	While WinExists($pdfReader)
		WinClose($pdfReader)
		Sleep(300)
	WEnd

   MsgBox(0, "Double check", "your files, dog")

	;Waits for file to be moved
	;FIX
	;$objFolderItem = Null
	;$objFolderItem = $oSHFolderView.Folder.ParseName($sFileToSelect)

	;MsgBox(0, "objFolderitem", $objFolderItem)

	;While (StringInStr($objFolderItem.Path, "Renamed") <> 0)
	;	Sleep(200)
	;	$objFolderItem = $oSHFolderView.Folder.ParseName($sFileToSelect)
	;WEnd
   ;END FIX

Next

If $perfect = 0 Then
	MsgBox(0, "Done", "Don't forget the files it didn't move")
Else
	MsgBox(0, "Done", "Finished without a hitch")
EndIf


;---------------------User hits cancel or the X on a box----------------
Func _BoxClosing()
	If $perfect = 0 Then
		MsgBox($MB_ICONWARNING, "Done", "You ended the program early - Don't forget the files it didn't move")
	Else
		MsgBox($MB_ICONWARNING, "Done", "You ended the program early")
	EndIf
	Exit
EndFunc   ;==>_BoxClosing

Func _VendorGuess($typed, $guess)
	$tryMe = 0
	If StringInStr($typed, $guess, 0) > 0 Then
		$tryMe = 1
	ElseIf StringInStr($guess, $typed, 0) > 0 Then
		$tryMe = 2
	EndIf

	If $tryMe = 0 Then
		$newTry = StringReplace($typed, "'", "")
		$newTry = StringReplace($typed, ".", "")
		If StringInStr($newTry, $guess, 0) > 0 Then
			$tryMe = 3
		ElseIf StringInStr($guess, $newTry, 0) > 0 Then
			$tryMe = 3
		EndIf
	EndIf

	Return $tryMe
EndFunc

Func _NumDashes($tempString)
	$num = 0
	$positionOfDash = StringInStr($tempString, "-")
	While $positionOfDash > 0
		$num = $num + 1
		$positionOfDash = StringInStr($tempString, "-", 0, $num+1)
	WEnd

	Return $num
EndFunc

; ==========================================================================================================================

; Func _ObjectSHFolderViewFromWin($hWnd)
;
; Returns an 'ShellFolderView' Object for the given Window handle
;
; Author: Ascend4nt, based on code by KaFu, klaus.s
; ==========================================================================================================================

Func _ObjectSHFolderViewFromWin($hWnd)
	If Not IsHWnd($hWnd) Then Return SetError(1, 0, 0)
	Local $oShell, $oShellWindows, $oIEObject, $oSHFolderView

	; Shell Object
	$oShell = ObjCreate("Shell.Application")
	If Not IsObj($oShell) Then Return SetError(2, 0, 0)

	;   Get a 'ShellWindows Collection' object
	$oShellWindows = $oShell.Windows()
	If Not IsObj($oShellWindows) Then Return SetError(3, 0, 0)

	;   Iterate through the collection - each of type 'InternetExplorer' Object

	For $oIEObject In $oShellWindows
		If $oIEObject.HWND = $hWnd Then
			; InternetExplorer->Document = ShellFolderView object
			$oSHFolderView = $oIEObject.Document
			If IsObj($oSHFolderView) Then Return $oSHFolderView
			Return SetError(4, 0, 0)
		EndIf
	Next

	Return SetError(-1, 0, 0)
EndFunc   ;==>_ObjectSHFolderViewFromWin

; ==========================================================================================================================
; Func _ExplorerWinGetSelectedItems($hWnd)
;
;
; Author: klaus.s, KaFu, Ascend4nt (consolidation & cleanup, Path name simplification)
; ==========================================================================================================================

Func _ExplorerWinGetSelectedItems($hWnd)
	If Not IsHWnd($hWnd) Then Return SetError(1, 0, '')
	Local $oSHFolderView
	Local $iSelectedItems, $iCounter = 2, $aSelectedItems[2] = [0, ""]

	$oSHFolderView = _ObjectSHFolderViewFromWin($hWnd)
	If @error Then Return SetError(@error, 0, '')

	;   SelectedItems = FolderItems Collection object->Count
	$iSelectedItems = $oSHFolderView.SelectedItems.Count

	Dim $aSelectedItems[$iSelectedItems + 2] ; 2 extra -> 1 for count [0], 1 for Folder path [1]

	$aSelectedItems[0] = $iSelectedItems
	;   ShellFolderView->Folder->Self as 'FolderItem'->Path
	$aSelectedItems[1] = $oSHFolderView.Folder.Self.Path

	;   ShellFolderView->SelectedItems = FolderItems Collection object
	$oSelectedFolderItems = $oSHFolderView.SelectedItems

	#cs
		; For ALL items in an Explorer window (not just the selected ones):
		$oSelectedFolderItems = $oSHFolderView.Folder.Items
		ReDim $aSelectedItems[$oSelectedFolderItems.Count+2]
	#ce

	For $oFolderItem In $oSelectedFolderItems
		$aSelectedItems[$iCounter] = $oFolderItem.Path
		$iCounter += 1
	Next

	Return SetExtended($iCounter - 2, $aSelectedItems)
EndFunc   ;==>_ExplorerWinGetSelectedItems

; ==========================================================================================================================
; ==========================================================================================================================
