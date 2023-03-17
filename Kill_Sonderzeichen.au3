#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Icons\ascii.ico
#AutoIt3Wrapper_Res_Comment=Replaces all characters with accents or non ascii. Only latin characters and numbers will stay.
#AutoIt3Wrapper_Res_Description=Replaces all characters with accents or non ascii. Only latin characters and numbers will stay.
#AutoIt3Wrapper_Res_Fileversion=1.0.0.42
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_CompanyName=Norddeutscher Rundfunk
#AutoIt3Wrapper_Res_LegalCopyright=Conrad Zelck
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Res_Language=1031
#AutoIt3Wrapper_Res_Field=Copyright|Conrad Zelck
#AutoIt3Wrapper_Res_Field=Compile Date|%date% %time%
#AutoIt3Wrapper_AU3Check_Parameters=-q -d -w 1 -w 2 -w 3 -w- 4 -w 5 -w 6 -w- 7
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/mo
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <File.au3>
#include <ProgressConstants.au3>
#include <SendMessage.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <Array.au3>
#include <Date.au3>
#include <TrayCox.au3> ; source: https://github.com/SimpelMe/TrayCox

Global Const $IS_FOLDER = True
Global $g_aDropFiles[1]
Local $aFilesAndFolders
Global $iSuccess
Global $g_bDryRun = False ; <== for debugging only
Global $g_sListOld
Global $g_sListNew
Global $g_sLogText
Global $g_bWriteToLog = False

; if parameter given via sendto or drag&drop onto AppIcon
ConsoleWrite("$CmdLineRaw: " & $CmdLineRaw & @CRLF)
If @Compiled Then
	If $CmdLineRaw <> "" Then
		$iSuccess = MsgBox($MB_TOPMOST + $MB_YESNOCANCEL, "ACHTUNG", "Es gibt kein UNDO." & @CRLF & "Bist Du sicher, dass Du dies umbenennen möchtest?" & @CRLF & @CRLF & $CmdLineRaw)
		If $iSuccess = $IDYES Then
			$aFilesAndFolders = $CmdLine
			ConsoleWrite("RENAMING ALL" & @CRLF)
			_RenamingAll($aFilesAndFolders)
			If $g_bWriteToLog Then FileWrite(@ScriptDir & "\Logfile.txt", _NowCalc() & @CRLF & $g_sLogText & @CRLF)
		EndIf
		Exit
	EndIf
EndIf

; if no parameters are give open a drag and drop gui
GUICreate("Kill_Chars", 300, 100, -1, -1, -1, $WS_EX_ACCEPTFILES)
GUICtrlCreateLabel("Drag&&Drop your files and folders here.", 20, 40, 260, 20, $SS_CENTER)
GUICtrlSetFont(-1, 10)
Local $FILES_DROPPED = GUICtrlCreateDummy()
GUISetState()
GUIRegisterMsg($WM_DROPFILES, 'WM_DROPFILES_FUNC')

While True
    Switch GUIGetMsg()
        Case $GUI_EVENT_CLOSE
            ExitLoop
        Case $FILES_DROPPED
            $aFilesAndFolders = $g_aDropFiles
			$iSuccess = MsgBox($MB_TOPMOST + $MB_YESNOCANCEL, "ACHTUNG", "Es gibt kein UNDO." & @CRLF & "Bist Du sicher, dass Du dies umbenennen möchtest?" & @CRLF & @CRLF & _ArrayToString($aFilesAndFolders, @CR, 1))
			If $iSuccess = $IDYES Then
				_RenamingAll($aFilesAndFolders)
				If $g_bWriteToLog Then FileWrite(@ScriptDir & "\Logfile.txt", _NowCalc() & @CRLF & $g_sLogText & @CRLF)
			EndIf
			If $g_bDryRun Then
				ConsoleWrite("! ----------------------------------------------------------------------" & @CRLF & $g_sListOld & @CRLF)
				ConsoleWrite("+ ----------------------------------------------------------------------" & @CRLF & $g_sListNew & @CRLF)
			EndIf
    EndSwitch
WEnd

Exit

#Region Funcs
Func _RenamingAll($aFilesAndFolders)
	; kind of progress bar
	Local $hGUI = GUICreate("running ...", 350, 40)
	GUICtrlCreateProgress(10, 10, 330, 20, $PBS_MARQUEE)
	_SendMessage(GUICtrlGetHandle(-1), $PBM_SETMARQUEE, True, 50) ; final parameter is update time in ms
	GUISetState()

	Local $aFiles, $aFolders
	$g_sLogText = ""
	$g_bWriteToLog = False
	For $i = 1 To $aFilesAndFolders[0] ; go through all given files and/or directories
		If FileGetAttrib($aFilesAndFolders[$i]) = "D" Then ; is a directory
			$aFiles = _RecFileListToArray($aFilesAndFolders[$i], "*", 1, 1, 1, 2) ; list files only in that directory
			If Not @error Then
				If $aFiles[0] > 0 Then ; go through all files in that directory
					For $ii = 1 To $aFiles[0]
						_Rename($aFiles[$ii]) ; rename current file
					Next
				EndIf
			EndIf
			$aFolders = _RecFileListToArray($aFilesAndFolders[$i], "*", 2, 1, 1, 2) ; list folders only in that directory
			If Not @error Then
				If $aFolders[0] > 0 Then ; go through all folders in that directory
					For $iii = $aFolders[0] To 1 Step - 1
						_Rename($aFolders[$iii], $IS_FOLDER) ; rename current folder
					Next
				EndIf
			EndIf
			_Rename($aFilesAndFolders[$i], $IS_FOLDER) ; rename that directory folder
		Else ; is only one file
			_Rename($aFilesAndFolders[$i]) ; rename that file
		EndIf
	Next
	GUIDelete($hGUI)
EndFunc

Func _Rename($sFile, $bFolder = False)
	Local $sDrive, $sDir, $sFName, $sExt
	Local $sFileNameOld, $sFileNameNew, $sFileOld, $sFileNew
	Local $aSplit = _PathSplit($sFile, $sDrive, $sDir, $sFName, $sExt)
	$sFileNameOld = $aSplit[3]
	$sFileNameNew = _StringReplaceAccent($sFileNameOld) ; replaces all characters with accent to pure latin characters
	$sFileNameNew = _StringReplaceNonAscii($sFileNameNew) ; replaces all non ascii characters with an underscore
	$sFileNameNew = _StringReplaceDoubleUnderline($sFileNameNew) ; replaces double underscores with just one
	; Don't remove characters if filename is very long
;~ 	$sFileNameNew = StringTrimRight($sFileNameNew, StringLen($sFileNameNew) - 56) ; max. 56 characters because of Sony Professional Disc
;~ 	$sFileNameNew = _StringReplaceDoubleUnderline($sFileNameNew) ; delete trailing underscore if truncating produces one
	$sFileOld = $aSplit[0]
	$sFileNew = _PathMake($aSplit[1], $aSplit[2], $sFileNameNew, $aSplit[4])
	If StringToBinary($sFileOld, $SB_UTF8) <> StringToBinary($sFileNew, $SB_UTF8) Then
		ConsoleWrite("Different" & @CRLF)
		While FileExists($sFileNew)
			$sFileNew = _PathMake($aSplit[1], $aSplit[2], $sFileNameNew & "_" & Random(10000, 99999, 1), $aSplit[4]) ; rename with a random extension
		WEnd
		If $g_bDryRun Then
			ConsoleWrite($sFileOld & " ==> " & $sFileNew & @CRLF)
			$g_sListOld &= $sFileOld & @CRLF
			$g_sListNew &= $sFileNew & @CRLF
		Else
			If $bFolder Then
				$iSuccess = DirMove($sFileOld, $sFileNew) ; rename directory
				ConsoleWrite("Success DirMove: " & $iSuccess & @CRLF)
				$g_sLogText &= $sFileOld & " ==> " & $sFileNew & @CRLF
				$g_bWriteToLog = True
			Else
				$iSuccess = FileMove($sFileOld, $sFileNew) ; rename file
				ConsoleWrite("Success FileMove: " & $iSuccess & @CRLF)
				$g_sLogText &= $sFileOld & " ==> " & $sFileNew & @CRLF
				$g_bWriteToLog = True
			EndIf
		EndIf
	EndIf
EndFunc

Func _StringReplaceAccent($sString) ; replaces all characters with accent to pure latin characters
    Local $exp, $rep
    Local $pattern[29][2] = [ _
            ["[ÀÁÂÃÅÆ]", "A"],["[àáâãå]", "a"],["Ä", "Ae"],["[æä]", "ae"], _
            ["Þ", "B"],["þ", "b"], _
            ["Ç", "C"],["ç", "c"], _
            ["[ÈÉÊË]", "E"],["[èéêë]", "e"], _
            ["[ÌÍÎÏ]", "I"],["[ìíîï]", "i"], _
            ["Ñ", "N"],["ñ", "n"], _
            ["[ÒÓÔÕØ]", "O"],["[ðòóôõø]", "o"],["Ö", "Oe"],["ö", "oe"], _
            ["[Š]", "S"],["[š]", "s"], _
            ["ß", "ss"], _
            ["[ÙÚÛ]", "U"],["[ùúû]", "u"],["Ü", "Ue"],["ü", "ue"], _
            ["Ý", "Y"],["[ýýÿ]", "y"], _
            ["Ž", "Z"],["ž", "z"]]

    For $i = 0 To (UBound($pattern) - 1)
        $exp = $pattern[$i][0]
        If $exp = "" Then ContinueLoop
        $rep = $pattern[$i][1]
        $sString = StringRegExpReplace($sString, $exp, $rep)
        If @error == 0 And @extended > 0 Then
            ConsoleWrite("Accend: " & $sString & @LF & "--> " & $exp & @LF)
        EndIf
    Next
    Return $sString
EndFunc   ;==>_StringReplaceAccent

Func _StringReplaceNonAscii($sString) ; replaces all non ascii characters with an underscore
;~ 	$sString = StringRegExpReplace ( $sString, "[\x00-\x2D\x2F\x3A-\x40\x5B-\x60\x7B-\xFFFF]", "_")
;~ 	$sString = StringRegExpReplace ( $sString, "\W", "_")
	$sString = StringRegExpReplace ( $sString, "[^\w\h\-]", "_")
	If @error == 0 And @extended > 0 Then
		ConsoleWrite("NonAscii: " & $sString & @LF)
	EndIf
	Return $sString
EndFunc   ;==>_StringReplaceNonAscii

Func _StringReplaceDoubleUnderline($sString) ; replaces double underscores with just one and deletes leading/trailing underscores
	$sString = StringRegExpReplace ($sString, "_{2,}", "_")
	If @error == 0 And @extended > 0 Then
		ConsoleWrite("Underline: " & $sString & @LF)
	EndIf
	If StringLeft($sString,1) = "_" Then $sString = StringTrimLeft($sString,1) ; deletes leading underscore
	If StringRight($sString,1) = "_" Then $sString = StringTrimRight($sString,1) ; deletes trailing underscore
	Return $sString
EndFunc   ;===>_StringReplaceDoubleUnderline

Func WM_DROPFILES_FUNC($hWnd, $msgID, $wParam, $lParam)
	If $bPaused Then Return
	#forceref $hWnd, $msgID, $wParam, $lParam
    Local $nSize, $pFileName
    Local $nAmt = DllCall('shell32.dll', 'int', 'DragQueryFileW', 'hwnd', $wParam, 'int', 0xFFFFFFFF, 'ptr', 0, 'int', 0)
    ReDim $g_aDropFiles[$nAmt[0]]
    For $i = 0 To $nAmt[0] - 1
        $nSize = DllCall('shell32.dll', 'int', 'DragQueryFileW', 'hwnd', $wParam, 'int', $i, 'ptr', 0, 'int', 0)
        $nSize = $nSize[0] + 1
        $pFileName = DllStructCreate('wchar[' & $nSize & ']')
        DllCall('shell32.dll', 'int', 'DragQueryFileW', 'hwnd', $wParam, 'int', $i, 'ptr', DllStructGetPtr($pFileName), 'int', $nSize)
        $g_aDropFiles[$i] = DllStructGetData($pFileName, 1)
        $pFileName = 0
    Next
	_ArrayInsert($g_aDropFiles, 0, UBound($g_aDropFiles))
    GUICtrlSendToDummy($FILES_DROPPED, $nAmt[0])
EndFunc   ;==>WM_DROPFILES_FUNC
#EndRegion

#region - Funcs RFLTA
Func _ErrorRFLTA($iError)
	Local $msg
	Switch $iError
		Case 1
			$msg = "Path not found or invalid"
		Case 2
			$msg = "Invalid $sInclude_List"
		Case 3
			$msg = "Invalid $iReturn"
		Case 4
			$msg = "Invalid $fRecur"
		Case 5
			$msg = "Invalid $fSort"
		Case 6
			$msg = "Invalid $iFullPath"
		Case 7
			$msg = "Invalid $sExclude_List"
		Case 8
			$msg = "No files/folders found"
	EndSwitch
	MsgBox(0, 'Fehler', "_RecFileListToArray: " & $msg)
	Return 666
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _RecFileListToArray
; Description ...: Lists files and\or folders in a specified path with optional recursion and sort.  Compatible with existing _FileListToArray syntax
; Syntax.........: _RecFileListToArray($sPath[, $sInclude_List = "*"[, $iReturn = 0[, $fRecur = 0[, $fSort = 0[, $sReturnPath = 1[, $sExclude_List = ""]]]]]])
; Parameters ....: $sPath - Initial path used to generate filelist.  If path ends in \ then folders will be returned with an ending \
;                  $sInclude_List - Optional: the filter for included results (default is "*"). Multiple filters must be separated by ";"
;                  $iReturn - Optional: specifies whether to return files, folders or both
;                  |$iReturn = 0 (Default) Return both files and folders
;                  |$iReturn = 1 Return files only
;                  |$iReturn = 2 Return folders only
;                  $fRecur - Optional: specifies whether to Search recursively in subfolders
;                  |$fRecur = 0 (Default) Do not Search in subfolders
;                  |$fRecur = 1 Search in subfolders
;                  $fSort - Optional: sort ordered in alphabetical and depth order
;                  |$fSort = 0 (Default) Not sorted
;                  |$fSort = 1 Sorted
;                  $sReturnPath - Optional: specifies displayed path of results
;                  |$sReturnPath = 0 File/folder name only
;                  |$sReturnPath = 1 (Default) Initial path not included
;                  |$sReturnPath = 2 Initial path included
;                  $sExclude_List - Optional: the filter for excluded results (default is ""). Multiple filters must be separated by ";"
; Requirement(s).: v3.3.1.1 or higher
; Return values .: Success: One-dimensional array made up as follows:
;                  |$array[0] = Number of Files\Folders returned
;                  |$array[1] = 1st File\Folder
;                  |$array[2] = 2nd File\Folder
;                  |...
;                  |$array[n] = nth File\Folder
;                   Failure: Null string and @error = 1 with @extended set as follows:
;                  |1 = Path not found or invalid
;                  |2 = Invalid $sInclude_List
;                  |3 = Invalid $iReturn
;                  |4 = Invalid $fRecur
;                  |5 = Invalid $fSort
;                  |6 = Invalid $iFullPath
;                  |7 = Invalid $sExclude_List
;                  |8 = No files/folders found
; Author ........: Melba23 using SRE code from forums
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; Yes
; ===============================================================================================================================
Func _RecFileListToArray($sInitialPath, $sInclude_List = "*", $iReturn = 0, $fRecur = 0, $fSort = 0, $sReturnPath = 1, $sExclude_List = "")

    Local $asReturnList[100] = [0], $asFileMatchList[100] = [0], $asRootFileMatchList[100] = [0], $asFolderMatchList[100] = [0], $asFolderList[100] = [1]
    Local $sFolderSlash = "", $sInclude_List_Mask, $sExclude_List_Mask, $hSearch, $fFolder, $sRetPath = "", $sCurrentPath, $sName
    ; Check valid path
    If Not FileExists($sInitialPath) Then Return SetError(1, 1, "")
    ; Check if folders should have trailing \ and ensure that $sInitialPath does have one
    If StringRight($sInitialPath, 1) = "\" Then
        $sFolderSlash = "\"
    Else
        $sInitialPath = $sInitialPath & "\"
    EndIf
    ; Add path to folder list
    $asFolderList[1] = $sInitialPath

    ; Determine Filter mask for SRE check
    If $sInclude_List = "*" Then
        $sInclude_List_Mask = ".+" ; Set mask to exclude base folder with NULL name
    Else
        If StringRegExp($sInclude_List, "\\|/|:|\<|\>|\|") Then Return SetError(1, 2, "") ; Check For invalid characters within $sInclude_List
        $sInclude_List = StringReplace(StringStripWS(StringRegExpReplace($sInclude_List, "\s*;\s*", ";"), 3), ";", "|") ; Strip WS and insert | for ;
        $sInclude_List_Mask = "(?i)^(" & StringReplace(StringReplace(StringRegExpReplace($sInclude_List, "(\^|\$|\.)", "\\$1"), "?", "."), "*", ".*?") & ")\z" ; Convert to SRE pattern
    EndIf

    ; Determine Exclude mask for SRE check
    If $sExclude_List = "" Then
        $sExclude_List_Mask = ":" ; Set unmatchable mask
    Else
        If StringRegExp($sExclude_List, "\\|/|:|\<|\>|\|") Then Return SetError(1, 7, "") ; Check For invalid characters within $sInclude_List
        $sExclude_List = StringReplace(StringStripWS(StringRegExpReplace($sExclude_List, "\s*;\s*", ";"), 3), ";", "|") ; Strip WS and insert | for ;
        $sExclude_List_Mask = "(?i)^(" & StringReplace(StringReplace(StringRegExpReplace($sExclude_List, "(\^|\$|\.)", "\\$1"), "?", "."), "*", ".*?") & ")\z" ; Convert to SRE pattern
    EndIf

    ; Verify other parameter values
    If Not ($iReturn = 0 Or $iReturn = 1 Or $iReturn = 2) Then Return SetError(1, 3, "")
    If Not ($fRecur = 0 Or $fRecur = 1) Then Return SetError(1, 4, "")
    If Not ($fSort = 0 Or $fSort = 1) Then Return SetError(1, 5, "")
    If Not ($sReturnPath = 0 Or $sReturnPath = 1 Or $sReturnPath = 2) Then Return SetError(1, 6, "")

    ; Search in listed folders
    While $asFolderList[0] > 0
        ; Set path to search
        $sCurrentPath = $asFolderList[$asFolderList[0]]
		; Reduce folder array count
        $asFolderList[0] -= 1

        ; Determine return path to add to file/folder name
        Switch $sReturnPath
            ; Case 0 ; Name only
            ; Leave as ""
            Case 1 ; Initial path not included
                $sRetPath = StringReplace($sCurrentPath, $sInitialPath, "")
            Case 2 ; Initial path included
                $sRetPath = $sCurrentPath
        EndSwitch

        If $fSort Then

            ; Get folder name
            $sName = StringRegExpReplace(StringReplace($sCurrentPath, $sInitialPath, ""), "(.+?\\)*(.+?)(\\.*?(?!\\))", "$2")

            ; Get Search handle
            $hSearch = FileFindFirstFile($sCurrentPath & "*")
            ; If folder empty move to next in list
            If $hSearch = -1 Then ContinueLoop

            ; Search folder
            While 1
                $sName = FileFindNextFile($hSearch)
                ; Check for end of folder
                If @error Then ExitLoop
                ; Check for file - @extended set for subfolder in 3.3.1.1 +
                If @extended Then
                    ; If recursive search, add subfolder to folder list
                    If $fRecur Then _RFLTA_AddToList($asFolderList, $sCurrentPath & $sName & "\")

                    ; Add folder name if matched against Include/Exclude masks
                    If StringRegExp($sName, $sInclude_List_Mask) And Not StringRegExp($sName, $sExclude_List_Mask) Then _
                            _RFLTA_AddToList($asFolderMatchList, $sRetPath & $sName & $sFolderSlash)

                Else
                    ; Add file name if matched against Include/Exclude masks
                    If StringRegExp($sName, $sInclude_List_Mask) And Not StringRegExp($sName, $sExclude_List_Mask) Then
                        If $sCurrentPath = $sInitialPath Then
                            _RFLTA_AddToList($asRootFileMatchList, $sRetPath & $sName)
                        Else
                            _RFLTA_AddToList($asFileMatchList, $sRetPath & $sName)
                        EndIf
                    EndIf
                EndIf
            WEnd

            ; Close current search
            FileClose($hSearch)

        Else ; No sorting required

            ; Get Search handle
            $hSearch = FileFindFirstFile($sCurrentPath & "*")
            ; If folder empty move to next in list
            If $hSearch = -1 Then ContinueLoop

            ; Search folder
            While 1
                $sName = FileFindNextFile($hSearch)
                ; Check for end of folder
                If @error Then ExitLoop
                ; Check for subfolder - @extended set in 3.3.1.1 +
                $fFolder = @extended

                ; If recursive search, add subfolder to folder list
                If $fRecur And $fFolder Then _RFLTA_AddToList($asFolderList, $sCurrentPath & $sName & "\")

                ; Check file/folder type against required return value and file/folder name against Include/Exclude masks
                If $fFolder + $iReturn <> 2 And StringRegExp($sName, $sInclude_List_Mask) And Not StringRegExp($sName, $sExclude_List_Mask) Then
                    ; Add final "\" to folders
                    If $fFolder Then $sName &= $sFolderSlash
                    _RFLTA_AddToList($asReturnList, $sRetPath & $sName)
                EndIf
            WEnd

            ; Close current search
            FileClose($hSearch)

        EndIf

    WEnd

    If $fSort Then

        ; Check if any file/folders have been added
        If $asRootFileMatchList[0] = 0 And $asFileMatchList[0] = 0 And $asFolderMatchList[0] = 0 Then Return SetError(1, 8, "")

        Switch $iReturn
            Case 2 ; Folders only
                ; Correctly size folder match list
                ReDim $asFolderMatchList[$asFolderMatchList[0] + 1]
                ; Copy size folder match array
                $asReturnList = $asFolderMatchList
                ; Simple sort list
                _RFLTA_ArraySort($asReturnList)
            Case 1 ; Files only
                If $sReturnPath = 0 Then ; names only so simple sort suffices
                    ; Combine file match lists
                    _RFLTA_AddFileLists($asReturnList, $asRootFileMatchList, $asFileMatchList)
                    ; Simple sort combined file list
                    _RFLTA_ArraySort($asReturnList)
                Else
                    ; Combine sorted file match lists
                    _RFLTA_AddFileLists($asReturnList, $asRootFileMatchList, $asFileMatchList, 1)
                EndIf
            Case 0 ; Both files and folders
                If $sReturnPath = 0 Then ; names only so simple sort suffices
                    ; Combine file match lists
                    _RFLTA_AddFileLists($asReturnList, $asRootFileMatchList, $asFileMatchList)
                    ; Set correct count for folder add
                    $asReturnList[0] += $asFolderMatchList[0]
                    ; Resize and add file match array
                    ReDim $asFolderMatchList[$asFolderMatchList[0] + 1]
                    _RFLTA_ArrayConcatenate($asReturnList, $asFolderMatchList)
                    ; Simple sort final list
                    _RFLTA_ArraySort($asReturnList)
                Else
                    ; Combine sorted file match lists
                    _RFLTA_AddFileLists($asReturnList, $asRootFileMatchList, $asFileMatchList, 1)
                    ; Add folder count
                    $asReturnList[0] += $asFolderMatchList[0]
                    ; Sort folder match list
                    ReDim $asFolderMatchList[$asFolderMatchList[0] + 1]
                    _RFLTA_ArraySort($asFolderMatchList)

                    ; Now add folders in correct place
                    Local $iLastIndex = $asReturnList[0]
                    For $i = $asFolderMatchList[0] To 1 Step -1
                        ; Find first filename containing folder name
                        Local $iIndex = _RFLTA_ArraySearch($asReturnList, $asFolderMatchList[$i])
                        If $iIndex = -1 Then
                            ; Empty folder so insert immediately above previous
                            _RFLTA_ArrayInsert($asReturnList, $iLastIndex, $asFolderMatchList[$i])
                        Else
                            ; Insert folder at correct point above files
                            _RFLTA_ArrayInsert($asReturnList, $iIndex, $asFolderMatchList[$i])
                            $iLastIndex = $iIndex
                        EndIf
                    Next
                EndIf
        EndSwitch

    Else ; No sort

        ; Check if any file/folders have been added
        If $asReturnList[0] = 0 Then Return SetError(1, 8, "")

        ; Remove any unused return list elements from last ReDim
        ReDim $asReturnList[$asReturnList[0] + 1]

    EndIf
    Return $asReturnList

EndFunc   ;==>_RecFileListToArray

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: _RFLTA_AddToList
; Description ...: Add element to list which is resized if necessary
; Syntax ........: _RFLTA_AddToList(ByRef $asList, $sValue)
; Parameters ....: $asList - List to be added to
;                  $sValue - Value to add
; Return values .: None - array modified ByRef
; Author ........: Melba23
; Remarks .......: This function is used internally by _RecFileListToArray
; ===============================================================================================================================
Func _RFLTA_AddToList(ByRef $asList, $sValue)

    ; Increase list count
    $asList[0] += 1
    ; Double list size if too small (fewer ReDim needed)
    If UBound($asList) <= $asList[0] Then ReDim $asList[UBound($asList) * 2]
    ; Add value
    $asList[$asList[0]] = $sValue

EndFunc   ;==>_RFLTA_AddToList

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: _RFLTA_AddFileLists
; Description ...: Add internal arrays after resizing and optional sorting
; Syntax ........: _RFLTA_AddFileLists(ByRef $asReturnList, $asRootFileMatchList, $asFileMatchList[, $iSort = 0])
; Parameters ....: $asReturnList - Base list
;                  $asRootFileMatchList - First list to add
;                  $asFileMatchList - Second list to add
;                  $iSort - (Optional) Whether to sort lists before adding
;                  |$iSort = 0 (Default) No sort
;                  |$iSort = 1 Sort in descending alphabetical order
; Return values .: None - array modified ByRef
; Author ........: Melba23
; Remarks .......: This function is used internally by _RecFileListToArray
; ===============================================================================================================================
Func _RFLTA_AddFileLists(ByRef $asReturnList, $asRootFileMatchList, $asFileMatchList, $iSort = 0)

    ; Correctly size root file match array
    ReDim $asRootFileMatchList[$asRootFileMatchList[0] + 1]
    ; Simple sort root file match array if required
    If $iSort = 1 Then _RFLTA_ArraySort($asRootFileMatchList)
    ; Copy root file match array
    $asReturnList = $asRootFileMatchList
    ; Add file match count
    $asReturnList[0] += $asFileMatchList[0]
    ; Correctly size file match array
    ReDim $asFileMatchList[$asFileMatchList[0] + 1]
    ; Simple sort file match array if required
    If $iSort = 1 Then _RFLTA_ArraySort($asFileMatchList)
    ; Add file match array
    _RFLTA_ArrayConcatenate($asReturnList, $asFileMatchList)

EndFunc   ;==>_RFLTA_AddFileLists

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: _RFLTA_ArraySearch
; Description ...: Search array downwards for partial match
; Syntax ........: _RFLTA_ArraySearch(Const ByRef $avArray, $vValue)
; Parameters ....: $avArray - Array to search
;                  $vValue - PValue to Search for
; Return values .: Success: Index of array in which element was found
;                  Failure: returns -1
; Author ........: SolidSnake, gcriaco, Ultima
; Modified.......: Melba23
; Remarks .......: This function is used internally by _RecFileListToArray
; ===============================================================================================================================
Func _RFLTA_ArraySearch(Const ByRef $avArray, $vValue)

    For $i = 1 To UBound($avArray) - 1
        If StringInStr($avArray[$i], $vValue) > 0 Then Return $i
    Next
    Return -1

EndFunc   ;==>_RFLTA_ArraySearch

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: _RFLTA_ArraySort
; Description ...: Wrapper for QuickSort function
; Syntax ........: _RFLTA_ArraySort(ByRef $avArray)
; Parameters ....: $avArray - Array to sort
;                  $pNew_WindowProc - Pointer to new WindowProc
; Return values .: None - array modified ByRef
; Author ........: Jos van der Zande, LazyCoder, Tylo, Ultima
; Modified.......: Melba23
; Remarks .......: This function is used internally by _RecFileListToArray
; ===============================================================================================================================
Func _RFLTA_ArraySort(ByRef $avArray)

    Local $iStart = 1, $iEnd = UBound($avArray) - 1
    _RFLTA_QuickSort($avArray, $iStart, $iEnd)

EndFunc   ;==>_RFLTA_ArraySort

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: _RFLTA_QuickSort
; Description ...: recursive array sort
; Syntax ........: _RFLTA_QuickSort(ByRef $avArray, ByRef $iStart, ByRef $iEnd)
; Parameters ....: $avArray - Array to sort in descending alphabetical order
;                  $iStart - Start index
;                  $iEnd - End index
; Return values .: None - array modified ByRef
; Author ........: Jos van der Zande, LazyCoder, Tylo, Ultima
; Modified.......: Melba23
; Remarks .......: This function is used internally by _RFLTA_ArraySort
; ===============================================================================================================================
Func _RFLTA_QuickSort(ByRef $avArray, ByRef $iStart, ByRef $iEnd)

    Local $vTmp
    If ($iEnd - $iStart) < 15 Then
        Local $i, $j, $vCur
        For $i = $iStart + 1 To $iEnd
            $vTmp = $avArray[$i]
            If IsNumber($vTmp) Then
                For $j = $i - 1 To $iStart Step -1
                    $vCur = $avArray[$j]
                    If ($vTmp >= $vCur And IsNumber($vCur)) Or (Not IsNumber($vCur) And StringCompare($vTmp, $vCur) >= 0) Then ExitLoop
                    $avArray[$j + 1] = $vCur
                Next
            Else
                For $j = $i - 1 To $iStart Step -1
                    If (StringCompare($vTmp, $avArray[$j]) >= 0) Then ExitLoop
                    $avArray[$j + 1] = $avArray[$j]
                Next
            EndIf
            $avArray[$j + 1] = $vTmp
        Next
        Return
    EndIf
    Local $L = $iStart, $R = $iEnd, $vPivot = $avArray[Int(($iStart + $iEnd) / 2)], $fNum = IsNumber($vPivot)
    Do
        If $fNum Then
            While ($avArray[$L] < $vPivot And IsNumber($avArray[$L])) Or (Not IsNumber($avArray[$L]) And StringCompare($avArray[$L], $vPivot) < 0)
                $L += 1
            WEnd
            While ($avArray[$R] > $vPivot And IsNumber($avArray[$R])) Or (Not IsNumber($avArray[$R]) And StringCompare($avArray[$R], $vPivot) > 0)
                $R -= 1
            WEnd
        Else
            While (StringCompare($avArray[$L], $vPivot) < 0)
                $L += 1
            WEnd
            While (StringCompare($avArray[$R], $vPivot) > 0)
                $R -= 1
            WEnd
        EndIf
        If $L <= $R Then
            $vTmp = $avArray[$L]
            $avArray[$L] = $avArray[$R]
            $avArray[$R] = $vTmp
            $L += 1
            $R -= 1
        EndIf
    Until $L > $R
    _RFLTA_QuickSort($avArray, $iStart, $R)
    _RFLTA_QuickSort($avArray, $L, $iEnd)

EndFunc   ;==>_RFLTA_QuickSort

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: _RFLTA_ArrayConcatenate
; Description ...: Joins 2 arrays
; Syntax ........: _RFLTA_ArrayConcatenate(ByRef $avArrayTarget, Const ByRef $avArraySource)
; Parameters ....: $avArrayTarget - Base array
;                  $avArraySource - Array to add from element 1 onwards
; Return values .: None - array modified ByRef
; Author ........: Ultima
; Modified.......: Melba23
; Remarks .......: This function is used internally by _RecFileListToArray
; ===============================================================================================================================
Func _RFLTA_ArrayConcatenate(ByRef $avArrayTarget, Const ByRef $avArraySource)

    Local $iUBoundTarget = UBound($avArrayTarget) - 1, $iUBoundSource = UBound($avArraySource)
    ReDim $avArrayTarget[$iUBoundTarget + $iUBoundSource]
    For $i = 1 To $iUBoundSource - 1
        $avArrayTarget[$iUBoundTarget + $i] = $avArraySource[$i]
    Next

EndFunc   ;==>_RFLTA_ArrayConcatenate

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: _RFLTA_ArrayInsert
; Description ...: Insert element into array
; Syntax ........: _RFLTA_ArrayInsert(ByRef $avArray, $iElement, $vValue = "")
; Parameters ....: $avArray - Array to modify
;                  $iElement - Index position for insertion
;                  $vValue - Value to insert
; Return values .: None - array modified ByRef
; Author ........: Jos van der Zande, Ultima
; Modified.......: Melba23
; Remarks .......: This function is used internally by _RecFileListToArray
; ===============================================================================================================================
Func _RFLTA_ArrayInsert(ByRef $avArray, $iElement, $vValue = "")

    Local $iUBound = UBound($avArray) + 1
    ReDim $avArray[$iUBound]
    For $i = $iUBound - 1 To $iElement + 1 Step -1
        $avArray[$i] = $avArray[$i - 1]
    Next
    $avArray[$iElement] = $vValue

EndFunc   ;==>_RFLTA_ArrayInsert
#endregion Funcs _RFLTA