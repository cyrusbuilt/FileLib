#include-once
Opt('MustDeclareVars', 1)

; #INDEX# =======================================================================================================================
; Name ...........: FileLib.au3
; AutoIt Version .: 3.3.2.0++
; Author .........: Chris Brunner <CyrusBuilt at gmail dot com>
; Language .......: English
; Library Version : 1.0.2.2
; Modified .......: 05/13/2014
; License ........: GPL v2
; Description ....: A collection of additional functions for manipulating files, directories, and disks.
; Remarks ........: Thanks to Livewire for code that some of these functions are based on.
; Dependencies ...: Microsoft Windows 2000/XP/2003/Vista/Win7/Server 2008. Not yet tested under Win8/8.1 or Server 2012.
; ================================================================================================================================


; #CURRENT# =====================================================================================================================
; _GetFileLibVersion
; _FileExtension
; _FileCompare
; _DiskExist
; _DiskInfoToArray
; _DiskIsNetwork
; _FileCopyDlg
; _FileMoveDlg
; _FileRenameDlg
; _FileDeleteDlg
; _FileIsDir
; _FileRecurseBuildList
; _FileGetFileName
; _FileCheckDirEmpty
; _FileGetDirFromPath
; _FileTruncateName
; _FileGetSystemDrive
; _FileGetProfilesDir
; _FileRemoveDirIfEmpty
; _DirectoryIsSymlink
; ===============================================================================================================================


; #INTERNAL_USE_ONLY#============================================================================================================
; __SHFileOperation
; __TrimTrailingDelimiter
; ===============================================================================================================================


; #VARIABLES# ======================================================================================================================
Global Const $CB_FILE_LIB_VER           = "1.0.2.2"
;Operation constants.
Global Const $FO_COPY                   = 0x0002        ; Copies the files specified in pFrom to the location specified in pTo.
Global Const $FO_DELETE                 = 0x0003        ; Deletes the files specified in pFrom (pTo is ignored).
Global Const $FO_MOVE                   = 0x0001        ; Moves the files specified in pFrom to the location specified in pTo.
Global Const $FO_RENAME                 = 0x0004        ; Renames the files specified in pFrom.
;Operation flags.
Global Const $FOF_ALLOWUNDO             = 0x0040        ; Preserve Undo information, if possible.
Global Const $FOF_CONFIRMMOUSE          = 0x0002        ; Not currently implemented.
Global Const $FOF_FILESONLY             = 0x0080        ; Perform the operation on files only if a wildcard file name (*.*) is specified.
Global Const $FOF_MULTIDESTFILES        = 0x0001        ; The pTo member specifies multiple destination files (one for each source file)
														; rather than one directory where all source files are to be deposited.
Global Const $FOF_NOCONFIRMATION        = 0x0010        ; Respond with "Yes to All" for any dialog box that is displayed.
Global Const $FOF_NOCONFIRMMKDIR        = 0x0200        ; Does not confirm the creation of a new directory if the operation requires one to be created.
Global Const $FOF_NOCOPYSECURITYATTRIBS = 0x0800        ; Do not copy NT file Security Attributes.
Global Const $FOF_NOERRORUI             = 0x0400        ; No user interface will be displayed if an error occurs.
Global Const $FOF_NORECURSION           = 0x1000        ; Do not recurse directories (i.e. no recursion into subdirectories).
Global Const $FOF_RENAMEONCOLLISION     = 0x0008        ; Give the file being operated on a new name in a move, copy, or rename operation if a file with the target name already exisits.
Global Const $FOF_SILENT                = 0x0004        ; Does not display a progress dialog box.
Global Const $FOF_SIMPLEPROGRESS        = 0x0100        ; Displays a progress box but does not show the file names.
Global Const $FOF_WANTMAPPINGHANDLE     = 0x0020        ; If $FOF_RENAMEONCOLLISION is specified, the hNameMappings member will be filled in if any files were renamed.

; These two apply in Internet Explorer 5 (or higher) Environments.
Global Const $FOF_NO_CONNECTED_ELEMENTS = 0x2000        ; Do not operate on connected elements.
Global Const $FOF_WANTNUKEWARNING       = 0x4000        ; During delete operations, warn if permanent deleting instead of placing in recycle bin (partially overrides $FOF_NOCONFIRMATION).

; ????
Global Const $FOF_NORECURSEREPARSE      = 0x8000        ;

Global Const $FILE_ATTRIBUTE_REPARSE_POINT = 0x400      ; Attribute indicating a directory is a reparse point (symlink).
; ===============================================================================================================================



; #FUNCTION# ====================================================================================================================
; Name...........: _GetFileLibVersion
; Description ...: Gets the current version of this library as defined by the $cbFileLibVer constant.
; Syntax.........: _GetFileLibVersion()
; Parameters ....: None.
; Return values .: A four-octet string representing the current version of this library.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/23/08
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _GetFileLibVersion()
	Return $CB_FILE_LIB_VER
EndFunc


; #FUNCTION# ====================================================================================================================
; Name...........: _FileExtension
; Description ...: Returns the file extension of the specified path.
; Syntax.........: _FileExtension($path)
; Parameters ....: $path - The name or full path of the file to get the extension of.
; Return values .: Success - Returns the extension of the file specified without the dot (".").
;                  Failure - If the specified name or path does not have an extension it will return an empty string.  If $path is empty
;                            or undefined then an empty string is returned and @error is set to 1.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/01/2010
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; Local $Extension = _FileExtension("My file.txt")  ;$Extension = "txt"
; ===============================================================================================================================
Func _FileExtension($path)
	If (StringLen($path) == 0) Then
		Return SetError(1, 0, "")
	EndIf

	Local $arrStrings = StringSplit($path, ".")
	If ((@error) Or ($arrStrings[0] == 0)) Then
		Return ""
	EndIf

	Local $i = 0
	Local $sExtension = ""
	For $i = 1 To $arrStrings[0]
		If ($i == $arrStrings[0]) Then
			$sExtension = $arrStrings[$i]
			ExitLoop
		EndIf
	Next

	$i = 0
	$arrStrings = 0
	$path = 0
	Return $sExtension
EndFunc   ;==>_FileExtension


; #FUNCTION# ====================================================================================================================
; Name...........: _FileCompare
; Description ...: Compares 2 files to see if they are the same.
; Syntax.........: _FileCompare($file1, $file2)
; Parameters ....: $file1 - First file to compare.
;                  $file2 - Second file to compare.
; Return values .: True  - Files are the same.
;                  False - Files are different.
;                  Sets @error to:
;                  |1 - $file1 not defined, is a directory, or does not exist.
;                  |2 - $file2 not defined, is a directory, or does not exist.
;                  |3 - Files did not match or 'FC' command error and sets @extended to the ERRORLEVEL returned by 'FC'.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/01/2010
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileCompare($file1, $file2)
	If ((StringLen($file1) == 0) Or (Not FileExists($file1)) Or (_FileIsDir($file1))) Then
		Return SetError(1, 0, False)
	EndIf

	If ((StringLen($file2) == 0) Or (Not FileExists($file2)) Or (_FileIsDir($file2))) Then
		Return SetError(2, 0, False)
	EndIf

	Local $nRet = RunWait(@ComSpec & ' /c fc /b /c "' & $file1 & '" "' & $file2 & '"', @ScriptDir, @SW_HIDE)
	If ($nRet == 0) Then
		Return True
	Else
		Return SetError(3, $nRet, False)
	EndIf
EndFunc   ;==>_FileCompare


; #FUNCTION# ====================================================================================================================
; Name...........: _DiskExist
; Description ...: Checks to see if a drive letter exists.
; Syntax.........: _DiskExist($drive)
; Parameters ....: $drive - A logical drive letter followed by a colon. i.e.: "D:"
; Return values .: True  - Drive letter exists and is valid.
;                  False - Drive does not exist.  Sets @error = 1 if parameter undefined.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/01/2010
; Remarks .......:
; Related .......: _DiskInfoToArray, _DiskIsNetwork
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _DiskExist($drive)
	If ((StringLen($drive) == 0) Or (StringLen($drive) > 2)) Then
		Return SetError(1, 0, False)
	EndIf

	$drive = StringStripWS($drive, 8)
	If (DriveStatus(StringUpper($drive)) == "INVALID") Then
		Return False
	EndIf
	Return True
EndFunc   ;==>_DiskExist


; #FUNCTION# ====================================================================================================================
; Name...........: _DiskInfoToArray
; Description ...: Returns an array containing info about a specified drive letter.
; Syntax.........: _DiskInfoToArray($drive)
; Parameters ....: $drive - A logical drive letter followed by a colon. i.e.: "D:"
; Return values .: Success - An array with the following elements:
;                            |0 - Drive status ("UNKNOWN", "READY", or "NOTREADY")
;                            |1 - Drive label
;                            |2 - Drive type ("Unknown", "Removable", "Fixed", "Network", "CDROM", or "RAMDisk")
;                            |3 - Drive serial number
;                            |4 - Filesystem ("FAT", "FAT32", "NTFS", "NWFS", "CDFS", "UDF", or 1 if RAW or volume not in drive)
;                            |5 - Volume free space (MB)
;                            |6 - Volume used space (MB)
;                            |7 - Volume total space (MB)
;                  Failure - Sets @error = 1 if specified drive does not exist.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/01/2010
; Remarks .......:
; Related .......: _DiskExist, _DiskIsNetwork
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _DiskInfoToArray($drive)
	Local $aInfo[8]
	Local $nFree
	Local $nUsed
	Local $nTotal

	$aInfo[0] = "INVAlID"
	If (Not _DiskExist($drive)) Then
		Return SetError(1, 0, $aInfo)
	EndIf

	$drive = StringUpper($drive) & "\"
	$aInfo[0] = DriveStatus($drive)
	$aInfo[1] = DriveGetLabel($drive)
	$aInfo[2] = DriveGetType($drive)
	$aInfo[3] = DriveGetSerial($drive)
	$aInfo[4] = DriveGetFileSystem($drive)

	$nFree = Floor(DriveSpaceFree($drive))
	$nTotal	= Floor(DriveSpaceTotal($drive))
	$nUsed = ($nTotal - $nFree)

	$aInfo[5] = $nFree
	$aInfo[6] = $nUsed
	$aInfo[7] = $nTotal
	Return $aInfo
EndFunc   ;==>_DiskInfoToArray


; #FUNCTION# ====================================================================================================================
; Name...........: _DiskIsNetwork
; Description ...: Determines if a specified drive letter is a network drive.
; Syntax.........: _DiskIsNetwork($drive)
; Parameters ....: $drive - A logical drive letter followed by a colon. i.e.: "D:"
; Return values .: True - Specified drive letter is a network drive.
;                  False - Specified drive is not a network drive.
;                  Sets @error:
;                        |1 - if drive does not exist
;                        |2 - if unable to determine drive type.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/01/2010
; Remarks .......:
; Related .......: _DiskExist, _DiskInfoToArray
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _DiskIsNetwork($drive)
	If _DiskExist($drive) = False Then
		Return SetError(1, 0, False)
	EndIf

	$drive = StringUpper($drive) & "\"
	If (DriveGetType($drive) == "Network") Then
		Return True
	ElseIf (@error) Then
		Return SetError(2, 0, False)
	Else
		Return False
	EndIf
EndFunc   ;==>_DiskIsNetwork


; #FUNCTION# ====================================================================================================================
; Name...........: _FileIsDir
; Description ...: Determines whether or not the supplied path is a directory.
; Syntax.........: _FileIsDir($path)
; Parameters ....: $path - The path to check.
; Return values .: True  - The specified path is a directory.
;                  False - The specified path is not a directory.  Sets @error = 1 if path does not exist.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/15/08
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileIsDir($path)
	If (FileExists($path)) Then
		If (StringInStr(FileGetAttrib($path), "D", 1) > 0) Then
			Return True
		EndIf
	Else
		Return SetError(1, 0, False)
	EndIf
	Return False
EndFunc   ;==>_FileIsDir


; #FUNCTION# ====================================================================================================================
; Name...........: _FileRecurseBuildList
; Description ...: Performs file recursion on the specified directory and builds an array containing the full path to every file
;                  matching the search parameter (all files by default) in the specified directory and every subdirectory.
; Syntax.........: _FileRecurseBuildList($path, ByRef $aFiles, $filename)
; Parameters ....: $path     - The path to check.
;                  $aFiles   - An array to contain the file paths and the count.
;                  $filename - The filename string to search for in $Path (* and ? wildcards accepted).
; Return values .: Success - $aFiles - The resulting path array with the path count at element 0.  If no file or directory exists within
;                  the parent path specified then $aFiles will be returned with only element 0 and it's value will be 0.
;                  Failure - If the parent path specified does not exist, the same will be true but it will also set @error to 1.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/01/2010
; Remarks .......: Only returns FILE paths.  If $Path contains any empty directories, then they will not be included in the list.
; Related .......: _FileIsDir
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileRecurseBuildList($path, ByRef $aFiles, $filename = "*.*")
	Local $sSearch
	Local $sFile
	Local $sFileFullPath

	If ((StringLen($path) == 0) Or (Not FileExists($path))) Then
		Return SetError(1, 0, $aFiles)
	EndIf

	If (Not IsArray($aFiles)) Then
		Return SetError(1, 0, $aFiles)
	EndIf

	If (StringRight($Path, 1) == "\") Then
		$path = StringLeft($path, StringLen($path) - 1)
	EndIf

	$sSearch = FileFindFirstFile($path & "\" & $filename)
	While 1
		If ($sSearch == -1) Then
			ExitLoop
		EndIf

		$sFile = FileFindNextFile($sSearch)
		If (@error) Then
			ExitLoop
		EndIf

		$sFileFullPath = $path & "\" & $sFile
		If (_FileIsDir($sFileFullPath)) Then
			_FileRecurseBuildList($sFileFullPath, $aFiles)
		Else
			$aFiles[0] = $aFiles[0] + 1
			ReDim $aFiles[$aFiles[0] + 1]
			$aFiles[$aFiles[0]] = $sFileFullPath
		EndIf
	WEnd

	FileClose($sSearch)
	Return $aFiles
EndFunc   ;==>_FileRecurseBuildList


; #FUNCTION# ====================================================================================================================
; Name...........: _FileGetFileName
; Description ...: Gets the file name from the specified file path.
; Syntax.........: _FileGetFileName($path)
; Parameters ....: $path   - A full file path (drive, directory, filename).
; Return values .: Success - The file name string (minus the rest of the path).
;                  Failure - Returns a blank string ( "" ) and sets @error to:
;                            |1 - The value specified for path is empty or not a string at all.
;                            |2 - Specified path is a directory.
;                            |3 - Could not split specified path into an array (no path separator detected).  This will occur
;                                 if *only* a file name is specified and not an actual path.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/17/2010
; Remarks .......:
; Related .......: _FileIsDir
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileGetFileName($path)
	If ((Not IsString($path)) Or (StringLen($path) == 0)) Then
		Return SetError(1, 0, "")
	EndIf

	If(_FileIsDir($path)) Then
		Return SetError(2, 0, "")
	EndIf

	Local $aPathStrings = StringSplit($path, "\")
	If ((@error) Or ($aPathStrings[0] == 0)) Then
		$aPathStrings = 0
		Return SetError(3, 0, "")
	EndIf
	Return $aPathStrings[$aPathStrings[0]]
EndFunc   ;==>_FileGetFileName


; #FUNCTION# ====================================================================================================================
; Name...........: _FileCopyDlg
; Description ...: Copys a file, multiple files, or directory to the specified destination.
; Syntax.........: _FileCopyDlg($source, $dest[, flags])
; Parameters ....: $source  - Full path to a file or directory.  Wildcards are supported. A list of individual file paths or
;                             directories can be specifed by seperating each path with @LF.
;                  $dest    - The destination directory.  By default, if the target destination does not exist, it will be created.
;                  $flags   - Supports one or more of the following optional flags (see globals section for flag definitions):
;                             $FOF_ALLOWUNDO
;                             $FOF_FILESONLY
;                             $FOF_MULTIDESTFILES
;                             $FOF_NOCONFIRMATION
;                             $FOF_NOCONFIRMMKDIR
;                             $FOF_NOCOPYSECURITYATTRIBS
;                             $FOF_NOERRORUI
;                             $FOF_NORECURSION
;                             $FOF_RENAMEONCOLLISION
;                             $FOF_SILENT
;                             $FOF_SIMPLEPROGRESS
;                             $FOF_WANTMAPPINGHANDLE   (requires $FOF_RENAMECOLLISION)
;                             $FOF_NO_CONNECTED_ELEMENTS  (requires IE5 or higher)
;                             $FOF_WANTNUKEWARNING  (requires IE5 or higher)
;                             $FOF_NORECURSEPARSE
; Return values .: Success - True.
;                  Failure - False.  Raises error message dialogs for the following errors:
;                            - $source not defined or does not exist.
;                            - $dest not defined.
;                            - Could not create SHFILEOPTSTRUCT structure.
;                            - _SHFileOperation() call failed and was unable to retrieve error info using GetLastError call,
;                              otherwise sets @error to the value returned by GetLastError.
;                            - Operation aborted.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/01/2010
; Remarks .......: Modified from original code by Livewire (see AutoIt Forum).
; Related .......: _SHFileOperation, _TrimTrailingDelimiter, _FileMoveDlg, _FileRenameDlg, _FileDeleteDlg
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileCopyDlg($source, $dest, $flags = 0)
    Local $SHFILEOPSTRUCT, $source_struct, $dest_struct
    Local $i, $aDllRet

	If (StringLen($source) == 0) Then
		MsgBox(16, "Error", "Source file or directory not defined or does not exist.")
		Return False
	EndIf

	If (StringLen($dest) == 0) Then
		MsgBox(16, "Error", "Destination not defined.")
		Return False
	EndIf

	$source = __TrimTrailingDelimiter($source)
	$dest = __TrimTrailingDelimiter($dest)
    Local Enum $hwnd = 1, $wFunc, $pFrom, $pTo, $fFlags, $fAnyOperationsAborted, $hNameMappings, $lpszProgressTitle

    $SHFILEOPSTRUCT = DllStructCreate("int;uint;ptr;ptr;uint;int;ptr;ptr")
    If (@error) Then
        MsgBox(16, "ERROR", "Error creating SHFILEOPSTRUCT structure")
        Return False
    EndIf

    $source_struct = DllStructCreate("char[" & StringLen($source) + 2 & "]")
    DllStructSetData($source_struct, 1, $source)

    For $i = 1 To StringLen($source) + 2
        If (DllStructGetData($source_struct, 1, $i) == 10) Then
			DllStructSetData($source_struct, 1, 0, $i)
		EndIf
	Next

    DllStructSetData($source_struct, 1, 0, StringLen($source) + 2)

    $dest_struct = DllStructCreate("char[" & StringLen($dest) + 2 & "]")
    DllStructSetData($dest_struct, 1, $dest)
    DllStructSetData($dest_struct, 1, 0, StringLen($dest) + 2)
    DllStructSetData($SHFILEOPSTRUCT, $hwnd, 0)
    DllStructSetData($SHFILEOPSTRUCT, $wFunc, $FO_COPY)
    DllStructSetData($SHFILEOPSTRUCT, $pFrom, DllStructGetPtr($source_struct))
    DllStructSetData($SHFILEOPSTRUCT, $pTo, DllStructGetPtr($dest_struct))
    DllStructSetData($SHFILEOPSTRUCT, $fFlags, $flags)
    DllStructSetData($SHFILEOPSTRUCT, $fAnyOperationsAborted, 0)
    DllStructSetData($SHFILEOPSTRUCT, $hNameMappings, 0)
    DllStructSetData($SHFILEOPSTRUCT, $lpszProgressTitle, 0)

    If (__SHFileOperation($SHFILEOPSTRUCT)) Then
        $aDllRet = DllCall("kernel32.dll", "long", "GetLastError")
        If (@error) Then
			MsgBox(16, "Error", "Error calling GetLastError")
		EndIf

        SetError($aDllRet[0])
        Return False
    ElseIf (DllStructGetData($SHFILEOPSTRUCT, $fAnyOperationsAborted)) Then
        MsgBox(16, "Error", "File Copy operation aborted!")
        Return False
    EndIf
    Return True
EndFunc   ;==>_FileCopyDlg


; #FUNCTION# ====================================================================================================================
; Name...........: _FileMoveDlg
; Description ...: Moves a file, multiple files, or directory to the specified destination.
; Syntax.........: _FileMoveDlg($source, $dest[, flags])
; Parameters ....: $source  - Full path to a file or directory.  Wildcards are supported. A list of individual file paths or
;                             directories can be specifed by seperating each path with @LF.
;                  $dest    - The destination directory.  By default, if the target destination does not exist, it will be created.
;                  $flags   - Supports one or more of the following optional flags (see globals section for flag definitions):
;                             $FOF_ALLOWUNDO
;                             $FOF_FILESONLY
;                             $FOF_MULTIDESTFILES
;                             $FOF_NOCONFIRMATION
;                             $FOF_NOCONFIRMMKDIR
;                             $FOF_NOCOPYSECURITYATTRIBS
;                             $FOF_NOERRORUI
;                             $FOF_NORECURSION
;                             $FOF_RENAMEONCOLLISION
;                             $FOF_SILENT
;                             $FOF_SIMPLEPROGRESS
;                             $FOF_WANTMAPPINGHANDLE   (requires $FOF_RENAMECOLLISION)
;                             $FOF_NO_CONNECTED_ELEMENTS  (requires IE5 or higher)
;                             $FOF_WANTNUKEWARNING  (requires IE5 or higher)
;                             $FOF_NORECURSEPARSE
; Return values .: Success - True.
;                  Failure - False.  Raises error message dialogs for the following errors:
;                            - $source not defined or does not exist.
;                            - $dest not defined.
;                            - Could not create SHFILEOPTSTRUCT structure.
;                            - _SHFileOperation() call failed and was unable to retrieve error info using GetLastError call,
;                              otherwise sets @error to the value returned by GetLastError.
;                            - Operation aborted.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/01/2010
; Remarks .......: Modified from original code by Livewire (see AutoIt Forum).
; Related .......: _SHFileOperation, _TrimTrailingDelimiter, _FileCopyDlg, _FileRenameDlg, _FileDeleteDlg
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileMoveDlg($source, $dest, $flags = 0)
    Local $SHFILEOPSTRUCT, $source_struct, $dest_struct
    Local $i, $aDllRet

	If (StringLen($source) == 0) Then
		MsgBox(16, "Error", "Source file or directory not defined or does not exist.")
		Return False
	EndIf

	If (StringLen($dest) == 0) Then
		MsgBox(16, "Error", "Destination not defined.")
		Return False
	EndIf

    $source = __TrimTrailingDelimiter($source)
	$dest = __TrimTrailingDelimiter($dest)
    Local Enum $hwnd = 1, $wFunc, $pFrom, $pTo, $fFlags, $fAnyOperationsAborted, $hNameMappings, $lpszProgressTitle

	$SHFILEOPSTRUCT = DllStructCreate("int;uint;ptr;ptr;uint;int;ptr;ptr")
    If (@error) Then
        MsgBox(4096, "ERROR", "Error creating SHFILEOPSTRUCT structure!")
        Return False
    EndIf

    $source_struct = DllStructCreate("char[" & StringLen($source)+2 & "]")
    DllStructSetData($source_struct, 1, $source)

    For $i = 1 To StringLen($source) + 2
        If (DllStructGetData($source_struct, 1, $i) == 10) Then
			DllStructSetData($source_struct, 1, 0, $i)
		EndIf
	Next

    DllStructSetData($source_struct, 1, 0, StringLen($source)+2)

    $dest_struct = DllStructCreate("char[" & StringLen($dest) + 2 & "]")
    DllStructSetData($dest_struct, 1, $dest)
    DllStructSetData($dest_struct, 1, 0, StringLen($dest)+2)
    DllStructSetData($SHFILEOPSTRUCT, $hwnd, 0)
    DllStructSetData($SHFILEOPSTRUCT, $wFunc, $FO_MOVE)
    DllStructSetData($SHFILEOPSTRUCT, $pFrom, DllStructGetPtr($source_struct))
    DllStructSetData($SHFILEOPSTRUCT, $pTo, DllStructGetPtr($dest_struct))
    DllStructSetData($SHFILEOPSTRUCT, $fFlags, $flags)
    DllStructSetData($SHFILEOPSTRUCT, $fAnyOperationsAborted, 0)
    DllStructSetData($SHFILEOPSTRUCT, $hNameMappings, 0)
    DllStructSetData($SHFILEOPSTRUCT, $lpszProgressTitle, 0)

    If (__SHFileOperation($SHFILEOPSTRUCT)) Then
        $aDllRet = DllCall("kernel32.dll", "long", "GetLastError")
        If (@error) Then
			MsgBox(16, "Error", "Error calling GetLastError")
		EndIf

        SetError($aDllRet[0])
        Return False
    ElseIf (DllStructGetData($SHFILEOPSTRUCT, $fAnyOperationsAborted)) Then
        MsgBox(16, "Error", "File/Directory Move operation aborted!")
        Return False
    EndIf

    Return True
EndFunc   ;==>_FileMoveDlg


; #FUNCTION# ====================================================================================================================
; Name...........: _FileDeleteDlg
; Description ...: Deletes a file, set of files, or directory while displaying a progress dialog.
; Syntax.........: _FileDeleteDlg($path[, flags])
; Parameters ....: $path    - The path to delete.  Wildcards are supported. A list of individual file paths or
;                             directories can be specifed by seperating each path with @LF.
;                  $flags   - Supports one or more of the following optional flags (see globals section for flag definitions):
;                             $FOF_ALLOWUNDO
;                             $FOF_FILESONLY
;                             $FOF_NOCONFIRMATION
;                             $FOF_NOERRORUI
;                             $FOF_NORECURSION
;                             $FOF_SILENT
;                             $FOF_SIMPLEPROGRESS
;                             $FOF_NO_CONNECTED_ELEMENTS  (requires IE5 or higher)
;                             $FOF_WANTNUKEWARNING  (requires IE5 or higher)
; Return values .: Success - True.
;                  Failure - False.  Raises error message dialogs for the following errors:
;                            - $path not defined or does not exist.
;                            - Could not create SHFILEOPTSTRUCT structure.
;                            - _SHFileOperation() call failed and was unable to retrieve error info using GetLastError call,
;                              otherwise sets @error to the value returned by GetLastError.
;                            - Operation aborted.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/01/2010
; Remarks .......: Modified from original code by Livewire (see AutoIt Forum).
; Related .......: _SHFileOperation, _TrimTrailingDelimiter, _FileCopyDlg, _FileMoveDlg, _FileRenameDlg
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileDeleteDlg($path, $flags = 0)
	If ((StringLen($path) == 0) Or (Not FileExists($path))) Then
		MsgBox(16, "Error", "Cannot delete! Path not defined or does not exist.")
		Return False
	EndIf

    Local $SHFILEOPSTRUCT, $path_struct
    Local $i

	$path = __TrimTrailingDelimiter($path)
    Local Enum $hwnd = 1, $wFunc, $pFrom, $pTo, $fFlags, $fAnyOperationsAborted, $hNameMappings, $lpszProgressTitle

    $SHFILEOPSTRUCT = DllStructCreate("int;uint;ptr;ptr;uint;int;ptr;ptr")
    If (@error) Then
        MsgBox(16, "ERROR", "Error creating SHFILEOPSTRUCT structure")
        Return False
    EndIf

    $path_struct = DllStructCreate("char[" & StringLen($path) + 2 & "]")
    DllStructSetData($path_struct, 1, $path)

    For $i = 1 To StringLen($path) + 2
        If (DllStructGetData($path_struct, 1, $i) == 10) Then
			DllStructSetData($path_struct, 1, 0, $i)
		EndIf
	Next

    DllStructSetData($path_struct, 1, 0, StringLen($path) + 2)
    DllStructSetData($SHFILEOPSTRUCT, $hwnd, 0)
    DllStructSetData($SHFILEOPSTRUCT, $wFunc, $FO_DELETE)
    DllStructSetData($SHFILEOPSTRUCT, $pFrom, DllStructGetPtr($path_struct))
    DllStructSetData($SHFILEOPSTRUCT, $pTo, 0)
    DllStructSetData($SHFILEOPSTRUCT, $fFlags, $flags)
    DllStructSetData($SHFILEOPSTRUCT, $fAnyOperationsAborted, 0)
    DllStructSetData($SHFILEOPSTRUCT, $hNameMappings, 0)
    DllStructSetData($SHFILEOPSTRUCT, $lpszProgressTitle, 0)

    If (__SHFileOperation($SHFILEOPSTRUCT)) Then
        Local $aDllRet = DllCall("kernel32.dll", "long", "GetLastError")
        If (@error) Then
			MsgBox(16, "Error", "Error calling GetLastError")
		EndIf

        SetError($aDllRet[0])
        Return False
    ElseIf (DllStructGetData($SHFILEOPSTRUCT, $fAnyOperationsAborted)) Then
        MsgBox(16, "Error", "File Delete operation aborted!")
        Return False
    EndIf
    Return True
EndFunc   ;==>_FileDeleteDlg


; #FUNCTION# ====================================================================================================================
; Name...........: _FileRenameDlg
; Description ...: Renames a file or directory and displays a progress dialog.
; Syntax.........: _FileRenameDlg($oldName, $newName[, flags])
; Parameters ....: $oldName  - An existing file or folder to rename.  Wildcards are supported.
;                  $newName  - The new name of the file or directory.  Target file or directory cannot already exist.
;                  $flags    - Supports one or more of the following optional flags (see globals section for flag definitions):
;                               $FOF_ALLOWUNDO
;                               $FOF_FILESONLY
;                               $FOF_NOERRORUI
;                               $FOF_NOCONFORMATION
;                               $FOF_RENAMEONCOLLISION
;                               $FOF_SILENT
;                               $FOF_SIMPLEPROGRESS
;                               $FOF_WANTMAPPINGHANDLE  (requires $FOF_RENAMECOLLISION)
; Return values .: Success - True.
;                  Failure - False.  Raises error message dialogs for the following errors:
;                            - $oldName not defined or does not exist.
;                            - $newName not defined.
;                            - Could not create SHFILEOPTSTRUCT structure.
;                            - _SHFileOperation() call failed and was unable to retrieve error info using GetLastError call,
;                              otherwise sets @error to the value returned by GetLastError.
;                            - Operation aborted.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/01/2010
; Remarks .......: Modified from original code by Livewire (see AutoIt Forum).
; Related .......: _SHFileOperation, _TrimTrailingDelimiter, _FileCopyDlg, _FileMoveDlg, _FileDeleteDlg
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileRenameDlg($oldName, $newName, $flags = 0)
    If ((StringLen($oldName) == 0) Or (Not FileExists($oldName))) Then
		MsgBox(16, "Error", "Source file name not defined or does not exist.")
		Return False
	EndIf

	If (StringLen($newName) == 0) Then
		MsgBox(16, "Error", "Target name not defined.")
		Return False
	EndIf

	Local $SHFILEOPSTRUCT, $old_name_struct, $new_name_struct
    Local $i
    Local Enum $hwnd = 1, $wFunc, $pFrom, $pTo, $fFlags, $fAnyOperationsAborted, $hNameMappings, $lpszProgressTitle

    $SHFILEOPSTRUCT = DllStructCreate("int;uint;ptr;ptr;uint;int;ptr;ptr")
    If (@error) Then
        MsgBox(4096, "ERROR", "Error creating SHFILEOPSTRUCT structure")
        Return False
    EndIf

	$oldName = __TrimTrailingDelimiter($oldName)
	$newName = __TrimTrailingDelimiter($newName)
    $old_name_struct = DllStructCreate("char[" & StringLen($oldName) + 2 & "]")
    DllStructSetData($old_name_struct, 1, $oldName)

    For $i = 1 To StringLen($oldName) + 2
        If (DllStructGetData($old_name_struct, 1, $i) == 10) Then
			DllStructSetData($old_name_struct, 1, 0, $i)
		EndIf
	Next

    DllStructSetData($old_name_struct, 1, 0, StringLen($oldName) + 2)

    $new_name_struct = DllStructCreate("char[" & StringLen($newName) + 2 & "]")
    DllStructSetData($new_name_struct, 1, $newName)
    DllStructSetData($new_name_struct, 1, 0, StringLen($newName) + 2)
    DllStructSetData($SHFILEOPSTRUCT, $hwnd, 0)
    DllStructSetData($SHFILEOPSTRUCT, $wFunc, $FO_RENAME)
    DllStructSetData($SHFILEOPSTRUCT, $pFrom, DllStructGetPtr($old_name_struct))
    DllStructSetData($SHFILEOPSTRUCT, $pTo, DllStructGetPtr($new_name_struct))
    DllStructSetData($SHFILEOPSTRUCT, $fFlags, $flags)
    DllStructSetData($SHFILEOPSTRUCT, $fAnyOperationsAborted, 0)
    DllStructSetData($SHFILEOPSTRUCT, $hNameMappings, 0)
    DllStructSetData($SHFILEOPSTRUCT, $lpszProgressTitle, 0)

    If (__SHFileOperation($SHFILEOPSTRUCT)) Then
        Local $aDllRet = DllCall("kernel32.dll", "long", "GetLastError")
        If (@error) Then
			MsgBox(16, "Error", "Error calling GetLastError")
		EndIf

        SetError($aDllRet[0])
        Return False
    ElseIf (DllStructGetData($SHFILEOPSTRUCT, $fAnyOperationsAborted)) Then
        MsgBox(16, "Error", "File Rename operation aborted!")
        Return False
    EndIf
    Return True
EndFunc   ;==>_FileRenameDlg


; #FUNCTION# ====================================================================================================================
; Name...........: _FileCheckDirEmpty
; Description ...: Checks to see if a specified directory is empty.
; Syntax.........: _FileCheckDirEmpty($path)
; Parameters ....: $path   - A full directory path (MUST be a directory).
; Return values .: Success - True. Directory is empty.
;                  Failure - False.  Sets @error to 1 if $Path is not a directory or does not exist.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/01/2010
; Remarks .......:
; Related .......: _FileIsDir
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileCheckDirEmpty($path)
	Local $hSearchHandle
	Local $bDirEmpty = False
	If (Not _FileIsDir($path)) Then
		Return SetError(1, 0, $bDirEmpty)
	EndIf

	If (StringRight($path, 1) <> "\") Then
		$path = $path & "\"
	EndIf

	$hSearchHandle = FileFindFirstFile($Path & "*.*")
	If ((@error) Or ($hSearchHandle == -1)) Then
		$bDirEmpty = True
	EndIf

	FileClose($hSearchHandle)
	Return $bDirEmpty
EndFunc   ;==>_FileCheckDirEmpty


; #FUNCTION# ====================================================================================================================
; Name...........: _FileGetDirFromPath
; Description ...: Returns the directory where a specified file (full path) exists.
; Syntax.........: _FileGetDirFromPath($path)
; Parameters ....: $path   - A full path to a specific file.
; Return values .: Success - Returns the path of the file without the filename. If the specified path does not already contain
;                  the filename, then the value will simply be returned unmolested.
;                  Failure - Returns and empty string and sets @error to:
;                  |1 - The value specified for path is an empty string or not a string at all.
;                  |2 - The specified path does not contain a valid path separator.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/17/2010
; Remarks .......:
; Related .......: _FileIsDir, _FileGetFileName
; Link ..........;
; Example .......; Include "FileLib.au3"
;                  $dir = _FileGetDirFromPath("C:\temp\myfile.txt")  ;$dir = "C:\temp\"
; ===============================================================================================================================
Func _FileGetDirFromPath($path)
	Local $sRet = ""
	If ((Not IsString($path)) Or (StringLen($path) == 0)) Then
		Return SetError(1, 0, $sRet)
	EndIf

	Local $sTempResult = _FileGetFileName($path)
	If (@error) Then
		Switch @error
			Case 2
				$sRet = $sTempResult
			Case 3
				SetError(2)
		EndSwitch
	EndIf
	$sRet = StringLeft($path, StringLen($path) - StringLen($sTempResult))
	Return $sRet
EndFunc   ;==>_FileGetDirFromPath


; #FUNCTION# ====================================================================================================================
; Name...........: _FileTruncateName
; Description ...: Returns a truncated file name in 8.3 (8 character name, 3 character extension) DOS style format.
; Syntax.........: _FileTruncateName($fileName)
; Parameters ....: $fileName   - A full path to a specific file.
; Return values .: Success - The truncated file name string.
;                  Failure - Returns the parameter and sets @error to:
;                            |1 - $fileName does not exist, is a directory, or not a fully-qualified file path.
;                            |2 - Unable to extract container directory from path.
;                            |3 - Exceeded 100 attempts to name the file to something that does not already exist in the path.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/01/2010
; Remarks .......: Does not truncate the full path.  Returns the truncated file name only.  The function does not actually rename
;                  the file, but it could be used in conjunction with FileMove() or _FileRenameDlg() to do so.
; Related .......: _FileGetFileName, _FileGetDirFromPath
; Link ..........;
; Example .......; $Result = _FileTruncateName("C:\temp\a long file name.txt")  ;$Result = "ALONGFIL.TXT"
; ===============================================================================================================================
Func _FileTruncateName($fileName)
	Local $justName = _FileGetFileName($fileName)		;Get just the name and discard the rest of the path.
	If ((@error) Or (StringLen($justName) == 0)) Then
		Return SetError(1, 0, $FileName)
	EndIf

	Local $justDir = _FileGetDirFromPath($FileName)     ;Get the directory where this file is located.
	If ((@error) Or (StringLen($justDir) == 0)) Then
		Return SetError(2, 0, $FileName)
	EndIf

	$justName = StringUpper($justName)					;Convert to uppercase.
	$justName = StringStripWS($justName, 8)             ;Strip all whitespaces.
	Local $i = 0
	Local $nPos = 0
	Local $newName = ""

	;Remove extra dots in the name, leaving only the file extension.
	Local $aSegments = StringSplit($justName, ".")
	If ((Not @error) And ($aSegments[0] > 0)) Then
		If ($aSegments[0] > 1) Then
			For $i = 1 To $aSegments[0]
				If ($i == $aSegments[0]) Then
					$newName = $newName & "." & $aSegments[$i]
					ExitLoop
				Else
					$newName &= $aSegments[$i]
				EndIf
			Next
		EndIf
	EndIf

	;Replace any and all illegal characters with underscores.
	Local $searchPattern = '+,;=[]\/:*?"<>|'
	Local $aChars = StringSplit($searchPattern, "")		;Build illegal character array.
	For $i = 1 To $aChars[0]
		$nPos = StringInStr($newName, $aChars[$i], 0, 1)
		If ($nPos == 0) Then
			ContinueLoop
		EndIf
		$newName = StringReplace($newName, $aChars[$i], "_", 0, 0)
	Next

	;Split name into segments again if a file extension exists, so we can evaluate them and then piece them back together later.
	Local $name = ""
	Local $extension = ""
	$aSegments = StringSplit($newName, ".")
	If ((Not @error) And ($aSegments[0] == 2)) Then
		$name = $aSegments[1]
		$extension = $aSegments[2]
	Else
		$name = $newName
	EndIf

	;If name is longer than 8 characters, discard the rest of the string after the 8th character.
	If (StringLen($name) > 8) Then
		$name = StringLeft($name, 8)
	EndIf

	;If extension exists and is longer than 3 characters, discard the rest of the string after the 3rd character.
	;Then piece the name and extension strings back together.
	If (StringLen($extension) > 0) Then
		$extension = StringLeft($extension, 3)
		$name &= "." & $extension
	EndIf

	;Check to see if the newly named file already exists.  If so, truncate an additional 2 characters from the name
	;and substitute with digits from 00 to 99. i.e. ismyfile.txt becomes ISMYFI00.TXT, or ISMYFI00.TXT becomes ISMYFI01.TXT, etc.
	$i = 0
	While FileExists($justDir & $name)
		;If there are 99 files with the same name, just give up and return an error.
		If ($i == 99) Then
			Return SetError(3, 0, $FileName)
		EndIf

		$aSegments = StringSplit($name, ".")
		If ($aSegments[0] == 2) Then
			$name = $aSegments[1]
		EndIf

		$name = StringLeft($name, 6)
		If ($i < 10) Then
			$name &= "0" & $i
		Else
			$name &= $i
		EndIf

		;If extension exists, piece name string and extension back together.
		If ($aSegments[0] == 2) Then
			$name &= "." & $aSegments[2]
		EndIf
		$i += 1
	WEnd

	;Return name only, without truncating the rest of the path.
	;This gives you the ability to truncate all names in a directory without truncating the directory. (Good for mass rename)
	Return $name
EndFunc   ;==>_FileTruncateName


; #FUNCTION# ====================================================================================================================
; Name...........: _FileGetSystemDrive
; Description ...: Gets the drive letter of the partition Windows is installed on.
; Syntax.........: _FileGetSystemDrive()
; Parameters ....: None.
; Return values .: Success - Returns the logical drive letter of the system partition. Example: "C:".
;                  Failure - Returns "" and sets @error = 1.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/01/2010
; Remarks .......: This function simply takes the return of the @SystemDir macro and splits the path, then returns the first
;                  element in the array.  Alternately, by setting Opt('ExpandEnvStrings', 1), you could just return the value of
;                  %SystemDrive%.  Since the value of this variable could be modified, I find my method to a little more intrinsic.
; Related .......:
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileGetSystemDrive()
	Local $arrStrings = StringSplit(@SystemDir, "\")
	If ($arrStrings[0] > 0) Then
		Return $arrStrings[1]
	EndIf
	Return SetError(1, 0, "")
EndFunc


; #FUNCTION# ====================================================================================================================
; Name...........: _FileGetProfilesDir
; Description ...: Gets the local profiles directory from the registry.
; Syntax.........: _FileGetProfilesDir()
; Parameters ....: None.
; Return values .: Success - Returns the full path of the local profiles directory.
;                  Failure - Returns "" and sets @error = 1.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/01/2010
; Remarks .......:
; Related .......: RegRead
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileGetProfilesDir()
	Local $key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	Local $val = "ProfilesDirectory"
	Local $result = RegRead($key, $val)
	If ((@error) Or (StringLen($result) == 0)) Then
		SetError(1)
	EndIf

	$key = 0
	$val = 0
	Return $result
EndFunc


; #FUNCTION# ====================================================================================================================
; Name...........: _FileRemoveDirIfEmpty
; Description ...: Removes a specified directory only if it is empty.
; Syntax.........: _FileRemoveDirIfEmpty($dirPath)
; Parameters ....: $dirPath - The directoy path to check and remove.
; Return values .: Success - Returns True.
;                  Failure - Returns False if there is an error removing the directory, directory does not exist, or directory
;                            is not empty.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/01/2010
; Remarks .......:
; Related .......: _FileCheckDirEmpty, DirRemove
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileRemoveDirIfEmpty($dirPath)
	If (_FileCheckDirEmpty($dirPath)) Then
		If (DirRemove($dirPath, False) == 1) Then
			Return True
		EndIf
	EndIf
	Return False
EndFunc   ;==>_FileRemoveDirIfEmpty


; #FUNCTION# ====================================================================================================================
; Name...........: _DirIsSymlink
; Description ...: Checks to see if the specified path is a symlink (reparse point).
; Syntax.........: _DirIsSymlink($dirPath)
; Parameters ....: $dirPath - The directoy path to check.
; Return values .: Success - Returns True.
;                  Failure - Returns False if path is empty, does not exist, or is not a symlink.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 05/13/2014
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _DirIsSymlink($dirPath)
	If ((StringLen($dirPath) == 0) Or (Not FileExists($dirPath))) Then
		Return False
	EndIf

	Local $rc = DllCall('kernel32.dll', 'Int', 'GetFileAttributes', 'str', $dirPath)
	If (IsArray($rc)) Then
		If (BitAND($rc[0], $FILE_ATTRIBUTE_REPARSE_POINT) == $FILE_ATTRIBUTE_REPARSE_POINT) Then
			Return True
		EndIf
	EndIf
	Return False
EndFunc   ;==>_DirIsSymlink


; #FUNCTION# ====================================================================================================================
; Name...........: _DirMakeSymlink
; Description ...: Creates a symlink to specified target directory.
; Syntax.........: _DirMakeSymlink($linkPath, $targetPath)
; Parameters ....: $linkPath   - The path to the link to be created.
;                  $targetPath - The path to the actual directory to link to.
; Return values .: Success - Returns True.
;                  Failure - Returns False and sets @error to one of the following:
;                            |1 - $linkPath is an empty string.
;                            |2 - $targetPath is an empty string.
;                            |3 - $linkPath already exists and is a symlink.
;                            |4 - $targetPath is not a directory or does not exist.
;                            |5 - Failed to create the link at the specified path.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 05/14/2014
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _DirMakeSymlink($linkPath, $targetPath)
	If (StringLen($linkPath) == 0) Then
		Return SetError(1, "", False)
	EndIf

	If (StringLen($targetPath) == 0) Then
		Return SetError(2, "", False)
	EndIf

	If (_DirIsSymlink($linkPath)) Then
		Return SetError(3, "", False)
	EndIf

	If (Not _FileIsDir($targetPath)) Then
		Return SetError(4, "", False)
	EndIf

	Local $result = RunWait(@ComSpec & ' /c mklink /d "' & $linkPath & '" "' & $targetPath, "", @SW_HIDE)
	If (($result <> 0) Or (@error)) Then
		Return SetError(5, "", False)
	EndIf
	Return True
EndFunc   ;==>_DirMakeSymlink


; #INTERNAL USE ONLY# ============================================================================================================
; Name ..........: __SHFileOperation
; Description ...: Asks the shell to perform the specified file operation.
; Syntax ........: __SHFileOperation(ByRef $lpFileOp)
; Parameters ....: $lpFileOp - The file operation to perform.  Use one of the following operation constants:
;                  $FO_COPY
;                  $FO_MOVE
;                  $FO_RENAME
;                  $FO_DELETE
; Return values .: If DllCall() succeeds, then the code from the DLL call is returned.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified ......: 12/01/2010
; Remarks .......: Modified from original code by Livewire (see AutoIt Forum).
; Related .......: See global constants section for definition of file operation constants.  Operation constants cannot be combined.
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func __SHFileOperation(ByRef $lpFileOp)
    Local $aDllRet = DllCall("shell32.dll", "int", "SHFileOperation", "ptr", DllStructGetPtr($lpFileOp))
    If (Not @error) Then
		Return $aDllRet[0]
	EndIf
EndFunc   ;==>_SHFileOperation


; #INTERNAL USE ONLY# ============================================================================================================
; Name ..........: __TrimTrailingDelimiter
; Description ...: Trims the trailing backslash (path delimiter) from the specified string (if one exists).
; Syntax ........: __TrimTrailingDelimiter($string)
; Parameters ....: $string - The string to check.
; Return values .: Success - Returns the original string minus the trailing backslash.
;                  Failure - If the original string is not greater than one character or if the string does not contain a trailing
;                            path delimiter ( "\" ) then the original string is returned.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func __TrimTrailingDelimiter($string)
	If ((StringLen($string) > 1) And (StringRight($string, 1) == "\")) Then
		Return StringLeft($string, StringLen($string) - 1)
	EndIf
	Return $string
EndFunc   ;==>_TrimTrailingDelimiter
