#include-once
Opt('MustDeclareVars', 1)

; #INDEX# =======================================================================================================================
; Name ...........: FileLib.au3
; AutoIt Version .: 3.2.12.0
; Author .........: Chris Brunner <CyrusBuilt at gmail dot com>
; Language .......: English
; Library Version : 1.0.0.8
; Modified .......: 01/15/09
; License ........: GPL v2
; Description ....: A collection of additional functions for manipulating files, directories, and disks.
; Remarks ........: Thanks to Livewire for code that some of these functions are based on.
; Dependencies ...: Microsoft Windows 2000/XP/2003.  Not yet tested with Windows Vista.
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
; ===============================================================================================================================


; #INTERNAL_USE_ONLY#============================================================================================================
;
; _SHFileOperation
; _TrimTrailingDelimiter
; ===============================================================================================================================


; #GLOBALS ======================================================================================================================
Global Const $CB_FILE_LIB_VER           = "1.0.0.8"
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
; Syntax.........: _FileExtension($Path)
; Parameters ....: $Path - The name or full path of the file to get the extension of.
; Return values .: Success - Returns the extension of the file specified without the dot (".").
;                  Failure - If the specified name or path does not have an extension it will return an empty string.  If $Path is empty
;                            or undefined then an empty string is returned and @error is set to 1.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/15/08
; Remarks .......: 
; Related .......: 
; Link ..........;
; Example .......; Local $Extension = _FileExtension("My file.txt")  ;$Extension = "txt"
; ===============================================================================================================================
Func _FileExtension($Path)
	If Not IsDeclared("Path") Or StringLen($Path) = 0 Then Return SetError(1, 0, "")
	Local $arrStrings = StringSplit($Path, ".")
	If @error Or $arrStrings[0] = 0 Then Return ""
	Local $i, $sExtension
	
	For $i = 1 To $arrStrings[0]
		If $i = $arrStrings[0] Then
			$sExtension = $arrStrings[$i]
			ExitLoop
		EndIf
	Next
	
	$i = 0
	$arrStrings = 0
	$Path = 0
	Return $sExtension
EndFunc   ;==>_FileExtension


; #FUNCTION# ====================================================================================================================
; Name...........: _FileCompare
; Description ...: Compares 2 files to see if they are the same.
; Syntax.........: _FileCompare($File1, $File2)
; Parameters ....: $File1 - First file to compare.
;                  $File2 - Second file to compare.
; Return values .: True  - Files are the same.
;                  False - Files are different.
;                  Sets @error to:
;                  |1 - $File1 not defined, is a directory, or does not exist.
;                  |2 - $File2 not defined, is a directory, or does not exist.
;                  |3 - Files did not match or 'FC' command error and sets @extended to the ERRORLEVEL returned by 'FC'.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/15/08
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileCompare($File1, $File2)
	If Not IsDeclared("File1") Or StringLen($File1) = 0 Or Not FileExists($File1) Or _FileIsDir($File1) Then Return SetError(1, 0, False)
	If Not IsDeclared("File2") Or StringLen($File2) = 0 Or Not FileExists($File2) Or _FileIsDir($File2) Then Return SetError(2, 0, False)
	Local $nRet = RunWait(@ComSpec & ' /c fc /b /c "' & $File1 & '" "' & $File2 & '"', @ScriptDir, @SW_HIDE) 
	
	If $nRet = 0 Then
		Return True
	Else
		Return SetError(3, $nRet, False)
	EndIf
EndFunc   ;==>_FileCompare


; #FUNCTION# ====================================================================================================================
; Name...........: _DiskExist
; Description ...: Checks to see if a drive letter exists.
; Syntax.........: _DiskExist($Drive)
; Parameters ....: $Drive - A logical drive letter followed by a colon. i.e.: "D:"
; Return values .: True  - Drive letter exists and is valid.
;                  False - Drive does not exist.  Sets @error = 1 if parameter undefined.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......:
; Remarks .......:
; Related .......: _DiskInfoToArray, _DiskIsNetwork
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _DiskExist($Drive)
	If Not IsDeclared("Drive") Or StringLen($Drive) = 0 Or StringLen($Drive) > 2 Then
		Return SetError(1, 0, False)
	EndIf
	
	$Drive = StringStripWS($Drive, 8)
	If DriveStatus(StringUpper($Drive)) = "INVALID" Then Return False
	Return True
EndFunc   ;==>_DiskExist


; #FUNCTION# ====================================================================================================================
; Name...........: _DiskInfoToArray
; Description ...: Returns an array containing info about a specified drive letter.
; Syntax.........: _DiskInfoToArray($Drive)
; Parameters ....: $Drive - A logical drive letter followed by a colon. i.e.: "D:"
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
; Modified.......:
; Remarks .......:
; Related .......: _DiskExist, _DiskIsNetwork
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _DiskInfoToArray($Drive)
	Local $aInfo[8]
	Local $nFree
	Local $nUsed
	Local $nTotal
	
	$aInfo[0] = "INVAlID"
	If _DiskExist($Drive) = False Then Return SetError(1, 0, $aInfo)
	
	$Drive = StringUpper($Drive) & "\"
	$aInfo[0] = DriveStatus($Drive)
	$aInfo[1] = DriveGetLabel($Drive)
	$aInfo[2] = DriveGetType($Drive)
	$aInfo[3] = DriveGetSerial($Drive)
	$aInfo[4] = DriveGetFileSystem($Drive)
	
	$nFree = Floor(DriveSpaceFree($Drive))
	$nTotal	= Floor(DriveSpaceTotal($Drive))
	$nUsed = ($nTotal - $nFree)
	
	$aInfo[5] = $nFree
	$aInfo[6] = $nUsed
	$aInfo[7] = $nTotal
	Return $aInfo
EndFunc   ;==>_DiskInfoToArray


; #FUNCTION# ====================================================================================================================
; Name...........: _DiskIsNetwork
; Description ...: Determines if a specified drive letter is a network drive.
; Syntax.........: _DiskIsNetwork($Drive)
; Parameters ....: $Drive - A logical drive letter followed by a colon. i.e.: "D:"
; Return values .: True - Specified drive letter is a network drive.
;                  False - Specified drive is not a network drive.
;                  Sets @error:
;                        |1 - if drive does not exist
;                        |2 - if unable to determine drive type.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......:
; Remarks .......:
; Related .......: _DiskExist, _DiskInfoToArray
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _DiskIsNetwork($Drive)
	If _DiskExist($Drive) = False Then Return SetError(1, 0, False)
	$Drive = StringUpper($Drive) & "\"
	
	If DriveGetType($Drive) = "Network" Then
		Return True
	ElseIf @error Then
		Return SetError(2, 0, False)
	Else
		Return False
	EndIf
EndFunc   ;==>_DiskIsNetwork


; #FUNCTION# ====================================================================================================================
; Name...........: _FileIsDir
; Description ...: Determines whether or not the supplied path is a directory.
; Syntax.........: _FileIsDir($Path)
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
Func _FileIsDir($Path)
	If FileExists($Path) Then
		If StringInStr(FileGetAttrib($Path), "D", 1) > 0 Then Return True
	Else
		Return SetError(1, 0, False)
	EndIf
	
	Return False
EndFunc   ;==>_FileIsDir


; #FUNCTION# ====================================================================================================================
; Name...........: _FileRecurseBuildList
; Description ...: Performs file recursion on the specified directory and builds an array containing the full path to every file
;                  the specified directory and every subdirectory.
; Syntax.........: _FileRecurseBuildList($Path, ByRef $aFiles)
; Parameters ....: $Path   - The path to check.
;                  $aFiles - An array to contain the file paths and the count.
; Return values .: Success - $aFiles - The resulting path array with the path count at element 0.  If no file or directory exists within
;                  the parent path specified then $aFiles will be returned with only element 0 and it's value will be 0.  
;                  Failure - If the parent path specified does not exist, the same will be true but it will also set @error to 1.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/27/08
; Remarks .......: Only returns FILE paths.  If $Path contains any empty directories, then they will not be included in the list.
; Related .......: _FileIsDir
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileRecurseBuildList($Path, ByRef $aFiles)
	Local $sSearch
	Local $sFile
	Local $sFileFullPath
    
	If StringLen($Path) = 0 Or Not FileExists($Path) Then Return SetError(1, 0, $aFiles)
	If Not IsArray($aFiles) Then SetError(1, 0, $aFiles)
	If StringRight($Path, 1) = "\" Then $Path = StringLeft($Path, StringLen($path) - 1)
	$sSearch = FileFindFirstFile($Path & "\*.*")
	
	While 1
		If $sSearch = -1 Then ExitLoop
		$sFile = FileFindNextFile($sSearch)
		If @error Then ExitLoop
		$sFileFullPath = $Path & "\" & $sFile
		
		If _FileIsDir($sFileFullPath) = True Then
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
; Syntax.........: _FileGetFileName($Path)
; Parameters ....: $Path   - A full file path (drive, directory, filename).
; Return values .: Success - The file name string (minus the rest of the path).
;                  Failure - Returns a blank string ( "" ) and sets @error to:
;                            |1 - Specified path does not exist.
;                            |2 - Specified path is a directory.
;                            |3 - Could not split specified path into an array (no path separator detected).  This will occur
;                                 if *only* a file name is specified and not an actual path.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 
; Remarks .......: 
; Related .......: _FileIsDir
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileGetFileName($Path)
	Local $bIsDir = _FileIsDir($Path)
	If @error Then Return SetError(1, 0, "")
	If $bIsDir = True Then Return SetError(2, 0, "")
	$bIsDir = 0
	Local $aPathStrings = StringSplit($Path, "\")
	
	If @error Or $aPathStrings[0] = 0 Then
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
; Modified.......: 12/26/08
; Remarks .......: Modified from original code by Livewire (see AutoIt Forum).
; Related .......: _SHFileOperation, _TrimTrailingDelimiter, _FileMoveDlg, _FileRenameDlg, _FileDeleteDlg
; Link ..........; 
; Example .......;
; ===============================================================================================================================
Func _FileCopyDlg($source, $dest, $flags = 0)
    Local $SHFILEOPSTRUCT, $source_struct, $dest_struct
    Local $i, $aDllRet
	
	If StringLen($source) = 0 Then
		MsgBox(16, "Error", "Source file or directory not defined or does not exist.")
		Return False
	EndIf
	
	If StringLen($dest) = 0 Then
		MsgBox(16, "Error", "Destination not defined.")
		Return False
	EndIf
   
	$source = _TrimTrailingDelimiter($source)
	$dest = _TrimTrailingDelimiter($dest)
    Local Enum $hwnd = 1, $wFunc, $pFrom, $pTo, $fFlags, $fAnyOperationsAborted, $hNameMappings, $lpszProgressTitle
    $SHFILEOPSTRUCT = DllStructCreate("int;uint;ptr;ptr;uint;int;ptr;ptr")
	
    If @error Then
        MsgBox(16, "ERROR", "Error creating SHFILEOPSTRUCT structure")
        Return False
    EndIf
   
    $source_struct = DllStructCreate("char[" & StringLen($source) + 2 & "]")
    DllStructSetData($source_struct, 1, $source)
	
    For $i = 1 To StringLen($source) + 2
        If DllStructGetData($source_struct, 1, $i) = 10 Then DllStructSetData($source_struct, 1, 0, $i)
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
   
    If _SHFileOperation($SHFILEOPSTRUCT) Then
        $aDllRet = DllCall("kernel32.dll", "long", "GetLastError")
        If @error Then MsgBox(16, "Error", "Error calling GetLastError")
        SetError($aDllRet[0])
        Return False
    ElseIf DllStructGetData($SHFILEOPSTRUCT,$fAnyOperationsAborted) Then
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
; Modified.......: 12/26/08
; Remarks .......: Modified from original code by Livewire (see AutoIt Forum).
; Related .......: _SHFileOperation, _TrimTrailingDelimiter, _FileCopyDlg, _FileRenameDlg, _FileDeleteDlg
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileMoveDlg($source, $dest, $flags = 0)
    Local $SHFILEOPSTRUCT, $source_struct, $dest_struct
    Local $i, $aDllRet
	
	If StringLen($source) = 0 Then
		MsgBox(16, "Error", "Source file or directory not defined or does not exist.")
		Return False
	EndIf
	
	If StringLen($dest) = 0 Then
		MsgBox(16, "Error", "Destination not defined.")
		Return False
	EndIf
   
    $source = _TrimTrailingDelimiter($source)
	$dest = _TrimTrailingDelimiter($dest)
    Local Enum $hwnd = 1, $wFunc, $pFrom, $pTo, $fFlags, $fAnyOperationsAborted, $hNameMappings, $lpszProgressTitle
    $SHFILEOPSTRUCT = DllStructCreate("int;uint;ptr;ptr;uint;int;ptr;ptr")
	
    If @error Then
        MsgBox(4096, "ERROR", "Error creating SHFILEOPSTRUCT structure!")
        Return False
    EndIf
   
    $source_struct = DllStructCreate("char[" & StringLen($source)+2 & "]")
    DllStructSetData($source_struct, 1, $source)
	
    For $i = 1 To StringLen($source) + 2
        If DllStructGetData($source_struct, 1, $i) = 10 Then DllStructSetData($source_struct, 1, 0, $i)
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
   
    If _SHFileOperation($SHFILEOPSTRUCT) Then
        $aDllRet = DllCall("kernel32.dll", "long", "GetLastError")
        If @error Then MsgBox(16, "Error", "Error calling GetLastError")
        SetError($aDllRet[0])
        Return False
    ElseIf DllStructGetData($SHFILEOPSTRUCT, $fAnyOperationsAborted) Then
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
; Modified.......: 12/26/08
; Remarks .......: Modified from original code by Livewire (see AutoIt Forum).
; Related .......: _SHFileOperation, _TrimTrailingDelimiter, _FileCopyDlg, _FileMoveDlg, _FileRenameDlg
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileDeleteDlg($path, $flags = 0)
	If StringLen($path) = 0 Or Not FileExists($path) Then
		MsgBox(16, "Error", "Cannot delete! Path not defined or does not exist.")
		Return False
	EndIf
	
    Local $SHFILEOPSTRUCT, $path_struct
    Local $nError = 0
    Local $i
   
	$path = _TrimTrailingDelimiter($path)
    Local Enum $hwnd = 1, $wFunc, $pFrom, $pTo, $fFlags, $fAnyOperationsAborted, $hNameMappings, $lpszProgressTitle
    $SHFILEOPSTRUCT = DllStructCreate("int;uint;ptr;ptr;uint;int;ptr;ptr")
	
    If @error Then
        MsgBox(16, "ERROR", "Error creating SHFILEOPSTRUCT structure")
        Return False
    EndIf
   
    $path_struct = DllStructCreate("char[" & StringLen($path) + 2 & "]")
    DllStructSetData($path_struct, 1, $path)
	
    For $i = 1 To StringLen($path) + 2
        If DllStructGetData($path_struct, 1, $i) = 10 Then DllStructSetData($path_struct, 1, 0, $i)
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
   
    If _SHFileOperation($SHFILEOPSTRUCT) Then
        Local $aDllRet = DllCall("kernel32.dll", "long", "GetLastError")
        If @error Then MsgBox(16, "Error", "Error calling GetLastError")
        SetError($aDllRet[0])
        Return False
    ElseIf DllStructGetData($SHFILEOPSTRUCT,$fAnyOperationsAborted) Then
        MsgBox(16, "Error", "File Delete operation aborted!")
        Return False
    EndIf
	
    Return True
EndFunc   ;==>_FileDeleteDlg


; #FUNCTION# ====================================================================================================================
; Name...........: _FileRenameDlg
; Description ...: Renames a file or directory and displays a progress dialog.
; Syntax.........: _FileRenameDlg($old_name, $new_name[, flags])
; Parameters ....: $old_name  - An existing file or folder to rename.  Wildcards are supported.
;                  $new_name  - The new name of the file or directory.  Target file or directory cannot already exist.
;                  $flags     - Supports one or more of the following optional flags (see globals section for flag definitions):
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
;                            - $old_name not defined or does not exist.
;                            - $new_name not defined.
;                            - Could not create SHFILEOPTSTRUCT structure.
;                            - _SHFileOperation() call failed and was unable to retrieve error info using GetLastError call,
;                              otherwise sets @error to the value returned by GetLastError.
;                            - Operation aborted.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 12/27/08
; Remarks .......: Modified from original code by Livewire (see AutoIt Forum).
; Related .......: _SHFileOperation, _TrimTrailingDelimiter, _FileCopyDlg, _FileMoveDlg, _FileDeleteDlg
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileRenameDlg($old_name, $new_name, $flags = 0)
    If StringLen($old_name) = 0 Or Not FileExists($old_name) Then
		MsgBox(16, "Error", "Source file name not defined or does not exist.")
		Return False
	EndIf
	
	If StringLen($new_name) = 0 Then
		MsgBox(16, "Error", "Target name not defined.")
		Return False
	EndIf
	
	Local $SHFILEOPSTRUCT, $old_name_struct, $new_name_struct
    Local $nError = 0
    Local $i
   
    Local Enum $hwnd = 1, $wFunc, $pFrom, $pTo, $fFlags, $fAnyOperationsAborted, $hNameMappings, $lpszProgressTitle
    $SHFILEOPSTRUCT = DllStructCreate("int;uint;ptr;ptr;uint;int;ptr;ptr")
	
    If @error Then
        MsgBox(4096, "ERROR", "Error creating SHFILEOPSTRUCT structure")
        Return False
    EndIf
   
	$old_name = _TrimTrailingDelimiter($old_name)
	$new_name = _TrimTrailingDelimiter($new_name)
    $old_name_struct = DllStructCreate("char[" & StringLen($old_name) + 2 & "]")
    DllStructSetData($old_name_struct, 1, $old_name)
	
    For $i = 1 To StringLen($old_name) + 2
        If DllStructGetData($old_name_struct, 1, $i) = 10 Then DllStructSetData($old_name_struct, 1, 0, $i)
	Next
	
    DllStructSetData($old_name_struct, 1, 0, StringLen($old_name) + 2)

    $new_name_struct = DllStructCreate("char[" & StringLen($new_name) + 2 & "]")
    DllStructSetData($new_name_struct, 1, $new_name)
    DllStructSetData($new_name_struct, 1, 0, StringLen($new_name) + 2)
   
    DllStructSetData($SHFILEOPSTRUCT, $hwnd, 0)
    DllStructSetData($SHFILEOPSTRUCT, $wFunc, $FO_RENAME)
    DllStructSetData($SHFILEOPSTRUCT, $pFrom, DllStructGetPtr($old_name_struct))
    DllStructSetData($SHFILEOPSTRUCT, $pTo, DllStructGetPtr($new_name_struct))
    DllStructSetData($SHFILEOPSTRUCT, $fFlags, $flags)
    DllStructSetData($SHFILEOPSTRUCT, $fAnyOperationsAborted, 0)
    DllStructSetData($SHFILEOPSTRUCT, $hNameMappings, 0)
    DllStructSetData($SHFILEOPSTRUCT, $lpszProgressTitle, 0)

    If _SHFileOperation($SHFILEOPSTRUCT) Then
        Local $aDllRet = DllCall("kernel32.dll", "long", "GetLastError")
        If @error Then MsgBox(16, "Error", "Error calling GetLastError")
        SetError($aDllRet[0])
        Return False
    ElseIf DllStructGetData($SHFILEOPSTRUCT, $fAnyOperationsAborted) Then
        MsgBox(16, "Error", "File Rename operation aborted!")
        Return False
    EndIf
	
    Return True
EndFunc   ;==>_FileRenameDlg


; #FUNCTION# ====================================================================================================================
; Name...........: _FileCheckDirEmpty
; Description ...: Checks to see if a specified directory is empty.
; Syntax.........: _FileCheckDirEmpty($Path)
; Parameters ....: $Path   - A full directory path (MUST be a directory).
; Return values .: Success - True. Directory is empty.
;                  Failure - False.  Sets @error to 1 if $Path is not a directory or does not exist.  Sets @error to 2 if unable
;                            to establish a search handle for $Path.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 
; Remarks .......: 
; Related .......: _FileIsDir
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _FileCheckDirEmpty($Path)
	Local $hSearchHandle
	Local $bDirEmpty = False
	If _FileIsDir($Path) = False Then Return SetError(1, 0, $bDirEmpty)
	If StringRight($Path, 1) <> "\" Then $Path = $Path & "\"
	$hSearchHandle = FileFindFirstFile($Path & "*.*")
	If @error Then $bDirEmpty = True
	If $hSearchHandle = -1 Then Return SetError(2, 0, False)
	FileClose($hSearchHandle)
	Return $bDirEmpty
EndFunc   ;==>_FileCheckDirEmpty


; #FUNCTION# ====================================================================================================================
; Name...........: _FileGetDirFromPath
; Description ...: Returns the directory where a specified file (full path) exists.
; Syntax.........: _FileGetDirFromPath($Path)
; Parameters ....: $Path   - A full path to a specific file.
; Return values .: Success - Returns the path of the file without the filename.
;                  Failure - Sets @error = 1 and returns the parameter if the path does not exist or is a directory.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 
; Remarks .......: 
; Related .......: _FileIsDir, _FileGetFileName
; Link ..........;
; Example .......; Include "FileLib.au3"
;                  $dir = _FileGetDirFromPath("C:\temp\myfile.txt")  ;$dir = "C:\temp\"
; ===============================================================================================================================
Func _FileGetDirFromPath($Path)
	If _FileIsDir($Path) Then Return SetError(1, 0, $Path)
	Return StringLeft($Path, StringLen($Path) - StringLen(_FileGetFileName($Path)))
EndFunc   ;==>_FileGetDirFromPath


; #FUNCTION# ====================================================================================================================
; Name...........: _FileTruncateName
; Description ...: Returns a truncated file name in 8.3 (8 character name, 3 character extension) DOS style format.
; Syntax.........: _FileTruncateName($FileName)
; Parameters ....: $FileName   - A full path to a specific file.
; Return values .: Success - The truncated file name string.
;                  Failure - Returns the parameter and sets @error to:
;                            |1 - $FileName does not exist, is a directory, or not a fully-qualified file path.
;                            |2 - Unable to extract container directory from path.
;                            |3 - Exceeded 100 attempts to name the file to something that does not already exist in the path.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified.......: 
; Remarks .......: Does not truncate the full path.  Returns the truncated file name only.  The function does not actually rename
;                  the file, but it could be used in conjunction with FileMove() or _FileRenameDlg() to do so.
; Related .......: _FileGetFileName, _FileGetDirFromPath
; Link ..........;
; Example .......; $Result = _FileTruncateName("C:\temp\a long file name.txt")  ;$Result = "ALONGFIL.TXT"
; ===============================================================================================================================
Func _FileTruncateName($FileName)
	Local $justName = _FileGetFileName($FileName)		;Get just the name and discard the rest of the path.
	If @error Or StringLen($justName) = 0 Then Return SetError(1, 0, $FileName)	
	Local $justDir = _FileGetDirFromPath($FileName)     ;Get the directory where this file is located.
	If @error Or StringLen($justDir) = 0 Then Return SetError(2, 0, $FileName)
	$justName = StringUpper($justName)					;Convert to uppercase.
	$justName = StringStripWS($justName, 8)             ;Strip all whitespaces.
	Local $i, $nPos, $newName
	Local $aSegments = StringSplit($justName, ".")
	
	;Remove extra dots in the name, leaving only the file extension.
	If Not @error And $aSegments[0] > 0 Then
		If $aSegments[0] > 1 Then
			For $i = 1 To $aSegments[0]
				If $i = $aSegments[0] Then
					$newName = $newName & "." & $aSegments[$i]
					ExitLoop
				Else
					$newName &= $aSegments[$i]
				EndIf
			Next
		EndIf
	EndIf
	
	Local $searchPattern = '+,;=[]\/:*?"<>|'
	Local $aChars = StringSplit($searchPattern, "")		;Build illegal character array.
	
	;Replace any and all illegal characters with underscores.
	For $i = 1 To $aChars[0]
		$nPos = StringInStr($newName, $aChars[$i], 0, 1)
		If $nPos = 0 Then ContinueLoop
		$newName = StringReplace($newName, $aChars[$i], "_", 0, 0)
	Next
	
	;Split name into segments again if a file extension exists, so we can evaluate them and then piece them back together later.
	$aSegments = StringSplit($newName, ".")
	Local $name, $extension
	
	If Not @error And $aSegments[0] = 2 Then
		$name = $aSegments[1]
		$extension = $aSegments[2]
	Else
		$name = $newName
	EndIf
	
	;If name is longer than 8 characters, discard the rest of the string after the 8th character.
	If StringLen($name) > 8 Then $name = StringLeft($name, 8)
	
	;If extension exists and is longer than 3 characters, discard the rest of the string after the 3rd character.
	;Then piece the name and extension strings back together.
	If StringLen($extension) > 0 Then 
		$extension = StringLeft($extension, 3)
		$name = $name & "." & $extension
	EndIf
	
	;Check to see if the newly named file already exists.  If so, truncate an additional 2 characters from the name
	;and substitute with digits from 00 to 99. i.e. ismyfile.txt becomes ISMYFI00.TXT, or ISMYFI00.TXT becomes ISMYFI01.TXT, etc.
	$i = 0
	
	While FileExists($justDir & $name)
		;If there are 99 files with the same name, just give up and return an error.
		If $i = 99 Then Return SetError(3, 0, $FileName)
		$aSegments = StringSplit($name, ".")
		If $aSegments[0] = 2 Then $name = $aSegments[1]
		$name = StringLeft($name, 6)
		
		If $i < 10 Then
			$name = $name & "0" & $i
		Else
			$name = $name & $i
		EndIf
		
		;If extension exists, piece name string and extension back together.
		If $aSegments[0] = 2 Then $name = $name & "." & $aSegments[2]
		$i = $i + 1
	WEnd
	
	;Return name only, without truncating the rest of the path.
	;This gives you the ability to truncate all names in a directory without truncating the directory. (Good for mass rename)
	Return $name
EndFunc   ;==>_FileTruncateName


; #INTERNAL USE ONLY# ============================================================================================================
; Name ..........: _SHFileOperation
; Description ...: Asks the shell to perform the specified file operation.
; Syntax ........: _SHFileOperation(ByRef $lpFileOp)
; Parameters ....: $lpFileOp - The file operation to perform.  Use one of the following operation constants:
;                  $FO_COPY
;                  $FO_MOVE
;                  $FO_RENAME
;                  $FO_DELETE
; Return values .: If DllCall() succeeds, then the code from the DLL call is returned.
; Author ........: Chris Brunner <CyrusBuilt at gmail dot com>
; Modified ......: 
; Remarks .......: Modified from original code by Livewire (see AutoIt Forum).
; Related .......: See global constants section for definition of file operation constants.  Operation constants cannot be combined.
; Link ..........;
; Example .......; 
; ===============================================================================================================================
Func _SHFileOperation(ByRef $lpFileOp)
    Local $aDllRet = DllCall("shell32.dll", "int", "SHFileOperation", "ptr", DllStructGetPtr($lpFileOp))
    If Not @error Then Return $aDllRet[0]
EndFunc   ;==>_SHFileOperation


; #INTERNAL USE ONLY# ============================================================================================================
; Name ..........: _TrimTrailingDelimiter
; Description ...: Trims the trailing backslash (path delimiter) from the specified string (if one exists).
; Syntax ........: _TrimTrailingDelimiter($string)
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
Func _TrimTrailingDelimiter($string)
	If StringLen($string) > 1 And StringRight($string, 1) = "\" Then
		Return StringLeft($string, StringLen($string) - 1)
	Else
		Return $string
	EndIf
EndFunc   ;==>_TrimTrailingDelimiter
