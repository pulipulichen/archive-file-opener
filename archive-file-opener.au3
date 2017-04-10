#pragma compile(Icon, 'icon.ico')
#include <File.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>

Local $aArray[0] = []

Local $fileOpenLimit = 3
Local $tmpDir = @TempDir & "/zip-opener"
;$tmpDir = @ScriptDir & "/temp"

; ---------------------------------------------------

Func openDir($path)
   ;MsgBox($MB_SYSTEMMODAL, "file" & FileExists($path), $path)
   If FileExists($path) == 1 Then
	  IF StringInStr(FileGetAttrib($path),"D")  Then
		 ;MsgBox($MB_SYSTEMMODAL, "dir", $path)
		 Local $aFileList = _FileListToArray($path, "*")
		 For $f = 1 To $aFileList[0]
			$filePath = $path & "/" & $aFileList[$f]

			IF StringInStr(FileGetAttrib($filePath),"D")  Then
			   openDir($filePath)
			Else
			   ;MsgBox($MB_SYSTEMMODAL, $fileName, $file)
			   ShellExecute('"' & $filePath & '"')
			EndIf
		 Next
	  Else
		 ;MsgBox($MB_SYSTEMMODAL, "file", $path)
		 ShellExecute('"' & $path & '"')
	  EndIf
   EndIf
EndFunc

Func countFiles($path, $count)
   ;MsgBox($MB_SYSTEMMODAL, "file" & FileExists($path), $path)
   If FileExists($path) == 1 Then
	  IF StringInStr(FileGetAttrib($path),"D")  Then
		 ;MsgBox($MB_SYSTEMMODAL, "dir", $path)
		 Local $aFileList = _FileListToArray($path, "*")
		 For $f = 1 To $aFileList[0]
			$filePath = $path & "/" & $aFileList[$f]

			IF StringInStr(FileGetAttrib($filePath),"D")  Then
			   $count = $count + countFiles($filePath, $count)
			Else
			   ;MsgBox($MB_SYSTEMMODAL, $fileName, $file)
			   $count = $count + 1
			EndIf
		 Next
	  Else
		 ;MsgBox($MB_SYSTEMMODAL, "file", $path)
		 $count = $count + 1
	  EndIf
   EndIf
   return $count
EndFunc

; ---------------------------------------------------

For $i = 1 To $CmdLine[0]

   Local $filePath = $CmdLine[$i]
   ;MsgBox($MB_SYSTEMMODAL, "", $filePath)
   If FileExists ($filePath) == 1 Then
	  Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
	  $pathSplit = _PathSplit($filePath, $sDrive, $sDir, $sFileName, $sExtension)
	  Local $zipFilePath = $filePath

	  $fileName = $pathSplit[3]

	  ;MsgBox($MB_SYSTEMMODAL, $fileName, $filePath)

	  $zipPath = $tmpDir & '/' & $fileName

	  If DirGetSize ($zipPath) <> -1 Then
		 DirRemove($zipPath, 1)
	  EndIf

	  If DirGetSize ($zipPath) == -1 Then
		 $unzipCmd = @ScriptDir & '/7zip/7z.exe x "' & $filePath & '" -o"' & $zipPath & '"'
		 ;MsgBox($MB_SYSTEMMODAL, $fileName, $unzipCmd)

		 RunWait($unzipCmd)

		 ; 查看這裡面有什麼檔案，然後全部開啟它
		 Local $aFileList = _FileListToArray($zipPath, "*")

		 Local $open = True

		 Local $fileCount = countFiles($zipPath, 0)
		 ;MsgBox($MB_SYSTEMMODAL, "Error ", $fileOpenLimit & "  " & $aFileList[0] & " " & countFiles($zipPath, 0) )

		 If $fileCount > $fileOpenLimit Then
			$open = False
			$confirm = MsgBox(4, "Confirmation", "Thie archive file: " & @CRLF & $zipFilePath & @CRLF & "contains " & $fileCount & " files. " & @CRLF & "Do you want to OPEN them?")
			If $confirm = 6 Then
			   $open = True
			Else
			   DirRemove($zipPath, 1)
			EndIf
		 EndIf

		 If $open == True Then
			For $f = 1 To $aFileList[0]
			   $unzipFilePath = $zipPath & "/" & $aFileList[$f]
			   ;MsgBox($MB_SYSTEMMODAL, $fileName, $file)
			   ;ShellExecute($unzipFilePath)
			   openDir($unzipFilePath)
			Next
		 EndIf
	  Else
		 MsgBox($MB_SYSTEMMODAL, "Error", "Cannot remove directory: " & @CRLF & $zipPath)
	  EndIf


   EndIf
Next

