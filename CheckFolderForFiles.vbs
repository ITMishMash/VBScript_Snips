FUNCTION CheckFolderForFiles(strFolderName)
'**************************************************************************************
'Returns int
Checks a folder for files
'Returns the number of files or -1 if the folder does not exists
'Sample Usage:
'intMyNumFiles = CheckFolderForFiles "C:\TEMP"
	
	Dim objFSO, objFolder
	'Check to see if there are files in the folder and if folder exists
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If (objFSO.FolderExists(strFolderName)) Then
		Set objFolder = objFSO.GetFolder(strFolderName)
		CheckFolderForFiles = objFolder.Files.Count	
	Else
		CheckFolderForFiles = -1
	End If
'**************************************************************************************
END FUNCTION
