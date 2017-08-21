'------------------------------------------------------------------------------
' Delete all the expired files in specific folder
' Please change the FOLDER NAME and PASSED DAYS in input Filder before use
'
'
' Date/By                    Changes                          Version
' 25.10.2010/dongdly          create                          0.1
' 25.08.2015/dongdly          add folder array                0.2
'
'------------------------------------------------------------------------------

'******************Input Field Start**************************

'set folder to check
Dim folderArray(2)
'folderPath = "D:\Cache\"
folderArray(0) = "D:\Cache\"
folderArray(1) = "D:\LTE\TrustFolder\"
'the passed days since the last modified date of the file
passedDays = 7 
'******************Input Field End*****************************

expiredDate = Now - passedDays
'Wscript.Echo "Files which older than " &expiredDate &" will be deleted"

On Error GoTo 0

'check folder format
For Each folderPath in folderArray
	If Len(folderPath) > 1 And Right(folderPath, 1) <> "\" Then
		'folderPath = Left(folderPath, Len(folderPath) - 1)
		folderPath = folderPath & "\"
	End If
	
	Set fs = CreateObject("Scripting.FileSystemObject")

	If fs.FolderExists(folderPath) = True Then
		'Wscript.Echo "loop files in " &folderPath
		Set folderObject = fs.GetFolder(folderPath)
		'Wscript.Echo "----------------------------------------------------------------"
		'Loop sub folders
		LoopSubFolders folderObject, expiredDate
		'delete root files
		'Wscript.Echo folderPath
		DeleteExpiredFiles folderObject, expiredDate
	Else
		'Wscript.Echo folderPath & " not exist!" 
	End If

Next

'------------------------------------------------------------------------------
'DeleteExpiredFiles()
'
' Date/By         	Changes                          Version
' 25.10.2010/dongdly  First version                    0.1
'
'
'INPUT: Folder Object, Expired Date
'OUTPUT: None
'
'CALLS:
'
'Comments: Delete all files which is older that expired date under Folder
'------------------------------------------------------------------------------
Function DeleteExpiredFiles(folderObject, expiredDate)
	For Each file In folderObject.Files
		lastModifiedDate = file.DateLastModified
		'Wscript.Echo file.Name & ":    lastModifiedDate " & lastModifiedDate
		If lastModifiedDate < expiredDate Then
			'file deleted
			'Wscript.Echo file.Name & " deleted.  " & lastModifiedDate
			file.Delete
		End If
	Next
End Function


'------------------------------------------------------------------------------
'LoopSubFolders()
'
' Date/By         	Changes                          Version
' 25.10.2010/dongdly  First version                    0.1
'
'
'INPUT: Folder Object, Expired Date
'OUTPUT: None
'
'CALLS:
'
'Comments: loop all sub folders in specific folder and delete all the expired files
'------------------------------------------------------------------------------
Function LoopSubFolders(folderObject, expiredDate)
	For Each subFolder In folderObject.SubFolders
		'recursive all the sub folders
		'Wscript.Echo "clear fils in sub folder: " & subFolder.Name
		LoopSubFolders subFolder, expiredDate
		'delete all files
		'Wscript.Echo subFolder.Name
		DeleteExpiredFiles subFolder, expiredDate
		
		If subFolder.Size < 1 Then
			subFolder.Delete
		End If
	Next
End Function