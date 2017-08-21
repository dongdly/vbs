' Author: Dong Fangjun
' create shortcut to desktop

backlogPath = "D:\"

Set fileSystem = CreateObject("Scripting.FileSystemObject") 
Set folder = fileSystem.GetFolder(backlogPath) 
Set recentFile = Nothing
For Each file In folder.Files      
  If (recentFile is Nothing) Then
    Set recentFile = file
  ElseIf InStr(file.Name,"ShortFileName") > 0 And (file.DateLastModified > recentFile.DateLastModified) Then
    Set recentFile = file
  End If
Next

'Debug
'If recentFile is Nothing Then
'  WScript.Echo "no recent files"
'Else
'  WScript.Echo "Recent file is " & recentFile.Name & " " & recentFile.DateLastModified
'End If

Set WshShell = CreateObject("Wscript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")

' CreateShortcut works like a GetShortcut when the shortcut already exists!
Set objShortcut = WshShell.CreateShortcut(strDesktop + "\ShortFileName.lnk")

'objShortcut.TargetPath = "%windir%\notepad.exe"
objShortcut.TargetPath = recentFile.Path
objShortcut.Save

Set objShortcut = Nothing
Set wshShell    = Nothing