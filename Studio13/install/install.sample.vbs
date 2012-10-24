Dim dropBoxBuildPath
Dim latestDate
Dim folderName
Dim dropBoxRoot
Dim scriptLocation
Dim installerName
Dim installResponseFileName
Dim uninstallResponseFileName
Dim installerLocation

scriptLocation = "C:\Users\regular-user\Documents\GitHub\desktop-Test\Studio13\install\"
dropBoxRoot = "C:\Dropbox\"

dropBoxBuildPath = dropBoxRoot & "Builds\Studio Debug\2.0\"
installerName = "studio40.exe"
uninstallResponseFileName = "uninstall.iss"
installResponseFileName = "install.iss"

Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'***** look in each subfolder for studio installer
Set objFolder = objFSO.GetFolder(scriptLocation)
Set colSubfolders = objFolder.Subfolders

'***** there should only be one folder under \install so use the first one
For Each objSubfolder in colSubfolders
	folderName = objSubfolder.Name
	If (objFSO.FileExists(scriptLocation & folderName & "\" & installerName)) Then
		objShell.Run scriptLocation & folderName & "\" & installerName & " -s -f1" & chr(34) & scriptLocation & uninstallResponseFileName & chr(34) & " -sms",,TRUE
		objFSO.DeleteFolder scriptLocation & folderName 
	End If
Next

'***** get latest build folder in dropbox folder
Set objFolder = objFSO.GetFolder(dropBoxBuildPath)
Set colSubfolders = objFolder.Subfolders
For Each objSubfolder in colSubfolders
	If latestDate < objSubfolder.DateCreated Then
		latestDate = objSubfolder.DateCreated
		folderName = objSubfolder.Name
	End If
Next

'***** set build path to latest build
dropBoxBuildPath = dropBoxBuildPath & folderName & "\install\" & installerName
installerLocation = scriptLocation & folderName & "\"
rem Wscript.Echo "dropBoxBuildPath = " & dropBoxBuildPath
rem Wscript.Echo "installerLocation = " & installerLocation
objFSO.CreateFolder installerLocation

'***** copy installer to test folder
objFSO.CopyFile dropBoxBuildPath, installerLocation

'***** run installer
objShell.Run installerLocation & installerName & " -s -f1" & chr(34) & scriptLocation & installResponseFileName & chr(34) & " -sms",,TRUE


