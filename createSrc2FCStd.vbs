Option Explicit

' License GPL-3.0: https://choosealicense.com/licenses/gpl-3.0/

' Creating a file document through Zip container based on the scheme
Const VERSION = "0.01.000"

Dim fso, scriptFolder, zipFile, shell, projectFolders, objFolder

Set fso = CreateObject("Scripting.FileSystemObject")
scriptFolder = fso.GetAbsolutePathName(".")

Set shell = CreateObject("Shell.Application")
Set projectFolders = shell.Namespace(scriptFolder).ParseName(InputBox("Выберите папку для сборки проектов", , "CNC"))

' Output file extension
Const EXTENTION = ".FCStd"


With fso
	If Not projectFolders Is Nothing Then
		Set projectFolders = projectFolders.GetFolder.Items()
		If projectFolders.Count < 1 Then MsgBox "No found projectFolders", 16: WScript.Quit
		
		For Each zipFile In projectFolders
			Set objFolder = shell.Namespace(zipFile.Path).ParentFolder.ParseName(zipFile.Name)
			zipFile = scriptFolder & "\" & zipFile.Name & ".zip"
			
			If .FileExists(Replace(zipFile, ".zip", EXTENTION)) = True Then
				' Clean up output File
				.GetFile(Replace(zipFile, ".zip", EXTENTION)).Delete
				WScript.Sleep 200
			End If
			
			With .CreateTextFile(zipFile, True)
				' Create an empty Zip container
				.Write "PK" & Chr(5) & Chr(6) & String(18, vbNullChar)
				.Close
			End With
			' Copy schema folder to Zip container
			shell.Namespace(Replace(zipFile, "\\", "\")).CopyHere objFolder, 16
			
			WScript.Sleep 600
			If .GetFile(zipFile).Size < &H4BA Then
				MsgBox "Not created Zip container """ & zipFile & """", 48
			End If
			' Rename the Zip container to change the file extension to the Extention
			.GetFile(zipFile).Move Replace(zipFile, ".zip", EXTENTION)
		Next
	End If
End With

Set projectFolders = Nothing
Set shell = Nothing
Set fso = Nothing
