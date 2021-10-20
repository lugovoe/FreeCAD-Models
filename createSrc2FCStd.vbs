Option Explicit

' License GPL-3.0: https://choosealicense.com/licenses/gpl-3.0/

' Creating a file document through Zip container based on the scheme
Const VERSION = "0.02.000"

Dim HEX, fso, scriptFolder, zipFile, shell, projectFolders, objFolder

Set fso = CreateObject("Scripting.FileSystemObject")
scriptFolder = fso.GetAbsolutePathName(".")

Set shell = CreateObject("Shell.Application")
Set projectFolders = shell.Namespace(scriptFolder).ParseName(InputBox("Выберите папку для сборки проектов", , "CNC"))

' Output file extension
Const EXTENTION = ".FCStd"
HEX = "PK" & Chr(3) & Chr(4) & Chr(20) & String(3, vbNullChar) & Chr(8) _
    & String(3, vbNullChar) & "TSg" & Chr(135) & "H" & Chr(177) & Chr(153) _
    & String(3, vbNullChar) & Chr(248) & String(3, vbNullChar) & Chr(12) _
    & String(3, vbNullChar) & "Document.xmlm" & Chr(207) & "=" & Chr(11) _
    & Chr(194) & "0" & Chr(16) & Chr(6) & Chr(224) & Chr(217) & Chr(254) _
    & Chr(138) & "#K&" & Chr(219) & Chr(6) & Chr(28) & Chr(20) & Chr(218) _
    & "t" & Chr(176) & Chr(184) & "*(" & Chr(238) & "1" & Chr(158) & "5" _
    & Chr(210) & "$" & Chr(146) & Chr(15) & Chr(241) & Chr(231) & Chr(155) _
    & "Bm" & Chr(17) & Chr(28) & Chr(239) & Chr(189) & Chr(231) & Chr(14) _
    & Chr(222) & Chr(170) & "y" & Chr(235) & Chr(30) & "^" & Chr(232) _
    & Chr(188) & Chr(178) & String(2, Chr(166)) & ",/)" & Chr(160) & Chr(145) _
    & Chr(246) & Chr(170) & "LW" & Chr(211) & Chr(24) & "n" & Chr(203) _
    & "5mxV" & Chr(181) & "VF" & Chr(141) & "&" & Chr(192) & "Q" & Chr(222) _
    & "Q" & Chr(139) & Chr(243) & "x@V" & Chr(4) & Chr(14) & Chr(206) & "vN" _
    & Chr(232) & ")*s" & Chr(182) & "!" & Chr(176) & "S=N" & Chr(17) & "#<[T" _
    & Chr(9) & ">" & Chr(209) & Chr(5) & Chr(133) & Chr(30) & Chr(182) & "6" _
    & Chr(154) & Chr(144) & "(" & Chr(129) & Chr(147) & Chr(19) & Chr(198) _
    & Chr(171) & Chr(244) & "y" & Chr(138) & Chr(6) & "Z" & Chr(204) & "v" _
    & Chr(24) & Chr(247) & Chr(151) & Chr(7) & Chr(202) & Chr(224) & Chr(225) _
    & Chr(215) & Chr(140) & Chr(233) & Chr(12) & "Z" & Chr(17) & Chr(196) _
    & "_3,R" & Chr(137) & Chr(226) & Chr(219) & Chr(130) & "g" & Chr(31)
HEX = HEX & "PK" & Chr(1) & Chr(2) & Chr(31) & vbNullChar & Chr(20) _
    & String(3, vbNullChar) & Chr(8) & String(3, vbNullChar) & "TSg" _
    & Chr(135) & "H" & Chr(177) & Chr(153) & String(3, vbNullChar) _
    & Chr(248) & String(3, vbNullChar) & Chr(12) & vbNullChar & "$" _
    & String(7, vbNullChar) & " " & String(7, vbNullChar) & "Document.xml" _
    & Chr(10) & vbNullChar & " " & String(5, vbNullChar) & Chr(1) _
    & vbNullChar & Chr(24) & vbNullChar & "[" & Chr(156) & Chr(209) & "G," _
    & Chr(197) & Chr(215) & Chr(1) & "[" & Chr(156) & Chr(209) & "G," _
    & Chr(197) & Chr(215) & Chr(1) & "m" & Chr(228) & "pj" & Chr(20) _
    & Chr(197) & Chr(215) & Chr(1) & "PK" & Chr(5) & Chr(6) _
    & String(4, vbNullChar) & Chr(1) & vbNullChar & Chr(1) & vbNullChar _
    & "^" & String(3, vbNullChar) & Chr(195) & String(3, vbNullChar) _
    & Chr(16) & vbNullChar & "FreeCAD Document"

With fso
	If Not projectFolders Is Nothing Then
		Set projectFolders = projectFolders.GetFolder.Items()
		If projectFolders.Count < 1 Then MsgBox "No found projectFolders", 16: WScript.Quit
		
		For Each zipFile In projectFolders
			Set objFolder = shell.Namespace(zipFile.Path).ParentFolder.ParseName(zipFile.Name).GetFolder.Items
			
			zipFile = scriptFolder & "\" & zipFile.Name & ".zip"
			
			If .FileExists(Replace(zipFile, ".zip", EXTENTION)) = True Then
				' Clean up output File
				.GetFile(Replace(zipFile, ".zip", EXTENTION)).Delete
				WScript.Sleep 200
			End If
			
			With .CreateTextFile(zipFile, True)
				' Create an empty Zip container
				.Write HEX
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
