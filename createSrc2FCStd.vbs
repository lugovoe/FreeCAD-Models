Option Explicit

' License GPL-3.0: https://choosealicense.com/licenses/gpl-3.0/

' Creating a file document through Zip container based on the scheme
Const VERSION = "0.03.000"

Dim HEX, fso, scriptFolder, zipFile, shell, projectFolders, objFolder

Set fso = CreateObject("Scripting.FileSystemObject")
scriptFolder = fso.GetAbsolutePathName(".")

Set shell = CreateObject("Shell.Application")
Set projectFolders = shell.Namespace(scriptFolder).ParseName(InputBox("Выберите папку для сборки проектов", , "CNC"))

' Output file extension
Const EXTENTION = ".FCStd"
HEX = "PK" & Chr(3) & Chr(4) & Chr(20) & String(3, vbNullChar) & Chr(8) _
	& String(3, vbNullChar) & "TSg" & Chr(135) & "H" & Chr(177) & Chr(154) _
	& String(3, vbNullChar) & Chr(248) & String(3, vbNullChar) & Chr(12) _
	& String(3, vbNullChar) & "Document.xmlm" & Chr(143) & "=" & Chr(11) _
	& Chr(194) & "0" & Chr(20) & "Eg" & Chr(243) & "+" & Chr(30) & "Y2" _
	& Chr(217) & "6" & Chr(224) & Chr(160) & Chr(208) & Chr(143) & Chr(193) _
	& Chr(226) & Chr(170) & Chr(160) & Chr(184) & Chr(199) & Chr(248) & "l#M" _
	& Chr(34) & "i*" & Chr(254) & "|S" & String(2, Chr(169)) & Chr(131) _
	& Chr(227) & ";" & Chr(239) & Chr(220) & Chr(11) & "7" & Chr(175) & "^" _
	& Chr(186) & Chr(131) & "'" & Chr(186) & "^YS0" & Chr(158) & "d" & Chr(12) _
	& Chr(208) & "H{U" & Chr(166) & ")" & Chr(216) & Chr(224) & "o" & Chr(203) _
	& "5" & Chr(171) & "J" & Chr(146) & Chr(215) & "V" & Chr(14) & Chr(26) _
	& Chr(141) & Chr(135) & Chr(163) & "lQ" & Chr(139) & Chr(243) & Chr(20) _
	& Chr(160) & "+" & Chr(10) & Chr(7) & "g" & Chr(27) & "'tDY" & Chr(194) & "7" _
	& Chr(20) & "v" & Chr(170) & Chr(195) & Chr(136) & "8-" & Chr(201) & Chr(34) _
	& Chr(15) & Chr(226) & Chr(3) & Chr(157) & "W" & Chr(216) & Chr(195) _
	& Chr(214) & Chr(14) & Chr(198) & Chr(7) & Chr(149) & Chr(194) & Chr(201) _
	& Chr(9) & Chr(211) & Chr(171) & Chr(208) & Chr(28) & Chr(209) & Chr(168) _
	& Chr(166) & Chr(179) & ";" & Chr(158) & Chr(251) & Chr(203) & Chr(29) _
	& Chr(165) & Chr(255) & Chr(137) & "}" & Chr(156) & Chr(137) & Chr(206) _
	& "B-" & Chr(188) & Chr(248) & Chr(235) & Chr(140) & Chr(143) & "0" & Chr(34) _
	& Chr(253) & Chr(174) & "(" & Chr(201) & Chr(27) 
HEX = HEX & "PK" & Chr(3) & Chr(4) & Chr(20) & String(3, vbNullChar) & Chr(8) _
	& vbNullChar & Chr(6) & vbNullChar & "TS" & Chr(203) & "KWp" & Chr(144) _
	& String(3, vbNullChar) & Chr(204) & String(3, vbNullChar) & Chr(15) _
	& String(3, vbNullChar) & "GuiDocument.xml" & Chr(179) & Chr(177) & Chr(175) _
	& Chr(200) & Chr(205) & "Q(K-*" & Chr(206) & Chr(204) & Chr(207) & Chr(179) _
	& "U7" & Chr(212) & "3PWH" & Chr(205) & "K" & Chr(206) & "O" & Chr(201) _
	& Chr(204) & "K" & Chr(183) & "U/-I" & Chr(211) & Chr(181) & "P" _
	& String(2, Chr(183)) & Chr(227) & Chr(178) & "q" & Chr(201) & "O." _
	& Chr(205) & "M" & Chr(205) & "+Q" & Chr(8) & "N" & Chr(206) & "H" & Chr(205) _
	& "M" & Chr(12) & Chr(131) & "jP2TR" & Chr(240) & "H,v" & Chr(173) & "(H" _
	& Chr(204) & Chr(131) & Chr(9) & Chr(216) & "qq" & Chr(218) & Chr(128) _
	& Chr(5) & "R" & String(2, Chr(20)) & Chr(146) & Chr(243) & "K" & Chr(243) _
	& "J" & Chr(160) & Chr(130) & "p" & Chr(209) & Chr(188) & Chr(196) & Chr(220) _
	& "T[%" & Chr(167) & Chr(252) & Chr(148) & "J%}" & Chr(144) & "b}" & Chr(136) _
	& "8" & Chr(136) & Chr(25) & Chr(150) & Chr(153) & "Z" & Chr(30) & "P" _
	& Chr(148) & "_" & Chr(150) & Chr(153) & Chr(146) & "Z" & Chr(228) & Chr(146) _
	& "X" & Chr(146) & Chr(168) & Chr(224) & Chr(12) & "1" & Chr(192) _
	& vbNullChar & "l" & Chr(170) & ">" & Chr(186) & "4" & Chr(208) & "a" _
	& Chr(250) & "0" & Chr(151) & Chr(217) & "q" & Chr(1) & vbNullChar
HEX = HEX & "PK" & Chr(1) & Chr(2) & Chr(31) & vbNullChar & Chr(20) _
	& String(3, vbNullChar) & Chr(8) & String(3, vbNullChar) & "TSg" & Chr(135) _
	& "H" & Chr(177) & Chr(154) & String(3, vbNullChar) & Chr(248) _
	& String(3, vbNullChar) & Chr(12) & String(9, vbNullChar) & " " _
	& String(7, vbNullChar) & "Document.xmlPK" & Chr(1) & Chr(2) & Chr(31) _
	& vbNullChar & Chr(20) & String(3, vbNullChar) & Chr(8) & vbNullChar & Chr(6) _
	& vbNullChar & "TS" & Chr(203) & "KWp" & Chr(144) & String(3, vbNullChar) _
	& Chr(204) & String(3, vbNullChar) & Chr(15) & String(9, vbNullChar) & " " _
	& String(3, vbNullChar) & Chr(196) & String(3, vbNullChar) & "GuiDocument.xml"
HEX = HEX & "PK" & Chr(5) & Chr(6) & String(4, vbNullChar) & Chr(2) & vbNullChar _
	& Chr(2) & vbNullChar & "w" & String(3, vbNullChar) & Chr(129) & Chr(1) _
	& String(2, vbNullChar) & Chr(16) & vbNullChar & "FreeCAD Document" & vbNullChar

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
