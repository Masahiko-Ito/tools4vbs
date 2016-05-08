'
' Easy encrypt(scramble?) tool for .zip in current directory
'
Option Explicit

'
' Global Variable
'
Dim objFso
Dim objFolder
Dim objFiles
Dim objFile
Dim objRegex
Dim objIn
Dim objOut
Dim bytesRec
Dim strPassword
Dim aAsciiPw()
Dim intPosition
Dim i

'
' get Objects for FileSystem
'
Set objFSo = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFso.GetFolder(".")
Set objFiles = objFolder.Files

'
' get Objects for Regex(foo.zip|foo.ZIP)
'
Set objRegex = CreateObject("VBScript.RegExp")
objRegex.IgnoreCase = True
objRegex.Pattern = ".zip$"

'
' get Password
'
If LCase(Right(WScript.FullName, 11)) = "wscript.exe" Then
	strPassword = InputBox("Password?:")
Else
	wscript.stdout.write("Password?:")
	strPassword = wscript.stdin.ReadLine
End If

'
' get ascii code of Password
'
For i = 1 to Len(strPassword)
	Redim preserve aAsciiPw(i - 1)
	aAsciiPw(i - 1) = Asc(Mid(strPassword, i, 1))
Next

'
' Main loop for FileObjects in Current Directory
'
For Each objFile in objFiles
	'
	' Filename is "foo.zip|foo.ZIP"
	'
	If objRegex.Test(objFile.Name) Then
		'
		' open Input as Binary mode
		'
		Set objIn = CreateObject("ADODB.Stream")
		objIn.Open
		objIn.Type = 1   ' StreamTypeEnum ‚Ì adTypeBinary
		objIn.LoadFromFile objFile.Name

		'
		' open Output as Binary mode
		'
		Set objOut = CreateObject("ADODB.Stream")
		objOut.Open
		objOut.Type = 1   ' StreamTypeEnum ‚Ì adTypeBinary

		'
		' get First to Intermediate Blocks
		'
		i = 0
		intPosition = objIn.Size - aAsciiPw(i)
		While intPosition > 0
			objIn.Position = intPosition
			bytesRec = objIn.Read(aAsciiPw(i))
			objOut.Write bytesRec

			i = i + 1
			If i > Ubound(aAsciiPw) Then
				i = 0
			End If
			intPosition = intPosition - aAsciiPw(i)
		Wend

		'
		' get Last Block
		'
		objIn.Position = 0
		bytesRec = objIn.Read(aAsciiPw(i) + intPosition)
		objOut.Write bytesRec

		'
		' save Output File
		'
		objOut.SaveToFile objFile.Name & ".enc", 2

		'
		' close Files
		'
		objIn.close
		objOut.close
	End If
Next
