'
' Easy decrypt(de-scramble?) tool for .zip in current directory
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
Dim intAsciiPwTotal
Dim intPosition
Dim intFirstBlockSize
Dim i

'
' get Objects for FileSystem
'
Set objFSo = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFso.GetFolder(".")
Set objFiles = objFolder.Files

'
' get Objects for Regex(foo.enc|foo.ENC)
'
Set objRegex = CreateObject("VBScript.RegExp")
objRegex.IgnoreCase = True
objRegex.Pattern = ".enc$"

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
' Caluculate total of ascii code of Password
'
intAsciiPwTotal = 0
For i = 1 to Len(strPassword)
	Redim preserve aAsciiPw(i - 1)
	aAsciiPw(i - 1) = Asc(Mid(strPassword, i, 1))
	intAsciiPwTotal = intAsciiPwTotal + aAsciiPw(i - 1)
Next

'
' Main loop for FileObjects in Current Directory
'
For Each objFile in objFiles
	'
	' Filename is "foo.enc|foo.ENC"
	'
	If objRegex.Test(objFile.Name) Then
		'
		' Open Input in Text mode for checking Password
		'
		Set objIn = CreateObject("ADODB.Stream")
		objIn.Charset = "ascii"
		objIn.Open
		objIn.Type = 2   ' StreamTypeEnum ‚Ì adTypeText
		objIn.LoadFromFile objFile.Name

		'
		' Caluculate First Block
		'
		i = 0
		intPosition = (ObjIn.Size Mod intAsciiPwTotal) - aAsciiPw(i)
		While intPosition > 0
			i = i + 1
			If i > Ubound(aAsciiPw) Then
				i = 0
			End If
			intPosition = intPosition - aAsciiPw(i)
		Wend
		intFirstBlockSize = aAsciiPw(i) + intPosition

		'
		' get First Block
		'
		If intFirstBlockSize > 0 Then
			intPosition = objIn.Size - intFirstBlockSize
			objIn.Position = intPosition
			bytesRec = objIn.ReadText(intFirstBlockSize)

			i = i - 1
			If i < 0 Then
				i = Ubound(aAsciiPw)
			End If
			intPosition = intPosition - aAsciiPw(i)
		Else
			bytesRec = ""
			i = 0
			intPosition = objIn.Size - aAsciiPw(i)
		End If

		'
		' get Second Block
		'
		objIn.Position = intPosition
		bytesRec = bytesRec & objIn.ReadText(aAsciiPw(i))
		objIn.close

		'
		' check Password
		'
		If Mid(bytesRec, 1, 2) <> "PK" Then
			If LCase(Right(WScript.FullName, 11)) = "wscript.exe" Then
				MsgBox("Wrong Password!")
			Else
				wscript.stdout.writeLine "Wrong Password!"
			End If
			WScript.Quit(1)
		End If

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
		' Caluculate First Block
		'
		i = 0
		intPosition = (ObjIn.Size Mod intAsciiPwTotal) - aAsciiPw(i)
		While intPosition > 0
			i = i + 1
			If i > Ubound(aAsciiPw) Then
				i = 0
			End If
			intPosition = intPosition - aAsciiPw(i)
		Wend
		intFirstBlockSize = aAsciiPw(i) + intPosition

		'
		' get First Block
		'
		If intFirstBlockSize > 0 Then
			intPosition = objIn.Size - intFirstBlockSize
			objIn.Position = intPosition
			bytesRec = objIn.Read(intFirstBlockSize)
			objOut.Write bytesRec

			i = i - 1
			If i < 0 Then
				i = Ubound(aAsciiPw)
			End If
			intPosition = intPosition - aAsciiPw(i)
		Else
			i = 0
			intPosition = objIn.Size - aAsciiPw(i)
		End If

		'
		' get Intermediate Blocks
		'
		While intPosition > 0
			objIn.Position = intPosition
			bytesRec = objIn.Read(aAsciiPw(i))
			objOut.Write bytesRec

			i = i - 1
			If i < 0 Then
				i = Ubound(aAsciiPw)
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
		objOut.SaveToFile objRegex.Replace(objFile.Name, ""), 2

		'
		' close Files
		'
		objIn.close
		objOut.close
	End If
Next
