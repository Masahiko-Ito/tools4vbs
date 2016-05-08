'
' Run as cscript //NoLogo tail.vbs [-l LINE_NUMBER] [input ...]
'
'--------------------------------------------------
' 共通処理
'
Option Explicit
Function include(filename)
	ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile(filename).ReadAll()
End Function
include("mytoolkit.vbs")
'--------------------------------------------------
' 大局変数
'
Dim objEndSw : Set objEndSw = new MySwitch
Dim objMisc : Set objMisc = new MyMisc
' User coding start
Dim objOpt : Set objOpt = new MyOption
Dim objStdio : Set objStdio = new MyStdio
Dim objFsoIn : Set objFsoIn = new MyFso
Dim strRec
Dim strArrayFilesEtc
Dim intArrayFilesEtcIndex : intArrayFilesEtcIndex = 0
Dim intLine : intLine = 10
Dim objSortA
Dim objSortD
Dim intCount : intCount = 0
Dim objTmpStr : Set objTmpStr = new MyString
Dim objReadEndSw : Set objReadEndSw = new MySwitch
' User coding end
'--------------------------------------------------
' 処理開始
'
sub_open		' オープン処理
sub_initialize		' 開始処理
While objEndSw.isOff
	sub_main	' 主処理
Wend
sub_terminate		' 終了処理
sub_close		' クローズ処理
objMisc.exitProg(0)
'--------------------------------------------------
' オープン処理
'
Sub sub_open
' User coding start
	objOpt.initialize("-h=n,--help=n,-l=y")

	If objOpt.isSpecified("-h") or objOpt.isSpecified("--help") Then
		objStdio.writeLine "Usage : cscript //NoLogo tail.vbs [-l LINE_NUMBER] [input ...]"
		objStdio.writeLine "Print the last 10 lines of each input to standard output."
		objStdio.writeLine "With no input, read standard input."
		objStdio.writeLine ""
		objStdio.writeLine " -l LINE_NUMBER        output the last LINE_NUMBER lines, instead of the last 10."
		objEndSw.turnOn
	Else
		objEndSw.turnOff
	End If

	strArrayFilesEtc = objOpt.getArrayNonOptions
	If objEndSw.isOff Then
		If objOpt.isSpecified("-l") Then
			intLine = CInt(objOpt.getValue("-l"))
		End If
	End If
' User coding end
End Sub
'--------------------------------------------------
' 開始処理
'
Sub sub_initialize
' User coding start
' User coding end
End Sub
'--------------------------------------------------
' 主処理
'
Sub sub_main
' User coding start
	Set objSortA = new MySort
	Set objSortD = new MySort

	objSortA.setDelimiter(" ")
	objSortA.setKey("0:A:N:16,32000")
	objSortD.setDelimiter(" ")
	objSortD.setKey("0:D:N:16,32000")

	openFile

	intCount = 0
	objReadEndSw.turnOff

	strRec = readRec
	While objReadEndSw.isOff
		intCount = intCount + 1
		objSortD.putRec(intCount & " " & strRec)
		strRec = readRec
	Wend

	if intCount > 0 Then
		objSortD.sort()
		intCount = 0
		While (Not objSortD.isEof) and intCount < CInt(intLine)
			objSortA.putRec(objSortD.getRec)
			intCount = intCount + 1
		Wend
		objSortA.sort()
		While Not objSortA.isEof
			objTmpStr.setValue objSortA.getRec
			objStdio.writeLine objTmpStr.getReplace("^[0-9]* ", "")
		Wend
	End If

	Set objSortA = Nothing
	Set objSortD = Nothing

	closeFile

	intArrayFilesEtcIndex = intArrayFilesEtcIndex + 1
	If intArrayFilesEtcIndex > UBound(strArrayFilesEtc) Then
		objEndSw.turnOn
	End If
' User coding end
End Sub
'--------------------------------------------------
' 終了処理
'
Sub sub_terminate
' User coding start
' User coding end
End Sub
'--------------------------------------------------
' クローズ処理
'
Sub sub_close
' User coding start
' User coding end
End Sub
'--------------------------------------------------
' その他の処理
'
' User coding start
Function openFile
	If UBound(strArrayFilesEtc) >= 0 Then
		objFsoIn.openInput(strArrayFilesEtc(intArrayFilesEtcIndex))
	End If
End Function

Function closeFile
	If UBound(strArrayFilesEtc) >= 0 Then
		objFsoIn.close
	End If
End Function

Function readRec
	If UBound(strArrayFilesEtc) >= 0 Then
		readRec = objFsoIn.readLine
		If objFsoIn.isReadFailure Then
			objReadEndSw.turnOn
		End If
	Else
		readRec = objStdio.readLine
		If objStdio.isReadFailure Then
			objReadEndSw.turnOn
		End If
	End If
End Function
' User coding end
