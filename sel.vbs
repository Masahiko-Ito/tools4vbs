'
' Run as cscript //NoLogo sel.vbs [-v] [-d DELIMITER] -c CONDITION [input ...]
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
Dim objSkipSw : Set objSkipSw = new MySwitch
Dim strCurrentFileName
Dim strArrayFilesEtc
Dim intArrayFilesEtcIndex : intArrayFilesEtcIndex = 0
Dim strDelimiter : strDelimiter = ","
Dim strCondition
Dim objOmitSw : Set objOmitSw = new MySwitch
Dim objTmpStr1 : Set objTmpStr1 = new MyString
Dim objTmpStr2 : Set objTmpStr2 = new MyString
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
	objOpt.initialize("-h=n,--help=n,-v=n,-d=y,-c=y")

	If objOpt.isSpecified("-h") or objOpt.isSpecified("--help") Then
		objStdio.writeLine "Usage : cscript //NoLogo sel.vbs [-v] [-d DELIMITER] -c CONDITION [input ...]"
		objStdio.writeLine "Search for CONDITION in each input or standard input."
		objStdio.writeLine ""
		objStdio.writeLine " -v                  select non-matching lines."
		objStdio.writeLine " -d DELIMITER        use DELIMITER instead of comma for field delimiter."
		objStdio.writeLine " -c CONDITION        ex. ""v(0) = 'foo' and v(2) > 'boo'"""
		objStdio.writeLine "                         ""grep(v(0),'^a') and grep(v(0),'z$')"""
		objEndSw.turnOn
	End If

	strArrayFilesEtc = objOpt.getArrayNonOptions
	If objEndSw.isOff Then
		If objOpt.isSpecified("-v") Then
			objOmitSw.turnOn
		End If
		If objOpt.isSpecified("-d") Then
			strDelimiter = objOpt.getValue("-d")
		End If
		If objOpt.isSpecified("-c") Then
			strCondition = objOpt.getValue("-c")
		End If
		openFile
	End If
' User coding end
End Sub
'--------------------------------------------------
' 開始処理
'
Sub sub_initialize
' User coding start
	If objEndSw.isOff Then
		strRec = readRec
	End If
' User coding end
End Sub
'--------------------------------------------------
' 主処理
'
Sub sub_main
' User coding start
        If objSkipSw.isOff Then
		Dim strInRec
		Dim v

		strInRec = strRec
		v = split(strInRec, strDelimiter)
		objTmpStr1.setValue strCondition
		If objOmitSw.isOff Then
			If Eval(objTmpStr1.getReplace("'", """")) Then
				If UBound(strArrayFilesEtc) >= 1 Then
					objStdio.writeLine strCurrentFileName & ":" & strRec
				Else
					objStdio.writeLine strInRec
				End If
			End If
		Else
			If Not Eval(objTmpStr1.getReplace("'", """")) Then
				If UBound(strArrayFilesEtc) >= 1 Then
					objStdio.writeLine strCurrentFileName & ":" & strRec
				Else
					objStdio.writeLine strInRec
				End If
			End If
		End IF
	End If

	strRec = readRec
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
	closeFile
' User coding end
End Sub
'--------------------------------------------------
' その他の処理
'
' User coding start
Function openFile
	If UBound(strArrayFilesEtc) >= 0 Then
		objEndSw.turnOff
		objFsoIn.openInput(strArrayFilesEtc(intArrayFilesEtcIndex))
		strCurrentFileName = strArrayFilesEtc(intArrayFilesEtcIndex)
		intArrayFilesEtcIndex = intArrayFilesEtcIndex + 1
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
			If intArrayFilesEtcIndex <= UBound(strArrayFilesEtc) Then
				closeFile
				openFile
				objSkipSw.turnOn
			Else
				objEndSw.turnOn
			End If
		Else
			objSkipSw.turnOff
		End If
	Else
		readRec = objStdio.readLine
		If objStdio.isReadFailure Then
			objEndSw.turnOn
		End If
	End If
End Function

Function grep(string, regex)
	objTmpStr2.setValue string
	grep = objTmpStr2.isMatch(regex)
End Function
' User coding end
