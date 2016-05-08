'
' Run as cscript //NoLogo selgrep.vbs [-v] [-d DELIMITER] -i INDEX REGEX [input ...]
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
Dim intArrayFilesEtcIndex : intArrayFilesEtcIndex = 1
Dim objOmitSw : Set objOmitSw = new MySwitch
Dim strDelimiter : strDelimiter = ","
Dim strIndex
Dim strRegex
Dim objTmpStr : Set objTmpStr = new MyString
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
	objOpt.initialize("-h=n,--help=n,-v=n,-d=y,-i=y")

	If objOpt.isSpecified("-h") or objOpt.isSpecified("--help") Then
		objStdio.writeLine "Usage : cscript //NoLogo selgrep.vbs [-v] [-d DELIMITER] -i INDEX REGEX [input ...]"
		objStdio.writeLine "Search for REGEX in INDEX of each input or standard input."
		objStdio.writeLine "REGEX is an regular expression."
		objStdio.writeLine ""
                objStdio.writeLine " -v                  select non-matching lines."
		objStdio.writeLine " -d DELIMITER        use DELIMITER instead of comma for field delimiter."
		objStdio.writeLine " -i INDEX            matching field(0 origin)."
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
		If objOpt.isSpecified("-i") Then
			strIndex = objOpt.getValue("-i")
		End If
		strRegex = strArrayFilesEtc(0)
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
		If CInt(strIndex) <= UBound(v) Then
			objTmpStr.setValue v(CInt(strIndex))
			If objOmitSw.isOff Then
				If objTmpStr.isMatch(strRegex) Then
					If UBound(strArrayFilesEtc) >= 2 Then
						objStdio.writeLine strCurrentFileName & ":" & strRec
					Else
						objStdio.writeLine strInRec
					End If
				End If
			Else
				If Not objTmpStr.isMatch(strRegex) Then
					If UBound(strArrayFilesEtc) >= 2 Then
						objStdio.writeLine strCurrentFileName & ":" & strRec
					Else
						objStdio.writeLine strInRec
					End If
				End If
			End If
		End If
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
	If UBound(strArrayFilesEtc) >= 1 Then
		objEndSw.turnOff
		objFsoIn.openInput(strArrayFilesEtc(intArrayFilesEtcIndex))
		strCurrentFileName = strArrayFilesEtc(intArrayFilesEtcIndex)
		intArrayFilesEtcIndex = intArrayFilesEtcIndex + 1
	End If
End Function

Function closeFile
	If UBound(strArrayFilesEtc) >= 1 Then
		objFsoIn.close
	End If
End Function

Function readRec
	If UBound(strArrayFilesEtc) >= 1 Then
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
' User coding end
