'
' Run as cscript //NoLogo uniq.vbs [-d|-c] [input ...]
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
Dim strArrayFilesEtc
Dim intArrayFilesEtcIndex : intArrayFilesEtcIndex = 0
Dim strOldRec : strOldRec = ""
Dim strNewRec : strNewRec = ""
Dim intCount : intCount = 0
Dim objFirstRecSw : Set objFirstRecSw = new MySwitch : objFirstRecSw.turnOn
Dim objFirstDupSw : Set objFirstDupSw = new MySwitch : objFirstDupSw.turnOn
Dim objUniqSw : Set objUniqSw = new MySwitch : objUniqSw.turnOn
Dim objDupSw : Set objDupSw = new MySwitch
Dim objCountSw : Set objCountSw = new MySwitch
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
	objOpt.initialize("-h=n,--help=n,-d=n,-c=n")

	If objOpt.isSpecified("-h") or objOpt.isSpecified("--help") Then
		objStdio.writeLine "Usage : cscript //NoLogo uniq.vbs [-d|-c] [input ...]"
		objStdio.writeLine "Filter adjacent matching lines from input (or standard input), writing to standard output."
		objStdio.writeLine "With no options, matching lines are merged to the first occurrence."
		objStdio.writeLine ""
		objStdio.writeLine " -d        only print duplicate lines."
		objStdio.writeLine " -c        prefix lines by the number of occurrences."
		objEndSw.turnOn
	End If

	strArrayFilesEtc = objOpt.getArrayNonOptions
	If objEndSw.isOff Then
		If objOpt.isSpecified("-d") Then
			objDupSw.turnOn
			objUniqSw.turnOff
		End If
		If objOpt.isSpecified("-c") Then
			objCountSw.turnOn
			objUniqSw.turnOff
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
		strNewRec = strRec
		If objUniqSw.isOn Then
			If objFirstRecSw.isOn Then
				objStdio.writeLine strNewRec
				objFirstRecSw.turnOff
			Else
				If strNewRec <> strOldRec Then
					objStdio.writeLine strNewRec
				End If
			End If
			strOldRec = strNewRec
		Elseif objDupSw.isOn Then
			If objFirstRecSw.isOn Then
				objFirstRecSw.turnOff
			Else
				If strNewRec = strOldRec Then
					If objFirstDupSw.isOn Then
						objStdio.writeLine strNewRec
						objFirstDupSw.turnOff
					End If
				Else
					objFirstDupSw.turnOn
				End If
			End If
			strOldRec = strNewRec
		Elseif objCountSw.isOn Then
			If objFirstRecSw.isOn Then
				intCount = intCount + 1
				objFirstRecSw.turnOff
			Else
				If strNewRec = strOldRec Then
					intCount = intCount + 1
				Else
					objStdio.writeLine intCount & " " & strOldRec
					intCount = 1
				End If
			End If
			strOldRec = strNewRec
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
	If objCountSw.isOn Then
		objStdio.writeLine intCount & " " & strOldRec
	End If
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
' User coding end
