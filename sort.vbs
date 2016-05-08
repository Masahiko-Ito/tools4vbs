'
' Run as cscript //NoLogo sort.vbs [-d DELIMITER] -k KEY_INFO [-r RECLEN] [input ...]
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
Dim objSort : Set objSort = new MySort
Dim objDataExistSw : Set objDataExistSw = new MySwitch
Dim strDelimiter : strDelimiter = ","
Dim strKeyInfo
Dim intRecLen : intRecLen = 1024
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
	objOpt.initialize("-h=n,--help=n,-d=y,-k=y,-r=y")

	If objOpt.isSpecified("-h") or objOpt.isSpecified("--help") Then
		objStdio.writeLine "Usage : cscript //NoLogo sort.vbs [-d DELIMITER] -k KEY_INFO [-r RECLEN] [input ...]"
		objStdio.writeLine "Write sorted concatenation of all input(s) to standard output."
		objStdio.writeLine ""
		objStdio.writeLine " -d DELIMITER       use DELIMITER instead of comma for field delimiter."
		objStdio.writeLine " -k KEY_INFO        index:seq:type:max_length[,index:seq:type:max_length ...]"
		objStdio.writeLine "   index               index number for sort field in input(0 origin)"
		objStdio.writeLine "   seq                 A or D (ascending or descending)"
		objStdio.writeLine "   type                S or N (string or number)"
		objStdio.writeLine "   max_length          max length for key"
		objStdio.writeLine " -r RECLEN          max length for input record(default 1024)."
		objEndSw.turnOn
		objDataExistSw.turnOff
	End If

	strArrayFilesEtc = objOpt.getArrayNonOptions
	If objEndSw.isOff Then
		If objOpt.isSpecified("-d") Then
			strDelimiter = objOpt.getValue("-d")
		End If
		If objOpt.isSpecified("-k") Then
			strKeyInfo = objOpt.getValue("-k")
		End If
		If objOpt.isSpecified("-r") Then
			intRecLen = CInt(objOpt.getValue("-r"))
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
		objSort.setDelimiter(strDelimiter)
		objSort.setKey(strKeyInfo & "," & intRecLen)
		strRec = readRec
		If objEndSw.isOn Then
			objDataExistSw.turnOff
		Else
			objDataExistSw.turnOn
		End If
	End If
' User coding end
End Sub
'--------------------------------------------------
' 主処理
'
Sub sub_main
' User coding start
	If objSkipSw.isOff Then
		objSort.putRec(strRec)
	End If

	strRec = readRec
' User coding end
End Sub
'--------------------------------------------------
' 終了処理
'
Sub sub_terminate
' User coding start
	If objDataExistSw.isOn Then
		objSort.sort()
		While Not objSort.isEof
			objStdio.writeLine objSort.getRec
		Wend
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
