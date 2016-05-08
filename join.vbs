'
' Run as cscript //NoLogo join.vbs [-d DELIMITER] [-a ACTION] [-1 KEY1] [-2 KEY2] input1 input2
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
Dim objFsoIn1 : Set objFsoIn1 = new MyFso
Dim objFsoIn2 : Set objFsoIn2 = new MyFso
Dim isRec1HighValue : isRec1HighValue = False
Dim isRec2HighValue : isRec2HighValue = False
Dim strRec1 : strRec1 = ""
Dim strRec2 : strRec2 = ""
Dim strOldRec1 : strOldRec1 = ""
Dim strOldRec2 : strOldRec2 = ""
Dim objSkipSw : Set objSkipSw = new MySwitch
Dim strArrayFilesEtc
Dim intArrayFilesEtcIndex : intArrayFilesEtcIndex = 0
Dim strDelimiter : strDelimiter = ","
Dim strAction : strAction = "m"
Dim strKey1 : strKey1 = "0"
Dim strKey2 : strKey2 = "0"
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
	objOpt.initialize("-h=n,--help=n,-d=y,-a=y,-1=y,-2=y")

	If objOpt.isSpecified("-h") or objOpt.isSpecified("--help") Then
		objStdio.writeLine "Usage : cscript //NoLogo join.vbs [-d DELIMITER] [-a ACTION] [-1 KEY1] [-2 KEY2] input1 input2"
		objStdio.writeLine "For each pair of input lines with identical join fields, write a line to standard output."
		objStdio.writeLine "The default join field is the first, delimited by comma."
		objStdio.writeLine "When input1 or input2 (not both) is -, read standard input."
		objStdio.writeLine ""
		objStdio.writeLine " -d DELIMITER        use DELIMITER instead of comma as input and output field separator."
		objStdio.writeLine " -a ACTION           1 means to print only unpairable lines from input1."
		objStdio.writeLine "                     2 means to print only unpairable lines from input2."
		objStdio.writeLine "                     m means to print pairable lines from input1 and input2(default)."
		objStdio.writeLine "                     1m means to print from input1 and pairable lines from input2."
		objStdio.writeLine "                     2m means to print from input2 and pairable lines from input1."
		objStdio.writeLine " -1 KEY1             join on this KEY1 of input1(default 0)."
		objStdio.writeLine " -2 KEY2             join on this KEY2 of input2(default 0)."
		objStdio.writeLine "                     ex. 0,2,4 means 1st, 3rd, 5th fields are used as matching key."
		objStdio.writeLine "                         0,2,4m means 1st, 3rd, 5th fields are used as matching key for multiple record."
                objStdio.writeLine "                         Do not specify ""m"" for both of KEY1 and KEY2."

		objEndSw.turnOn
	End If

	strArrayFilesEtc = objOpt.getArrayNonOptions
	If objEndSw.isOff Then
		If objOpt.isSpecified("-d") Then
			strDelimiter = objOpt.getValue("-d")
		End If
		If objOpt.isSpecified("-a") Then
			strAction = objOpt.getValue("-a")
		End If
		If objOpt.isSpecified("-1") Then
			strKey1 = objOpt.getValue("-1")
		End If
		If objOpt.isSpecified("-2") Then
			strKey2 = objOpt.getValue("-2")
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
		strRec1 = readRec1
		strRec2 = readRec2
		If isRec1HighValue and isRec2HighValue Then
			objEndSw.turnOn
		End If
	End If
' User coding end
End Sub
'--------------------------------------------------
' 主処理
'
Sub sub_main
' User coding start
	If isLessThan(strRec1, strKey1, isRec1HighValue, strRec2, strKey2, isRec2HighValue) Then
		if strAction = "1" or strAction = "1m" Then
			If getKey(strRec1, strKey1) = getKey(strOldRec2, strKey2) Then
			Else
				objStdio.writeLine strRec1
			End If
		End If
		strOldRec1 = strRec1
		strRec1 = readRec1
	Elseif isGreaterThan(strRec1, strKey1, isRec1HighValue, strRec2, strKey2, isRec2HighValue) Then
		if strAction = "2" or strAction = "2m" Then
			If getKey(strOldRec1, strKey1) = getKey(strRec2, strKey2) Then
			Else
				objStdio.writeLine strRec2
			End If
		End If
		strOldRec2 = strRec2
		strRec2 = readRec2
	Else
		If strAction = "1m" Then
			objStdio.writeLine strRec1 & strDelimiter & strRec2
		Elseif strAction = "2m" Then
			objStdio.writeLine strRec2 & strDelimiter & strRec1
		Elseif strAction = "m" Then
			objStdio.writeLine strRec1 & strDelimiter & strRec2
		End If
		if isMultiKey(strKey1) and isMultiKey(strKey2) Then
			strOldRec1 = strRec1
			strRec1 = readRec1
			strOldRec2 = strRec2
			strRec2 = readRec2
		Else
			if isMultiKey(strKey2) Then
			Else
				strOldRec1 = strRec1
				strRec1 = readRec1
			End If
			if isMultiKey(strKey1) Then
			Else
				strOldRec2 = strRec2
				strRec2 = readRec2
			End If
		End If
	End If
	If isRec1HighValue and isRec2HighValue Then
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
	closeFile
' User coding end
End Sub
'--------------------------------------------------
' その他の処理
'
' User coding start
Function openFile
	If strArrayFilesEtc(0) <> "-" Then
		objFsoIn1.openInput(strArrayFilesEtc(0))
	End If
	If strArrayFilesEtc(1) <> "-" Then
		objFsoIn2.openInput(strArrayFilesEtc(1))
	End If
End Function

Function closeFile
	If UBound(strArrayFilesEtc) >= 0 Then
		If strArrayFilesEtc(0) <> "-" Then
			objFsoIn1.close
		End If
	End If
	If UBound(strArrayFilesEtc) >= 1 Then
		If strArrayFilesEtc(1) <> "-" Then
			objFsoIn2.close
		End If
	End If
End Function

Function readRec1
	If strArrayFilesEtc(0) = "-" Then
		If objStdio.isEof Then
			isRec1HighValue = True
		Else
			readRec1 = objStdio.readLine
		End If
	Else
		If objFsoIn1.isEof Then
			isRec1HighValue = True
		Else
			readRec1 = objFsoIn1.readLine
		End If
	End If
End Function

Function readRec2
	If strArrayFilesEtc(1) = "-" Then
		If objStdio.isEof Then
			isRec2HighValue = True
		Else
			readRec2 = objStdio.readLine
		End If
	Else
		If objFsoIn2.isEof Then
			isRec2HighValue = True
		Else
			readRec2 = objFsoIn2.readLine
		End If
	End If
End Function

Function isLessThan(rec1, key1, ishv1, rec2, key2, ishv2)
	If ishv1 and ishv2 Then
		isLessThan = False
	Elseif ishv1 Then
		isLessThan = False
	Elseif ishv2 Then
		isLessThan = True
	Else 
		If getKey(rec1, key1) < getKey(rec2, key2) Then
			isLessThan = True
		Else
			isLessThan = False
		End If
	End If
End Function

Function isGreaterThan(rec1, key1, ishv1, rec2, key2, ishv2)
	If ishv1 and ishv2 Then
		isGreaterThan = False
	Elseif ishv1 Then
		isGreaterThan = True
	Elseif ishv2 Then
		isGreaterThan = False
	Else 
		If getKey(rec1, key1) > getKey(rec2, key2) Then
			isGreaterThan = True
		Else
			isGreaterThan = False
		End If
	End If
End Function

Function getKey(rec, keys)
	Dim aRec
	Dim aKey
	Dim strKey
	Dim i

	objTmpStr.setValue keys
	aRec = split(rec, strDelimiter)
	aKey = split(objTmpStr.getReplace("m$", ""), ",")
	strKey = ""
	i = 0
	While i <= UBound(aKey)
		If CInt(aKey(i)) <= UBound(aRec) Then
			strKey = strKey & aRec(aKey(i))
		End if
		i = i + 1
	Wend
	objTmpStr.setValue strKey
	getKey = objTmpStr.getHexString
End Function

Function isMultiKey(keys)
	objTmpStr.setValue keys
	isMultiKey = objTmpStr.isMatch("m$")
End Function
' User coding end
