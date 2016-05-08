'
' Run as cscript //NoLogo tee.vbs output
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
Dim objFsoOut : Set objFsoOut = new MyFso
Dim strRec
Dim strArrayFilesEtc
Dim intArrayFilesEtcIndex : intArrayFilesEtcIndex = 0
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
        objOpt.initialize("-h=n,--help=n")

        If objOpt.isSpecified("-h") or objOpt.isSpecified("--help") Then
                objStdio.writeLine "Usage : cscript //NoLogo tee.vbs output"
                objStdio.writeLine "Copy standard input to output, and also to standard output."
                objEndSw.turnOn
        End If

        strArrayFilesEtc = objOpt.getArrayNonOptions
        If objEndSw.isOff Then
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
	writeRec strRec
	objStdio.writeLine strRec

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
		objFsoOut.openOutput(strArrayFilesEtc(0))
	End If
End Function

Function closeFile
	If UBound(strArrayFilesEtc) >= 0 Then
		objFsoOut.close
	End If
End Function

Function readRec
	readRec = objStdio.readLine
	If objStdio.isReadFailure Then
		objEndSw.turnOn
	End If
End Function

Function writeRec(rec)
	objFsoOut.writeLine rec
End Function

Function isEof
	isEof = objStdio.isEof
End Function
' User coding end
