'
' Run as cscript //NoLogo cl.vbs [input ...]
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
Dim intCount : intCount = 0
Dim intTotalCount : intTotalCount = 0
Dim strCurrentFileName
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
                objStdio.writeLine "Usage : cscript //NoLogo cl.vbs [input ...]"
                objStdio.writeLine "Print newline counts for each input, and a total line if more than one input is specified."
                objStdio.writeLine "With no input, read standard input."
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
	If objSkipSw.isOff Then
		Dim strInRec

		strInRec = strRec
		intCount = intCount + 1
	End If

	strRec = readRec
' User coding end
End Sub
'--------------------------------------------------
' 終了処理
'
Sub sub_terminate
' User coding start
        If objOpt.isSpecified("-h") or objOpt.isSpecified("--help") Then
		' do nothing
	Else
		If UBound(strArrayFilesEtc) >= 0 Then
			objStdio.writeLine intCount & " " & strCurrentFileName
			intTotalCount = intTotalCount + intCount
			If UBound(strArrayFilesEtc) >= 1 Then
				objStdio.writeLine intTotalCount & " Total"
			End If
		Else
			objStdio.writeLine intCount
		End If
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
				objStdio.writeLine intCount & " " & strCurrentFileName
				intTotalCount = intTotalCount + intCount
                                closeFile
				intCount = 0
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
