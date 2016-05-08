'
' Run as cscript //NoLogo cat.vbs [-n] [input ...]
'
'--------------------------------------------------
' ���ʏ���
'
Option Explicit
Function include(filename)
	ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile(filename).ReadAll()
End Function
include("mytoolkit.vbs")
'--------------------------------------------------
' ��Ǖϐ�
'
Dim objEndSw : Set objEndSw = new MySwitch
Dim objMisc : Set objMisc = new MyMisc
' User coding start
Dim objOpt : Set objOpt = new MyOption
Dim objStdio : Set objStdio = new MyStdio
Dim objFsoIn : Set objFsoIn = new MyFso
Dim strRec
Dim objCountSw : Set objCountSw = new MySwitch 
Dim objSkipSw : Set objSkipSw = new MySwitch 
Dim strArrayFilesEtc
Dim intArrayFilesEtcIndex : intArrayFilesEtcIndex = 0
Dim intCount : intCount = 0
' User coding end
'--------------------------------------------------
' �����J�n
'
sub_open		' �I�[�v������
sub_initialize		' �J�n����
While objEndSw.isOff
	sub_main	' �又��
Wend
sub_terminate		' �I������
sub_close		' �N���[�Y����
objMisc.exitProg(0)
'--------------------------------------------------
' �I�[�v������
'
Sub sub_open
' User coding start
	objOpt.initialize("-h=n,--help=n,-n=n")

	If objOpt.isSpecified("-h") or objOpt.isSpecified("--help") Then
		objStdio.writeLine "Usage : cscript //NoLogo cat.vbs [-n] [input ...]"
		objStdio.writeLine "Concatenate input(s), or standard input, to standard output."
		objStdio.writeLine ""
		objStdio.writeLine " -n        number all output lines"
		objEndSw.turnOn
	End If

	strArrayFilesEtc = objOpt.getArrayNonOptions
	If objEndSw.isOff Then
		If objOpt.isSpecified("-n") Then
			objCountSw.turnOn
		End If
		openFile
	End If
' User coding end
End Sub
'--------------------------------------------------
' �J�n����
'
Sub sub_initialize
' User coding start
	If objEndSw.isOff Then
		strRec = readRec
	End If
' User coding end
End Sub
'--------------------------------------------------
' �又��
'
Sub sub_main
' User coding start
	If objSkipSw.isOff Then
		If objCountSw.isOn Then
			intCount = intCount + 1
			objStdio.writeLine intCount & " " & strRec
		Else
			objStdio.writeLine strRec
		End If
	End If

	strRec = readRec
' User coding end
End Sub
'--------------------------------------------------
' �I������
'
Sub sub_terminate
' User coding start
' User coding end
End Sub
'--------------------------------------------------
' �N���[�Y����
'
Sub sub_close
' User coding start
	closeFile
' User coding end
End Sub
'--------------------------------------------------
' ���̑��̏���
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
