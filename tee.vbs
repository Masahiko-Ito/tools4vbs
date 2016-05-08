'
' Run as cscript //NoLogo tee.vbs output
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
Dim objFsoOut : Set objFsoOut = new MyFso
Dim strRec
Dim strArrayFilesEtc
Dim intArrayFilesEtcIndex : intArrayFilesEtcIndex = 0
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
	writeRec strRec
	objStdio.writeLine strRec

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
