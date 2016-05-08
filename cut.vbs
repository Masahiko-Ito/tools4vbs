'
' Run as cscript //NoLogo cut.vbs -d "delimiter" -i "idx0,idx1, ..." [input ...]
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
Dim objSkipSw : Set objSkipSw = new MySwitch
Dim strArrayFilesEtc
Dim intArrayFilesEtcIndex : intArrayFilesEtcIndex = 0
Dim strDelimiter : strDelimiter = ","
Dim strArrayIndex
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
	objOpt.initialize("-h=n,--help=n,-d=y,-i=y")

	If objOpt.isSpecified("-h") or objOpt.isSpecified("--help") Then
		objStdio.writeLine "Usage : cscript //NoLogo cut.vbs [-d DELIMITER] -i INDEXIES [input ...]"
		objStdio.writeLine "Print selected parts of lines from each input to standard output."
		objStdio.writeLine ""
		objStdio.writeLine " -d DELIMITER        use DELIMITER instead of comma for field delimiter."
		objStdio.writeLine " -i INDEXIES         select only these fields(0 origin)."
		objStdio.writeLine "                     ex. 0,2,4 means 1st, 3rd, 5th fields should be selected."
		objEndSw.turnOn
	End If

	strArrayFilesEtc = objOpt.getArrayNonOptions
	If objEndSw.isOff Then
		If objOpt.isSpecified("-d") Then
			strDelimiter = objOpt.getValue("-d")
		End If
		strArrayIndex = Split(objOpt.getValue("-i"), ",")
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
		Dim strOutRec : strOutRec = ""
		Dim strArrayInValues
		Dim i

		strArrayInValues = Split(strRec, strDelimiter)
		i = 0
		While i < UBound(strArrayIndex)
			If CInt(strArrayIndex(i)) > UBound(strArrayInValues) Then
				strOutRec = strOutRec & "" & strDelimiter
			Else
				strOutRec = strOutRec & strArrayInValues(strArrayIndex(i)) & strDelimiter
			End If
			i = i + 1
		Wend
		If CInt(strArrayIndex(i)) > UBound(strArrayInValues) Then
			' do nothing
		Else
			strOutRec = strOutRec & strArrayInValues(strArrayIndex(i))
		End If
		objStdio.writeLine  strOutRec
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
