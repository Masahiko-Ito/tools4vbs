'
' Run as cscript //NoLogo prov.vbs [-p] [-d DELIMITER] -o OVERLAY.XLS [-i INPUT.CSV] -f FORMAT.TXT
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
Dim objOverlay : Set objOverlay = new MyExcel
Dim objInput : Set objInput = new MyFso
Dim objFormat : Set objFormat = new MyFso
Dim objPreviewSw : Set objPreviewSw = new MySwitch
Dim strDelimiter : strDelimiter = ","
Dim strOverlayFilename : strOverlayFilename = ""
Dim strInputFilename : strInputFilename = ""
Dim strFormatFilename : strFormatFilename = ""
Dim arrayFormatRec : arrayFormatRec = Array()
Dim strRec
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
	objOpt.initialize("-h=n,--help=n,-p=n,-d=y,-o=y,-i=y,-f=y")

	If objOpt.isSpecified("-h") or objOpt.isSpecified("--help") Then
		objStdio.writeLine "Usage : cscript //NoLogo prov.vbs [-p] [-d DELIMITER] -o OVERLAY.XLS [-i INPUT.CSV] -f FORMAT.TXT"
		objStdio.writeLine "Print formatted data with overlay."
		objStdio.writeLine ""
		objStdio.writeLine " -p                    preview mode."
		objStdio.writeLine " -d DELIMITER          use DELIMITER instead of comma for INPUT.CSV."
		objStdio.writeLine " -o OVERLAY.XLS        overlay definition by excel."
		objStdio.writeLine " -i INPUT.CSV          input data in csv format. If omitted then read stdin."
		objStdio.writeLine " -f FORMAT.TXT         format definition for INPUT.CSV."
		objStdio.writeLine "                       each line should have like ""1=A1"""
		objStdio.writeLine "                       ""1=A1"" means ""1st column"" in INPUT.CSV should be placed into ""A1"" cell in OVERLAY.XLS"
		objEndSw.turnOn
	End If

	If objEndSw.isOff Then
		If objOpt.isSpecified("-p") Then
			objPreviewSw.turnOn
		End If

		If objOpt.isSpecified("-d") Then
			strDelimiter = objOpt.getValue("-d")
		End If

		strOverlayFilename = objOpt.getValue("-o")
	        objOverlay.open(strOverlayFilename)

		strInputFilename = objOpt.getValue("-i")
		If strInputFilename <> "" Then
			objInput.openInput(strInputFilename)
		End If

		strFormatFilename = objOpt.getValue("-f")
		objFormat.openInput(strFormatFilename)

		Redim arrayFormatRec(-1)
	End If
' User coding end
End Sub
'--------------------------------------------------
' �J�n����
'
Sub sub_initialize
' User coding start
	Dim i
	Dim objStrFormatRec : Set objStrFormatRec = new MyString

	If objEndSw.isOff Then
		i = 0
		objStrFormatRec.setValue objFormat.readLine
		While objFormat.isReadSuccess
			If objStrFormatRec.isMatch("^[0-9][0-9]*=[A-za-z][A-Za-z]*[0-9][0-9]*$") Then
				Redim Preserve arrayFormatRec(i)
				arrayFormatRec(i) = objStrFormatRec.getValue
				i = i + 1
			End If
			objStrFormatRec.setValue objFormat.readLine
		Wend
		strRec = readRec
	End If

	Set objStrFormatRec = Nothing
' User coding end
End Sub
'--------------------------------------------------
' �又��
'
Sub sub_main
' User coding start
	Dim i
	Dim arrayIndexCell
	Dim objStr : Set objStr = new MyString
	Dim objCsv : Set objCsv = new MyCsv
	
	objCsv.setDelimiter strDelimiter
	objCsv.setRec strRec

	i = 0
	While i <= Ubound(arrayFormatRec)
		objStr.setValue arrayFormatRec(i)
		arrayIndexCell = Split(arrayFormatRec(i), "=")
		objOverlay.setCell 1, arrayIndexCell(1), objCsv.getValueByIndex(CInt(arrayIndexCell(0))-1)
		i = i + 1
	Wend

	If objPreviewSw.isOn Then
		objOverlay.preview 1
	Else
		objOverlay.print 1
	End If

	Set objStr = Nothing
	Set objCsv = Nothing

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
	If objOpt.isSpecified("-h") or objOpt.isSpecified("--help") Then
		' Do nothing
	Else
		objOverlay.close
		If strInputFilename <> "" Then
			objInput.close
		End If
		objFormat.close
	End If
' User coding end
End Sub
'--------------------------------------------------
' ���̑��̏���
'
' User coding start
Function readRec
	If strInputFilename = "" Then
		readRec = objStdio.readLine
		If objStdio.isReadFailure Then
			objEndSw.turnOn
		End If
	Else
		readRec = objInput.readLine
		If objInput.isReadFailure Then
			objEndSw.turnOn
		End If
	End If
End Function
' User coding end
