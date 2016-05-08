'
' Run as "cscript //NoLogo xls2csv.vbs [-sn SheetName|-si SheetIndex] [-i input.xls] [-o output.csv]"
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
Dim strInFile : strInFile = ""
Dim strOutFile : strOutFile = ""
Dim objDir : Set objDir = new MyDir
Dim objExcel : Set objExcel = new MyExcel
Dim objFilename : Set objFilename = new MyString
Dim objFsOpe : Set objFsOpe = new MyFsOpe
Dim strSheetName : strSheetName = ""
Dim intSheetIndex : intSheetIndex = 1
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
	objOpt.initialize("-h=n,--help=n,-sn=y,-si=y,-i=y,-o=y")

	If objOpt.isSpecified("-h") or objOpt.isSpecified("--help") Then
		objStdio.writeLine "Usage : cscript //NoLogo xls2csv.vbs [-sn SheetName|-si SheetIndex] [-i input.xls] [-o output.csv]"
		objStdio.writeLine "Convert a sheet of excel to csv."
		objStdio.writeLine ""
		objStdio.writeLine " -sn SheetName        SheetName must be specified in STRING."
		objStdio.writeLine " -si SheetIndex       SheetIndex must be specified in INTEGER."
		objStdio.writeLine "     If SheetName and SheetIndex are not specified, it is assumed that ""-si 1"" is specified."
		objStdio.writeLine " -i input.xls         specify input excel file."
		objStdio.writeLine " -o output.csv        specify output csv file."
		objStdio.writeLine "     If both of input.xls and output.csv are not specified, all *.xls and *.xlsx in current directory will be converted into *.csv."
		objStdio.writeLine "     If input.xls is specified and output.csv is not specified, input.xls will be converted into stdout."
		objEndSw.turnOn
	End If

	strArrayFilesEtc = objOpt.getArrayNonOptions
	If objEndSw.isOff Then
		If objOpt.isSpecified("-sn") Then
			strSheetName = objOpt.getValue("-sn")
		End If
		If objOpt.isSpecified("-si") Then
			intSheetIndex = CInt(objOpt.getValue("-si"))
		End If
		If objOpt.isSpecified("-i") Then
			strInFile = objOpt.getValue("-i")
		End If
		If objOpt.isSpecified("-o") Then
			strOutFile = objOpt.getValue("-o")
		End If
	End If

' User coding end
End Sub
'--------------------------------------------------
' 開始処理
'
Sub sub_initialize
' User coding start
	If objEndSw.isOff Then
		If strInFile = "" Then
			objDir.setDir(".")
			objFilename.setValue objDir.getFirstFilename
			If objFilename.getValue = "" Then
				objEndSw.turnOn
			End If
		End If
	End If
' User coding end
End Sub
'--------------------------------------------------
' 主処理
'
Sub sub_main
' User coding start
	Dim strCsvFilename
	If strInFile = "" Then

		If objFilename.isMatch("\.[Xx][Ll][Ss][XxMm]*$") Then
'			On Error Resume Next
			objExcel.open objFilename.getValue
			strCsvFilename = objFilename.getReplace("\.[Xx][Ll][Ss][XxMm]*$", ".csv")
			If strSheetName = "" Then
				objExcel.saveAsCsv intSheetIndex, strCsvFilename
			Else
				objExcel.saveAsCsv strSheetName, strCsvFilename
			End If
			objExcel.close
'			On Error Goto 0
			objStdio.writeLine objFilename.getValue & " -> " & strCsvFilename
		End If

		objFilename.setValue objDir.getNextFilename
		If objFilename.getValue = "" Then
			objEndSw.turnOn
		End If
	Else
'		On Error Resume Next
		objExcel.open strInFile
		If strOutFile = "" Then
			strCsvFilename = objFsOpe.getTempFileName
		Else
			strCsvFilename = strOutFile
		End If
		If strSheetName = "" Then
			objExcel.saveAsCsv intSheetIndex, strCsvFilename
		Else
			objExcel.saveAsCsv strSheetName, strCsvFilename
		End If
		objExcel.close
'		On Error Goto 0
		If strOutFile = "" Then
			Dim objTemp : Set objTemp = new MyFso
			Dim strRec
			objTemp.openInput strCsvFilename
			strRec = objTemp.readLine
			While objTemp.isReadSuccess
				objStdio.writeLine strRec
				strRec = objTemp.readLine
			Wend
			objTemp.close
			objFsOpe.deleteFile strCsvFilename
			Set objTemp = Nothing
		End If
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
' User coding end
End Sub
'--------------------------------------------------
' その他の処理
'
' User coding start
' User coding end
