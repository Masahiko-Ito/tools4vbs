Option Explicit
''==================================================
''
'' �W�����o�̓N���X
''
Class MyStdio
	Dim objStdin
	Dim objStdout
	Dim objStderr
	Dim bReadSuccess

	Sub Class_initialize()
		Set objStdin = WScript.StdIn
		Set objStdout = WScript.StdOut
		Set objStderr = WScript.StdErr
		bReadSuccess = false
	End Sub
'
' �@�@�\�F�W�����͂��P�s�擾
' �߂�l�F���R�[�h
' ��@�@�FstrRec = MyStdio.readLine
'
	Function readLine()
		If objStdin.AtEndOfStream Then
			readLine = ""
			bReadSuccess = false
		Else
			On Error Resume Next
			readLine = objStdin.ReadLine()
			If Err.Number = 0 Then
				bReadSuccess = true
			Else
				bReadSuccess = false
			End If
			On Error Goto 0
		End If
	End Function
'
' �@�@�\�F�W�����͂��w�蕶�������擾
' ���@���Fcount		�擾���镶����
' �߂�l�F������
' ��@�@�FstrRec = MyStdio.read(128)
'
	Function read(count)
		If objStdin.AtEndOfStream Then
			read = ""
			bReadSuccess = false
		Else
			On Error Resume Next
			read = objStdin.Read(count)
			If Err.Number = 0 Then
				bReadSuccess = true
			Else
				bReadSuccess = false
			End If
			On Error Goto 0
		End If
	End Function
'
' �@�@�\�F�W�����͓ǂݎ�萬���`�F�b�N
' �߂�l�Ftrue or false
' ��@�@�FWhile MyStdio.isRaedSuccess
'
	Function isReadSuccess()
		isReadSuccess = bReadSuccess
	End Function
'
' �@�@�\�F�W�����͓ǂݎ�莸�s�`�F�b�N
' �߂�l�Ftrue or false
' ��@�@�FWhile Not MyStdio.isRaedFailure
'
	Function isReadFailure()
		isReadFailure = Not isReadSuccess
	End Function
'
' �@�@�\�F�W���o�͂ɂP�s�o��
' ���@���Fstr		���R�[�h
' ��@�@�FMyStdio.writeLine strRec
'
	Function writeLine(str)
		objStdout.WriteLine str
	End Function
'
' �@�@�\�F�W���o�͂Ɏw�蕶������o��
' ���@���Fstr		������
' ��@�@�FMyStdio.write strString
'
	Function write(str)
		objStdout.Write str
	End Function
'
' �@�@�\�F�W���G���[�o�͂ɂP�s�o��
' ���@���Fstr		���R�[�h
' ��@�@�FMyStdio.writeErrorLine strRec
'
	Function writeErrorLine(str)
		objStderr.WriteLine str
	End Function
'
' �@�@�\�F�W���G���[�o�͂Ɏw�蕶������o��
' ���@���Fstr		������
' ��@�@�FMyStdio.writeError strString
'
	Function writeError(str)
		objStderr.Write str
	End Function
'
' �@�@�\�F�W�����͂̏I�[���擾
' �߁@�l�FTrue(�I�[��) or False(��I�[��)
' ��@�@�FWhile Not isEof
'
	Function isEof()
		isEof = objStdin.AtEndOfStream
	End Function
End Class
''==================================================
''
'' ������N���X
''
Class MyString
	Dim value
	Dim booleanHighValue
	Dim booleanLowValue
	Dim objRegex

	Sub Class_initialize()
		value = ""
		booleanHighValue = false
		booleanLowValue = false
		Set objRegex = CreateObject("VBScript.RegExp")
	End Sub
'
' �@�@�\�F�������ݒ�
' ���@���Fstr		������
' ��@�@�FMyString.setValue strString
'
	Function setValue(str)
		value = CStr(str)
		booleanHighValue = false
		booleanLowValue = false
	End Function
'
' �@�@�\�F��������擾
' �߁@�l�F������
' ��@�@�FstrString = MyString.getValue
'
	Function getValue()
		getValue = CStr(value)
	End Function
'
' �@�@�\�F���K�\���Ƃ̈�v���擾
' ���@���Fregex		���K�\��
' �߁@�l�FTrue(��v��) or False(�s��v��)
' ��@�@�FIf MyString.isMatch("^A.*Z$")
'
	Function isMatch(regex)
		objRegex.Pattern = regex
		isMatch = objRegex.Test(value)
	End Function
'
' �@�@�\�F���K�\���Ɉ�v����������ʂ̕�����ɒu�������邽������̎擾
' ���@���Fregex		���K�\��
'	  str		�u����̕�����
' �߁@�l�F������
' ��@�@�FstrString = MyString.getReplace("^A.*Z$", "A to Z")
'
	Function getReplace(regex, str)
		objRegex.Global = True
		objRegex.Pattern = regex
		getReplace = objRegex.Replace(value, str)
	End Function
'
' �@�@�\�F�w�肵����؂蕶���ŋ�؂��Ĕz��Ŏ擾
' ���@���Fdel		��؂蕶��
' �߁@�l�F������z��
' ��@�@�FstrArray = MyString.getSplit(",")
'
	Function getSplit(del)
		getSplit = split(value, del)
	End Function
'
' �@�@�\�F�����񒷂��擾
' �߁@�l�F������
' ��@�@�FintLength = MyString.getLength
'
	Function getLength()
		getLength = len(value)
	End Function
'
' �@�@�\�F������������擾
' ���@���Fstart		�J�n�ʒu(1�I���W��)
'	  length	������
' �߁@�l�F������
' ��@�@�FstrString = MyString.Substr(1,5)
'
	Function getSubstr(start, length)
		getSubstr = Mid(value, start, length)
	End Function

'
' �@�@�\�F�������16�i���R�[�h�ɕϊ�
' �߁@�l�F������
' ��@�@�FstrHexString = MyString.getHexString
'
	Function getHexString()
	    Dim strChar
	    Dim strHex
	    Dim intLen
	    Dim intCnt
	
	    strHex = ""
	    intLen = Len(value)
	    For intCnt = 1 To intLen
	        strChar = Mid(value, intCnt, 1)
	        strHex = strHex & Hex(Asc(strChar))
	    Next
	    getHexString = strHex
	End Function
'
' �@�@�\�FHIGH-VALUE���Z�b�g����
' ��@�@�FMyString.setHighValue
'
	Function setHighValue()
		booleanHighValue = true
		value = ""
	End Function
'
' �@�@�\�FHIGH-VALUE���m�F����
' �߁@�l�FTrue or False
' ��@�@�FMyString.isHighValue
'
	Function isHighValue()
		isHighValue = booleanHighValue
	End Function
'
' �@�@�\�FLOW-VALUE���Z�b�g����
' ��@�@�FMyString.setLowValue
'
	Function setLowValue()
		booleanLowValue = true
		value = ""
	End Function
'
' �@�@�\�FLOW-VALUE���m�F����
' �߁@�l�FTrue or False
' ��@�@�FMyString.isLowValue
'
	Function isLowValue()
		isLowValue = booleanLowValue
	End Function
'
' �@�@�\�F���������m�F����
' �߁@�l�FTrue or False
' ��@�@�FMyString.isEqual(objString)
'
	Function isEqual(objString)
		Dim stat
		If isHighValue Then
			If objString.isHighValue Then
				stat = true
			Else
				stat = false
			End If
		ElseIf isLowValue Then
			If objString.isLowValue Then
				stat = true
			Else
				stat = false
			End If
		Else
			If objString.isHighValue or objString.isLowValue Then
				stat = false
			Else
				If CStr(value) = CStr(objString.getValue) Then
					stat = true
				Else
					stat = false
				End If
			End If
		End If
		isEqual = stat
	End Function
'
' �@�@�\�F�傫�����m�F����
' �߁@�l�FTrue or False
' ��@�@�FMyString.isGreater(objString)
'
	Function isGreater(objString)
		Dim stat
		If isHighValue Then
			If objString.isHighValue Then
				stat = false
			Else
				stat = true
			End If
		ElseIf isLowValue Then
			stat = false
		Else
			If objString.isHighValue Then
				stat = false
			ElseIf  objString.isLowValue Then
				stat = True
			Else
				If CStr(value) > CStr(objString.getValue) Then
					stat = true
				Else
					stat = false
				End If
			End If
		End If
		isGreater = stat
	End Function
'
' �@�@�\�F���������m�F����
' �߁@�l�FTrue or False
' ��@�@�FMyString.isLess(objString)
'
	Function isLess(objString)
		Dim stat
		If isEqual(objString) Then
			stat = false
		ElseIf isGreater(objString) Then
			stat = false
		Else
			stat = true
		End If
		isLess = stat
	End Function
End Class
''==================================================
''
'' �����N���X
''
Class MyArg
	Dim objArg

	Sub Class_initialize()
		Set objArg = WScript.Arguments
	End Sub

'
' �@�@�\�F�����̌����擾
' �߁@�l�F�����̌�
' ��@�@�FintCount= MyArg.getCount
'
	Function getCount()
		getCount = objArg.Count
	End Function
'
' �@�@�\�F�w�肵�������̎擾
' ���@���F�C���f�b�N�X(0�I���W��)
' �߁@�l�F������
' ��@�@�FstrArg= MyArg.getValue(0)
'
	Function getValue(idx)
		getValue = objArg(idx)
	End Function
End Class
''==================================================
''
'' FileSystemObject�N���X(Shift_JIS)
''
Class MyFso
	Dim objFso
	Dim objFile
	Dim strFilename
	Dim strMode
	Dim bReadSuccess

	Sub Class_initialize()
		Set objFile = Nothing
		Set objFso = Nothing
		strFilename = ""
		strMode = ""
		bReadSuccess = false
	End Sub
'
' �@�@�\�FFSO���J��
' ���@���Ffilename		�t�@�C����
'	  mode			"r"(�ǂݍ���) or "w"(��������)
' ��@�@�FMyFso.open "INPUT.TXT", "r"
'
	Function open(filename, mode)
		Set objFso = CreateObject("Scripting.FileSystemObject")
		If mode = "r" Then
			Set objFile = objFso.OpenTextFile(filename, 1, False)
		Elseif mode = "w" Then
			Set objFile = objFso.OpenTextFile(filename, 2, True)
		End If
		strFilename = filename
		strMode = mode
	End Function
'
' �@�@�\�FFSO����̓��[�h�ŊJ��
' ���@���Ffilename		�t�@�C����
' ��@�@�FMyFso.openInput "INPUT.TXT"
'
	Function openInput(filename)
		open filename, "r"
	End Function
'
' �@�@�\�FFSO���o�̓��[�h�ŊJ��
' ���@���Ffilename		�t�@�C����
' ��@�@�FMyFso.openOutput "OUTPUT.TXT"
'
	Function openOutput(filename)
		open filename, "w"
	End Function
'
' �@�@�\�F�P�s�擾
' �߂�l�F���R�[�h
' ��@�@�FstrRec = MyFso.readLine
'
	Function readLine()
		If objFile.AtEndOfStream Then
			readLine = ""
			bReadSuccess = false
		Else
			On Error Resume Next
			readLine = objFile.ReadLine
			If Err.Number = 0 Then
				bReadSuccess = true
			Else
				bReadSuccess = false
			End If
			On Error Goto 0
		End If
	End Function
'
' �@�@�\�F�w�蕶�������擾
' ���@���Fcount		�擾���镶����
' �߂�l�F������
' ��@�@�FstrString = MyFso.read(128)
'
	Function read(count)
		If objFile.AtEndOfStream Then
			read = ""
			bReadSuccess = false
		Else
			On Error Resume Next
			read = objFile.Read(count)
			If Err.Number = 0 Then
				bReadSuccess = true
			Else
				bReadSuccess = false
			End If
			On Error Goto 0
		End If
	End Function
'
' �@�@�\�F�ǂݎ�萬���`�F�b�N
' �߂�l�Ftrue or false
' ��@�@�FWhile MyFso.isReadSuccess
'
	Function isReadSuccess()
		isReadSuccess = bReadSuccess
	End Function
'
' �@�@�\�F�ǂݎ�莸�s�`�F�b�N
' �߂�l�Ftrue or false
' ��@�@�FWhile Not MyFso.isReadFailure
'
	Function isReadFailure()
		isReadFailure = Not isReadSuccess
	End Function
'
' �@�@�\�F�P�s�o��
' ���@���Fstr		���R�[�h
' ��@�@�FMyFso.writeLine strRec
'
	Function writeLine(str)
		objFile.WriteLine str
	End Function
'
' �@�@�\�F�w�蕶������o��
' ���@���Fstr		������
' ��@�@�FMyFso.write strString
'
	Function write(str)
		objFile.Write str
	End Function
'
' �@�@�\�F�I�[���擾
' �߁@�l�FTrue(�I�[��) or False(��I�[��)
' ��@�@�FWhile Not MyFso.isEof
'
	Function isEof()
		isEof = objFile.AtEndOfStream
	End Function
'
' �@�@�\�FFSO�����
' ��@�@�FMyFso.close
'
	Function close()
		objFile.close
		Set objFile = Nothing
		Set objFso = Nothing
		strFilename = ""
		strMode = ""
	End Function
End Class
''==================================================
''
'' ActiveX Data Object�N���X(�����R�[�h�w��\)
''
Class MyAdo
	Dim objStream
	Dim strFilename
	Dim strMode
	Dim strCharset
	Dim booleanBomForUtf8
	Dim bReadSuccess

	Sub Class_initialize()
		Set objStream = Nothing
		strFilename = ""
		strMode = ""
		strCharset = ""
		booleanBomForUtf8 = False
		bReadSuccess = false
	End Sub
'
' �@�@�\�FADO���J��
' ���@���Ffilename		�t�@�C����
'	  mode			"r"(�ǂݍ���) or "w"(��������)
'	  charset		"UTF-8" or "Shift_JIS" or "ASCII" etc
'	  bom			True(UTF-8�̎�BOM�t���ŏo��) or
'				False(UTF-8�̎�BOM�����ŏo��)
' ��@�@�FMyAdo.open "INPUT.TXT", "r", "UTF-8", False
'
	Function open(filename, mode, charset, bom)
		Set objStream = CreateObject("ADODB.Stream")
		objStream.charset = charset
		objStream.type = 2	' text
		objStream.open
		If mode = "r" Then
			objStream.LoadFromFile filename
		End If
		strFilename = filename
		strMode = mode
		strCharset = charset
		booleanBomForUtf8 = bom
	End Function
'
' �@�@�\�FADO����̓��[�h�ŊJ��
' ���@���Ffilename		�t�@�C����
'	  charset		"UTF-8" or "Shift_JIS" or "ASCII" etc
' ��@�@�FMyAdo.openInput "INPUT.TXT", "UTF-8"
'
	Function openInput(filename, charset)
		open filename, "r", charset, False
	End Function
'
' �@�@�\�FADO���o�̓��[�h�ŊJ��
' ���@���Ffilename		�t�@�C����
'	  charset		"UTF-8" or "Shift_JIS" or "ASCII" etc
' ��@�@�FMyAdo.openOutput "OUTPUT.TXT", "UTF-8"
'
	Function openOutput(filename, charset)
		open filename, "w", charset, False
	End Function
'
' �@�@�\�F�P�s�擾
' �߂�l�F���R�[�h
' ��@�@�FstrRec = MyAdo.readLine
'
	Function readLine()
		If objStream.EOS Then
			readLine = ""
			bReadSuccess = false
		Else
			On Error Resume Next
			readLine = objStream.ReadText(-2) 
			If Err.Number = 0 Then
				bReadSuccess = true
			Else
				bReadSuccess = false
			End If
			On Error Goto 0
		End If
	End Function
'
' �@�@�\�F�ǂݎ�萬���`�F�b�N
' �߂�l�Ftrue or false
' ��@�@�FWhile MyAdo.isReadSuccess
'
	Function isReadSuccess()
		isReadSuccess = bReadSuccess
	End Function
'
' �@�@�\�F�ǂݎ�莸�s�`�F�b�N
' �߂�l�Ftrue or false
' ��@�@�FWhile Not MyAdo.isReadFailure
'
	Function isReadFailure()
		isReadFailure = Not isReadSuccess
	End Function
'
' �@�@�\�F�P�s�o��
' ���@���Fstr		���R�[�h
' ��@�@�FMyAdo.writeLine strRec
'
	Function writeLine(record)
		objStream.WriteText record, 1
	End Function
'
' �@�@�\�FADO�����
' ��@�@�FMyAdo.close
'
	Function close()
		If strMode = "w" Then
			If strCharset = "UTF-8" Then
				If booleanBomForUtf8 Then
					save()
				Else
					saveWithoutBom()
				End If
			Else
				save()
			End If
		End If
		objStream.close
		Set objStream = Nothing
	End Function
'
' �@�@�\�F�I�[���擾
' �߁@�l�FTrue(�I�[��) or False(��I�[��)
' ��@�@�FWhile Not MyAdo.isEof
'
	Function isEof()
		isEof = objStream.EOS
	End Function
'
' �@�@�\�F�ۑ�����
' ��@�@�FMyAdo.save
'
	Private Function save()
		objStream.SaveToFile strFilename, 2
	End Function
'
' �@�@�\�F�擪��3bytes�������ĕۑ�����
' ��@�@�FMyAdo.saveWithoutBom
'
	Private Function saveWithoutBom()
		objStream.Position = 0
		objStream.Type = 1
		objStream.Position = 3
		Dim bin : bin = objStream.Read()
		Dim stm : Set stm = CreateObject("ADODB.Stream")
		stm.Type = 1
		stm.Open()
		stm.Write(bin)
		stm.SaveToFile strFilename, 2
		stm.Close()
	End Function
End Class
''==================================================
''
'' CSV�N���X
''
Class MyCsv
	Dim strValues
	Dim strNames
	Dim strDel
	Dim objStr
	Dim objDict

	Sub Class_initialize()
		strValues = Array()
		strNames = ""
		strDel = ","
		Set objStr = new MyString
		Set objDict = CreateObject("Scripting.Dictionary")
	End Sub
'
' �@�@�\�F���R�[�h�̏�����
' ���@���Fcount		���R�[�h�̍��ڐ�
' ��@�@�FMyCsv.initRec 16
'
	Function initRec(count)
		ReDim strValues(count - 1)
	End Function
'
' �@�@�\�F���R�[�h�i��؂蕶���܂ށj��ݒ�
' ���@���Fstr		������i���ږ��P�o��؂蕶���p���ږ��Q�o��؂蕶���p...�j
' ��@�@�FMyCsv.setRec strRec
'
	Function setRec(str)
		objStr.setValue(str)
		strValues = objStr.getSplit(strDel)
	End Function
'
' �@�@�\�F���R�[�h���擾
' �߁@�l�F���R�[�h�i��؂蕶���܂ށj
' ��@�@�FstrRec = MyCsv.getRec
'
	Function getRec()
		Dim strRec
		Dim i

		strRec = ""
		If UBound(strValues) >= 0 Then
			strRec = strValues(0)
			For i = 1 to UBound(strValues)
				strRec = strRec & strDel & strValues(i)
			Next
		End If
		getRec = strRec
	End Function
'
' �@�@�\�F���ږ����R�[�h��ݒ�
' ���@���Fstr		������i���ږ��P�o��؂蕶���p���ږ��Q�o��؂蕶���p...�j
' ��@�@�FMyCsv.setNameRec strNameRec
'
	Function setNameRec(str)
		Dim strArray
		Dim i
		strArray = Split(str, strDel)
		For i = 0 to UBound(strArray)
			objDict.Add strArray(i), i
		Next
		strNames = strDel
	End Function
'
' �@�@�\�F���ږ����R�[�h���擾
' �߂�l�F������i���ږ��P�o��؂蕶���p���ږ��Q�o��؂蕶���p...�j
' ��@�@�FstrNameRec = MyCsv.getNameRec
'
	Function getNameRec()
		getNameRec = strNames
	End Function
'
' �@�@�\�F��؂蕶����ݒ�
' ���@���Fdel		��؂蕶����
' ��@�@�FMyCsv.setDelimiter ","
'
	Function setDelimiter(del)
		strDel = del
	End Function
'
' �@�@�\�F��؂蕶�����擾
' �߂�l�F����
' ��@�@�FcharDel = MyCsv.getDelimiter
'
	Function getDelimiter()
		getDelimiter = strDel
	End Function
'
' �@�@�\�F�C���f�b�N�X�w��ō��ڂ��擾
' ���@���Fi		�C���f�b�N�X�i�O�I���W���j
' �߂�l�F������
' ��@�@�FstrString = MyCsv.getValueByIndex(0)
'
	Function getValueByIndex(i)
		If 0 <= i And i <= UBound(strValues) Then
			getValueByIndex = strValues(i)
		Else
			getValueByIndex = ""
		End If
	End Function
'
' �@�@�\�F�C���f�b�N�X�w��ō��ڂ�ݒ�
' ���@���Fi		�C���f�b�N�X�i�O�I���W���j
'	  value		������
' �߂�l�F������
' ��@�@�FMyCsv.setValueByIndex 0, "STRING"
'
	Function setValueByIndex(i, value)
		If 0 <= i And i <= UBound(strValues) Then
			strValues(i) = value
		End If
	End Function
'
' �@�@�\�F�ŏI�C���f�b�N�X���擾
' �߂�l�F�����i�O�I���W���j
' ��@�@�FintIndex = MyCsv.getLastIndex
'
	Function getLastIndex()
		getLastIndex = UBound(strValues)
	End Function
'
' �@�@�\�F���ږ��ɑΉ������C���f�b�N�X���擾
' ���@���Fname		���ږ�
' �߂�l�F�����i�O�I���W���j
' ��@�@�FintIndex = MyCsv.getIndexByName("ELEMENT_01")
'
	Function getIndexByName(name)
		If objDict.Exists(name) Then
			getIndexByName = objDict(name)
		Else
			getIndexByName = -1
		End If
	End Function
'
' �@�@�\�F���ږ��w��ō��ڂ��擾
' ���@���Fname		���ږ�
' �߂�l�F������
' ��@�@�FstrString = MyCsv.getValueByName("ELEMENT_01")
'
	Function getValueByName(name)
		getValueByName = getValueByIndex(getIndexByName(name))
	End Function
'
' �@�@�\�F���ږ��w��ō��ڂ�ݒ�
' ���@���Fname		���ږ�
'	  value		������
' �߂�l�F������
' ��@�@�FMyCsv.setValueByName "ELEMENT_01", 123
'
	Function setValueByName(name, value)
		setValueByIndex getIndexByName(name), value
	End Function
End Class
''==================================================
''
'' �W�����o��CSV�N���X
''
Class MyStdioCsv
	Dim objStdio
	Dim objCsv

	Sub Class_initialize()
		Set objStdio = new MyStdio
		Set objCsv = new MyCsv
	End Sub
'
' �@�@�\�F���R�[�h�̏�����
' ���@���Fcount		���R�[�h�̍��ڐ�
' ��@�@�FMyStdioCsv.initRec 16
'
	Function initRec(count)
		objCsv.initRec count
	End Function
'
' �@�@�\�F���R�[�h�i��؂蕶���܂ށj��ݒ�
' ���@���Fstr		������i���ږ��P�o��؂蕶���p���ږ��Q�o��؂蕶���p...�j
' ��@�@�FMyStdioCsv.setRec strRec
'
	Function setRec(str)
		objCsv.setRec str
	End Function
'
' �@�@�\�F���R�[�h���擾
' �߁@�l�F���R�[�h�i��؂蕶���܂ށj
' ��@�@�FstrRec = MyStdioCsv.getRec
'
	Function getRec()
		getRec = objCsv.getRec
	End Function
'
' �@�@�\�F���ږ����R�[�h��ݒ�
' ���@���Fstr		������i���ږ��P�o��؂蕶���p���ږ��Q�o��؂蕶���p...�j
' ��@�@�FMyStdioCsv.setNameRec strNameRec
'
	Function setNameRec(str)
		objCsv.setNameRec str
	End Function
'
' �@�@�\�F���ږ����R�[�h���擾
' �߂�l�F������i���ږ��P�o��؂蕶���p���ږ��Q�o��؂蕶���p...�j
' ��@�@�FstrNameRec = MyStdioCsv.getNameRec
'
	Function getNameRec()
		getNameRec = objCsv.getNameRec
	End Function
'
' �@�@�\�F��؂蕶����ݒ�
' ���@���Fdel		��؂蕶����
' ��@�@�FMyStdioCsv.setDelimiter ","
'
	Function setDelimiter(del)
		objCsv.setDelimiter del
	End Function
'
' �@�@�\�F��؂蕶�����擾
' �߂�l�F����
' ��@�@�FcharDel = MyStdioCsv.getDelimiter
'
	Function getDelimiter()
		getDelimiter = objCsv.getDelimiter
	End Function
'
' �@�@�\�F�C���f�b�N�X�w��ō��ڂ��擾
' ���@���Fi		�C���f�b�N�X�i�O�I���W���j
' �߂�l�F������
' ��@�@�FstrString = MyStdioCsv.getValueByIndex(0)
'
	Function getValueByIndex(i)
		getValueByIndex = objCsv.getValueByIndex(i)
	End Function
'
' �@�@�\�F�C���f�b�N�X�w��ō��ڂ�ݒ�
' ���@���Fi		�C���f�b�N�X�i�O�I���W���j
'	  value		������
' �߂�l�F������
' ��@�@�FMyStdioCsv.setValueByIndex 0, "STRING"
'
	Function setValueByIndex(i, value)
		objCsv.setValueByIndex i, value
	End Function
'
' �@�@�\�F�ŏI�C���f�b�N�X���擾
' �߂�l�F�����i�O�I���W���j
' ��@�@�FintIndex = MyStdioCsv.getLastIndex
'
	Function getLastIndex()
		getLastIndex = objCsv.getLastIndex
	End Function
'
' �@�@�\�F���ږ��ɑΉ������C���f�b�N�X���擾
' ���@���Fname		���ږ�
' �߂�l�F�����i�O�I���W���j
' ��@�@�FintIndex = MyStdioCsv.getIndexByName("ELEMENT_01")
'
	Function getIndexByName(name)
		getIndexByName = objCsv.getIndexByName(name)
	End Function
'
' �@�@�\�F���ږ��w��ō��ڂ��擾
' ���@���Fname		���ږ�
' �߂�l�F������
' ��@�@�FstrString = MyStdioCsv.getValueByName("ELEMENT_01")
'
	Function getValueByName(name)
		getValueByName = objCsv.getValueByName(name)
	End Function
'
' �@�@�\�F���ږ��w��ō��ڂ�ݒ�
' ���@���Fname		���ږ�
'	  value		������
' �߂�l�F������
' ��@�@�FMyStdioCsv.setValueByName "ELEMENT_01", 123
'
	Function setValueByName(name, value)
		objCsv.setValueByName name, value
	End Function
'
' �@�@�\�F�W�����͂��P�sCSV���R�[�h�擾���擾���AobjCsv�֕ۑ�
' �߂�l�FTrue(�ǂ߂�) or False(�ǂ߂Ȃ�����)
' ��@�@�FMyStdioCsv.readLine
'
	Function readLine()
		Dim strRec
		strRec = objStdio.readLine
		If objStdio.isReadSuccess Then
			objCsv.setRec strRec
			readLine = True
		Else
			objCsv.setRec ""
			readLine = False
		End If
	End Function
'
' �@�@�\�FobjCsv����W���o�͂ɂP�sCSV���R�[�h�o��
' ��@�@�FMyStdioCsv.writeLine
'
	Function writeLine()
		objStdio.writeLine objCsv.getRec
	End Function
'
' �@�@�\�FobjCsv����W���G���[�o�͂�CSV���R�[�h�o��
' ���@���Fstr		���R�[�h
' ��@�@�FMyStdioCsv.writeErrorLine
'
	Function writeErrorLine
		objStdio.writeErrorLine objCsv.getRec
	End Function
'
' �@�@�\�F�W�����͓ǂݎ�萬���`�F�b�N
' �߂�l�Ftrue or false
' ��@�@�FWhile MyStdioCsv.isRaedSuccess
'
	Function isReadSuccess()
		isReadSuccess = objStdio.isReadSuccess
	End Function
'
' �@�@�\�F�W�����͓ǂݎ�莸�s�`�F�b�N
' �߂�l�Ftrue or false
' ��@�@�FWhile Not MyStdioCsv.isRaedFailure
'
	Function isReadFailure()
		isReadFailure = objStdio.isReadFailure
	End Function
'
' �@�@�\�F�W�����͂̏I�[���擾
' �߁@�l�FTrue(�I�[��) or False(��I�[��)
' ��@�@�FWhile Not MyStdioCsv.isEof
'
	Function isEof()
		isEof = objStdio.isEof
	End Function
End Class
''==================================================
''
'' FileSystemObjectCSV�N���X(Shift_JIS)
''
Class MyFsoCsv
	Dim objFso
	Dim objCsv

	Sub Class_initialize()
		Set objFso = new MyFso
		Set objCsv = new MyCsv
	End Sub
'
' �@�@�\�F���R�[�h�̏�����
' ���@���Fcount		���R�[�h�̍��ڐ�
' ��@�@�FMyFsoCsv.initRec 16
'
	Function initRec(count)
		objCsv.initRec count
	End Function
'
' �@�@�\�F���R�[�h�i��؂蕶���܂ށj��ݒ�
' ���@���Fstr		������i���ږ��P�o��؂蕶���p���ږ��Q�o��؂蕶���p...�j
' ��@�@�FMyFsoCsv.setRec strRec
'
	Function setRec(str)
		objCsv.setRec str
	End Function
'
' �@�@�\�F���R�[�h���擾
' �߁@�l�F���R�[�h�i��؂蕶���܂ށj
' ��@�@�FstrRec = MyFsoCsv.getRec
'
	Function getRec()
		getRec = objCsv.getRec
	End Function
'
' �@�@�\�F���ږ����R�[�h��ݒ�
' ���@���Fstr		������i���ږ��P�o��؂蕶���p���ږ��Q�o��؂蕶���p...�j
' ��@�@�FMyFsoCsv.setNameRec strNameRec
'
	Function setNameRec(str)
		objCsv.setNameRec str
	End Function
'
' �@�@�\�F���ږ����R�[�h���擾
' �߂�l�F������i���ږ��P�o��؂蕶���p���ږ��Q�o��؂蕶���p...�j
' ��@�@�FstrNameRec = MyFsoCsv.getNameRec
'
	Function getNameRec()
		getNameRec = objCsv.getNameRec
	End Function
'
' �@�@�\�F��؂蕶����ݒ�
' ���@���Fdel		��؂蕶����
' ��@�@�FMyFsoCsv.setDelimiter ","
'
	Function setDelimiter(del)
		objCsv.setDelimiter del
	End Function
'
' �@�@�\�F��؂蕶�����擾
' �߂�l�F����
' ��@�@�FcharDel = MyFsoCsv.getDelimiter
'
	Function getDelimiter()
		getDelimiter = objCsv.getDelimiter
	End Function
'
' �@�@�\�F�C���f�b�N�X�w��ō��ڂ��擾
' ���@���Fi		�C���f�b�N�X�i�O�I���W���j
' �߂�l�F������
' ��@�@�FstrString = MyFsoCsv.getValueByIndex(0)
'
	Function getValueByIndex(i)
		getValueByIndex = objCsv.getValueByIndex(i)
	End Function
'
' �@�@�\�F�C���f�b�N�X�w��ō��ڂ�ݒ�
' ���@���Fi		�C���f�b�N�X�i�O�I���W���j
'	  value		������
' �߂�l�F������
' ��@�@�FMyFsoCsv.setValueByIndex 0, "STRING"
'
	Function setValueByIndex(i, value)
		objCsv.setValueByIndex i, value
	End Function
'
' �@�@�\�F�ŏI�C���f�b�N�X���擾
' �߂�l�F�����i�O�I���W���j
' ��@�@�FintIndex = MyFsoCsv.getLastIndex
'
	Function getLastIndex()
		getLastIndex = objCsv.getLastIndex
	End Function
'
' �@�@�\�F���ږ��ɑΉ������C���f�b�N�X���擾
' ���@���Fname		���ږ�
' �߂�l�F�����i�O�I���W���j
' ��@�@�FintIndex = MyFsoCsv.getIndexByName("ELEMENT_01")
'
	Function getIndexByName(name)
		getIndexByName = objCsv.getIndexByName(name)
	End Function
'
' �@�@�\�F���ږ��w��ō��ڂ��擾
' ���@���Fname		���ږ�
' �߂�l�F������
' ��@�@�FstrString = MyFsoCsv.getValueByName("ELEMENT_01")
'
	Function getValueByName(name)
		getValueByName = objCsv.getValueByName(name)
	End Function
'
' �@�@�\�F���ږ��w��ō��ڂ�ݒ�
' ���@���Fname		���ږ�
'	  value		������
' �߂�l�F������
' ��@�@�FMyFsoCsv.setValueByName "ELEMENT_01", 123
'
	Function setValueByName(name, value)
		objCsv.setValueByName name, value
	End Function
'
' �@�@�\�F�J��
' ���@���Ffilename		�t�@�C����
'	  mode			"r"(�ǂݍ���) or "w"(��������)
' ��@�@�FMyFsoCsv.open "INPUT.TXT", "r"
'
	Function open(filename, mode)
		objFso.open filename, mode
	End Function
'
' �@�@�\�F���̓��[�h�ŊJ��
' ���@���Ffilename		�t�@�C����
' ��@�@�FMyFsoCsv.openInput "INPUT.TXT"
'
	Function openInput(filename)
		objFso.openInput filename
	End Function
'
' �@�@�\�F�o�̓��[�h�ŊJ��
' ���@���Ffilename		�t�@�C����
' ��@�@�FMyFsoCsv.openOutput "OUTPUT.TXT"
'
	Function openOutput(filename)
		objFso.openOutput filename
	End Function
'
' �@�@�\�FFSO���P�sCSV���R�[�h�擾���AobjCsv�֕ۑ�
' �߂�l�FTrue(�ǂ߂�) or False(�ǂ߂Ȃ�����)
' ��@�@�FMyFsoCsv.readLine
'
	Function readLine()
		Dim strRec
		strRec = objFso.readLine
		If objFso.isReadSuccess Then
			objCsv.setRec strRec
			readLine = True
		Else
			objCsv.setRec ""
			readLine = False
		End If
	End Function
'
' �@�@�\�F�ǂݎ�萬���`�F�b�N
' �߂�l�Ftrue or false
' ��@�@�FWhile MyFsoCsv.isReadSuccess
'
	Function isReadSuccess()
		isReadSuccess = objFso.isReadSuccess
	End Function
'
' �@�@�\�F�ǂݎ�莸�s�`�F�b�N
' �߂�l�Ftrue or false
' ��@�@�FWhile Not MyFsoCsv.isReadFailure
'
	Function isReadFailure()
		isReadFailure = objFso.isReadFailure
	End Function
'
' �@�@�\�FobjCsv����FSO�ɂP�sCSV���R�[�h�o��
' ��@�@�FMyFsoCsv.writeLine
'
	Function writeLine()
		objFso.writeLine objCsv.getRec
	End Function
'
' �@�@�\�F�I�[���擾
' �߁@�l�FTrue(�I�[��) or False(��I�[��)
' ��@�@�FWhile Not MyFsoCsv.isEof
'
	Function isEof()
		isEof = objFso.isEof
	End Function
'
' �@�@�\�F����
' ��@�@�FMyFsoCsv.close
'
	Function close()
		objFso.close
	End Function
End Class
''==================================================
''
'' ActiveX Data Object CSV�N���X(�����R�[�h�w��\)
''
Class MyAdoCsv
	Dim objAdo
	Dim objCsv

	Sub Class_initialize()
		Set objAdo = new MyAdo
		Set objCsv = new MyCsv
	End Sub
'
' �@�@�\�F���R�[�h�̏�����
' ���@���Fcount		���R�[�h�̍��ڐ�
' ��@�@�FMyAdoFso.initRec 16
'
	Function initRec(count)
		objCsv.initRec count
	End Function
'
' �@�@�\�F���R�[�h�i��؂蕶���܂ށj��ݒ�
' ���@���Fstr		������i���ږ��P�o��؂蕶���p���ږ��Q�o��؂蕶���p...�j
' ��@�@�FMyAdoFso.setRec strRec
'
	Function setRec(str)
		objCsv.setRec str
	End Function
'
' �@�@�\�F���R�[�h���擾
' �߁@�l�F���R�[�h�i��؂蕶���܂ށj
' ��@�@�FstrRec = MyAdoFso.getRec
'
	Function getRec()
		getRec = objCsv.getRec
	End Function
'
' �@�@�\�F���ږ����R�[�h��ݒ�
' ���@���Fstr		������i���ږ��P�o��؂蕶���p���ږ��Q�o��؂蕶���p...�j
' ��@�@�FMyAdoFso.setNameRec strNameRec
'
	Function setNameRec(str)
		objCsv.setNameRec str
	End Function
'
' �@�@�\�F���ږ����R�[�h���擾
' �߂�l�F������i���ږ��P�o��؂蕶���p���ږ��Q�o��؂蕶���p...�j
' ��@�@�FstrNameRec = MyAdoFso.getNameRec
'
	Function getNameRec()
		getNameRec = objCsv.getNameRec
	End Function
'
' �@�@�\�F��؂蕶����ݒ�
' ���@���Fdel		��؂蕶����
' ��@�@�FMyAdoFso.setDelimiter ","
'
	Function setDelimiter(del)
		objCsv.setDelimiter del
	End Function
'
' �@�@�\�F��؂蕶�����擾
' �߂�l�F����
' ��@�@�FcharDel = MyAdoFso.getDelimiter
'
	Function getDelimiter()
		getDelimiter = objCsv.getDelimiter
	End Function
'
' �@�@�\�F�C���f�b�N�X�w��ō��ڂ��擾
' ���@���Fi		�C���f�b�N�X�i�O�I���W���j
' �߂�l�F������
' ��@�@�FstrString = MyAdoFso.getValueByIndex(0)
'
	Function getValueByIndex(i)
		getValueByIndex = objCsv.getValueByIndex(i)
	End Function
'
' �@�@�\�F�C���f�b�N�X�w��ō��ڂ�ݒ�
' ���@���Fi		�C���f�b�N�X�i�O�I���W���j
'	  value		������
' �߂�l�F������
' ��@�@�FMyAdoFso.setValueByIndex 0, "STRING"
'
	Function setValueByIndex(i, value)
		objCsv.setValueByIndex i, value
	End Function
'
' �@�@�\�F�ŏI�C���f�b�N�X���擾
' �߂�l�F�����i�O�I���W���j
' ��@�@�FintIndex = MyAdoFso.getLastIndex
'
	Function getLastIndex()
		getLastIndex = objCsv.getLastIndex
	End Function
'
' �@�@�\�F���ږ��ɑΉ������C���f�b�N�X���擾
' ���@���Fname		���ږ�
' �߂�l�F�����i�O�I���W���j
' ��@�@�FintIndex = MyAdoFso.getIndexByName("ELEMENT_01")
'
	Function getIndexByName(name)
		getIndexByName = objCsv.getIndexByName(name)
	End Function
'
' �@�@�\�F���ږ��w��ō��ڂ��擾
' ���@���Fname		���ږ�
' �߂�l�F������
' ��@�@�FstrString = MyAdoFso.getValueByName("ELEMENT_01")
'
	Function getValueByName(name)
		getValueByName = objCsv.getValueByName(name)
	End Function
'
' �@�@�\�F���ږ��w��ō��ڂ�ݒ�
' ���@���Fname		���ږ�
'	  value		������
' �߂�l�F������
' ��@�@�FMyAdoFso.setValueByName "ELEMENT_01", 123
'
	Function setValueByName(name, value)
		objCsv.setValueByName name, value
	End Function
'
' �@�@�\�F�J��
' ���@���Ffilename		�t�@�C����
'	  mode			"r"(�ǂݍ���) or "w"(��������)
'	  charset		"UTF-8" or "Shift_JIS" or "ASCII" etc
'	  bom			True(UTF-8�̎�BOM�t���ŏo��) or
'				False(UTF-8�̎�BOM�����ŏo��)
' ��@�@�FMyAdoCsv.open "INPUT.TXT", "r", "UTF-8", False
'
	Function open(filename, mode, charset, bom)
		objAdo.open filename, mode, charset, bom
	End Function
'
' �@�@�\�F���̓��[�h�ŊJ��
' ���@���Ffilename		�t�@�C����
'	  charset		"UTF-8" or "Shift_JIS" or "ASCII" etc
' ��@�@�FMyAdoCsv.openInput "INPUT.TXT", "UTF-8"
'
	Function openInput(filename, charset)
		objAdo.openInput filename, charset
	End Function
'
' �@�@�\�F�o�̓��[�h�ŊJ��
' ���@���Ffilename		�t�@�C����
'	  charset		"UTF-8" or "Shift_JIS" or "ASCII" etc
' ��@�@�FMyAdoCsv.openOutput "OUTPUT.TXT", "UTF-8"
'
	Function openOutput(filename, charset)
		objAdo.openOutput filename, charset
	End Function
'
' �@�@�\�FADO���P�sCSV���R�[�h�擾���AobjCsv�֕ۑ�
' �߂�l�FTrue(�ǂ߂�) or False(�ǂ߂Ȃ�����)
' ��@�@�FMyAdoCsv.readLine
'
	Function readLine()
		Dim strRec
		strRec = objAdo.readLine
		If objAdo.isReadSuccess Then
			objCsv.setRec strRec
			readLine = True
		Else
			objCsv.setRec ""
			readLine = False
		End If
	End Function
'
' �@�@�\�F�ǂݎ�萬���`�F�b�N
' �߂�l�Ftrue or false
' ��@�@�FWhile MyAdoCsv.isReadSuccess
'
	Function isReadSuccess()
		isReadSuccess = objAdo.isReadSuccess
	End Function
'
' �@�@�\�F�ǂݎ�莸�s�`�F�b�N
' �߂�l�Ftrue or false
' ��@�@�FWhile Not MyAdoCsv.isReadFailure
'
	Function isReadFailure()
		isReadFailure = objAdo.isReadFailure
	End Function
'
' �@�@�\�FobjCsv����ADO�ɂP�sCSV���R�[�h�o��
' ��@�@�FMyAdoCsv.writeLine
'
	Function writeLine()
		objAdo.writeLine objCsv.getRec
	End Function
'
' �@�@�\�F����
' ��@�@�FMyAdoCsv.close
'
	Function close()
		objAdo.close
	End Function
'
' �@�@�\�F�I�[���擾
' �߁@�l�FTrue(�I�[��) or False(��I�[��)
' ��@�@�FWhile Not MyAdoCsv.isEof
'
	Function isEof()
		objAdo.isEof
	End Function
'
' �@�@�\�F�ۑ�����
' ��@�@�FMyAdoCsv.save
'
	Private Function save()
		objAdo.save
	End Function
'
' �@�@�\�F�擪��3bytes�������ĕۑ�����
' ��@�@�FMyAdoCsv.saveWithoutBom
'
	Private Function saveWithoutBom()
		objAdo.saveWithoutBom
	End Function
End Class
''==================================================
''
'' SORT�N���X
''
Class MySort
	Dim objRs
	Dim strDelimiter
	Dim strKey
	Dim bGetSuccess

	Sub Class_initialize()
		Set objRs = CreateObject("ADODB.Recordset")
		bGetSuccess = false
	End Sub
'
' �@�@�\�F���͂b�r�u���R�[�h�̋�؂蕶�����w�肷��
' ���@���Fstr		��؂蕶��
' ��@�@�FMySort.setDelimiter ","
'
	Function setDelimiter(str)
		strDelimiter = str
	End Function
'
' �@�@�\�F�\�[�g�L�[���ڂ�ݒ肷��
' ���@���Fstr		"index:seq:type:max_len, ... ,max_rec_len"
' ��@�@�FMySort.setKey "0:A:H:256,1:D:H:256,65535"
'         index               index number for sort-key in csv format(0 origin)
'         seq                 A or D (ascending or descending)
'         type                H or X or S for raw character code
'                             200         for VarChar
'                             14 or N     for Decimal
'                             4           for Single
'                             5           for Double
'                             3           for Integer
'                             20          for BigInt
'         max_length          max length for key
'         max_rec_length      max length for input rec
'
	Function setKey(str)
		Dim strKeys
		Dim strKeyDetails
		Dim i

		strKey = str
		strKeys = split(strKey, ",")
		i = 0
		While i < UBound(strKeys)
			strKeyDetails = split(strKeys(i), ":")
			If strKeyDetails(2) = "X" or strKeyDetails(2) = "H" or strKeyDetails(2) = "S" Then
				If strKeyDetails(3) = "" Then
					objRs.Fields.Append "k" & strKeyDetails(0), 200
				Else
					objRs.Fields.Append "k" & strKeyDetails(0), 200, Cint(strKeyDetails(3)) * 2
				End If
			ElseIf strKeyDetails(2) = "N" Then
				If strKeyDetails(3) = "" Then
					objRs.Fields.Append "k" & strKeyDetails(0), 14
				Else
					objRs.Fields.Append "k" & strKeyDetails(0), 14, strKeyDetails(3)
				End If
			Else
				If strKeyDetails(3) = "" Then
					objRs.Fields.Append "k" & strKeyDetails(0), strKeyDetails(2)
				Else
					objRs.Fields.Append "k" & strKeyDetails(0), strKeyDetails(2), strKeyDetails(3)
				End If
			End If
			i = i + 1
		Wend
		strKeyDetails = split(strKeys(i), ":")
		objRs.Fields.Append "data", 200, strKeyDetails(0)
		objRs.Open
	End Function
'
' �@�@�\�F�\�[�g�����ɓ��̓��R�[�h��n��
' ���@���Fstr		���̓��R�[�h
' ��@�@�FMySort.putRec strRec
'
	Function putRec(str)
		Dim strRecs
		Dim strKeys
		Dim strKeyDetails
		Dim i
		Dim objStr : Set objStr = new MyString

		objRs.AddNew
		strRecs = split(str & strDelimiter, strDelimiter)
		strKeys = split(strKey, ",")
		i = 0
		While i < UBound(strKeys)
			strKeyDetails = split(strKeys(i), ":")
			If strKeyDetails(2) = "X" or strKeyDetails(2) = "H" or strKeyDetails(2) = "S" Then
				objStr.setValue(strRecs(strKeyDetails(0)))
				objRs.Fields("k" & strKeyDetails(0)).Value = objStr.getHexString
			Else
				objRs.Fields("k" & strKeyDetails(0)).Value = strRecs(strKeyDetails(0))
			End If
			i = i + 1
		Wend
		objRs.Fields("data").Value = str
	End Function
'
' �@�@�\�F�\�[�g�������s��
' ��@�@�FMySort.sort
'
	Function sort()
		Dim strKeys
		Dim strKeyDetails
		Dim i
		Dim strSortKey
		Dim strSeq

		strKeys = split(strKey, ",")
		i = 0
		While i < UBound(strKeys)
			strKeyDetails = split(strKeys(i), ":")
			If strKeyDetails(1) = "A" Then
				strSeq = "ASC"
			Else
				strSeq = "DESC"
			End If
			If strSortKey = "" Then
				strSortKey = "k" & strKeyDetails(0) & " " & strSeq
			Else
				strSortKey = strSortKey & "," & "k" & strKeyDetails(0) & " " & strSeq
			End If
			i = i + 1
		Wend
		objRs.Sort = strSortKey
		objRs.MoveFirst
	End Function
'
' �@�@�\�F�\�[�g���ʂ����o��
' ��@�@�FstrRec = MySort.getRec
'
	Function getRec()
		If objRs.EOF Then
			getRec = ""
			bGetSuccess = false
		Else
			getRec = objRs.Fields("data").Value
			objRs.MoveNext
			bGetSuccess = true
		End If
	End Function
'
' �@�@�\�F�\�[�g���ʓǂݎ�萬���`�F�b�N
' �߂�l�Ftrue or false
' ��@�@�FWhile isGetSuccess
'
	Function isGetSuccess()
		isGetSuccess = bGetSuccess
	End Function
'
' �@�@�\�F�\�[�g���ʓǂݎ�莸�s�`�F�b�N
' �߂�l�Ftrue or false
' ��@�@�FWhile Not isGetFailure
'
	Function isGetFailure()
		isGetFailure = Not isGetSuccess
	End Function
'
' �@�@�\�F�\�[�g���ʂ̏I�[���擾
' �߁@�l�FTrue(�I�[��) or False(��I�[��)
' ��@�@�FWhile Not isEof
'
	Function isEof()
		isEof = objRs.EOF
	End Function
End Class
''==================================================
''
'' �X�C�b�`�������N���X
''
Class MySwitch
	Dim switch

	Sub Class_Initialize()
		switch = 0
	End Sub
'
' �@�@�\�F�X�C�b�`�n�m
' ��@�@�FMySwitch.turnOn
'
	Function turnOn
		switch = 1
	End Function
'
' �@�@�\�F�X�C�b�`�n�e�e
' ��@�@�FMySwitch.turnOff
'
	Function turnOff
		switch = 0
	End Function
'
' �@�@�\�F�n�m���`�F�b�N
' ��@�@�FIf MySwitch.isOn Then
'
	Function isOn
		Dim stat
		If switch = 1 Then
			stat = true
		Else
			stat = false
		End If
		isOn = stat
	End Function
'
' �@�@�\�F�n�e�e���`�F�b�N
' ��@�@�FIf MySwitch.isOff Then
'
	Function isOff
		isOff = (Not isOn)
	End Function
End Class
''==================================================
''
'' �f�B���N�g���������N���X
''
Class MyDir
	Dim objFso
	Dim strDir
	Dim strFiles
	Dim intIndexStrFiles
	Dim strDirs
	Dim intIndexStrDirs

	Sub Class_Initialize()
		Set objFSo = CreateObject("Scripting.FileSystemObject")
		strDir = "."
		intIndexStrFiles = -1
		intIndexStrDirs = -1
	End Sub

'
' �@�@�\�F���΃p�X�Ńf�B���N�g�����w�肷��
' ���@���FstrDir	�f�B���N�g��
' ��@�@�FMyDir.setDir "."
'
	Function setDir(str)
		strDir = str
	End Function
'
' �@�@�\�F�f�B���N�g���̐�΃p�X�𓾂�
' �߂�l�F������
' ��@�@�FMyDir.getDirPath
'
	Function getDirPath()
		Dim objFolder
		Set objFolder = objFso.GetFolder(strDir)
		getDirPath = objFso.BuildPath(objFolder, "")
	End Function
'
' �@�@�\�F�f�B���N�g�����̍ŏ��̃t�@�C�����𓾂�
' �߂�l�F������
' ��@�@�FMyDir.getFirstFilename
'
	Function getFirstFilename()
		Dim objFolder : Set objFolder = objFso.GetFolder(strDir)
		Dim objFiles : Set objFiles = objFolder.Files
		Dim objFile
		Dim strFilesRec
		strFilesRec = ""
		For Each objFile in objFiles
			If strFilesRec = "" Then
			        strFilesRec = objFile.Name
			Else
			        strFilesRec = strFilesRec & Chr(9) & objFile.Name
			End If
		Next
		strFiles = Split(strFilesRec, Chr(9))
		intIndexStrFiles = 0
		If intIndexStrFiles <= UBound(strFiles) Then
			getFirstFilename = strFiles(intIndexStrFiles)
		Else
			getFirstFilename = ""
		End If
	End Function
'
' �@�@�\�F���̃t�@�C�����𓾂�
' �߂�l�F������
' ��@�@�FMyDir.getNextFilename
'
	Function getNextFilename()
		If intIndexStrFiles < UBound(strFiles) Then
			intIndexStrFiles = intIndexStrFiles + 1
			getNextFilename = strFiles(intIndexStrFiles)
		Else
			getNextFilename = ""
		End If
	End Function
'
' �@�@�\�F�f�B���N�g�����̍ŏ��̃T�u�f�B���N�g�����𓾂�
' �߂�l�F������
' ��@�@�FMyDir.getFirstDirname
'
	Function getFirstDirname()
		Dim objFolder : Set objFolder = objFso.GetFolder(strDir)
		Dim objDirs : Set objDirs = objFolder.SubFolders
		Dim objDir
		Dim strDirsRec
		strDirsRec = ""
		For Each objDir in objDirs
			If strDirsRec = "" Then
			        strDirsRec = objDir.Name
			Else
			        strDirsRec = strDirsRec & Chr(9) & objDir.Name
			End If
		Next
		strDirs = Split(strDirsRec, Chr(9))
		intIndexStrDirs = 0
		If intIndexStrDirs <= UBound(strDirs) Then
			getFirstDirname = strDirs(intIndexStrDirs)
		Else
			getFirstDirname = ""
		End If
	End Function
'
' �@�@�\�F���̃T�u�f�B���N�g�����𓾂�
' �߂�l�F������
' ��@�@�FMyDir.getNextDirname
'
	Function getNextDirname()
		If intIndexStrDirs < UBound(strDirs) Then
			intIndexStrDirs = intIndexStrDirs + 1
			getNextDirname = strDirs(intIndexStrDirs)
		Else
			getNextDirname = ""
		End If
	End Function
End Class
''==================================================
''
'' Excel�������N���X
''
Class MyExcel
	Dim objExcel
	Dim strBookPath

	Sub Class_Initialize()
		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible = False
		objExcel.DisplayAlerts = False
	End Sub

'
' �@�@�\�F�u�b�N���J��
'         ���݂��Ȃ��u�b�N�̏ꍇ�́A���̏�ō쐬
' ���@���FstrBook	�u�b�N(��΃p�X�ł����΃p�X�ł��n�j�j
' ��@�@�FMyExcel.open "foo.xlsx"
'
	Function open(strBook)
		Dim objDir : Set objDir = new MyDir
		Dim objStr : Set objStr = new MyString

		objDir.setDir(".")
		objStr.setValue strBook
		If objStr.isMatch("^[A-Za-z]:") Then
			strBookPath = strBook
		Else
			strBookPath = objDir.getDirPath() & "\" & strBook
		End If

		On Error Resume Next
		objExcel.Workbooks.Open(strBookPath)
		If Err.Number = 1004 Then
			objExcel.Workbooks.Add
			saveAs strBookPath
			objExcel.Workbooks.close
			objExcel.Workbooks.Open(strBookPath)
		End If
		On Error Goto 0
	End Function
'
' �@�@�\�F�u�b�N�����
' ��@�@�FMyExcel.close
'
	Function close()
		objExcel.Workbooks.Close
	End Function
'
' �@�@�\�F�u�b�N���̎w�肵���V�[�g��csv�`���ŕۑ�����
' ���@���FintstrSheet	�V�[�g
'	  strCsvFile	csv�t�@�C��(��΃p�X�ł����΃p�X�ł��n�j�j
' ��@�@�FMyExcel.saveAsCsv 1, "foo.csv" or MyExcel.saveAsCSv "Sheet1", "foo.csv"
'
	Function saveAsCsv(intstrSheet, strCsvFile)
		Dim objDir : Set objDir = new MyDir
		Dim objStr : Set objStr = new MyString
		Dim strOutFullPath
		Dim objSheet

		objStr.setValue strCsvFile
		If objStr.isMatch("^[A-Za-z]:") Then
			strOutFullPath = strCsvFile
		Else
			strOutFullPath = objDir.getDirPath() & "\" & strCsvFile
		End If

		Set objSheet = objExcel.Worksheets(intstrSheet)
		objSheet.SaveAs strOutFullPath, 6	' 6 means to save as CSV
	End Function
'
' �@�@�\�F�u�b�N���w�肵���t�@�C�����ŕۑ�����
' ���@���FstrBook	�u�b�N(��΃p�X�ł����΃p�X�ł��n�j�j
' ��@�@�FMyExcel.saveAs "foo.xlsx"
'
	Function saveAs(strBook)
		Dim objDir : Set objDir = new MyDir
		Dim objStr : Set objStr = new MyString
		Dim strOutBookPath

		objDir.setDir(".")
		objStr.setValue strBook
		If objStr.isMatch("^[A-Za-z]:") Then
			strOutBookPath = strBook
		Else
			strOutBookPath = objDir.getDirPath() & "\" & strBook
		End If
		objExcel.Workbooks(1).SaveAs strOutBookPath
	End Function
'
' �@�@�\�F�u�b�N���㏑���ۑ�����
' ���@���FstrBook	�u�b�N
' ��@�@�FMyExcel.save
'
	Function save()
		On Error Resume Next
		objExcel.Workbooks(1).Save
		On Error Goto 0
	End Function
'
' �@�@�\�F�Z���̒l���擾����
' ���@���FintstrSheet	�V�[�g
'         strRange	�����W
' �߂�l�F������
' ��@�@�FstrStr = MyExcel.getCell(1, "B3") or strStr = MyExcel.getCell("Sheet1", "B3")
'
	Function getCell(intstrSheet, strRange)
		getCell = objExcel.Worksheets(intstrSheet).Range(strRange).Value
	End Function
'
' �@�@�\�F�Z���ɒl��ݒ肷��
' ���@���FintstrSheet	�V�[�g
'         strRange	�����W
'         strValue	�ݒ�l
' ��@�@�FMyExcel.setCell 1, "B3", 123 or MyExcel.setCell "Sheet1", "B3", "�ݒ�l"
'
	Function setCell(intstrSheet, strRange, strValue)
		objExcel.Worksheets(intstrSheet).Range(strRange).Value = strValue
	End Function
'
' �@�@�\�F�Z���̌v�Z�����擾����
' ���@���FintstrSheet	�V�[�g
'         strRange	�����W
' �߂�l�F������
' ��@�@�FstrStr = MyExcel.getFormula(1, "B3") or strStr = MyExcel.getFormula("Sheet1", "B3")
'
	Function getFormula(intstrSheet, strRange)
		getFormula = objExcel.Worksheets(intstrSheet).Range(strRange).Formula
	End Function
'
' �@�@�\�F�Z���Ɍv�Z����ݒ肷��
' ���@���FintstrSheet	�V�[�g
'         strRange	�����W
'         strFormula	�v�Z��
' ��@�@�FMyExcel.setFormula 1, "B3", "=A10+B10" or MyExcel.setFormula "Sheet1", "B3", "=A10+B10"
'
	Function setFormula(intstrSheet, strRange, strFormula)
		objExcel.Worksheets(intstrSheet).Range(strRange).Formula = strFormula
	End Function
'
' �@�@�\�F�Z���̃t�H���g�����擾����
' ���@���FintstrSheet	�V�[�g
'         strRange	�����W
' �߂�l�F������
' ��@�@�FstrStr = MyExcel.getFontName(1, "B3") or strStr = MyExcel.getFontName("Sheet1", "B3")
'
	Function getFontName(intstrSheet, strRange)
		getFontName = objExcel.Worksheets(intstrSheet).Range(strRange).Font.Name
	End Function
'
' �@�@�\�F�Z���̃t�H���g����ݒ肷��
' ���@���FintstrSheet	�V�[�g
'         strRange	�����W
'         strFontName	�t�H���g��
' ��@�@�FMyExcel.setFontName 1, "B3", "�l�r�@�S�V�b�N" or MyExcel.setFontName "Sheet1", "B3", "�l�r�@�S�V�b�N"
'
	Function setFontName(intstrSheet, strRange, strFontName)
		objExcel.Worksheets(intstrSheet).Range(strRange).Font.Name = strFontName
	End Function
'
' �@�@�\�F�Z���̃t�H���g�T�C�Y���擾����
' ���@���FintstrSheet	�V�[�g
'         strRange	�����W
' �߂�l�F���l
' ��@�@�FrealNum = MyExcel.getFontSize(1, "B3") or realNum = MyExcel.getFontSize("Sheet1", "B3")
'
	Function getFontSize(intstrSheet, strRange)
		getFontSize = objExcel.Worksheets(intstrSheet).Range(strRange).Font.Size
	End Function
'
' �@�@�\�F�Z���̃t�H���g�T�C�Y��ݒ肷��
' ���@���FintstrSheet	�V�[�g
'         strRange	�����W
'         realFontSize	�t�H���g�T�C�Y
' ��@�@�FMyExcel.setFontSize 1, "B3", 10.5 or MyExcel.setFontSize "Sheet1", "B3", 10.5
'
	Function setFontSize(intstrSheet, strRange, realFontSize)
		objExcel.Worksheets(intstrSheet).Range(strRange).Font.Size = realFontSize
	End Function
'
' �@�@�\�F�Z���̕����F���擾����
' ���@���FintstrSheet	�V�[�g
'         strRange	�����W
' �߂�l�F���l
' ��@�@�FintNum = MyExcel.getForegroundColor(1, "B3") or intNum = MyExcel.getForegroundColor("Sheet1", "B3")
'
	Function getForegroundColor(intstrSheet, strRange)
		getForegroundColor = objExcel.Worksheets(intstrSheet).Range(strRange).Font.Color
	End Function
'
' �@�@�\�F�Z���̕����F��ݒ肷��
' ���@���FintstrSheet	�V�[�g
'         strRange	�����W
'         intColor	�����F
' ��@�@�FMyExcel.setForegroundColor 1, "B3", &HFF0000 or MyExcel.setForegroundColor "Sheet1", "B3", RGB(255,0,0)
'
	Function setForegroundColor(intstrSheet, strRange, intColor)
		objExcel.Worksheets(intstrSheet).Range(strRange).Font.Color = intColor
	End Function
'
' �@�@�\�F�Z���̔w�i�F���擾����
' ���@���FintstrSheet	�V�[�g
'         strRange	�����W
' �߂�l�F���l
' ��@�@�FintNum = MyExcel.getBackgroundColor(1, "B3") or intNum = MyExcel.getBackgroundColor("Sheet1", "B3")
'
	Function getBackgroundColor(intstrSheet, strRange)
		getBackgroundColor = objExcel.Worksheets(intstrSheet).Range(strRange).Interior.Color
	End Function
'
' �@�@�\�F�Z���̔w�i�F��ݒ肷��
' ���@���FintstrSheet	�V�[�g
'         strRange	�����W
'         intColor	�w�i�F(&HBBGGRR)
' ��@�@�FMyExcel.setBackgroundColor 1, "B3", &HFF0000 or MyExcel.setBackgroundColor "Sheet1", "B3", RGB(255,0,0)
'
	Function setBackgroundColor(intstrSheet, strRange, intColor)
		objExcel.Worksheets(intstrSheet).Range(strRange).Interior.Color = intColor
	End Function
'
' �@�@�\�F�Z���̕����擾����
' ���@���FintstrSheet	�V�[�g
'         strColumn	�J����
' �߂�l�F���l
' ��@�@�FrealNum = MyExcel.getColumnWidth(1, "B") or realNum = MyExcel.getColumnWidth("Sheet1", "B")
'
	Function getColumnWidth(intstrSheet, strColumn)
		getColumnWidth = objExcel.Worksheets(intstrSheet).Columns(strColumn).ColumnWidth
	End Function
'
' �@�@�\�F�Z���̕���ݒ肷��
' ���@���FintstrSheet	�V�[�g
'         strColumn	�J����
'         realWidth	��
' ��@�@�FMyExcel.setColumnWidth 1, "B3", 10.5 or MyExcel.setColumnWidth "Sheet1", "B3", 10.5
'
	Function setColumnWidth(intstrSheet, strColumn, realWidth)
		objExcel.Worksheets(intstrSheet).Columns(strColumn).ColumnWidth = realWidth
	End Function
'
' �@�@�\�F�Z���̍������擾����
' ���@���FintstrSheet	�V�[�g
'         intRow	�s
' �߂�l�F���l
' ��@�@�FrealNum = MyExcel.getRowHeight(1, 2) or realNum = MyExcel.getRowHeight("Sheet1", 2)
'
	Function getRowHeight(intstrSheet, intRow)
		getRowHeight = objExcel.Worksheets(intstrSheet).Rows(intRow).RowHeight
	End Function
'
' �@�@�\�F�Z���̍�����ݒ肷��
' ���@���FintstrSheet	�V�[�g
'         intRow	�s
'         realHeight	����
' ��@�@�FMyExcel.setRowHeight 1, 2, 10.5 or MyExcel.setRowHeight "Sheet1", 2, 10.5
'
	Function setRowHeight(intstrSheet, intRow, realHeight)
		objExcel.Worksheets(intstrSheet).Rows(intRow).RowHeight = realHeight
	End Function
'
' �@�@�\�F�Z���̏����`�����擾����
' ���@���FintstrSheet	�V�[�g
'         strRange	�����W
' �߂�l�F������
' ��@�@�FstrStr = MyExcel.getFormat(1, "B3") or strStr = MyExcel.getFormat("Sheet1", "B3")
'
	Function getFormat(intstrSheet, strRange)
		getFormat = objExcel.Worksheets(intstrSheet).Range(strRange).NumberFormatLocal
	End Function
'
' �@�@�\�F�Z���̏����`����ݒ肷��
' ���@���FintstrSheet	�V�[�g
'         strRange	�����W
'         strFormat	�����`��
' ��@�@�FMyExcel.setFormat 1, "B3", "000000" or MyExcel.setFormat "Sheet1", "B3", "000000"
'
	Function setFormat(intstrSheet, strRange, strFormat)
		objExcel.Worksheets(intstrSheet).Range(strRange).NumberFormatLocal = strFormat
	End Function
'
' �@�@�\�F�w�肵���V�[�g�̈���v���r���[��\������
' ���@���FintstrSheet	�V�[�g
' ��@�@�FMyExcel.preview 1 or MyExcel.preview "Sheet1"
'
	Function preview(intstrSheet)
		objExcel.Visible = True
		objExcel.Worksheets(intstrSheet).PrintPreview
		objExcel.Visible = False
	End Function
'
' �@�@�\�F�w�肵���V�[�g���������
' ���@���FintstrSheet	�V�[�g
' ��@�@�FMyExcel.print 1 or MyExcel.print "Sheet1"
'
	Function print(intstrSheet)
		objExcel.Worksheets(intstrSheet).PrintOut
	End Function
'
' �@�@�\�F�u�b�N��\������
' ��@�@�FMyExcel.onVisible
'
	Function onVisible()
		objExcel.Visible = True
	End Function
'
' �@�@�\�F�u�b�N���\���ɂ���
' ��@�@�FMyExcel.offVisible
'
	Function offVisible()
		objExcel.Visible = False
	End Function
'
' �@�@�\�F�A���[�g��\������
' ��@�@�FMyExcel.onAlert
'
	Function onAlert()
		objExcel.DisplayAlerts = True
	End Function
'
' �@�@�\�F�A���[�g���\���ɂ���
' ��@�@�FMyExcel.offAlert
'
	Function offAlert()
		objExcel.DisplayAlerts = False
	End Function
End Class
''==================================================
''
'' �I�v�V�����������N���X
''
Class MyOption
	Dim objArg
	Dim objDict
	Dim strArrayNonOpt()

	Sub Class_initialize()
		Set objArg = new MyArg
                Set objDict = CreateObject("Scripting.Dictionary")
		ReDim strArrayNonOpt(-1)
	End Sub

'
' �@�@�\�F�����I�v�V������ݒ肷��
' ���@���FstrAllOption	"-option1=[y|n],..."
'                                 y means option take value like "-option1 value"
'                                 n means option do not take value"
' ��@�@�FMyOption.initialize("-h=n,--help=n,-i=y,--input=y,-o=y,--output=y")
'
	Function initialize(strAllOption)
		Dim array
		Dim i

		array = Split(strAllOption, ",")
		i = 0
		While i <= UBound(array)
			Dim arrayKv

			arrayKv = Split(array(i), "=")
			objDict.add arrayKv(0), arrayKv(1)
			i = i + 1
		Wend
	End Function

'
' �@�@�\�F�I�v�V�����̒l���擾����
' ���@���FstrOptions	�I�v�V����������
' �߂�l�F������
' ��@�@�FstrOpt = MyOption.getValue("--input")
'
	Function getValue(strOption)
		Dim intIndex
		Dim strValue

		strValue = ""
		intIndex = 0
		While intIndex < objArg.getCount
			If strOption = objArg.getValue(intIndex) Then
				intIndex = intIndex + 1
				strValue = objArg.getValue(intIndex) 
				intIndex = objArg.getCount
			End If
			intIndex = intIndex + 1
		Wend
		getValue = strValue
	End Function

'
' �@�@�\�F�I�v�V�������w�肳��Ă��邩�擾����
' ���@���FstrOptions	�I�v�V����������
' �߂�l�FTrue or False
' ��@�@�FIf MyOption.isSpecified("--help") Then
'
	Function isSpecified(strOption)
		Dim intIndex
		Dim booleanRet

		booleanRet = false
		intIndex = 0
		While intIndex < objArg.getCount
			If strOption = objArg.getValue(intIndex) Then
				booleanRet = true
				intIndex = objArg.getCount
			End If
			intIndex = intIndex + 1
		Wend
		isSpecified = booleanRet
	End Function

'
' �@�@�\�F�I�v�V�����ȊO�̕�������擾����
' �߂�l�F������(�󔒋�؂�)
' ��@�@�FstrNonOpt = MyOption.getNonOptions()
'
	Function getNonOptions()
		Dim intIndex
		Dim strValue

		strValue = ""
		intIndex = 0
		While intIndex < objArg.getCount
			If objDict.Exists(objArg.getValue(intIndex)) Then
				If objDict(objArg.getValue(intIndex)) = "y" Then
					intIndex = intIndex + 1
				End If
			Else
				If strValue = "" Then
					strValue = objArg.getValue(intIndex)
				Else
					strValue = strValue & " " & objArg.getValue(intIndex)
				End If
			End If
			intIndex = intIndex + 1
		Wend
		getNonOptions = strValue
	End Function

'
' �@�@�\�F�I�v�V�����ȊO�̕������z��Ŏ擾����
' �߂�l�F������z��
' ��@�@�FstrArraynNonOpt = MyOption.getArrayNonOptions()
'
	Function getArrayNonOptions()
		Dim intIndex
		Dim strValue
		Dim intNonOptIndex

		strValue = ""
		intIndex = 0
		intNonOptIndex = 0
		While intIndex < objArg.getCount
			If objDict.Exists(objArg.getValue(intIndex)) Then
				If objDict(objArg.getValue(intIndex)) = "y" Then
					intIndex = intIndex + 1
				End If
			Else
				ReDim Preserve strArrayNonOpt(intNonOptIndex)
				strArrayNonOpt(intNonOptIndex) = objArg.getValue(intIndex)
				intNonOptIndex = intNonOptIndex + 1
			End If
			intIndex = intIndex + 1
		Wend
		getArrayNonOptions = strArrayNonOpt
	End Function
End Class
''==================================================
''
'' �t�@�C���V�X�e������������N���X
''
Class MyFsOpe
	Sub Class_Initialize()
	End Sub

'
' �@�@�\�F�ꎞ�t�@�C�����̎擾
' �߂�l�F������
' ��@�@�FstrTempFile = MyFsOpe.getTempFileName
'
	Function getTempFileName()
		getTempFileName = CreateObject("Scripting.FileSystemObject").GetTempName
	End Function

'
' �@�@�\�F�f�B���N�g���̍쐬
' ���@���FstrDirName
' ��@�@�FMyFsOpe.createFolder "TEMP.dir"
'
	Function createFolder(strFileName)
		CreateObject("Scripting.FileSystemObject").createFolder strDirName
	End Function

'
' �@�@�\�F�t�@�C���̍폜
' ���@���FstrFileName
' ��@�@�FMyFsOpe.deleteFile "TEMP.tmp"
'
	Function deleteFile(strFileName)
		CreateObject("Scripting.FileSystemObject").DeleteFile strFileName, True
	End Function

'
' �@�@�\�F�f�B���N�g���̍폜
' ���@���FstrDirName
' ��@�@�FMyFsOpe.deleteFolder "TEMP.dir"
'
	Function deleteFolder(strDirName)
		CreateObject("Scripting.FileSystemObject").DeleteFolder strDirName, True
	End Function
End Class
''==================================================
''
'' ���X�������N���X
''
Class MyMisc
	Sub Class_Initialize()
	End Sub

'
' �@�@�\�F�v���O�����̏I��
' ���@���F�I���X�e�[�^�X
' ��@�@�FMyMisc.exitProg 0
'
	Function exitProg(stat)
		WScript.Quit(stat)
	End Function
End Class
'==================================================
'
' �f�o�b�O�\���p
'
Function Debug(str)
	Dim objStdio : Set objStdio = new MyStdio
	objStdio.writeLine(str)
End Function
