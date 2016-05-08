Option Explicit
''==================================================
''
'' 標準入出力クラス
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
' 機　能：標準入力より１行取得
' 戻り値：レコード
' 例　　：strRec = MyStdio.readLine
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
' 機　能：標準入力より指定文字数分取得
' 引　数：count		取得する文字数
' 戻り値：文字列
' 例　　：strRec = MyStdio.read(128)
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
' 機　能：標準入力読み取り成功チェック
' 戻り値：true or false
' 例　　：While MyStdio.isRaedSuccess
'
	Function isReadSuccess()
		isReadSuccess = bReadSuccess
	End Function
'
' 機　能：標準入力読み取り失敗チェック
' 戻り値：true or false
' 例　　：While Not MyStdio.isRaedFailure
'
	Function isReadFailure()
		isReadFailure = Not isReadSuccess
	End Function
'
' 機　能：標準出力に１行出力
' 引　数：str		レコード
' 例　　：MyStdio.writeLine strRec
'
	Function writeLine(str)
		objStdout.WriteLine str
	End Function
'
' 機　能：標準出力に指定文字列を出力
' 引　数：str		文字列
' 例　　：MyStdio.write strString
'
	Function write(str)
		objStdout.Write str
	End Function
'
' 機　能：標準エラー出力に１行出力
' 引　数：str		レコード
' 例　　：MyStdio.writeErrorLine strRec
'
	Function writeErrorLine(str)
		objStderr.WriteLine str
	End Function
'
' 機　能：標準エラー出力に指定文字列を出力
' 引　数：str		文字列
' 例　　：MyStdio.writeError strString
'
	Function writeError(str)
		objStderr.Write str
	End Function
'
' 機　能：標準入力の終端を取得
' 戻　値：True(終端時) or False(非終端時)
' 例　　：While Not isEof
'
	Function isEof()
		isEof = objStdin.AtEndOfStream
	End Function
End Class
''==================================================
''
'' 文字列クラス
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
' 機　能：文字列を設定
' 引　数：str		文字列
' 例　　：MyString.setValue strString
'
	Function setValue(str)
		value = CStr(str)
		booleanHighValue = false
		booleanLowValue = false
	End Function
'
' 機　能：文字列を取得
' 戻　値：文字列
' 例　　：strString = MyString.getValue
'
	Function getValue()
		getValue = CStr(value)
	End Function
'
' 機　能：正規表現との一致を取得
' 引　数：regex		正規表現
' 戻　値：True(一致時) or False(不一致時)
' 例　　：If MyString.isMatch("^A.*Z$")
'
	Function isMatch(regex)
		objRegex.Pattern = regex
		isMatch = objRegex.Test(value)
	End Function
'
' 機　能：正規表現に一致した部分を別の文字列に置き換えるた文字列の取得
' 引　数：regex		正規表現
'	  str		置換後の文字列
' 戻　値：文字列
' 例　　：strString = MyString.getReplace("^A.*Z$", "A to Z")
'
	Function getReplace(regex, str)
		objRegex.Global = True
		objRegex.Pattern = regex
		getReplace = objRegex.Replace(value, str)
	End Function
'
' 機　能：指定した区切り文字で区切って配列で取得
' 引　数：del		区切り文字
' 戻　値：文字列配列
' 例　　：strArray = MyString.getSplit(",")
'
	Function getSplit(del)
		getSplit = split(value, del)
	End Function
'
' 機　能：文字列長を取得
' 戻　値：文字列長
' 例　　：intLength = MyString.getLength
'
	Function getLength()
		getLength = len(value)
	End Function
'
' 機　能：部分文字列を取得
' 引　数：start		開始位置(1オリジン)
'	  length	文字列長
' 戻　値：文字列
' 例　　：strString = MyString.Substr(1,5)
'
	Function getSubstr(start, length)
		getSubstr = Mid(value, start, length)
	End Function

'
' 機　能：文字列を16進数コードに変換
' 戻　値：文字列
' 例　　：strHexString = MyString.getHexString
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
' 機　能：HIGH-VALUEをセットする
' 例　　：MyString.setHighValue
'
	Function setHighValue()
		booleanHighValue = true
		value = ""
	End Function
'
' 機　能：HIGH-VALUEか確認する
' 戻　値：True or False
' 例　　：MyString.isHighValue
'
	Function isHighValue()
		isHighValue = booleanHighValue
	End Function
'
' 機　能：LOW-VALUEをセットする
' 例　　：MyString.setLowValue
'
	Function setLowValue()
		booleanLowValue = true
		value = ""
	End Function
'
' 機　能：LOW-VALUEか確認する
' 戻　値：True or False
' 例　　：MyString.isLowValue
'
	Function isLowValue()
		isLowValue = booleanLowValue
	End Function
'
' 機　能：等しいか確認する
' 戻　値：True or False
' 例　　：MyString.isEqual(objString)
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
' 機　能：大きいか確認する
' 戻　値：True or False
' 例　　：MyString.isGreater(objString)
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
' 機　能：小さいか確認する
' 戻　値：True or False
' 例　　：MyString.isLess(objString)
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
'' 引数クラス
''
Class MyArg
	Dim objArg

	Sub Class_initialize()
		Set objArg = WScript.Arguments
	End Sub

'
' 機　能：引数の個数を取得
' 戻　値：引数の個数
' 例　　：intCount= MyArg.getCount
'
	Function getCount()
		getCount = objArg.Count
	End Function
'
' 機　能：指定した引数の取得
' 引　数：インデックス(0オリジン)
' 戻　値：文字列
' 例　　：strArg= MyArg.getValue(0)
'
	Function getValue(idx)
		getValue = objArg(idx)
	End Function
End Class
''==================================================
''
'' FileSystemObjectクラス(Shift_JIS)
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
' 機　能：FSOを開く
' 引　数：filename		ファイル名
'	  mode			"r"(読み込み) or "w"(書き込み)
' 例　　：MyFso.open "INPUT.TXT", "r"
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
' 機　能：FSOを入力モードで開く
' 引　数：filename		ファイル名
' 例　　：MyFso.openInput "INPUT.TXT"
'
	Function openInput(filename)
		open filename, "r"
	End Function
'
' 機　能：FSOを出力モードで開く
' 引　数：filename		ファイル名
' 例　　：MyFso.openOutput "OUTPUT.TXT"
'
	Function openOutput(filename)
		open filename, "w"
	End Function
'
' 機　能：１行取得
' 戻り値：レコード
' 例　　：strRec = MyFso.readLine
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
' 機　能：指定文字数分取得
' 引　数：count		取得する文字数
' 戻り値：文字列
' 例　　：strString = MyFso.read(128)
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
' 機　能：読み取り成功チェック
' 戻り値：true or false
' 例　　：While MyFso.isReadSuccess
'
	Function isReadSuccess()
		isReadSuccess = bReadSuccess
	End Function
'
' 機　能：読み取り失敗チェック
' 戻り値：true or false
' 例　　：While Not MyFso.isReadFailure
'
	Function isReadFailure()
		isReadFailure = Not isReadSuccess
	End Function
'
' 機　能：１行出力
' 引　数：str		レコード
' 例　　：MyFso.writeLine strRec
'
	Function writeLine(str)
		objFile.WriteLine str
	End Function
'
' 機　能：指定文字列を出力
' 引　数：str		文字列
' 例　　：MyFso.write strString
'
	Function write(str)
		objFile.Write str
	End Function
'
' 機　能：終端を取得
' 戻　値：True(終端時) or False(非終端時)
' 例　　：While Not MyFso.isEof
'
	Function isEof()
		isEof = objFile.AtEndOfStream
	End Function
'
' 機　能：FSOを閉じる
' 例　　：MyFso.close
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
'' ActiveX Data Objectクラス(文字コード指定可能)
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
' 機　能：ADOを開く
' 引　数：filename		ファイル名
'	  mode			"r"(読み込み) or "w"(書き込み)
'	  charset		"UTF-8" or "Shift_JIS" or "ASCII" etc
'	  bom			True(UTF-8の時BOM付きで出力) or
'				False(UTF-8の時BOM無しで出力)
' 例　　：MyAdo.open "INPUT.TXT", "r", "UTF-8", False
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
' 機　能：ADOを入力モードで開く
' 引　数：filename		ファイル名
'	  charset		"UTF-8" or "Shift_JIS" or "ASCII" etc
' 例　　：MyAdo.openInput "INPUT.TXT", "UTF-8"
'
	Function openInput(filename, charset)
		open filename, "r", charset, False
	End Function
'
' 機　能：ADOを出力モードで開く
' 引　数：filename		ファイル名
'	  charset		"UTF-8" or "Shift_JIS" or "ASCII" etc
' 例　　：MyAdo.openOutput "OUTPUT.TXT", "UTF-8"
'
	Function openOutput(filename, charset)
		open filename, "w", charset, False
	End Function
'
' 機　能：１行取得
' 戻り値：レコード
' 例　　：strRec = MyAdo.readLine
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
' 機　能：読み取り成功チェック
' 戻り値：true or false
' 例　　：While MyAdo.isReadSuccess
'
	Function isReadSuccess()
		isReadSuccess = bReadSuccess
	End Function
'
' 機　能：読み取り失敗チェック
' 戻り値：true or false
' 例　　：While Not MyAdo.isReadFailure
'
	Function isReadFailure()
		isReadFailure = Not isReadSuccess
	End Function
'
' 機　能：１行出力
' 引　数：str		レコード
' 例　　：MyAdo.writeLine strRec
'
	Function writeLine(record)
		objStream.WriteText record, 1
	End Function
'
' 機　能：ADOを閉じる
' 例　　：MyAdo.close
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
' 機　能：終端を取得
' 戻　値：True(終端時) or False(非終端時)
' 例　　：While Not MyAdo.isEof
'
	Function isEof()
		isEof = objStream.EOS
	End Function
'
' 機　能：保存する
' 例　　：MyAdo.save
'
	Private Function save()
		objStream.SaveToFile strFilename, 2
	End Function
'
' 機　能：先頭の3bytesを除いて保存する
' 例　　：MyAdo.saveWithoutBom
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
'' CSVクラス
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
' 機　能：レコードの初期化
' 引　数：count		レコードの項目数
' 例　　：MyCsv.initRec 16
'
	Function initRec(count)
		ReDim strValues(count - 1)
	End Function
'
' 機　能：レコード（区切り文字含む）を設定
' 引　数：str		文字列（項目名１｛区切り文字｝項目名２｛区切り文字｝...）
' 例　　：MyCsv.setRec strRec
'
	Function setRec(str)
		objStr.setValue(str)
		strValues = objStr.getSplit(strDel)
	End Function
'
' 機　能：レコードを取得
' 戻　値：レコード（区切り文字含む）
' 例　　：strRec = MyCsv.getRec
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
' 機　能：項目名レコードを設定
' 引　数：str		文字列（項目名１｛区切り文字｝項目名２｛区切り文字｝...）
' 例　　：MyCsv.setNameRec strNameRec
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
' 機　能：項目名レコードを取得
' 戻り値：文字列（項目名１｛区切り文字｝項目名２｛区切り文字｝...）
' 例　　：strNameRec = MyCsv.getNameRec
'
	Function getNameRec()
		getNameRec = strNames
	End Function
'
' 機　能：区切り文字を設定
' 引　数：del		区切り文字列
' 例　　：MyCsv.setDelimiter ","
'
	Function setDelimiter(del)
		strDel = del
	End Function
'
' 機　能：区切り文字を取得
' 戻り値：文字
' 例　　：charDel = MyCsv.getDelimiter
'
	Function getDelimiter()
		getDelimiter = strDel
	End Function
'
' 機　能：インデックス指定で項目を取得
' 引　数：i		インデックス（０オリジン）
' 戻り値：文字列
' 例　　：strString = MyCsv.getValueByIndex(0)
'
	Function getValueByIndex(i)
		If 0 <= i And i <= UBound(strValues) Then
			getValueByIndex = strValues(i)
		Else
			getValueByIndex = ""
		End If
	End Function
'
' 機　能：インデックス指定で項目を設定
' 引　数：i		インデックス（０オリジン）
'	  value		文字列
' 戻り値：文字列
' 例　　：MyCsv.setValueByIndex 0, "STRING"
'
	Function setValueByIndex(i, value)
		If 0 <= i And i <= UBound(strValues) Then
			strValues(i) = value
		End If
	End Function
'
' 機　能：最終インデックスを取得
' 戻り値：整数（０オリジン）
' 例　　：intIndex = MyCsv.getLastIndex
'
	Function getLastIndex()
		getLastIndex = UBound(strValues)
	End Function
'
' 機　能：項目名に対応したインデックスを取得
' 引　数：name		項目名
' 戻り値：整数（０オリジン）
' 例　　：intIndex = MyCsv.getIndexByName("ELEMENT_01")
'
	Function getIndexByName(name)
		If objDict.Exists(name) Then
			getIndexByName = objDict(name)
		Else
			getIndexByName = -1
		End If
	End Function
'
' 機　能：項目名指定で項目を取得
' 引　数：name		項目名
' 戻り値：文字列
' 例　　：strString = MyCsv.getValueByName("ELEMENT_01")
'
	Function getValueByName(name)
		getValueByName = getValueByIndex(getIndexByName(name))
	End Function
'
' 機　能：項目名指定で項目を設定
' 引　数：name		項目名
'	  value		文字列
' 戻り値：文字列
' 例　　：MyCsv.setValueByName "ELEMENT_01", 123
'
	Function setValueByName(name, value)
		setValueByIndex getIndexByName(name), value
	End Function
End Class
''==================================================
''
'' 標準入出力CSVクラス
''
Class MyStdioCsv
	Dim objStdio
	Dim objCsv

	Sub Class_initialize()
		Set objStdio = new MyStdio
		Set objCsv = new MyCsv
	End Sub
'
' 機　能：レコードの初期化
' 引　数：count		レコードの項目数
' 例　　：MyStdioCsv.initRec 16
'
	Function initRec(count)
		objCsv.initRec count
	End Function
'
' 機　能：レコード（区切り文字含む）を設定
' 引　数：str		文字列（項目名１｛区切り文字｝項目名２｛区切り文字｝...）
' 例　　：MyStdioCsv.setRec strRec
'
	Function setRec(str)
		objCsv.setRec str
	End Function
'
' 機　能：レコードを取得
' 戻　値：レコード（区切り文字含む）
' 例　　：strRec = MyStdioCsv.getRec
'
	Function getRec()
		getRec = objCsv.getRec
	End Function
'
' 機　能：項目名レコードを設定
' 引　数：str		文字列（項目名１｛区切り文字｝項目名２｛区切り文字｝...）
' 例　　：MyStdioCsv.setNameRec strNameRec
'
	Function setNameRec(str)
		objCsv.setNameRec str
	End Function
'
' 機　能：項目名レコードを取得
' 戻り値：文字列（項目名１｛区切り文字｝項目名２｛区切り文字｝...）
' 例　　：strNameRec = MyStdioCsv.getNameRec
'
	Function getNameRec()
		getNameRec = objCsv.getNameRec
	End Function
'
' 機　能：区切り文字を設定
' 引　数：del		区切り文字列
' 例　　：MyStdioCsv.setDelimiter ","
'
	Function setDelimiter(del)
		objCsv.setDelimiter del
	End Function
'
' 機　能：区切り文字を取得
' 戻り値：文字
' 例　　：charDel = MyStdioCsv.getDelimiter
'
	Function getDelimiter()
		getDelimiter = objCsv.getDelimiter
	End Function
'
' 機　能：インデックス指定で項目を取得
' 引　数：i		インデックス（０オリジン）
' 戻り値：文字列
' 例　　：strString = MyStdioCsv.getValueByIndex(0)
'
	Function getValueByIndex(i)
		getValueByIndex = objCsv.getValueByIndex(i)
	End Function
'
' 機　能：インデックス指定で項目を設定
' 引　数：i		インデックス（０オリジン）
'	  value		文字列
' 戻り値：文字列
' 例　　：MyStdioCsv.setValueByIndex 0, "STRING"
'
	Function setValueByIndex(i, value)
		objCsv.setValueByIndex i, value
	End Function
'
' 機　能：最終インデックスを取得
' 戻り値：整数（０オリジン）
' 例　　：intIndex = MyStdioCsv.getLastIndex
'
	Function getLastIndex()
		getLastIndex = objCsv.getLastIndex
	End Function
'
' 機　能：項目名に対応したインデックスを取得
' 引　数：name		項目名
' 戻り値：整数（０オリジン）
' 例　　：intIndex = MyStdioCsv.getIndexByName("ELEMENT_01")
'
	Function getIndexByName(name)
		getIndexByName = objCsv.getIndexByName(name)
	End Function
'
' 機　能：項目名指定で項目を取得
' 引　数：name		項目名
' 戻り値：文字列
' 例　　：strString = MyStdioCsv.getValueByName("ELEMENT_01")
'
	Function getValueByName(name)
		getValueByName = objCsv.getValueByName(name)
	End Function
'
' 機　能：項目名指定で項目を設定
' 引　数：name		項目名
'	  value		文字列
' 戻り値：文字列
' 例　　：MyStdioCsv.setValueByName "ELEMENT_01", 123
'
	Function setValueByName(name, value)
		objCsv.setValueByName name, value
	End Function
'
' 機　能：標準入力より１行CSVレコード取得を取得し、objCsvへ保存
' 戻り値：True(読めた) or False(読めなかった)
' 例　　：MyStdioCsv.readLine
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
' 機　能：objCsvから標準出力に１行CSVレコード出力
' 例　　：MyStdioCsv.writeLine
'
	Function writeLine()
		objStdio.writeLine objCsv.getRec
	End Function
'
' 機　能：objCsvから標準エラー出力にCSVレコード出力
' 引　数：str		レコード
' 例　　：MyStdioCsv.writeErrorLine
'
	Function writeErrorLine
		objStdio.writeErrorLine objCsv.getRec
	End Function
'
' 機　能：標準入力読み取り成功チェック
' 戻り値：true or false
' 例　　：While MyStdioCsv.isRaedSuccess
'
	Function isReadSuccess()
		isReadSuccess = objStdio.isReadSuccess
	End Function
'
' 機　能：標準入力読み取り失敗チェック
' 戻り値：true or false
' 例　　：While Not MyStdioCsv.isRaedFailure
'
	Function isReadFailure()
		isReadFailure = objStdio.isReadFailure
	End Function
'
' 機　能：標準入力の終端を取得
' 戻　値：True(終端時) or False(非終端時)
' 例　　：While Not MyStdioCsv.isEof
'
	Function isEof()
		isEof = objStdio.isEof
	End Function
End Class
''==================================================
''
'' FileSystemObjectCSVクラス(Shift_JIS)
''
Class MyFsoCsv
	Dim objFso
	Dim objCsv

	Sub Class_initialize()
		Set objFso = new MyFso
		Set objCsv = new MyCsv
	End Sub
'
' 機　能：レコードの初期化
' 引　数：count		レコードの項目数
' 例　　：MyFsoCsv.initRec 16
'
	Function initRec(count)
		objCsv.initRec count
	End Function
'
' 機　能：レコード（区切り文字含む）を設定
' 引　数：str		文字列（項目名１｛区切り文字｝項目名２｛区切り文字｝...）
' 例　　：MyFsoCsv.setRec strRec
'
	Function setRec(str)
		objCsv.setRec str
	End Function
'
' 機　能：レコードを取得
' 戻　値：レコード（区切り文字含む）
' 例　　：strRec = MyFsoCsv.getRec
'
	Function getRec()
		getRec = objCsv.getRec
	End Function
'
' 機　能：項目名レコードを設定
' 引　数：str		文字列（項目名１｛区切り文字｝項目名２｛区切り文字｝...）
' 例　　：MyFsoCsv.setNameRec strNameRec
'
	Function setNameRec(str)
		objCsv.setNameRec str
	End Function
'
' 機　能：項目名レコードを取得
' 戻り値：文字列（項目名１｛区切り文字｝項目名２｛区切り文字｝...）
' 例　　：strNameRec = MyFsoCsv.getNameRec
'
	Function getNameRec()
		getNameRec = objCsv.getNameRec
	End Function
'
' 機　能：区切り文字を設定
' 引　数：del		区切り文字列
' 例　　：MyFsoCsv.setDelimiter ","
'
	Function setDelimiter(del)
		objCsv.setDelimiter del
	End Function
'
' 機　能：区切り文字を取得
' 戻り値：文字
' 例　　：charDel = MyFsoCsv.getDelimiter
'
	Function getDelimiter()
		getDelimiter = objCsv.getDelimiter
	End Function
'
' 機　能：インデックス指定で項目を取得
' 引　数：i		インデックス（０オリジン）
' 戻り値：文字列
' 例　　：strString = MyFsoCsv.getValueByIndex(0)
'
	Function getValueByIndex(i)
		getValueByIndex = objCsv.getValueByIndex(i)
	End Function
'
' 機　能：インデックス指定で項目を設定
' 引　数：i		インデックス（０オリジン）
'	  value		文字列
' 戻り値：文字列
' 例　　：MyFsoCsv.setValueByIndex 0, "STRING"
'
	Function setValueByIndex(i, value)
		objCsv.setValueByIndex i, value
	End Function
'
' 機　能：最終インデックスを取得
' 戻り値：整数（０オリジン）
' 例　　：intIndex = MyFsoCsv.getLastIndex
'
	Function getLastIndex()
		getLastIndex = objCsv.getLastIndex
	End Function
'
' 機　能：項目名に対応したインデックスを取得
' 引　数：name		項目名
' 戻り値：整数（０オリジン）
' 例　　：intIndex = MyFsoCsv.getIndexByName("ELEMENT_01")
'
	Function getIndexByName(name)
		getIndexByName = objCsv.getIndexByName(name)
	End Function
'
' 機　能：項目名指定で項目を取得
' 引　数：name		項目名
' 戻り値：文字列
' 例　　：strString = MyFsoCsv.getValueByName("ELEMENT_01")
'
	Function getValueByName(name)
		getValueByName = objCsv.getValueByName(name)
	End Function
'
' 機　能：項目名指定で項目を設定
' 引　数：name		項目名
'	  value		文字列
' 戻り値：文字列
' 例　　：MyFsoCsv.setValueByName "ELEMENT_01", 123
'
	Function setValueByName(name, value)
		objCsv.setValueByName name, value
	End Function
'
' 機　能：開く
' 引　数：filename		ファイル名
'	  mode			"r"(読み込み) or "w"(書き込み)
' 例　　：MyFsoCsv.open "INPUT.TXT", "r"
'
	Function open(filename, mode)
		objFso.open filename, mode
	End Function
'
' 機　能：入力モードで開く
' 引　数：filename		ファイル名
' 例　　：MyFsoCsv.openInput "INPUT.TXT"
'
	Function openInput(filename)
		objFso.openInput filename
	End Function
'
' 機　能：出力モードで開く
' 引　数：filename		ファイル名
' 例　　：MyFsoCsv.openOutput "OUTPUT.TXT"
'
	Function openOutput(filename)
		objFso.openOutput filename
	End Function
'
' 機　能：FSOより１行CSVレコード取得し、objCsvへ保存
' 戻り値：True(読めた) or False(読めなかった)
' 例　　：MyFsoCsv.readLine
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
' 機　能：読み取り成功チェック
' 戻り値：true or false
' 例　　：While MyFsoCsv.isReadSuccess
'
	Function isReadSuccess()
		isReadSuccess = objFso.isReadSuccess
	End Function
'
' 機　能：読み取り失敗チェック
' 戻り値：true or false
' 例　　：While Not MyFsoCsv.isReadFailure
'
	Function isReadFailure()
		isReadFailure = objFso.isReadFailure
	End Function
'
' 機　能：objCsvからFSOに１行CSVレコード出力
' 例　　：MyFsoCsv.writeLine
'
	Function writeLine()
		objFso.writeLine objCsv.getRec
	End Function
'
' 機　能：終端を取得
' 戻　値：True(終端時) or False(非終端時)
' 例　　：While Not MyFsoCsv.isEof
'
	Function isEof()
		isEof = objFso.isEof
	End Function
'
' 機　能：閉じる
' 例　　：MyFsoCsv.close
'
	Function close()
		objFso.close
	End Function
End Class
''==================================================
''
'' ActiveX Data Object CSVクラス(文字コード指定可能)
''
Class MyAdoCsv
	Dim objAdo
	Dim objCsv

	Sub Class_initialize()
		Set objAdo = new MyAdo
		Set objCsv = new MyCsv
	End Sub
'
' 機　能：レコードの初期化
' 引　数：count		レコードの項目数
' 例　　：MyAdoFso.initRec 16
'
	Function initRec(count)
		objCsv.initRec count
	End Function
'
' 機　能：レコード（区切り文字含む）を設定
' 引　数：str		文字列（項目名１｛区切り文字｝項目名２｛区切り文字｝...）
' 例　　：MyAdoFso.setRec strRec
'
	Function setRec(str)
		objCsv.setRec str
	End Function
'
' 機　能：レコードを取得
' 戻　値：レコード（区切り文字含む）
' 例　　：strRec = MyAdoFso.getRec
'
	Function getRec()
		getRec = objCsv.getRec
	End Function
'
' 機　能：項目名レコードを設定
' 引　数：str		文字列（項目名１｛区切り文字｝項目名２｛区切り文字｝...）
' 例　　：MyAdoFso.setNameRec strNameRec
'
	Function setNameRec(str)
		objCsv.setNameRec str
	End Function
'
' 機　能：項目名レコードを取得
' 戻り値：文字列（項目名１｛区切り文字｝項目名２｛区切り文字｝...）
' 例　　：strNameRec = MyAdoFso.getNameRec
'
	Function getNameRec()
		getNameRec = objCsv.getNameRec
	End Function
'
' 機　能：区切り文字を設定
' 引　数：del		区切り文字列
' 例　　：MyAdoFso.setDelimiter ","
'
	Function setDelimiter(del)
		objCsv.setDelimiter del
	End Function
'
' 機　能：区切り文字を取得
' 戻り値：文字
' 例　　：charDel = MyAdoFso.getDelimiter
'
	Function getDelimiter()
		getDelimiter = objCsv.getDelimiter
	End Function
'
' 機　能：インデックス指定で項目を取得
' 引　数：i		インデックス（０オリジン）
' 戻り値：文字列
' 例　　：strString = MyAdoFso.getValueByIndex(0)
'
	Function getValueByIndex(i)
		getValueByIndex = objCsv.getValueByIndex(i)
	End Function
'
' 機　能：インデックス指定で項目を設定
' 引　数：i		インデックス（０オリジン）
'	  value		文字列
' 戻り値：文字列
' 例　　：MyAdoFso.setValueByIndex 0, "STRING"
'
	Function setValueByIndex(i, value)
		objCsv.setValueByIndex i, value
	End Function
'
' 機　能：最終インデックスを取得
' 戻り値：整数（０オリジン）
' 例　　：intIndex = MyAdoFso.getLastIndex
'
	Function getLastIndex()
		getLastIndex = objCsv.getLastIndex
	End Function
'
' 機　能：項目名に対応したインデックスを取得
' 引　数：name		項目名
' 戻り値：整数（０オリジン）
' 例　　：intIndex = MyAdoFso.getIndexByName("ELEMENT_01")
'
	Function getIndexByName(name)
		getIndexByName = objCsv.getIndexByName(name)
	End Function
'
' 機　能：項目名指定で項目を取得
' 引　数：name		項目名
' 戻り値：文字列
' 例　　：strString = MyAdoFso.getValueByName("ELEMENT_01")
'
	Function getValueByName(name)
		getValueByName = objCsv.getValueByName(name)
	End Function
'
' 機　能：項目名指定で項目を設定
' 引　数：name		項目名
'	  value		文字列
' 戻り値：文字列
' 例　　：MyAdoFso.setValueByName "ELEMENT_01", 123
'
	Function setValueByName(name, value)
		objCsv.setValueByName name, value
	End Function
'
' 機　能：開く
' 引　数：filename		ファイル名
'	  mode			"r"(読み込み) or "w"(書き込み)
'	  charset		"UTF-8" or "Shift_JIS" or "ASCII" etc
'	  bom			True(UTF-8の時BOM付きで出力) or
'				False(UTF-8の時BOM無しで出力)
' 例　　：MyAdoCsv.open "INPUT.TXT", "r", "UTF-8", False
'
	Function open(filename, mode, charset, bom)
		objAdo.open filename, mode, charset, bom
	End Function
'
' 機　能：入力モードで開く
' 引　数：filename		ファイル名
'	  charset		"UTF-8" or "Shift_JIS" or "ASCII" etc
' 例　　：MyAdoCsv.openInput "INPUT.TXT", "UTF-8"
'
	Function openInput(filename, charset)
		objAdo.openInput filename, charset
	End Function
'
' 機　能：出力モードで開く
' 引　数：filename		ファイル名
'	  charset		"UTF-8" or "Shift_JIS" or "ASCII" etc
' 例　　：MyAdoCsv.openOutput "OUTPUT.TXT", "UTF-8"
'
	Function openOutput(filename, charset)
		objAdo.openOutput filename, charset
	End Function
'
' 機　能：ADOより１行CSVレコード取得し、objCsvへ保存
' 戻り値：True(読めた) or False(読めなかった)
' 例　　：MyAdoCsv.readLine
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
' 機　能：読み取り成功チェック
' 戻り値：true or false
' 例　　：While MyAdoCsv.isReadSuccess
'
	Function isReadSuccess()
		isReadSuccess = objAdo.isReadSuccess
	End Function
'
' 機　能：読み取り失敗チェック
' 戻り値：true or false
' 例　　：While Not MyAdoCsv.isReadFailure
'
	Function isReadFailure()
		isReadFailure = objAdo.isReadFailure
	End Function
'
' 機　能：objCsvからADOに１行CSVレコード出力
' 例　　：MyAdoCsv.writeLine
'
	Function writeLine()
		objAdo.writeLine objCsv.getRec
	End Function
'
' 機　能：閉じる
' 例　　：MyAdoCsv.close
'
	Function close()
		objAdo.close
	End Function
'
' 機　能：終端を取得
' 戻　値：True(終端時) or False(非終端時)
' 例　　：While Not MyAdoCsv.isEof
'
	Function isEof()
		objAdo.isEof
	End Function
'
' 機　能：保存する
' 例　　：MyAdoCsv.save
'
	Private Function save()
		objAdo.save
	End Function
'
' 機　能：先頭の3bytesを除いて保存する
' 例　　：MyAdoCsv.saveWithoutBom
'
	Private Function saveWithoutBom()
		objAdo.saveWithoutBom
	End Function
End Class
''==================================================
''
'' SORTクラス
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
' 機　能：入力ＣＳＶレコードの区切り文字を指定する
' 引　数：str		区切り文字
' 例　　：MySort.setDelimiter ","
'
	Function setDelimiter(str)
		strDelimiter = str
	End Function
'
' 機　能：ソートキー項目を設定する
' 引　数：str		"index:seq:type:max_len, ... ,max_rec_len"
' 例　　：MySort.setKey "0:A:H:256,1:D:H:256,65535"
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
' 機　能：ソート処理に入力レコードを渡す
' 引　数：str		入力レコード
' 例　　：MySort.putRec strRec
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
' 機　能：ソート処理を行う
' 例　　：MySort.sort
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
' 機　能：ソート結果を取り出す
' 例　　：strRec = MySort.getRec
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
' 機　能：ソート結果読み取り成功チェック
' 戻り値：true or false
' 例　　：While isGetSuccess
'
	Function isGetSuccess()
		isGetSuccess = bGetSuccess
	End Function
'
' 機　能：ソート結果読み取り失敗チェック
' 戻り値：true or false
' 例　　：While Not isGetFailure
'
	Function isGetFailure()
		isGetFailure = Not isGetSuccess
	End Function
'
' 機　能：ソート結果の終端を取得
' 戻　値：True(終端時) or False(非終端時)
' 例　　：While Not isEof
'
	Function isEof()
		isEof = objRs.EOF
	End Function
End Class
''==================================================
''
'' スイッチを扱うクラス
''
Class MySwitch
	Dim switch

	Sub Class_Initialize()
		switch = 0
	End Sub
'
' 機　能：スイッチＯＮ
' 例　　：MySwitch.turnOn
'
	Function turnOn
		switch = 1
	End Function
'
' 機　能：スイッチＯＦＦ
' 例　　：MySwitch.turnOff
'
	Function turnOff
		switch = 0
	End Function
'
' 機　能：ＯＮかチェック
' 例　　：If MySwitch.isOn Then
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
' 機　能：ＯＦＦかチェック
' 例　　：If MySwitch.isOff Then
'
	Function isOff
		isOff = (Not isOn)
	End Function
End Class
''==================================================
''
'' ディレクトリを扱うクラス
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
' 機　能：相対パスでディレクトリを指定する
' 引　数：strDir	ディレクトリ
' 例　　：MyDir.setDir "."
'
	Function setDir(str)
		strDir = str
	End Function
'
' 機　能：ディレクトリの絶対パスを得る
' 戻り値：文字列
' 例　　：MyDir.getDirPath
'
	Function getDirPath()
		Dim objFolder
		Set objFolder = objFso.GetFolder(strDir)
		getDirPath = objFso.BuildPath(objFolder, "")
	End Function
'
' 機　能：ディレクトリ内の最初のファイル名を得る
' 戻り値：文字列
' 例　　：MyDir.getFirstFilename
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
' 機　能：次のファイル名を得る
' 戻り値：文字列
' 例　　：MyDir.getNextFilename
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
' 機　能：ディレクトリ内の最初のサブディレクトリ名を得る
' 戻り値：文字列
' 例　　：MyDir.getFirstDirname
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
' 機　能：次のサブディレクトリ名を得る
' 戻り値：文字列
' 例　　：MyDir.getNextDirname
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
'' Excelを扱うクラス
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
' 機　能：ブックを開く
'         存在しないブックの場合は、その場で作成
' 引　数：strBook	ブック(絶対パスでも相対パスでもＯＫ）
' 例　　：MyExcel.open "foo.xlsx"
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
' 機　能：ブックを閉じる
' 例　　：MyExcel.close
'
	Function close()
		objExcel.Workbooks.Close
	End Function
'
' 機　能：ブック内の指定したシートをcsv形式で保存する
' 引　数：intstrSheet	シート
'	  strCsvFile	csvファイル(絶対パスでも相対パスでもＯＫ）
' 例　　：MyExcel.saveAsCsv 1, "foo.csv" or MyExcel.saveAsCSv "Sheet1", "foo.csv"
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
' 機　能：ブックを指定したファイル名で保存する
' 引　数：strBook	ブック(絶対パスでも相対パスでもＯＫ）
' 例　　：MyExcel.saveAs "foo.xlsx"
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
' 機　能：ブックを上書き保存する
' 引　数：strBook	ブック
' 例　　：MyExcel.save
'
	Function save()
		On Error Resume Next
		objExcel.Workbooks(1).Save
		On Error Goto 0
	End Function
'
' 機　能：セルの値を取得する
' 引　数：intstrSheet	シート
'         strRange	レンジ
' 戻り値：文字列
' 例　　：strStr = MyExcel.getCell(1, "B3") or strStr = MyExcel.getCell("Sheet1", "B3")
'
	Function getCell(intstrSheet, strRange)
		getCell = objExcel.Worksheets(intstrSheet).Range(strRange).Value
	End Function
'
' 機　能：セルに値を設定する
' 引　数：intstrSheet	シート
'         strRange	レンジ
'         strValue	設定値
' 例　　：MyExcel.setCell 1, "B3", 123 or MyExcel.setCell "Sheet1", "B3", "設定値"
'
	Function setCell(intstrSheet, strRange, strValue)
		objExcel.Worksheets(intstrSheet).Range(strRange).Value = strValue
	End Function
'
' 機　能：セルの計算式を取得する
' 引　数：intstrSheet	シート
'         strRange	レンジ
' 戻り値：文字列
' 例　　：strStr = MyExcel.getFormula(1, "B3") or strStr = MyExcel.getFormula("Sheet1", "B3")
'
	Function getFormula(intstrSheet, strRange)
		getFormula = objExcel.Worksheets(intstrSheet).Range(strRange).Formula
	End Function
'
' 機　能：セルに計算式を設定する
' 引　数：intstrSheet	シート
'         strRange	レンジ
'         strFormula	計算式
' 例　　：MyExcel.setFormula 1, "B3", "=A10+B10" or MyExcel.setFormula "Sheet1", "B3", "=A10+B10"
'
	Function setFormula(intstrSheet, strRange, strFormula)
		objExcel.Worksheets(intstrSheet).Range(strRange).Formula = strFormula
	End Function
'
' 機　能：セルのフォント名を取得する
' 引　数：intstrSheet	シート
'         strRange	レンジ
' 戻り値：文字列
' 例　　：strStr = MyExcel.getFontName(1, "B3") or strStr = MyExcel.getFontName("Sheet1", "B3")
'
	Function getFontName(intstrSheet, strRange)
		getFontName = objExcel.Worksheets(intstrSheet).Range(strRange).Font.Name
	End Function
'
' 機　能：セルのフォント名を設定する
' 引　数：intstrSheet	シート
'         strRange	レンジ
'         strFontName	フォント名
' 例　　：MyExcel.setFontName 1, "B3", "ＭＳ　ゴシック" or MyExcel.setFontName "Sheet1", "B3", "ＭＳ　ゴシック"
'
	Function setFontName(intstrSheet, strRange, strFontName)
		objExcel.Worksheets(intstrSheet).Range(strRange).Font.Name = strFontName
	End Function
'
' 機　能：セルのフォントサイズを取得する
' 引　数：intstrSheet	シート
'         strRange	レンジ
' 戻り値：数値
' 例　　：realNum = MyExcel.getFontSize(1, "B3") or realNum = MyExcel.getFontSize("Sheet1", "B3")
'
	Function getFontSize(intstrSheet, strRange)
		getFontSize = objExcel.Worksheets(intstrSheet).Range(strRange).Font.Size
	End Function
'
' 機　能：セルのフォントサイズを設定する
' 引　数：intstrSheet	シート
'         strRange	レンジ
'         realFontSize	フォントサイズ
' 例　　：MyExcel.setFontSize 1, "B3", 10.5 or MyExcel.setFontSize "Sheet1", "B3", 10.5
'
	Function setFontSize(intstrSheet, strRange, realFontSize)
		objExcel.Worksheets(intstrSheet).Range(strRange).Font.Size = realFontSize
	End Function
'
' 機　能：セルの文字色を取得する
' 引　数：intstrSheet	シート
'         strRange	レンジ
' 戻り値：数値
' 例　　：intNum = MyExcel.getForegroundColor(1, "B3") or intNum = MyExcel.getForegroundColor("Sheet1", "B3")
'
	Function getForegroundColor(intstrSheet, strRange)
		getForegroundColor = objExcel.Worksheets(intstrSheet).Range(strRange).Font.Color
	End Function
'
' 機　能：セルの文字色を設定する
' 引　数：intstrSheet	シート
'         strRange	レンジ
'         intColor	文字色
' 例　　：MyExcel.setForegroundColor 1, "B3", &HFF0000 or MyExcel.setForegroundColor "Sheet1", "B3", RGB(255,0,0)
'
	Function setForegroundColor(intstrSheet, strRange, intColor)
		objExcel.Worksheets(intstrSheet).Range(strRange).Font.Color = intColor
	End Function
'
' 機　能：セルの背景色を取得する
' 引　数：intstrSheet	シート
'         strRange	レンジ
' 戻り値：数値
' 例　　：intNum = MyExcel.getBackgroundColor(1, "B3") or intNum = MyExcel.getBackgroundColor("Sheet1", "B3")
'
	Function getBackgroundColor(intstrSheet, strRange)
		getBackgroundColor = objExcel.Worksheets(intstrSheet).Range(strRange).Interior.Color
	End Function
'
' 機　能：セルの背景色を設定する
' 引　数：intstrSheet	シート
'         strRange	レンジ
'         intColor	背景色(&HBBGGRR)
' 例　　：MyExcel.setBackgroundColor 1, "B3", &HFF0000 or MyExcel.setBackgroundColor "Sheet1", "B3", RGB(255,0,0)
'
	Function setBackgroundColor(intstrSheet, strRange, intColor)
		objExcel.Worksheets(intstrSheet).Range(strRange).Interior.Color = intColor
	End Function
'
' 機　能：セルの幅を取得する
' 引　数：intstrSheet	シート
'         strColumn	カラム
' 戻り値：数値
' 例　　：realNum = MyExcel.getColumnWidth(1, "B") or realNum = MyExcel.getColumnWidth("Sheet1", "B")
'
	Function getColumnWidth(intstrSheet, strColumn)
		getColumnWidth = objExcel.Worksheets(intstrSheet).Columns(strColumn).ColumnWidth
	End Function
'
' 機　能：セルの幅を設定する
' 引　数：intstrSheet	シート
'         strColumn	カラム
'         realWidth	幅
' 例　　：MyExcel.setColumnWidth 1, "B3", 10.5 or MyExcel.setColumnWidth "Sheet1", "B3", 10.5
'
	Function setColumnWidth(intstrSheet, strColumn, realWidth)
		objExcel.Worksheets(intstrSheet).Columns(strColumn).ColumnWidth = realWidth
	End Function
'
' 機　能：セルの高さを取得する
' 引　数：intstrSheet	シート
'         intRow	行
' 戻り値：数値
' 例　　：realNum = MyExcel.getRowHeight(1, 2) or realNum = MyExcel.getRowHeight("Sheet1", 2)
'
	Function getRowHeight(intstrSheet, intRow)
		getRowHeight = objExcel.Worksheets(intstrSheet).Rows(intRow).RowHeight
	End Function
'
' 機　能：セルの高さを設定する
' 引　数：intstrSheet	シート
'         intRow	行
'         realHeight	高さ
' 例　　：MyExcel.setRowHeight 1, 2, 10.5 or MyExcel.setRowHeight "Sheet1", 2, 10.5
'
	Function setRowHeight(intstrSheet, intRow, realHeight)
		objExcel.Worksheets(intstrSheet).Rows(intRow).RowHeight = realHeight
	End Function
'
' 機　能：セルの書式形式を取得する
' 引　数：intstrSheet	シート
'         strRange	レンジ
' 戻り値：文字列
' 例　　：strStr = MyExcel.getFormat(1, "B3") or strStr = MyExcel.getFormat("Sheet1", "B3")
'
	Function getFormat(intstrSheet, strRange)
		getFormat = objExcel.Worksheets(intstrSheet).Range(strRange).NumberFormatLocal
	End Function
'
' 機　能：セルの書式形式を設定する
' 引　数：intstrSheet	シート
'         strRange	レンジ
'         strFormat	書式形式
' 例　　：MyExcel.setFormat 1, "B3", "000000" or MyExcel.setFormat "Sheet1", "B3", "000000"
'
	Function setFormat(intstrSheet, strRange, strFormat)
		objExcel.Worksheets(intstrSheet).Range(strRange).NumberFormatLocal = strFormat
	End Function
'
' 機　能：指定したシートの印刷プレビューを表示する
' 引　数：intstrSheet	シート
' 例　　：MyExcel.preview 1 or MyExcel.preview "Sheet1"
'
	Function preview(intstrSheet)
		objExcel.Visible = True
		objExcel.Worksheets(intstrSheet).PrintPreview
		objExcel.Visible = False
	End Function
'
' 機　能：指定したシートを印刷する
' 引　数：intstrSheet	シート
' 例　　：MyExcel.print 1 or MyExcel.print "Sheet1"
'
	Function print(intstrSheet)
		objExcel.Worksheets(intstrSheet).PrintOut
	End Function
'
' 機　能：ブックを表示する
' 例　　：MyExcel.onVisible
'
	Function onVisible()
		objExcel.Visible = True
	End Function
'
' 機　能：ブックを非表示にする
' 例　　：MyExcel.offVisible
'
	Function offVisible()
		objExcel.Visible = False
	End Function
'
' 機　能：アラートを表示する
' 例　　：MyExcel.onAlert
'
	Function onAlert()
		objExcel.DisplayAlerts = True
	End Function
'
' 機　能：アラートを非表示にする
' 例　　：MyExcel.offAlert
'
	Function offAlert()
		objExcel.DisplayAlerts = False
	End Function
End Class
''==================================================
''
'' オプションを扱うクラス
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
' 機　能：扱うオプションを設定する
' 引　数：strAllOption	"-option1=[y|n],..."
'                                 y means option take value like "-option1 value"
'                                 n means option do not take value"
' 例　　：MyOption.initialize("-h=n,--help=n,-i=y,--input=y,-o=y,--output=y")
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
' 機　能：オプションの値を取得する
' 引　数：strOptions	オプション文字列
' 戻り値：文字列
' 例　　：strOpt = MyOption.getValue("--input")
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
' 機　能：オプションが指定されているか取得する
' 引　数：strOptions	オプション文字列
' 戻り値：True or False
' 例　　：If MyOption.isSpecified("--help") Then
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
' 機　能：オプション以外の文字列を取得する
' 戻り値：文字列(空白区切り)
' 例　　：strNonOpt = MyOption.getNonOptions()
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
' 機　能：オプション以外の文字列を配列で取得する
' 戻り値：文字列配列
' 例　　：strArraynNonOpt = MyOption.getArrayNonOptions()
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
'' ファイルシステム操作を扱うクラス
''
Class MyFsOpe
	Sub Class_Initialize()
	End Sub

'
' 機　能：一時ファイル名の取得
' 戻り値：文字列
' 例　　：strTempFile = MyFsOpe.getTempFileName
'
	Function getTempFileName()
		getTempFileName = CreateObject("Scripting.FileSystemObject").GetTempName
	End Function

'
' 機　能：ディレクトリの作成
' 引　数：strDirName
' 例　　：MyFsOpe.createFolder "TEMP.dir"
'
	Function createFolder(strFileName)
		CreateObject("Scripting.FileSystemObject").createFolder strDirName
	End Function

'
' 機　能：ファイルの削除
' 引　数：strFileName
' 例　　：MyFsOpe.deleteFile "TEMP.tmp"
'
	Function deleteFile(strFileName)
		CreateObject("Scripting.FileSystemObject").DeleteFile strFileName, True
	End Function

'
' 機　能：ディレクトリの削除
' 引　数：strDirName
' 例　　：MyFsOpe.deleteFolder "TEMP.dir"
'
	Function deleteFolder(strDirName)
		CreateObject("Scripting.FileSystemObject").DeleteFolder strDirName, True
	End Function
End Class
''==================================================
''
'' 諸々を扱うクラス
''
Class MyMisc
	Sub Class_Initialize()
	End Sub

'
' 機　能：プログラムの終了
' 引　数：終了ステータス
' 例　　：MyMisc.exitProg 0
'
	Function exitProg(stat)
		WScript.Quit(stat)
	End Function
End Class
'==================================================
'
' デバッグ表示用
'
Function Debug(str)
	Dim objStdio : Set objStdio = new MyStdio
	objStdio.writeLine(str)
End Function
