Option Explicit

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

Class Config
	Dim dic
	Dim ConfigFile

' ------------------------------------------------------------
' I/O
' ------------------------------------------------------------
	Private Function AdjustName(filename)
		Dim path, name
		path = FSO.GetParentFolderName(filename)
		If path="" Then
			path = FSO.GetParentFolderName(WScript.ScriptFullName)
		End If
		name = FSO.GetFileName(filename)
		AdjustName = FSO.BuildPath(path, name)
	End Function

	' XMLファイルから取り込み
	Public Function LoadXML(filename)
		Init
		Dim fullname
		fullname = AdjustName(filename)

		' 無かったら初期作成
		If Not FSO.FileExists(fullname) Then
			Add "PARAMETER", "debug", "false"
			SaveXML fullname
		End If

		' 内容
'		With FSO.OpenTextFile(FSO.BuildPath(path, name))
'			CreateObject("WScript.Shell").Popup .ReadAll(), 3
'			.Close
'		End With

		' XML
		Dim Source, sec, secname, elem, key, value
		Set Source = CreateObject("Msxml2.DOMDocument")
		Source.load fullname
		If Source.parseError.errorCode = 0 Then
			LoadXML = ""
			For Each sec in Source.getElementsByTagName("section")
				secname = sec.getAttribute("name")
				For Each elem in sec.getElementsByTagName("value")
					key = elem.getAttribute("key")
					value = elem.getAttribute("value")
					If VarType(value)<>vbNull Then
						Append secname, key, value
					Else
						Dim item
						For Each item In elem.getElementsByTagName("item")
							value = item.getAttribute("value")
							If VarType(value)<>vbNull Then
								Append secname, key, value
							End If
						Next
					End If
				Next
			Next
		Else
			With Source.parseError
				LoadXML = fullname & "(" & .line & ", " & .linepos & ") : " & .reason
			End With
		End If

		ConfigFile = fullname
	End Function

	' XMLファイルに保存 (項目の順番や、手書きで書いたコメント <!-- --> などは消えますので注意)
	Public Sub SaveXML(filename)
		Init
		Dim fullname, ts
		fullname = AdjustName(filename)
		If FSO.FileExists(fullname) Then
			Dim regEx
			Set regEx = New RegExp
			regEx.Pattern = "[/\:\s]+"
			regEx.Global = true
			Dim fil, bckname
			bckname = fullname & "." & regEx.Replace(FormatDateTime(Now,0),"") & ".bck"
			Set fil = FSO.GetFile(fullname)
			fso.MoveFile fullname, fullname & "." & regEx.Replace(FormatDateTime(Now,0),"") & ".bck"
			Set fil = Nothing
		End If
		Set ts = FSO.CreateTextFile(fullname)
		ts.WriteLine "<?xml version='1.0' encoding='Shift_JIS' ?>"
		ts.WriteLine "<config>"
		Dim secName, keyName, value
		For Each secName In dic.Keys
			ts.WriteLine vbTab & "<section name='" & secName & "'>"
			With dic.Item(secName)
				For Each keyName IN .Keys
					value = .Item(keyName)
					If IsArray(value) Then
						ts.WriteLine vbTab & vbTab & "<value key='" & keyName & "'/>"
						Dim item
						For Each item In value
							ts.WriteLine vbTab & vbTab & "<item value='" & value & "'/>"
						Next
						ts.WriteLine vbTab & vbTab & "<value/>"
					Else
						ts.WriteLine vbTab & vbTab & "<value key='" & keyName & "' value='" & value & "'/>"
					End If
				Next
			End With
			ts.WriteLine vbTab & "</section>"
		Next
		ts.WriteLine "</config>"
		ts.Close
		Set ts = Nothing
	End Sub

' ------------------------------------------------------------
' 内部操作
' ------------------------------------------------------------

	' 値を取得する
	Public property Get Value(sec, key, defaultValue)
		Dim ret
		If Not ContainsKey(sec, key) Then
			ret = defaultValue
		Else
			ret = GetSection(sec).item(key)
		End If
		If IsArray(ret) Then
			ret = Join(ret, ",")
		End If
		Value = ret
	End property

	' 値を配列で取得する
	Public property Get ArrayValue(sec, key, defaultValue)
		Dim ret
		If Not ContainsKey(sec, key) Then
			ret = defaultValue
		Else
			ret = GetSection(sec).item(key)
		End If
		If IsArray(ret) Then
			ArrayValue = ret
		Else
			ArrayValue = Array(ret)
		End If
	End property

	' セクション名の配列
	Public Function Sections
		Init
		Sections = dic.Keys
	End Function

	' 指定のセクションに含まれるキーの配列
	Public Function Keys(sec)
		Init
		If ContainsSection(sec) Then
			Keys = GetSection(sec).Keys
		End If
	End Function

	' 要素を追加
	' 既に存在していたら置き換え
	Public Sub Add(sec, key, item)
		Init
		Dim secDic
		If ContainsSection(sec) Then
			Set secDic = GetSection(sec)
		Else
			Set secDic = CreateObject("Scripting.Dictionary")
			dic.Add sec, secDic
		End If
		If secDic.Exists(key) Then secDic.Remove key
		secDic.Add key, item
	End Sub

	' 要素を追加
	' 既に存在していたら配列化
	Public Sub Append(sec, key, item)
		If ContainsKey(sec, key) Then
			Dim val
			val = GetSection(sec).item(key)
			item = MergeArray(val, item)
		End If
		Add sec, key, item
	End Sub

	' セクションが存在するかチェック
	Public Function ContainsSection(sec)
		Init
		ContainsSection = dic.Exists(sec)
	End Function

	' キーが存在するかチェック
	Public Function ContainsKey(sec, key)
		ContainsKey = False
		if Not ContainsSection(sec) Then Exit Function
		ContainsKey = GetSection(sec).Exists(key)
	End Function

	' 全クリア
	Public Sub Clear
		Init
		Dim sec
		For Each sec in dic.Keys
			GetSection(sec).RemoveAll
		Next
		dic.RemoveAll
	End Sub

	' セクション削除
	Public Sub RemoveSection(sec)
		Init
		If ContainsSection(sec) Then dic.Remove sec
	End Sub

	' キー削除
	Public Sub RemoveKey(sec, key)
		Init
		If ContainsKey(sec, key) Then GetSection(sec).Remove key
	End Sub

' ------------------------------------------------------------
' Private
' ------------------------------------------------------------

	' 初期化
	Private Sub Init()
		If IsEmpty(dic) Then Set dic = CreateObject("Scripting.Dictionary")
	End Sub

	' セクションのオブジェクト (dictionary)
	Private Function GetSection(sec)
		Set GetSection = Nothing
		Init
		If Not ContainsSection(sec) Then Exit Function
		Set GetSection = dic.Item(sec)
	End Function

' ------------------------------------------------------------
' Common
' ------------------------------------------------------------

	' val と item を結合した配列を返します
	Public Function MergeArray(val, item)
		Dim arr()
		If Not IsArray(val) Then
			MergeArray = MergeArray(Array(val), item)
		ElseIf Not IsArray(item) Then
			MergeArray = MergeArray(val, Array(item))
		Else
			Dim count, inc, idx
			count = UBound(val) + 1
			inc = UBound(item) + 1
			' 格納に必要な新規の配列を作成
			ReDim arr(count + inc - 1)
			' val をコピー
			For idx=0 To count-1
				arr(idx) = val(idx)
			Next
			' item をコピー
			For idx=0 To inc-1
				arr(count + idx) = item(idx)
			Next
			MergeArray = arr
		End If
	End Function

End Class
' --------------------------------------------------------------------------------
