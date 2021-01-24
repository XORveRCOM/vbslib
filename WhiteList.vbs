Option Explicit

	' 正規表現のホワイトリスト
	Class WhiteList
		Dim dict
		Function GetDic
			If IsEmpty(dict) Then
				Set dict = CreateObject("Scripting.Dictionary")
			End If
			Set GetDic = dict
		End Function
		Sub AddPattern(pat)
			Dim regex
			Set regex = New RegExp
			regex.Pattern = pat
			regex.IgnoreCase = True
			GetDic.Add pat, regex
		End Sub
		' 登録
		Sub Init(arr)
			Dim item
			For Each item in arr
				AddPattern item
			Next
		End Sub
		' 正規表現リストと一致するかチェック
		Function IsMatch(check)
			Dim regex
			For Each regex In GetDic.Items
				If Not regex.Test(check) Then
					IsMatch = False
					Exit Function
				End If
			Next
			IsMatch = True
		End Function
		' 一致した正規表現の個数
		Function MatchCount(check)
			Dim regex, cnt
			cnt = 0
			For Each regex In GetDic.Items
				If regex.Test(check) Then
					cnt = cnt + 1
				End If
			Next
			MatchCount = cnt
		End Function
	End Class
