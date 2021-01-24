Option Explicit

	' ���K�\���̃z���C�g���X�g
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
		' �o�^
		Sub Init(arr)
			Dim item
			For Each item in arr
				AddPattern item
			Next
		End Sub
		' ���K�\�����X�g�ƈ�v���邩�`�F�b�N
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
		' ��v�������K�\���̌�
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
