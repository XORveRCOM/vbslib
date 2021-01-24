	' 追加記述の楽な配列
	Class Vector
		Dim pArr()
		Dim pIsEmpty
		Dim pCount

		' 初期化
		Sub Init
			If IsEmpty(pIsEmpty) Then
				pIsEmpty = True
				Clear
			End If
		End Sub

		' クリア
		Public Sub Clear
			Init
			pIsEmpty = True
			ReDim pArr(0)
			pArr(0) = Empty
			pCount = 0
		End Sub

		' 配列
		Public Property Get Array
			Init
			Array = pArr
		End Property

		' 配列要素数
		Public Function Count
			Init
			If pIsEmpty Then
				Count = 0
			Else
				Count = pCount
			End If
		End Function

		' 配列末尾に要素を追加
		Public Function Append(val)
			Init
			If Not pIsEmpty Then
				ReDim Preserve pArr(pCount)
			End If
			pIsEmpty = False
			If IsObject(val) Then
				Set pArr(pCount) = val
			Else
				pArr(pCount) = val
			End If
			pCount = pCount + 1
			Set Append = Me
		End Function
	End Class
