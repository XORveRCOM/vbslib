	' �ǉ��L�q�̊y�Ȕz��
	Class Vector
		Dim pArr()
		Dim pIsEmpty
		Dim pCount

		' ������
		Sub Init
			If IsEmpty(pIsEmpty) Then
				pIsEmpty = True
				Clear
			End If
		End Sub

		' �N���A
		Public Sub Clear
			Init
			pIsEmpty = True
			ReDim pArr(0)
			pArr(0) = Empty
			pCount = 0
		End Sub

		' �z��
		Public Property Get Array
			Init
			Array = pArr
		End Property

		' �z��v�f��
		Public Function Count
			Init
			If pIsEmpty Then
				Count = 0
			Else
				Count = pCount
			End If
		End Function

		' �z�񖖔��ɗv�f��ǉ�
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
