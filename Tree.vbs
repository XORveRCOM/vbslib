Option Explicit

Class Tree
	Public Value

	Dim branches
	Dim namestr
	Dim itPos

	Private Sub Class_Initialize
		Set branches = CreateObject("Scripting.Dictionary")
	End Sub
	Private Sub Class_Terminate
		RemoveBranchAll
	End Sub

	Public Property Get Name
		Name = namestr
	End Property
	Public Property Let Name(nam)
		If namestr="" Then
			namestr = nam
		End If
	End Property

	Public Function countBranch()
		countBranch = branches.Count
	End Function

	Public Function branch(ByVal nam)
		If IsString(nam) Then
			If Not branches.Exists(nam) Then
				Set branch = Nothing
				Exit Function
			End If
			Set branch = branches.Item(nam)
		Else
			If nam<0 Or branches.Count<=nam Then
				Set branch = Nothing
				Exit Function
			End If
			Dim a
			a = branches.Items
			Set branch = a(nam)
		End If
	End Function
	Public Function branchValue(ByVal nam, ByVal defval)
		If branches.Exists(nam) Then
			Substitution branchValue, branch(nam).Value
		Else
			Substitution branchValue, defval
		End If
	End Function

	Public Function addBranch(ByVal nam)
		If branches.Exists(nam) Then
			Set addBranch = branches.Item(nam)
		Else
			Set addBranch = New Tree
			addBranch.Name = nam
			branches.Add nam, addBranch
		End If
	End Function

	' 枝の問い合わせ
	Public Function Contains(ByVal index)
		If IsString(index) Then
			Contains = branches.Exists(index)
		Else
			Contains = ( 0<=index And index<branches.Count )
		End If
	End Function

	' 枝の除去
	Public Function removeBranch(ByVal index)
		If Contains(index) Then
			branch(index).RemoveBranchAll
			If IsString(index) Then
				branches.Remove index
			Else
				Dim keys
				keys = branches.Keys
				branches.Remove keys(index)
			End If
		End If
	End Function
	Public Sub RemoveBranchAll()
		Dim i, keys
		keys = branches.Keys
		For i = 0 To branches.Count-1
			removeBranch keys(i)
		Next
	End Sub

	Function IsString(ByVal str)
		IsString = Not IsNumeric(str)
	End Function

' ------------------------------------------------------------
	' tree の追加
	' [自分]
	'   [自分.branches]
	'   ・・・
	'   +
	' [br]
	'   [br.branches]
	'   ・・・
	'   =
	' [自分]
	'   [自分.branches]
	'   [br]
	'       [br.branches]
	'       ・・・
	Public Sub addBranchs(br)
		If br.name = "" Then br.name = Format(countBranch, String(15, "0"))
		branches.Add br, br.name
	End Sub

	' tree の挿入
	' [自分]
	'   [自分.branches]
	'   ・・・
	'   +
	' [br]
	'   [br.branches]
	'   ・・・
	'   =
	' [自分]
	'   [自分.branches]
	'   ・・・
	'   [br.branches]
	'   ・・・
	Public Sub importBranch(tr)
		With tr.Clone
			.MoveFirst
			Do While .HasNext
				addBranchs .MoveNext
			Loop
		End With
	End Sub

	' 枝の複製
	Public Function Clone()
		Set Clone = New Tree
		If IsObject(value) Then
			Set Clone.value = value
			' Clone メソッドを持っていたら使う
			On Error Resume Next
			Set Clone.value = value.Clone
			On Error GoTo 0
		Else
			Clone.value = value
		End If
		Dim posSav
		posSav = itPos
		MoveFirst
		Do While HasNext
			With MoveNext
				Dim tr
				Set tr = .Clone
				tr.Name = .Name
				Clone.addBranchs tr
			End With
		Loop
		itPos = posSav
	End Function

	Public Sub MoveFirst()
		itPos = 0
	End Sub
	Public Function HasNext()
		HasNext = itPos<branches.Count
	End Function
	Public Function MoveNext()
		Dim a
		a = branches.Items
		Set MoveNext = a(itPos)
		itPos = itPos + 1
	End Function
End Class

' ------------------------------------------------------------
