Option Explicit

Dim fso, Shell
Set fso = CreateObject("Scripting.FileSystemObject")
Set Shell = CreateObject("Shell.Application")

class ZIP
	Dim dic

	Sub Init
		If IsEmpty(dic) Then Set dic = CreateObject("Scripting.Dictionary")
	End Sub

	' ���k����t�@�C����o�^
	public Sub Add(byval filename)
		Init
		Dim name, path
		path = fso.GetParentFolderName(filename)
		If path="" Then path = fso.GetParentFolderName(WScript.ScriptFullName)
		name = FSO.GetFileName(filename)
		filename = fso.BuildPath(path, name)

		If fso.FileExists(filename) Then
			dic.Add name, path
		Else
			' �t�@�C�����Ȃ�����
			MsgBox filename & " is not exist,"
		End If
	End Sub

	' zip�t�@�C���Ƃ��Ĉ��k
	public Function Save(byval filename, IsForceReplace, timeout)
		Init
		Dim name, path
		path = fso.GetParentFolderName(filename)
		If path="" Then path = fso.GetParentFolderName(WScript.ScriptFullName)
		name = FSO.GetFileName(filename)
		filename = fso.BuildPath(path, name)

		' �i�[�悪���݂��Ȃ��ꍇ�A���zip�t�@�C�����쐬����
		If Not fso.FileExists(filename) Or IsForceReplace=True Then
			With Fso.CreateTextFile(filename, True)
				.Write Chr(&H50)
				.Write Chr(&H4B)
				.Write Chr(&H05)
				.Write Chr(&H06)
				.Write String(18, Chr(0))
				.Close
			End With
		End If

		' zip�t�@�C�������k�t�H���_�Ƃ��ĊJ���A�i�[���t�@�C�����R�s�[����
		Dim zip_ns, srcname
		Set zip_ns = Shell.NameSpace(filename)
		For Each name in dic.Keys
			path = dic.item(name)
			srcname = fso.BuildPath(path, name)
			zip_ns.CopyHere srcname, CLng(16)
		Next
		WScript.Sleep 1000

		' �R�s�[�͔񓯊��Ȃ̂ŁA�I�[�v���ł���܂ő҂�
		Save = WaitForUnlock(filename, timeout)

		' zip�t�@�C�������
		Set zip_ns = Nothing
	End function

	' �I�[�v���ł���܂ő҂�
	Function WaitForUnlock(filename, timeout)
		Dim toTime
		toTime = DateAdd("s", timeout, Now)
		Dim dummy
		Const ForAppending = 8
		Set dummy = Nothing
		Do While dummy Is Nothing
			If timeout>0 And toTime<Now Then
				WaitForUnlock = False
				Exit Function
			End If
			On Error Resume Next
			Set dummy = fso.OpenTextFile(filename, ForAppending, False)
			On Error Goto 0
			WScript.Sleep 500
		Loop
		dummy.Close
		Set dummy = Nothing
		WaitForUnlock = True
	End Function
End class
