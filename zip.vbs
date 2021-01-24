Option Explicit

Dim fso, Shell
Set fso = CreateObject("Scripting.FileSystemObject")
Set Shell = CreateObject("Shell.Application")

class ZIP
	Dim dic

	Sub Init
		If IsEmpty(dic) Then Set dic = CreateObject("Scripting.Dictionary")
	End Sub

	' 圧縮するファイルを登録
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
			' ファイルがなかった
			MsgBox filename & " is not exist,"
		End If
	End Sub

	' zipファイルとして圧縮
	public Function Save(byval filename, IsForceReplace, timeout)
		Init
		Dim name, path
		path = fso.GetParentFolderName(filename)
		If path="" Then path = fso.GetParentFolderName(WScript.ScriptFullName)
		name = FSO.GetFileName(filename)
		filename = fso.BuildPath(path, name)

		' 格納先が存在しない場合、空のzipファイルを作成する
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

		' zipファイルを圧縮フォルダとして開き、格納元ファイルをコピーする
		Dim zip_ns, srcname
		Set zip_ns = Shell.NameSpace(filename)
		For Each name in dic.Keys
			path = dic.item(name)
			srcname = fso.BuildPath(path, name)
			zip_ns.CopyHere srcname, CLng(16)
		Next
		WScript.Sleep 1000

		' コピーは非同期なので、オープンできるまで待つ
		Save = WaitForUnlock(filename, timeout)

		' zipファイルを閉じる
		Set zip_ns = Nothing
	End function

	' オープンできるまで待つ
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
