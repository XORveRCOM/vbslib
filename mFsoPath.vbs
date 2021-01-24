Option Explicit

Public Const PathSeparator = "\"
Public Const RelativePathHeader = ".\"

Class FsoPath

	Dim fso_impl

	Function fso()
		If IsEmpty(fso_impl) Then
			Set fso_impl = CreateObject("Scripting.FileSystemObject")
		End If
		Set fso = fso_impl
	End Function

	' --------------------------------------------------
	Public Function CurrentScriptName
		CurrentScriptName = WScript.ScriptFullName
	End Function
	Public Function CurrentScriptPath
		CurrentScriptPath = fso.GetParentFolderName(CurrentScriptName)
	End Function

	' --------------------------------------------------
	' 相対パス -> 絶対パス
	Public Function AbsolutePath(ByVal path)
		If path = "" Then path = RelativePathHeader
		If left(path, 1) = "." Then path = fso.BuildPath(CurrentPath, path)
		AbsolutePath = fso.GetAbsolutePathName(path) & PathSeparator
	End Function

	' カレントパス
	Public Function CurrentPath()
		CurrentPath = CurrentScriptPath & PathSeparator
	End Function

	' ファイルの絶対パス
	Public Function AbsoluteFile(fileName)
		AbsoluteFile = ExtractPathName(fileName) & ExtractFileName(fileName)
	End Function

	' --------------------------------------------------
	' パスの存在確認
	Public Function DirExists(fnam)
		DirExists = fso.FolderExists(fnam)
	End Function
	' ファイルの存在確認
	Public Function FileExists(fnam)
		FileExists = fso.FileExists(fnam)
	End Function

	' --------------------------------------------------
	' パス名作成
	Public Function CombinePath(path, path2)
		CombinePath = AbsolutePath(fso.BuildPath(path, path2))
	End Function
	' ファイル名作成
	Public Function CombineFile(path, file)
		CombineFile = AbsoluteFile(fso.BuildPath(path, file))
	End Function

	' --------------------------------------------------
	' 深いパスを一括作成
	Public Sub CreatePathDir(pathdir)
		Dim path
		path = ""
		Dim dnam
		For Each dnam In SplitPath(pathdir)
			path = path & dnam
			If Not DirExists(path) Then MkDir path
			path = path & PathSeparator
		Next
	End Sub

	Public Function SplitPath(ByVal path)
		path = AbsolutePath(path)
		If Right(path, 1) = PathSeparator Then path = left(path, Len(path) - 1)
		SplitPath = Split(path, PathSeparator)
	End Function

	Public Function FileNamePos(ByVal FullName)
		FileNamePos = 1
		Do Until 0 = InStr(FileNamePos, FullName, "\")
			FileNamePos = InStr(FileNamePos, FullName, "\") + 1
		Loop
	End Function

	' --------------------------------------------------
	' パスの取り出し
	Public Function ExtractPathName(ByVal FullName)
		If fso.GetParentFolderName(FullName) = "" Then
			ExtractPathName = CurrentPath
		Else
			If fso.FolderExists(FullName) Then
				ExtractPathName = FullName
			Else
				ExtractPathName = fso.GetParentFolderName(FullName)
			End If
			ExtractPathName = AbsolutePath(ExtractPathName)
		End If
	End Function

	' ファイル名の取り出し
	Public Function ExtractFileName(ByVal FullName)
		ExtractFileName = fso.GetFileName(FullName)
	End Function

	' ファイルBASE名の取り出し
	Public Function ExtractBaseName(ByVal FullName)
		ExtractBaseName = fso.GetBaseName(ExtractFileName(FullName))
	End Function

	' ファイル拡張子の取り出し
	Public Function ExtractExtensionName(ByVal FullName)
		ExtractExtensionName = fso.GetExtensionName(ExtractFileName(FullName))
	End Function

	' --------------------------------------------------

	Public Function DateCreated(fileName)
		DateCreated = fso.GetFile(AbsoluteFile(fileName)).DateCreated
	End Function
	Public Function DateUpdated(fileName)
		DateUpdated = fso.GetFile(AbsoluteFile(fileName)).DateLastModified
	End Function

' --------------------------------------------------
End Class
