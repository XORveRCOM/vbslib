		' フォルダ用のショートカットを作成
		Sub CreateExplorerFolderShortcut(lnkPath, pathname, overwrite)
			Dim fol
			Set fol = fso.GetFolder(pathname)

			Dim LinkPath, lnkName, ExplorerPath
			LinkPath = fol.Path
			lnkName = fol.Name
			ExplorerPath = Shell.ExpandEnvironmentStrings("%SystemRoot%") & "\explorer.exe"

			Dim Shortcut
			Set Shortcut = NewShortcut(lnkPath, lnkName, overwrite)
			If Shortcut Is Nothing Then Exit Sub
			With Shortcut
				.Arguments = " /e,/n,""" & LinkPath & """"
				.Description = "folder:" & LinkPath
				.IconLocation = ExplorerPath & ",0"
				.TargetPath = ExplorerPath
				.WindowStyle = 1
				.WorkingDirectory = LinkPath
				.Save
			End With
		End Sub

		' ファイル用のショートカットを作成
		Sub CreateExplorerFileShortcut(lnkPath, pathname, overwrite)
			Dim file
			Set file = fso.GetFile(pathname)

			Dim LinkPath, lnkName, ExplorerPath
			LinkPath = file.Path
			lnkName = file.Name

			Dim Shortcut
			Set Shortcut = NewShortcut(lnkPath, lnkName, overwrite)
			If Shortcut Is Nothing Then Exit Sub
			With Shortcut
				.Description = pathname
				.TargetPath = LinkPath
				.WindowStyle = 1
				.WorkingDirectory = file.ParentFolder.Path
				.Save
			End With
		End Sub

		' フォルダ／ファイル用のショートカット作成共通処理
		Function NewShortcut(lnkPath, lnkName, overwrite)
			Set NewShortcut = Nothing
			Dim Shortcut, lnkFilePath
			lnkFilePath = fso.BuildPath(lnkPath, lnkName & ".lnk")
			if fso.FileExists(lnkFilePath) then
				If overwrite Then
					fso.DeleteFile(lnkFilePath)
				Else
					Set Shortcut = Shell.CreateShortcut(lnkFilePath)
					lnkName = InputBox( _
						"Original Path : """ & Shortcut.WorkingDirectory & """" _
						& vbLf & "名前を変更してください" _
						& vbLf & "(Cancelならば実行されません)" _
						, "ショートカットは存在します", lnkName)
					If lnkName<>"" Then
						Set NewShortcut = NewShortcut(lnkPath, lnkName, overwrite)
					End If
					Exit Function
				end if
			end if
			Set NewShortcut = Shell.CreateShortcut(lnkFilePath)
		End Function

		' 現在のユーザの Quick Launch フォルダを返す
		Function QuickLaunchPath
			' AppData は隠し指定
			QuickLaunchPath = fso.GetAbsolutePathName( _
				fso.BuildPath( _
					Shell.SpecialFolders("AppData") _
					, "Microsoft\Internet Explorer\Quick Launch" _
				) _
			)
		End Function
