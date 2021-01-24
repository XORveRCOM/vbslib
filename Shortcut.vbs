		' �t�H���_�p�̃V���[�g�J�b�g���쐬
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

		' �t�@�C���p�̃V���[�g�J�b�g���쐬
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

		' �t�H���_�^�t�@�C���p�̃V���[�g�J�b�g�쐬���ʏ���
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
						& vbLf & "���O��ύX���Ă�������" _
						& vbLf & "(Cancel�Ȃ�Ύ��s����܂���)" _
						, "�V���[�g�J�b�g�͑��݂��܂�", lnkName)
					If lnkName<>"" Then
						Set NewShortcut = NewShortcut(lnkPath, lnkName, overwrite)
					End If
					Exit Function
				end if
			end if
			Set NewShortcut = Shell.CreateShortcut(lnkFilePath)
		End Function

		' ���݂̃��[�U�� Quick Launch �t�H���_��Ԃ�
		Function QuickLaunchPath
			' AppData �͉B���w��
			QuickLaunchPath = fso.GetAbsolutePathName( _
				fso.BuildPath( _
					Shell.SpecialFolders("AppData") _
					, "Microsoft\Internet Explorer\Quick Launch" _
				) _
			)
		End Function
