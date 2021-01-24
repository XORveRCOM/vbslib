Option Explicit

			Dim ScriptInfo
			Set ScriptInfo = New ScriptInfoClass
			ScriptInfo.Init

	Class ScriptInfoClass
			Dim Name, Path
			Dim Dom, fso

		Sub Init
			Set Dom = CreateObject("Msxml2.DOMDocument.5.0")
			Set fso = CreateObject("Scripting.FileSystemObject")
			Name = WScript.ScriptFullName
			Path = fso.GetParentFolderName(Name)
			Dom.Load Name
			With Dom.parseError
				If .errorCode <> 0 Then
						MsgBox "(" & .line & ", " & .linepos & ") : " & .reason, , Name
				End If
			End With
		End Sub

		' パラメータ部の設定値を取得します
		Function GetParam(name, defaultValue)
			Dim elem
			Set elem = Dom.SelectSingleNode("//package/parameter/resource[@id='" & name & "']")
			If IsNull(elem) Then
				GetParam = defaultValue
			Else
				GetParam = elem.Text
			End If
		End Function

		' スクリプトからの相対パスを絶対パスに変換します
		Function RelativeFile(file)
			RelativeFile = fso.BuildPath(Path, file)
		End Function
	End Class
