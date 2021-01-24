Option Explicit
'http://plaza.rakuten.co.jp/issosakura/diary/
' ---------------------------------------------------------------------------

Public Function encodeURI(str)
	With CreateObject("ScriptControl")
		.Language = "JScript"
		encodeURI = .CodeObject.encodeURI(str)
	End With
End Function

Public Function encodeURIComponent(str)
	With CreateObject("ScriptControl")
		.Language = "JScript"
		encodeURIComponent = .CodeObject.encodeURIComponent(str)
	End With
End Function

class IEwindow
	Dim exist
	Public IEobject
	Dim VisMode

	' --------------------------------------------------------------------------------
	' 初期化
	' --------------------------------------------------------------------------------
	Public Sub Init
		If IsEmpty(IEobject) Then
			Set IEobject = CreateObject("InternetExplorer.Application")
			IEobject.Visible = False
'			WScript.Sleep 2000	' [ms]
		ElseIf IEobject Is Nothing Then
			Set IEobject = CreateObject("InternetExplorer.Application")
			IEobject.Visible = False
'			WScript.Sleep 2000	' [ms]
		End If
		IEobject.Visible = Visible
	End Sub

	' --------------------------------------------------------------------------------
	' 閉じる
	' --------------------------------------------------------------------------------
	Public Sub Close
		If IsEmpty(IEobject) Then
		ElseIf IEobject Is Nothing Then
		Else
			IEobject.Quit
			Set IEobject = Nothing
		End If
	End Sub

	' --------------------------------------------------------------------------------
	' OpenしたURLは、既存のIEの再利用だった
	' --------------------------------------------------------------------------------
	Public Property Get isAlreadyExists
		isAlreadyExists = exist
	End Property

	' --------------------------------------------------------------------------------
	' 可視性
	' --------------------------------------------------------------------------------
	Public property Get Visible
'		Init
'		Visible = IEobject.Visible
		If IsEmpty(VisMode) Then VisMode = True
		Visible = VisMode
	End Property
	Public property Let Visible(value)
		VisMode = value
		Init
	End Property

	' --------------------------------------------------------------------------------
	' 対象URLに移動する
	' --------------------------------------------------------------------------------
	Public Function GoToURL(turl)
		Init
		IEobject.Navigate turl
		Do While (IEobject.busy Or IEobject.readyState <> 4)
			WScript.Sleep 500	' [ms]
		Loop
		GoToURL = IEobject.LocationURL
	End Function
	Public Function Navigate(turl, timeout)
		Init
		Dim toTime
		toTime = DateAdd("s", timeout, Now)
		'対象URLが非表示→新しく画面を開く
		IEobject.Navigate turl
		'待機
		If Not Wait(turl, timeout) Then
			Navigate = False
			Exit Function
		End If
		'表示URLチェック
		If IEobject.LocationURL<>turl Then
			MsgBox IEobject.LocationURL & vbLf & "<> " & turl, , "IEwindow.Navigate"
		End If
		Navigate = True
	End Function

	Public Function Search(turl)
		'対象画面を検索
		Dim Shell, Window
		Set Shell = CreateObject("Shell.Application")
		For Each Window In Shell.Windows
			Dim str
			' InternetExplorerObject を探す
			On Error Resume Next
			str = TypeName(Window.Document)
			On Error Goto 0
			If Err.Number<>0 Then
				MsgBox Err.Description
				Err.Clear
			End If
			If str = "HTMLDocument" Then
				' 指定のURLを探す
				On Error Resume Next
				str = ""
				str = Window.Document.url
				On Error Goto 0
				If Err.Number<>0 Then
					MsgBox Err.Description, , "IEwindow.Open"
					Err.Clear
				End If
				' 検索指定のurlと比較してみる
				if str=turl then
					Dim vis
					vis = Visible
					Close					' 以前に管理していた画面を閉じる
					Set IEobject = Window	' 対象URLが表示→その画面を使う
					IEobject.Visible = vis
					exist = true
					Search = True
					Exit Function
				end if
			end if
		next
		Set Shell = Nothing
		Search = false
	End Function

	' --------------------------------------------------------------------------------
	' 対象URLを開く
	' --------------------------------------------------------------------------------
	Public Function Open(turl, timeout)
		'対象画面を検索
		Open = Search(turl)
		' 無かったので開く
		If Not Open Then Open = Navigate(turl, timeout)
	End Function

	' --------------------------------------------------------------------------------
	' 処理が終わるのを待つ
	' --------------------------------------------------------------------------------
	Public Function Wait(turl, timeout)
		Init
		Dim toTime
		toTime = DateAdd("s", timeout, Now)
'		Do While IEobject.document.readyState <> "complete"
		Do While (IEobject.busy Or IEobject.readyState <> 4 Or (turl<>"" And IEobject.LocationURL<>turl))
			If timeout>0 And toTime<Now Then
				If turl="" Then
					MSgBox "Timeout" & vbLf & toTime & vbLf & Now, , "IEwindow.Wait"
				Else
					Dim place
					place = "期待:" & turl & vbLf & "現在:" & IEobject.LocationURL
					MSgBox "Timeout" & vbLf & toTime & vbLf & Now & vbLf & place, , "IEwindow.Wait"
				End If
				Close
				Wait = False
				Exit Function
			End If
			WScript.Sleep 500	' [ms]
		Loop
		Wait = True
	End Function

	' --------------------------------------------------------------------------------
	' Config 定義に従って Form の Submit を行う
	' --------------------------------------------------------------------------------
	Public Sub Submit(conf, secname)
			If Not conf.ContainsSection(secname) Then
				MsgBox "セクション " & secname & " が設定ファイルにありません。"
				Exit Sub
			End If

			Dim account, password, account_name, password_name, submit_name, submit_type, baseURL, nextURL
			account = conf.Value(secname, "user", "")
			password = conf.Value(secname, "password", "")
			account_name = conf.Value(secname, "user_name", "")
			password_name = conf.Value(secname, "password_name", "")
			submit_name = conf.Value(secname, "submit_name", "")
			submit_type = conf.Value(secname, "submit_type", "submit")
			baseURL = conf.Value(secname, "baseURL", "")
			nextURL = conf.Value(secname, "nextURL", "")

			Const maxWaitSec = 60
' --------------------
			' ログイン
			If Not Open(baseURL, maxWaitSec) Then Exit Sub
			With IEobject.document.all
				Err.Clear
				On Error Resume Next
				.namedItem(account_name).value = account
				If Err.Number<>0 Then WScript.Quit
				On Error Goto 0
				.namedItem(password_name).value = password
				If submit_type="click" Then
					.namedItem(submit_name).Click
				Else
					.namedItem(submit_name).Submit
				End If
			End With

			' ログイン処理完了待ち
			if nextURL<>"" THen
				If Not Wait("", maxWaitSec) Then Exit Sub
				Navigate nextURL, maxWaitSec
			End If
	End Sub
End Class
