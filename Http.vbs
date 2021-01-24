Option Explicit

	Const HTTP_READSTATE_UNINITIALIZED = 0
	Const HTTP_READSTATE_LOADING = 1
	Const HTTP_READSTATE_LOADED = 2
	Const HTTP_READSTATE_INTERACTIVE = 3
	Const HTTP_READSTATE_COMPLETED = 4

	Class HttpRequest
		Dim http
		Dim user, pass, method

		Dim mime, boundary

		Sub SetLogin(user, pass)
			Me.user = user
			Me.pass = pass
		End Sub

		Sub Init
			If IsEmpty(http) Then
				Set http = CreateObject("Msxml2.ServerXMLHTTP.6.0")
			End If
		End Sub

		Sub InitMime
			If IsEmpty(mime) Then
				Set mime = New BinaryStream
				boundary = "++" & YMDHMS(Now) & "++"
			End If
		End Sub

		Sub Open(method, url)
			Init
			Me.method = method
			If IsEmpty(user) Or user="" Then
				http.Open method, url
			Else
				http.Open method, url, False, user, pass
			End If
		End Sub

		' プロキシ
		Sub SetProxy(proxySetting, varProxyServer, varBypassList)
			Init
			http.setProxy proxySetting, varProxyServer, varBypassList
		End Sub
		Sub setProxyCredentials(username, password)
			Init
			http.SetProxyCredentials username, password
		End Sub

		' リクエストヘッダの追加
		Sub SetRequestHeader(label, value)
			Init
			http.setRequestHeader label, value
		End Sub

		' multipart/form-data
		Sub AddMIMEData(name, value)
			InitMime
			mime.AppendTextAsSJIS "--" & boundary & vbCrLf
			mime.AppendTextAsSJIS "Content-Disposition: form-data; name=""" & name & """" & vbCrLf
			mime.AppendTextAsSJIS vbCrLf
			mime.AppendTextAsSJIS value
			mime.AppendTextAsSJIS vbCrLf
		End Sub
		Sub AddMIMEFile(filename, encoding, bin)
			InitMime
			mime.AppendTextAsSJIS "--" & boundary & vbCrLf
			mime.AppendTextAsSJIS "Content-Disposition: form-data; name=""file""; filename=""" & filename & """" & vbCrLf
			mime.AppendTextAsSJIS "Content-Type: " & ContentType(filename) & vbCrLf
			If encoding="base64" Then
				mime.AppendTextAsSJIS "Content-Transfer-Encoding: base64" & vbCrLf
				mime.AppendTextAsSJIS vbCrLf
				mime.AppendTextAsSJIS Bytes.ByteArrayToBase64(bin)
				mime.AppendTextAsSJIS vbCrLf
			Else
				mime.AppendTextAsSJIS vbCrLf
				mime.AppendByteArray bin
				mime.AppendTextAsSJIS vbCrLf
			End If
		End Sub
		Function ContentType(filename)
			ContentType = "application/octet-stream"
			On Error Resume Next
			ContentType = shell.RegRead("HKCR\." & fso.GetExtensionName(filename) & "\Content Type")
		End Function

		' 送信
		Sub Send(content)
			Init
			If IsEmpty(mime) Then
				http.Send content
			Else
				mime.AppendTextAsSJIS "--" & boundary & "--" & vbCrLf
				mime.AppendTextAsSJIS vbCrLf
				http.SetRequestHeader "content-type", "multipart/form-data; boundary=" & boundary
				http.SetRequestHeader "content-length", mime.Size
				http.Send mime.ByteArray
			End IF
		End Sub
		Function IsRequestCompleted
			Init
			IsRequestCompleted = http.readyState = HTTP_READSTATE_COMPLETED
		End Function

		Sub SendWait(content)
			Send content
			Do Until IsRequestCompleted
				http.waitForResponse 100
			Loop
		End Sub

		' リクエストステータス
		Function RequestStatus
			Init
			RequestStatus = http.status
		End Function
		' レスポンス
		Function ResponseText
			Init
			ResponseText = http.responseText
		End Function
		Function getAllResponseHeaders
			Init
			getAllResponseHeaders = http.getAllResponseHeaders
		End Function
	End Class
