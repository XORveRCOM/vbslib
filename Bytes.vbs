Option Explicit

			Dim Bytes
			Set Bytes = New BytesUtil
			Bytes.Init

	Class BytesUtil
			Dim xmldom, node
		Sub Init
			If IsEmpty(xmldom) Then
				Set xmldom = CreateObject("Microsoft.XMLDOM")
				Set node = xmldom.CreateElement("work")
			End If
		End Sub

		' --------------------
		' ファイル
		' --------------------
		Function FileToByteArray(name)
			Const adTypeBinary = 1
			Const adTypeText = 2
			With CreateObject("ADODB.Stream")
				.Open
				.Type = adTypeBinary
				.LoadFromFile name
				FileToByteArray = .Read
				.Close
			End With
		End Function

		' --------------------
		' 文字列を UTF8 のバイト配列に変換
		' --------------------
		Function StringToUTF8ByteArray(str)
			Const adTypeBinary = 1
			Const adTypeText = 2
			With CreateObject("ADODB.Stream")
				.Open
				.Type = adTypeText
				.Charset = "utf-8"
				.WriteText str
				.Position = 0
				.Type = adTypeBinary
				.Position = 3	' BOM をスキップ
				StringToUTF8ByteArray = .Read
				.Close
			End With
		End Function

		' --------------------
		' HEXTEXT
		' --------------------
		Function ByteArrayToHexText(bin)
			node.Text = ""
			node.DataType = "bin.hex"
			node.NodeTypedValue = bin
			ByteArrayToHexText = node.Text
		End Function

		Function HexTextToByteArray(HexText)
			node.Text = ""
			node.DataType = "bin.hex"
			node.Text = HexText
			HexTextToByteArray = node.NodeTypedValue
		End Function

		' --------------------
		' BASE64
		' --------------------
		Function ByteArrayToBase64(bin)
			node.Text = ""
			node.DataType = "bin.base64"
			node.NodeTypedValue = bin
			ByteArrayToBase64 = Replace(node.Text, vbLf, vbCrLf)
		End Function

		Function Base64ToByteArray(base64)
			node.Text = ""
			node.DataType = "bin.base64"
			node.Text = base64
			Base64ToByteArray = node.NodeTypedValue
		End Function

		Function EncodeBase64(bin)
			EncodeBase64 = ByteArrayToBase64(bin)
		End Function

		Function DecodeBase64(base64)
			DecodeBase64 = Base64ToByteArray(base64)
		End Function

		' --------------------
		' ISO-2022-JP
		' --------------------
		Function EncodeMIMEHeader(str)
			Dim enc, substr, i, ch
			enc = ""
			substr = ""
			For i=1 To Len(str)
				ch = Mid(str, i, 1)
				If AscW(ch)>255 Then
					substr = substr & ch
				Else
					If substr<>"" Then
						enc = enc & EncodePartISO2022JP(substr)
						substr = ""
					End If
					enc = enc & ch
				End If
			Next
			If substr<>"" Then
				enc = enc & EncodePartISO2022JP(substr)
				substr = ""
			End If
			EncodeMIMEHeader = enc
		End Function
		' ISO-2022-JP
		Function EncodePartISO2022JP(str)
			Dim enc
			enc = ""
			enc = enc & "=?ISO-2022-JP?B?"
			With New BinaryStream
				.AppendTextAsISO2022JP str
				enc = enc & ByteArrayToBase64(.ByteArray)
				.Close
			End With
			enc = enc & "?="
			EncodePartISO2022JP = enc
		End Function
	End Class

	Class BinaryStream
		Dim st
		Sub Init
			Const adTypeBinary = 1
			If IsEmpty(st) Then
				Set st = CreateObject("ADODB.Stream")
				st.Open
				st.Type = adTypeBinary
			End If
		End Sub
		Sub Clear
			Init
			st.Size = 0
		End Sub
		Sub Close
			Init
			st.Close
			st = Empty
		End Sub

		' --------------------
		' バイナリ配列
		' --------------------
		Function ByteArray
			Const adTypeBinary = 1
			Init
			st.Position = 0
			st.Type = adTypeBinary
			ByteArray = st.Read
			st.Position = st.Size
		End Function

		' --------------------
		' サイズ
		' --------------------
		Function Size
			Size = st.Size
		End Function

		' --------------------
		' テキスト追加
		' --------------------
		Sub AppendTextAsSJIS(str)
			AppendText str, "shift_jis"
		End Sub
		Sub AppendTextAsAscii(str)
			AppendText str, "ascii"
		End Sub
		Sub AppendTextAsUTF8(str)
			AppendText str, "utf-8"
		End Sub
		Sub AppendTextAsISO2022JP(str)
			AppendText str, "iso-2022-jp"
		End Sub
		Sub AppendText(str, charset)
			Const adTypeBinary = 1
			Const adTypeText = 2
			Init
			If st.Type = adTypeBinary Then
				st.Position = 0
				st.Type = adTypeText
				st.Charset = charset
				st.Position = st.Size
			End If
			st.WriteText str
		End Sub
		' --------------------
		' バイナリ追加
		' --------------------
		Sub AppendByteArray(bin)
			Const adTypeBinary = 1
			Const adTypeText = 2
			Init
			If st.Type = adTypeText Then
				st.Position = 0
				st.Type = adTypeBinary
				st.Position = st.Size
			End If
			st.Write bin
		End Sub
	End Class
