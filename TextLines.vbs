Option Explicit

Class TextLines
	Dim rs
	Private Sub Class_Initialize
		Init
	End Sub
	Private Sub Class_Terminate
		rs.Close
	End Sub

	Private Sub Init
		Const adTypeText = 2
		Set rs = CreateObject("ADODB.Stream")
		rs.type = adTypeText
		rs.Charset = "UTF-8"
		rs.Open
	End Sub


	Public Sub Write(str)
		rs.WriteText str
	End Sub
	Public Sub WriteLine(str)
		rs.WriteText str
		rs.WriteText vbCrLf
	End Sub
	Public Sub Add(str)
		WriteLine str
	End Sub
	Public Sub NewLine
		WriteLine ""
	End Sub

	Public Sub Load(filename)
		rs.loadFromFile filename
	End Sub

	Public Property Get Text
		Dim p
		p = rs.Position
		rs.Position = 0
		Text = rs.ReadText
		rs.Position = p
	End Property

	Public Sub Clear
		rs.Close
		Init
	End Sub
End Class
