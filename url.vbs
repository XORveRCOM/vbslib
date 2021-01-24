Option Explicit

Sub UrlTest
	Dim u
	Set u = New Url
	u.Init "https://scnet.cybozu.com/o/ag.cgi?page=BulletinView&bid=19828&gid=0&cid=144&cp=blc&tp=t#"

	WScript.Echo u.Scheme
	WScript.Echo u.Authority
	WScript.Echo u.Path
	WScript.Echo u.Query
	WScript.Echo u.Fragment
End Sub

' RFC 3986 ベースの URL 解析を行います
Class Url
	Dim regEx

	Private Sub Class_Initialize
		Set regEx = New RegExp
		regEx.Pattern = "^(([^:/?#]+):)?(//([^/?#]*))?([^?#]*)(\?([^#]*))?(#(.*))?"
		regEx.IgnoreCase = True
		regEx.Global = True
	End Sub

	Dim Match

	Public Sub Init(url_str)
		Dim Matches
		Set Matches = regEx.Execute(url_str)
		Set Match = Matches(0)
	End Sub

	Public Property Get Scheme
		Scheme = Match.SubMatches(1)
	End Property

	Public Property Get Authority
		Authority = Match.SubMatches(3)
	End Property

	Public Property Get Path
		Path = Match.SubMatches(4)
	End Property

	Public Property Get Query
		Query = Match.SubMatches(6)
	End Property

	Public Property Get Fragment
		Fragment = Match.SubMatches(8)
	End Property
End Class
