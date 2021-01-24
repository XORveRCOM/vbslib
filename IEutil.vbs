Option Explicit

' ------------------------------------------------------------
' 
' ------------------------------------------------------------
Public Sub HtmlDump(ts, element, byval lv)
	If Not element.hasChildNodes() then Exit Sub
	Dim node, idt
	idt = String(lv, vbTab)
	For Each node in element.childNodes
		DispNode ts, node, lv
		If node.hasChildNodes() Then HtmlDump ts, node, lv+1
	Next
End Sub

' --------------------------------------------------------------------------------

Private Sub DispNode(ts, node, byval lv)
	If left(node.nodename,1)="#" Then Exit Sub
	Dim idt, n, i
	idt = String(lv, vbTab)
	n = trim(node.getAttribute("name"))
	i = trim(node.getAttribute("id"))
	If (IsNull(n) or n="") And (IsNull(i) or i="") Then
		ts.WriteLine idt & node.nodename
	ElseIf not(IsNull(n) or n="") And (IsNull(i) or i="") Then
		ts.WriteLine idt & node.nodename & " (name:" & n & ")"
	ElseIf (IsNull(n) or n="") And not (IsNull(i) or i="") Then
		ts.WriteLine idt & node.nodename & " (id:" & i & ")"
	Else
		ts.WriteLine idt & node.nodename & " (name:" & n & ", id:" & i & ")"
	End If
End Sub
' --------------------------------------------------------------------------------
