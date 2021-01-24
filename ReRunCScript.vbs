Option Explicit

ReRunCScript

Public Sub ReRunCScript()
	Dim args, i, str
	Set args = WScript.Arguments
	str = ""
	For i = 0 to args.Count - 1
		str = str & " " & args.Item(i)
	Next
'	WScript.Echo str
'	WScript.Echo WScript.FullName
'	WScript.Echo WScript.ScriptFullName
	Dim Fs, WshShell
	Set Fs = WScript.CreateObject("Scripting.FileSystemObject")
	Set WshShell = WScript.CreateObject("WScript.Shell")
	If LCase(Fs.GetFileName(WScript.FullName)) = "wscript.exe" Then
	    WshShell.Run "cmd /k cscript """ & WScript.ScriptFullName & """" & str,1,False
	    WScript.Quit
	End If
End Sub
