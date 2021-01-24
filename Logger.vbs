Option Explicit

	' ----------------------------------------
	' ÉçÉKÅ[
	Class Logger
		Dim file
		Dim stdout
		Sub Init(path)
			If Not IsEmpty(file) Then
				Close
			End If
			Set file = fso.CreateTextFile(path)
			If LCase(fso.GetBaseName(WScript.FullName))="cscript" Then
				Set stdout = WScript.StdOut
			End If
		End Sub
		' èIóπéûÇ…é©ìÆâï˙
		Private Sub Class_Terminate
			Close
		End Sub

		Sub WriteLine(str)
			If Not IsEmpty(file) Then
				If Not IsEmpty(stdout) Then
					stdout.WriteLine str
				End If
				file.WriteLine str
			End If
		End Sub
		Sub Write(str)
			If Not IsEmpty(file) Then
				If Not IsEmpty(stdout) Then
					stdout.Write str
				End If
				file.Write str
			End If
		End Sub
		Sub WriteBlankLines(count)
			If Not IsEmpty(file) Then
				If Not IsEmpty(stdout) Then
					stdout.WriteBlankLines count
				End If
				file.WriteBlankLines count
			End If
		End Sub
		Sub Close
			If Not IsEmpty(file) Then
				file.Close
				file = Empty
				stdout = Empty
			End If
		End Sub
	End Class
