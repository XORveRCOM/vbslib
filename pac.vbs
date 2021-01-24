' PAC スクリプトを実行するために必要な関数
'   isResolvable() など他の関数が pac スクリプトで使用されていたら適宜追加
'		<!-- PAC -->
'		<script language="VBScript" src="Lib/pac.vbs"/>
'		<script language="JScript" src="http://xxxxxxxx/proxy.pac"/>
Option Explicit

		' シェル表現(Like 演算子)マッチング
		Function shExpMatch(str, shexp)
			Dim pat
			pat = Trim(shexp)
			pat = Replace(pat, vbTab, "")
			pat = Replace(pat, ".", "\.")
			pat = Replace(pat, ":", "\:")
			pat = Replace(pat, "?", ".+")
			pat = Replace(pat, "*", vbTab)
			pat = Replace(pat, vbTab, ".*")
			Dim reg
			Set reg = New RegExp
			reg.Pattern = pat
			shExpMatch = reg.Test(str)
		End Function

		' ホスト名が "単純" (ドメイン名が含まれていない) な名前かどうかを判定します
		Function isPlainHostName(host)
			isPlainHostName = InStr(host, ".") = 0
		End Function

	' ----------------------------------------
	' Proxy の情報を PAC から取得
	' ----------------------------------------
	Class ProxyInfo
			Dim proxy
		Sub Init(site)
			proxy = FindProxyForURL("", site)
		End Sub
		Function IsUsedProxy
			IsUsedProxy = proxy<>"DIRECT"
		End Function
		Function ProxyAddress
			ProxyAddress = Split(proxy, " ")(1)
		End Function
	End Class
