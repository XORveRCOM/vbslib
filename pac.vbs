' PAC �X�N���v�g�����s���邽�߂ɕK�v�Ȋ֐�
'   isResolvable() �ȂǑ��̊֐��� pac �X�N���v�g�Ŏg�p����Ă�����K�X�ǉ�
'		<!-- PAC -->
'		<script language="VBScript" src="Lib/pac.vbs"/>
'		<script language="JScript" src="http://xxxxxxxx/proxy.pac"/>
Option Explicit

		' �V�F���\��(Like ���Z�q)�}�b�`���O
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

		' �z�X�g���� "�P��" (�h���C�������܂܂�Ă��Ȃ�) �Ȗ��O���ǂ����𔻒肵�܂�
		Function isPlainHostName(host)
			isPlainHostName = InStr(host, ".") = 0
		End Function

	' ----------------------------------------
	' Proxy �̏��� PAC ����擾
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
