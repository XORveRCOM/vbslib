Option Explicit

			Dim Capicom
			Set Capicom = New CapicomUtil
			Capicom.Init

			Const CAPICOM_HASH_ALGORITHM_SHA1 = 0
			Const CAPICOM_HASH_ALGORITHM_MD2 = 1
			Const CAPICOM_HASH_ALGORITHM_MD4 = 2
			Const CAPICOM_HASH_ALGORITHM_MD5 = 3
			Const CAPICOM_HASH_ALGORITHM_SHA_256 = 4
			Const CAPICOM_HASH_ALGORITHM_SHA_384 = 5
			Const CAPICOM_HASH_ALGORITHM_SHA_512 = 6

	Class CapicomUtil
			Dim capi, hash
		Sub Init
			If IsEmpty(capi) Then
				Set capi = CreateObject("CAPICOM.Utilities")
				Set hash = CreateObject("CAPICOM.HashedData")
			End If
		End Sub

		' --------------------
		' BinaryString
		' --------------------
		Function ByteArrayToBinaryString(bin)
			ByteArrayToBinaryString = capi.ByteArrayToBinaryString(bin)
		End Function

		Function BinaryStringToByteArray(bstr)
			BinaryStringToByteArray = capi.BinaryStringToByteArray(bstr)
		End Function

		' --------------------
		' SHA256
		' --------------------
		Function ByteArrayToSHA256(bin)
			hash.Algorithm = CAPICOM_HASH_ALGORITHM_SHA_256
			Dim bstr
			bstr = capi.ByteArrayToBinaryString(bin)
			hash.Hash bstr
			ByteArrayToSHA256 = UCase(hash.Value)
		End Function

		' --------------------
		' MD5
		' --------------------
		Function ByteArrayToMD5(bin)
			hash.Algorithm = CAPICOM_HASH_ALGORITHM_MD5
			Dim bstr
			bstr = capi.ByteArrayToBinaryString(bin)
			hash.Hash bstr
			ByteArrayToMD5 = UCase(hash.Value)
		End Function
	End Class
