Option Explicit

Function XsltExec(fileXML, fileXSL)
	Dim xslTemp	'As Msxml2.XSLTemplate
	Dim xslDoc	'As Msxml2.FreeThreadedDOMDocument
	Dim xmlDoc	'As Msxml2.DOMDocument
	Dim xslProc	'As IXSLProcessor

	' 変換元のXMLを読み込み
	Set xmlDoc = CreateObject("Msxml2.DOMDocument")
	xmlDoc.async = False
	xmlDoc.Load fileXML

	' XSL 読み込み
	Set xslDoc = CreateObject("Msxml2.FreeThreadedDOMDocument")
	xslDoc.async = False
	xslDoc.Load fileXSL

	' XSLT として適用
	Set xslTemp = CreateObject("Msxml2.XSLTemplate")
	Set xslTemp.stylesheet = xslDoc

	' XSLT プロセッサ
	Set xslProc = xslTemp.createProcessor()

	xslProc.input = xmlDoc
	xslProc.Transform

	XsltExec = xslProc.output
End Function
