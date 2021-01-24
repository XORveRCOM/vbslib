Option Explicit

Function XsltExec(fileXML, fileXSL)
	Dim xslTemp	'As Msxml2.XSLTemplate
	Dim xslDoc	'As Msxml2.FreeThreadedDOMDocument
	Dim xmlDoc	'As Msxml2.DOMDocument
	Dim xslProc	'As IXSLProcessor

	' �ϊ�����XML��ǂݍ���
	Set xmlDoc = CreateObject("Msxml2.DOMDocument")
	xmlDoc.async = False
	xmlDoc.Load fileXML

	' XSL �ǂݍ���
	Set xslDoc = CreateObject("Msxml2.FreeThreadedDOMDocument")
	xslDoc.async = False
	xslDoc.Load fileXSL

	' XSLT �Ƃ��ēK�p
	Set xslTemp = CreateObject("Msxml2.XSLTemplate")
	Set xslTemp.stylesheet = xslDoc

	' XSLT �v���Z�b�T
	Set xslProc = xslTemp.createProcessor()

	xslProc.input = xmlDoc
	xslProc.Transform

	XsltExec = xslProc.output
End Function
