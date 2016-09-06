' _0Main: The main module for the statistics macros
'   by imacat <imacat@mail.imacat.idv.tw>, 2016-08-10

Option Explicit

' subMain: The main program
Sub subMain
	BasicLibraries.loadLibrary "XrayTool"
	Dim dStart As Date
	dStart = Now
	
	'MsgBox InStr (1, "abca", "ad")
	'Xray ThisComponent.getSheets.getByIndex (0).getCellByPosition (0, 0)
	'subTestCorrelation
	subTestChi2GoodnessOfFit
	
	MsgBox "Done.  " & Format (Now - dStart, "mm:ss") & " elapsed."
End Sub

' fnQueryFormat: Returns the index of the number format, and creates the number format if required.
Function fnQueryFormat (oDoc As Object, sFormat As String) As Integer
	Dim oFormats As Object, nIndex As Integer
	Dim aLocale As New com.sun.star.lang.Locale
	
	oFormats = oDoc.getNumberFormats
	nIndex = oFormats.queryKey (sFormat, aLocale, True)
	If nIndex = -1 Then
		oFormats.addNew (sFormat, aLocale)
		nIndex = oFormats.queryKey (sFormat, aLocale, True)
	End If
	fnQueryFormat = nIndex
End Function

' fnGetRangeName: Obtains the name of a spreadsheet cell range
Function fnGetRangeName (oRange As Object) As String
	Dim nPos As Integer, sName As String
	
	sName = oRange.getPropertyValue ("AbsoluteName")
	nPos = InStr (sName, "$")
	Do While nPos <> 0
		sName = Left (sName, nPos - 1) & Right (sName, Len (sName) - nPos)
		nPos = InStr (sName, "$")
	Loop
	fnGetRangeName = sName
End Function

' fnGetLocalRangeName: Obtains the name of a local spreadsheet cell range
Function fnGetLocalRangeName (oRange As Object) As String
	Dim nPos As Integer, sName As String
	
	sName = fnGetRangeName (oRange)
	nPos = InStr (sName, ".")
	If nPos <> 0 Then
		sName = Right (sName, Len (sName) - nPos)
	End If
	fnGetLocalRangeName = sName
End Function

' fnFindStatsTestDocument: Finds the statistics test document.
Function fnFindStatsTestDocument As Object
	Dim oEnum As Object, oDoc As Object
	
	oEnum = StarDesktop.getComponents.createEnumeration
	Do While oEnum.hasMoreElements
		oDoc = oEnum.nextElement
		If oDoc.supportsService ("com.sun.star.document.OfficeDocument") Then
			If Right (oDoc.getLocation, Len ("/statstest.ods")) = "/statstest.ods" Then
				fnFindStatsTestDocument = oDoc
				Exit Function
			End If
		End If
	Loop
	fnFindStatsTestDocument = Null
End Function
