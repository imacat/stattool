' _3ITTest: The macros to for generating the report of independent T-Test
'   by imacat <imacat@mail.imacat.idv.tw>, 2016-08-24
Option Explicit

' subTestIndependentTTest: Tests the independent T-test report
Sub subTestIndependentTTest
	Dim oSheets As Object, sSheetName As String
	Dim oSheet As Object, oRange As Object
	
	sSheetName = "ittest"
	oSheets = ThisComponent.getSheets
	If Not oSheets.hasByName (sSheetName) Then
		MsgBox "Data sheet """ & sSheetName & """ not found"
		Exit Sub
	End If
	If oSheets.hasByName (sSheetName & "_ttest") Then
		oSheets.removeByName (sSheetName & "_ttest")
	End If
	If oSheets.hasByName (sSheetName & "_ttesttmp") Then
		oSheets.removeByName (sSheetName & "_ttesttmp")
	End If
	oSheet = ThisComponent.getSheets.getByName (sSheetName)
	oRange = oSheet.getCellRangeByName ("A15:B34")
	subReportIndependentTTest (ThisComponent, oRange)
End Sub

' subReportIndependentTTest: Reports the independent T-test
Sub subReportIndependentTTest (oDoc As Object, oDataRange As Object)
	Dim oSheets As Object, sSheetName As String
	Dim mNames () As String, nI As Integer, nSheetIndex As Integer
	Dim oSheet As Object, oColumns As Object, nRow As Integer
	Dim oCell As Object, oCells As Object, oCursor As Object, oTempDataRange As Object
	Dim nN As Integer, sFormula As String, sSP2 As String
	Dim sNotes As String, nPos As Integer
	Dim nFormatN As Integer, nFormatF As Integer, nFormatP As Integer
	Dim aBorderSingle As New com.sun.star.table.BorderLine
	Dim aBorderDouble As New com.sun.star.table.BorderLine
	Dim sCellXLabel As String, sCellsXData As String
	Dim sCellXN As String, sCellXMean As String, sCellXS As String
	Dim sCellYLabel As String, sCellsYData As String
	Dim sCellYN As String, sCellYMean As String, sCellYS As String
	Dim sCellF As String, sCellsN As String, sCellN As String
	
	oSheets = oDoc.getSheets
	sSheetName = oDataRange.getSpreadsheet.getName
	mNames = oSheets.getElementNames
	For nI = 0 To UBound (mNames)
		If mNames (nI) = sSheetName Then
			nSheetIndex = nI
		End If
	Next nI
	
	oSheets.insertNewByName (sSheetName & "_ttesttmp", nSheetIndex + 1)
	oSheet = oSheets.getByName (sSheetName & "_ttesttmp")
	oTempDataRange = fnCollectIndependentTTestData (oDataRange, oSheet)
	
	oSheets.insertNewByName (sSheetName & "_ttest", nSheetIndex + 1)
	oSheet = oSheets.getByName (sSheetName & "_ttest")
	
	sCellXLabel = fnGetRangeName (oTempDataRange.getCellByPosition (0, 0))
	nN = oTempDataRange.getCellByPosition (0, oTempDataRange.getRows.getCount - 3).getValue
	oCells = oTempDataRange.getCellRangeByPosition (0, 1, 0, nN)
	sCellsXData = fnGetRangeName (oCells)
	sCellYLabel = fnGetRangeName (oTempDataRange.getCellByPosition (1, 0))
	nN = oTempDataRange.getCellByPosition (1, oTempDataRange.getRows.getCount - 3).getValue
	oCells = oTempDataRange.getCellRangeByPosition (1, 1, 1, nN)
	sCellsYData = fnGetRangeName (oCells)
	
	' Obtains the format parameters for the report.
	nFormatN = fnQueryFormat (oDoc, "#,##0")
	nFormatF = fnQueryFormat (oDoc, "#,###.000")
	nFormatP = fnQueryFormat (oDoc, "[<0.01]#.000""**"";[<0.05]#.000""*"";#.000")
	
	aBorderSingle.OuterLineWidth = 2
	aBorderDouble.OuterLineWidth = 2
	aBorderDouble.InnerLineWidth = 2
	aBorderDouble.LineDistance = 2
	
	' Sets the column widths of the report.
	oColumns = oSheet.getColumns
	oColumns.getByIndex (0).setPropertyValue ("Width", 3060)
	oColumns.getByIndex (1).setPropertyValue ("Width", 2080)
	oColumns.getByIndex (2).setPropertyValue ("Width", 2080)
	oColumns.getByIndex (3).setPropertyValue ("Width", 2080)
	oColumns.getByIndex (4).setPropertyValue ("Width", 2080)
	
	nRow = -2
	
	' Group description
	nRow = nRow + 2
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Group Description")
	oCell.setPropertyValue ("CellStyle", "Result2")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 4, nRow)
	oCells.merge (True)
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Group")
	oCell = oSheet.getCellByPosition (1, nRow)
	oCell.setString ("N")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCell = oSheet.getCellByPosition (2, nRow)
	oCell.setString ("X")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharOverline", com.sun.star.awt.FontUnderline.SINGLE)
	oCell = oSheet.getCellByPosition (3, nRow)
	oCell.setString ("s")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCell = oSheet.getCellByPosition (4, nRow)
	oCell.setString ("sX")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.collapseToEnd
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharEscapement", -33)
	oCursor.setPropertyValue ("CharEscapementHeight", 58)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharOverline", com.sun.star.awt.FontUnderline.SINGLE)
	
	' The first group
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	sFormula = "=" & sCellXLabel
	oCell.setFormula (sFormula)
	oCell = oSheet.getCellByPosition (1, nRow)
	sFormula = "=COUNT(" & sCellsXData & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatN)
	sCellXN = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (2, nRow)
	sFormula = "=AVERAGE(" & sCellsXData & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellXMean = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (3, nRow)
	sFormula = "=STDEV(" & sCellsXData & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellXS = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (4, nRow)
	sFormula = "=" & sCellXS & "/SQRT(" & sCellXN & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	
	' The second group
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	sFormula = "=" & sCellYLabel
	oCell.setFormula (sFormula)
	oCell = oSheet.getCellByPosition (1, nRow)
	sFormula = "=COUNT(" & sCellsYData & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatN)
	sCellYN = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (2, nRow)
	sFormula = "=AVERAGE(" & sCellsYData & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellYMean = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (3, nRow)
	sFormula = "=STDEV(" & sCellsYData & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellYS = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (4, nRow)
	sFormula = "=" & sCellYS & "/SQRT(" & sCellYN & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	
	' Draws the table borders.
	oCells = oSheet.getCellByPosition (0, nRow - 2)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellRangeByPosition (1, nRow - 2, 4, nRow - 2)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellByPosition (0, nRow - 1)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells = oSheet.getCellByPosition (0, nRow)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	oCells = oSheet.getCellRangeByPosition (1, nRow, 4, nRow)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	
	' Levene's test for homogeneity of variances
	nRow = nRow + 2
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Test for Homogeneity of Variances")
	oCell.setPropertyValue ("CellStyle", "Result2")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 4, nRow)
	oCells.merge (True)
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Test")
	oCell = oSheet.getCellByPosition (1, nRow)
	oCell.setString ("F")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCell = oSheet.getCellByPosition (2, nRow)
	oCell.setString ("p")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	
	' The test result.
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Levene’s Test")
	oCell = oSheet.getCellByPosition (1, nRow)
	sFormula = "=" & fnGetLeveneTest (oTempDataRange)
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellF = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (2, nRow)
	sCellsN = fnGetRangeName (oTempDataRange.getCellRangeByPosition (0, oTempDataRange.getRows.getCount - 3, 1, oTempDataRange.getRows.getCount - 3))
	sCellN = fnGetRangeName (oTempDataRange.getCellByPosition (4, oTempDataRange.getRows.getCount - 3))
	sFormula = "=FDIST(" & sCellF & ";COUNT(" & sCellsN & ")-1;" & sCellN & "-COUNT(" & sCellsN & "))"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatP)
	
	' The foot notes of the test.
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Note: *: p<.05, **: p<.01" & Chr (10) & _
		"H0: σ1=σ2 (homogeneity; the populations of the two groups have the same variances)." & Chr (10) & _
		"H1: σ1≠σ2 (heterogeneity; the populations of the two groups have different variances) if the probability (p) is small enough.")
	oCell.setPropertyValue ("IsTextWrapped", True)
	oCells = oSheet.getCellRangeByPosition (0, nRow, 4, nRow)
	oCells.merge (True)
	sNotes = oCell.getString
	oCursor = oCell.createTextCursor
	nPos = InStr (sNotes, "p<")
	Do While nPos <> 0
		oCursor.gotoStart (False)
		oCursor.goRight (nPos - 1, False)
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
		nPos = InStr (nPos + 1, sNotes, "p<")
	Loop
	nPos = InStr (sNotes, "(p)")
	oCursor.gotoStart (False)
	oCursor.goRight (nPos, False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	nPos = InStr (sNotes, "σ")
	Do While nPos <> 0
		oCursor.gotoStart (False)
		oCursor.goRight (nPos - 1, False)
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.collapseToEnd
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharEscapement", -33)
		oCursor.setPropertyValue ("CharEscapementHeight", 58)
		nPos = InStr (nPos + 1, sNotes, "σ")
	Loop
	nPos = InStr (sNotes, "H0")
	oCursor.gotoStart (False)
	oCursor.goRight (nPos - 1, False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.collapseToEnd
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharEscapement", -33)
	oCursor.setPropertyValue ("CharEscapementHeight", 58)
	nPos = InStr (sNotes, "H1")
	oCursor.gotoStart (False)
	oCursor.goRight (nPos - 1, False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.collapseToEnd
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharEscapement", -33)
	oCursor.setPropertyValue ("CharEscapementHeight", 58)
	
	' Draws the table borders.
	oCells = oSheet.getCellByPosition (0, nRow - 2)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellRangeByPosition (1, nRow - 2, 2, nRow - 2)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellByPosition (0, nRow - 1)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	oCells = oSheet.getCellRangeByPosition (1, nRow - 1, 2, nRow - 1)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	
	' The independent samples T-test
	nRow = nRow + 2
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Independent Samples T-Test")
	oCell.setPropertyValue ("CellStyle", "Result2")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 4, nRow)
	oCells.merge (True)
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Type")
	oCell = oSheet.getCellByPosition (1, nRow)
	oCell.setString ("t")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCell = oSheet.getCellByPosition (2, nRow)
	oCell.setString ("df")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCell = oSheet.getCellByPosition (3, nRow)
	oCell.setString ("p")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCell = oSheet.getCellByPosition (4, nRow)
	oCell.setString ("X1-X2")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharOverline", com.sun.star.awt.FontUnderline.SINGLE)
	oCursor.collapseToEnd
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharEscapement", -33)
	oCursor.setPropertyValue ("CharEscapementHeight", 58)
	oCursor.collapseToEnd
	oCursor.goRight (1, False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharOverline", com.sun.star.awt.FontUnderline.SINGLE)
	oCursor.collapseToEnd
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharEscapement", -33)
	oCursor.setPropertyValue ("CharEscapementHeight", 58)
	
	' The test of the homogeneity of variances.
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Homogeneity")
	oCell = oSheet.getCellByPosition (1, nRow)
	sSP2 = "((SUMPRODUCT(" & sCellsXData & ";" & sCellsXData & ")-POWER(SUM(" & sCellsXData & ");2)/" & sCellXN & "+SUMPRODUCT(" & sCellsYData & ";" & sCellsYData & ")-POWER(SUM(" & sCellsYData & ");2)/" & sCellYN & ")/(" & sCellXN & "+" & sCellYN & "-2))"
	sFormula = "=(" & sCellXMean & "-" & sCellYMean & ")/SQRT(" & sSP2 & "*(1/" & sCellXN & "+1/" & sCellYN & "))"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	oCell = oSheet.getCellByPosition (2, nRow)
	sFormula = "=" &  sCellXN & "+" & sCellYN & "-2"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatN)
	oCell = oSheet.getCellByPosition (3, nRow)
	sFormula = "=TTEST(" &  sCellsXData & ";" & sCellsYData & ";2;2)"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatP)
	oCell = oSheet.getCellByPosition (4, nRow)
	sFormula = "=" &  sCellXMean & "-" & sCellYMean
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	
	' The test of the heterogeneity of variances.
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Heterogeneity")
	oCell = oSheet.getCellByPosition (1, nRow)
	sFormula = "=(" & sCellXMean & "-" & sCellYMean & ")/SQRT(POWER(" & sCellXS & ";2)/" & sCellXN & "+POWER(" & sCellYS & ";2)/" & sCellYN & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	oCell = oSheet.getCellByPosition (2, nRow)
	sFormula = "=POWER(POWER(" & sCellXS & ";2)/" & sCellXN & "+POWER(" & sCellYS & ";2)/" & sCellYN & ";2)/(POWER(" & sCellXS & ";4)/(POWER(" & sCellXN & ";2)*(" & sCellXN & "-1))+POWER(" & sCellYS & ";4)/(POWER(" & sCellYN & ";2)*(" & sCellYN & "-1)))"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	oCell = oSheet.getCellByPosition (3, nRow)
	sFormula = "=TTEST(" &  sCellsXData & ";" & sCellsYData & ";2;3)"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatP)
	oCell = oSheet.getCellByPosition (4, nRow)
	sFormula = "=" &  sCellXMean & "-" & sCellYMean
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	
	' The foot notes of the test.
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Note: *: p<.05, **: p<.01" & Chr (10) & _
		"H0: μ1=μ2 (the populations of the two groups have the same means)." & Chr (10) & _
		"H1: μ1≠μ2 (the populations of the two groups have different means) if the probability (p) is small enough.")
	oCell.setPropertyValue ("IsTextWrapped", True)
	oCells = oSheet.getCellRangeByPosition (0, nRow, 4, nRow)
	oCells.merge (True)
	sNotes = oCell.getString
	oCursor = oCell.createTextCursor
	nPos = InStr (sNotes, "p<")
	Do While nPos <> 0
		oCursor.gotoStart (False)
		oCursor.goRight (nPos - 1, False)
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
		nPos = InStr (nPos + 1, sNotes, "p<")
	Loop
	nPos = InStr (sNotes, "(p)")
	oCursor.gotoStart (False)
	oCursor.goRight (nPos, False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	nPos = InStr (sNotes, "μ")
	Do While nPos <> 0
		oCursor.gotoStart (False)
		oCursor.goRight (nPos - 1, False)
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.collapseToEnd
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharEscapement", -33)
		oCursor.setPropertyValue ("CharEscapementHeight", 58)
		nPos = InStr (nPos + 1, sNotes, "μ")
	Loop
	nPos = InStr (sNotes, "H0")
	oCursor.gotoStart (False)
	oCursor.goRight (nPos - 1, False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.collapseToEnd
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharEscapement", -33)
	oCursor.setPropertyValue ("CharEscapementHeight", 58)
	nPos = InStr (sNotes, "H1")
	oCursor.gotoStart (False)
	oCursor.goRight (nPos - 1, False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.collapseToEnd
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharEscapement", -33)
	oCursor.setPropertyValue ("CharEscapementHeight", 58)
	
	' Draws the table borders.
	oCells = oSheet.getCellByPosition (0, nRow - 3)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellRangeByPosition (1, nRow - 3, 4, nRow - 3)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellByPosition (0, nRow - 2)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells = oSheet.getCellByPosition (0, nRow - 1)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	oCells = oSheet.getCellRangeByPosition (1, nRow - 1, 4, nRow - 1)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
End Sub

' fnCollectIndependentTTestData: Collects the data for the independent T-test.
Function fnCollectIndependentTTestData (oDataRange As Object, oReportSheet As Object) As Object
	Dim nRow As Integer, nNRow As Integer, sCellZMean As String, sCellsN As String
	Dim oCell As Object, oCells As Object, oCursor As Object
	Dim sCell As String, sLabel As String, sFormula As String
	Dim sCellXLabel As String, sCellsXData As String, sCellXMean As String
	Dim sXLabel As String, nNX As Integer
	Dim sCellsXZData As String, sCellXZMean As String
	Dim sCellYLabel As String, sCellsYData As String, sCellYMean As String
	Dim sYLabel As String, nNY As Integer
	Dim sCellsYZData As String, sCellYZMean As String
	
	sCellXLabel = ""
	sCellYLabel = ""
	For nRow = 1 To oDataRange.getRows.getCount - 1
		oCell = oDataRange.getCellByPosition (0, nRow)
		sLabel = oCell.getString
		If sLabel <> "" Then
			If sCellXLabel = "" Then
				sCellXLabel = fnGetRangeName (oCell)
				sXLabel = sLabel
			Else
				If sLabel <> sXLabel And sCellYLabel = "" Then
					sCellYLabel = fnGetRangeName (oCell)
					sYLabel = sLabel
					nRow = oDataRange.getRows.getCount - 1
				End If
			End If
		End If
	Next nRow
	
	' The data labels
	oCell = oReportSheet.getCellByPosition (0, 0)
	sFormula = "=" & sCellXLabel
	oCell.setFormula (sFormula)
	oCell = oReportSheet.getCellByPosition (1, 0)
	sFormula = "=" & sCellYLabel
	oCell.setFormula (sFormula)
	
	' The data
	nNX = 0
	nNY = 0
	For nRow = 1 To oDataRange.getRows.getCount - 1
		If oDataRange.getCellByPosition (0, nRow).getString = sXLabel Then
			nNX = nNX + 1
			sFormula = "=" & fnGetRangeName (oDataRange.getCellByPosition (1, nRow))
			oReportSheet.getCellByPosition (0, nNX).setFormula (sFormula)
		Else
			If oDataRange.getCellByPosition (0, nRow).getString = sYLabel Then
				nNY = nNY + 1
				sFormula = "=" & fnGetRangeName (oDataRange.getCellByPosition (1, nRow))
				oReportSheet.getCellByPosition (1, nNY).setFormula (sFormula)
			End If
		End If
	Next nRow
	
	' Collects the data
	sCellsXData = fnGetLocalRangeName (oReportSheet.getCellRangeByPosition (0, 1, 0, nNX))
	sCellsYData = fnGetLocalRangeName (oReportSheet.getCellRangeByPosition (1, 1, 1, nNY))
	If nNX > nNY Then
		nNRow = nNX + 1
	Else
		nNRow = nNY + 1
	End If
	oCell = oReportSheet.getCellByPosition (0, nNRow)
	sFormula = "=COUNT(" & sCellsXData & ")"
	oCell.setFormula (sFormula)
	oCell = oReportSheet.getCellByPosition (1, nNRow)
	sFormula = "=COUNT(" & sCellsYData & ")"
	oCell.setFormula (sFormula)
	oCell = oReportSheet.getCellByPosition (0, nNRow + 1)
	sFormula = "=AVERAGE(" & sCellsXData & ")"
	oCell.setFormula (sFormula)
	sCellXMean = fnGetLocalRangeName (oCell)
	oCell = oReportSheet.getCellByPosition (1, nNRow + 1)
	sFormula = "=AVERAGE(" & sCellsYData & ")"
	oCell.setFormula (sFormula)
	sCellYMean = fnGetLocalRangeName (oCell)
	oCells = oReportSheet.getCellRangeByPosition (0, nNRow, 1, nNRow)
	sCellsN = fnGetLocalRangeName (oCells)
	
	' Calculates the Z values
	sCell = fnGetLocalRangeName (oReportSheet.getCellByPosition (0, 0))
	sFormula = "=""Z""&" & sCell
	oCell = oReportSheet.getCellByPosition (2, 0)
	oCell.setFormula (sFormula)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.collapseToEnd
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharEscapement", -33)
	oCursor.setPropertyValue ("CharEscapementHeight", 58)
	For nRow = 1 To nNX
		sCell = fnGetLocalRangeName (oReportSheet.getCellByPosition (0, nRow))
		sFormula = "=ABS(" & sCell & "-" & sCellXMean & ")"
		oCell = oReportSheet.getCellByPosition (2, nRow)
		oCell.setFormula (sFormula)
	Next nRow
	sCellsXZData = fnGetLocalRangeName (oReportSheet.getCellRangeByPosition (2, 1, 2, nNX))
	oCell = oReportSheet.getCellByPosition (2, nNRow + 1)
	sFormula = "=AVERAGE(" & sCellsXZData & ")"
	oCell.setFormula (sFormula)
	sCellXZMean = fnGetLocalRangeName (oCell)
	
	sCell = fnGetLocalRangeName (oReportSheet.getCellByPosition (1, 0))
	sFormula = "=""Z""&" & sCell
	oCell = oReportSheet.getCellByPosition (3, 0)
	oCell.setFormula (sFormula)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.collapseToEnd
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharEscapement", -33)
	oCursor.setPropertyValue ("CharEscapementHeight", 58)
	For nRow = 1 To nNY
		sCell = fnGetLocalRangeName (oReportSheet.getCellByPosition (1, nRow))
		sFormula = "=ABS(" & sCell & "-" & sCellYMean & ")"
		oCell = oReportSheet.getCellByPosition (3, nRow)
		oCell.setFormula (sFormula)
	Next nRow
	sCellsYZData = fnGetLocalRangeName (oReportSheet.getCellRangeByPosition (3, 1, 3, nNY))
	oCell = oReportSheet.getCellByPosition (3, nNRow + 1)
	sFormula = "=AVERAGE(" & sCellsYZData & ")"
	oCell.setFormula (sFormula)
	sCellYZMean = fnGetLocalRangeName (oCell)
	
	' Calculates the total average
	oCell = oReportSheet.getCellByPosition (4, nNRow)
	sFormula = "=SUM(" & sCellsN & ")"
	oCell.setFormula (sFormula)
	oCell = oReportSheet.getCellByPosition (4, nNRow + 1)
	sFormula = "=AVERAGE(" & sCellsXZData & ";" & sCellsYZData & ")"
	oCell.setFormula (sFormula)
	sCellZMean = fnGetLocalRangeName (oCell)
	
	' Calculates the difference of the Z values to their means
	sCell = fnGetLocalRangeName (oReportSheet.getCellByPosition (0, 0))
	sFormula = "=""dZ""&" & sCell
	oCell = oReportSheet.getCellByPosition (4, 0)
	oCell.setFormula (sFormula)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.collapseToEnd
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharEscapement", -33)
	oCursor.setPropertyValue ("CharEscapementHeight", 58)
	oCursor.collapseToEnd
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharEscapement", -44)
	oCursor.setPropertyValue ("CharEscapementHeight", 34)
	For nRow = 1 To nNX
		sCell = fnGetLocalRangeName (oReportSheet.getCellByPosition (2, nRow))
		sFormula = "=" & sCell & "-" & sCellXZMean
		oCell = oReportSheet.getCellByPosition (4, nRow)
		oCell.setFormula (sFormula)
	Next nRow
	
	sCell = fnGetLocalRangeName (oReportSheet.getCellByPosition (1, 0))
	sFormula = "=""dZ""&" & sCell
	oCell = oReportSheet.getCellByPosition (5, 0)
	oCell.setFormula (sFormula)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.collapseToEnd
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharEscapement", -33)
	oCursor.setPropertyValue ("CharEscapementHeight", 58)
	oCursor.collapseToEnd
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharEscapement", -44)
	oCursor.setPropertyValue ("CharEscapementHeight", 34)
	For nRow = 1 To nNY
		sCell = fnGetLocalRangeName (oReportSheet.getCellByPosition (3, nRow))
		sFormula = "=" & sCell & "-" & sCellYZMean
		oCell = oReportSheet.getCellByPosition (5, nRow)
		oCell.setFormula (sFormula)
	Next nRow
	
	' Calculates the difference of the Z means to the total mean
	sCell = fnGetLocalRangeName (oReportSheet.getCellByPosition (2, nNRow + 1))
	sFormula = "=" & sCell & "-" & sCellZMean
	oCell = oReportSheet.getCellByPosition (2, nNRow + 2)
	oCell.setFormula (sFormula)
	sCell = fnGetLocalRangeName (oReportSheet.getCellByPosition (3, nNRow + 1))
	sFormula = "=" & sCell & "-" & sCellZMean
	oCell = oReportSheet.getCellByPosition (3, nNRow + 2)
	oCell.setFormula (sFormula)
	
	fnCollectIndependentTTestData = oReportSheet.getCellRangeByPosition (0, 0, 5, nNRow + 2)
End Function

' fnGetLeveneTest: Returns the Levene's test result.
Function fnGetLeveneTest (oZDataRange As Object) As String
	Dim nK As Integer, nRows As Integer
	Dim oCell As Object, oCells As Object
	Dim sCellN As String, sCellsN As String
	Dim sCellsDZMean As String, sCellsDZData As String
	
	nRows = oZDataRange.getRows.getCount
	nK = oZDataRange.getColumns.getCount / 3
	oCell = oZDataRange.getCellByPosition (nK * 2, nRows - 3)
	sCellN = fnGetRangeName (oCell)
	oCells = oZDataRange.getCellRangeByPosition (0, nRows - 3, nK - 1, nRows - 3)
	sCellsN = fnGetRangeName (oCells)
	oCells = oZDataRange.getCellRangeByPosition (nK, nRows - 1, nK * 2 - 1, nRows - 1)
	sCellsDZMean = fnGetRangeName (oCells)
	oCells = oZDataRange.getCellRangeByPosition (nK * 2, 1, nK * 3 - 1, nRows - 4)
	sCellsDZData = fnGetRangeName (oCells)
	fnGetLeveneTest = "((" & sCellN & "-COUNT(" & sCellsN & "))/(COUNT(" & sCellsN & ")-1))*(SUMPRODUCT(" & sCellsN & ";" & sCellsDZMean & ";" & sCellsDZMean & ")/SUMPRODUCT(" & sCellsDZData & ";" & sCellsDZData & "))"
End Function
