' Copyright (c) 2016 imacat.
' 
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
' 
'     http://www.apache.org/licenses/LICENSE-2.0
' 
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.

' 2PTTest: The macros to for generating the report of paired T-Test
'   by imacat <imacat@mail.imacat.idv.tw>, 2016-08-11

Option Explicit

' subRunPairedTTest: Runs the paired T-test.
Sub subRunPairedTTest As Object
	Dim oRange As Object
	Dim oSheets As Object, sSheetName As String
	Dim oSheet As Object, mRanges As Object
	Dim sExisted As String, nResult As Integer
	
	DialogLibraries.loadLibrary "Stats"
	
	' Asks the user for the data range
	oRange = fnAskDataRange (ThisComponent)
	If IsNull (oRange) Then
		Exit Sub
	End If
	
	' Specifies the data
	mRanges = fnSpecifyData (oRange, _
		"&3.Dlg2SpecData.txtPrompt1.Label1CorRel", _
		"&6.Dlg2SpecData.txtPrompt2.Label1CorRel")
	If IsNull (mRanges) Then
		Exit Sub
	End If
	
	' Checks the existing report
	oSheets = ThisComponent.getSheets
	sSheetName = oRange.getSpreadsheet.getName
	If oSheets.hasByName (sSheetName & "_ttest") Then
		sExisted = "Spreadsheet """ & sSheetName & "_ttest"" exists.  Overwrite?"
		nResult = MsgBox (sExisted, MB_YESNO + MB_DEFBUTTON2 + MB_ICONQUESTION)
		If nResult = IDNO Then
			Exit Sub
		End If
		' Drops the existing report
		oSheets.removeByname (sSheetName & "_ttest")
	End If
	
	' Reports the paired T-test.
	subReportPairedTTest (ThisComponent, mRanges (0), mRanges (1))
	oSheet = oSheets.getByName (sSheetName & "_ttest")
	
	' Adds an X-Y diagram.
	subAddChart (oSheet, mRanges (0), mRanges (1))
	
	' Makes the report sheet active.
	ThisComponent.getCurrentController.setActiveSheet (oSheet)
End Sub

' subAddChart: Adds a chart for the data
Sub subAddChart (oSheet As Object, oDataXRange As Object, oDataYRange As Object)
	Dim oCharts As Object, oChart As Object
	Dim oChartDoc As Object, oDiagram As Object
	Dim aPos As New com.sun.star.awt.Rectangle
	Dim mAddrs (1) As New com.sun.star.table.CellRangeAddress
	Dim sTitle As String
	Dim oProvider As Object, oData As Object
	Dim sRange As String, mData () As Object
	
	' Adds the chart
	With aPos
		.X = 0
		.Y = 10000
		.Width = 10000
		.Height = 10000
	End With
	mAddrs (0) = oDataXRange.getRangeAddress
	mAddrs (1) = oDataYRange.getRangeAddress
	oCharts = oSheet.getCharts
	oCharts.addNewByName (oSheet.getName, aPos, mAddrs, True, False)
	oChart = oCharts.getByName (oSheet.getName)
	oChartDoc = oChart.getEmbeddedObject
	
	BasicLibraries.loadLibrary "XrayTool"
	oDiagram = oChartDoc.createInstance ( _
		"com.sun.star.chart.XYDiagram")
	oDiagram.setPropertyValue ("Lines", False)
	oDiagram.setPropertyValue ("HasXAxisGrid", False)
	oDiagram.setPropertyValue ("HasYAxisGrid", False)
	sTitle = oDataXRange.getCellByPosition (0, 0).getString
	oDiagram.getXAxisTitle.setPropertyValue ("String", sTitle)
	sTitle = oDataYRange.getCellByPosition (0, 0).getString
	oDiagram.getYAxisTitle.setPropertyValue ("String", sTitle)
	oDiagram.getXAxis.setPropertyValue ("Min", 0)
	oDiagram.getYAxis.setPropertyValue ("Min", 0)
	oChartDoc.setDiagram (oDiagram)
	
	oProvider = oChartDoc.getDataProvider
	mData = oChartDoc.getDataSequences
	sRange = oDataXRange.getCellByPosition(0, 0).getPropertyValue ("AbsoluteName")
	oData = oProvider.createDataSequenceByRangeRepresentation (sRange)
	mData (0).setLabel (oData)
	sRange = oDataXRange.getCellRangeByPosition(0, 1, 0, oDataXRange.getRows.getCount - 1).getPropertyValue ("AbsoluteName")
	oData = oProvider.createDataSequenceByRangeRepresentation (sRange)
	oData.Role = "values-x"
	mData (0).setValues (oData)
	sRange = oDataYRange.getCellByPosition(0, 0).getPropertyValue ("AbsoluteName")
	oData = oProvider.createDataSequenceByRangeRepresentation (sRange)
	mData (1).setLabel (oData)
	sRange = oDataYRange.getCellRangeByPosition(0, 1, 0, oDataYRange.getRows.getCount - 1).getPropertyValue ("AbsoluteName")
	oData = oProvider.createDataSequenceByRangeRepresentation (sRange)
	oData.Role = "values-y"
	mData (1).setValues (oData)
	
	oChartDoc.setPropertyValue ("HasLegend", False)
End Sub

' subReportPairedTTest: Reports the paired T-test
Sub subReportPairedTTest (oDoc As Object, oDataXRange As Object, oDataYRange As Object)
	Dim oSheets As Object, sSheetName As String
	Dim mNames () As String, nI As Integer, nSheetIndex As Integer
	Dim oSheet As Object, oColumns As Object, nRow As Integer
	Dim oCell As Object, oCells As Object, oCursor As Object
	Dim nN As Integer, sFormula As String
	Dim sNotes As String, nPos As Integer
	Dim nFormatN As Integer, nFormatF As Integer, nFormatP As Integer
	Dim aBorderSingle As New com.sun.star.table.BorderLine
	Dim aBorderDouble As New com.sun.star.table.BorderLine
	Dim sCellXLabel As String, sCellsXData As String
	Dim sCellXN As String, sCellXMean As String, sCellXS As String
	Dim sCellYLabel As String, sCellsYData As String
	Dim sCellYN As String, sCellYMean As String, sCellYS As String
	Dim sCellN As String, sCellXYS As String, sCellR As String
	
	oSheets = oDoc.getSheets
	sSheetName = oDataXRange.getSpreadsheet.getName
	mNames = oSheets.getElementNames
	For nI = 0 To UBound (mNames)
		If mNames (nI) = sSheetName Then
			nSheetIndex = nI
		End If
	Next nI
	oSheets.insertNewByName (sSheetName & "_ttest", nSheetIndex + 1)
	oSheet = oSheets.getByName (sSheetName & "_ttest")
	
	nN = oDataXRange.getRows.getCount - 1
	sCellXLabel = fnGetRangeName (oDataXRange.getCellByPosition (0, 0))
	sCellsXData = fnGetRangeName (oDataXRange.getCellRangeByPosition (0, 1, 0, nN))
	sCellYLabel = fnGetRangeName (oDataYRange.getCellByPosition (0, 0))
	sCellsYData = fnGetRangeName (oDataYRange.getCellRangeByPosition (0, 1, 0, nN))
	
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
	oColumns.getByIndex (1).setPropertyValue ("Width", 3060)
	oColumns.getByIndex (2).setPropertyValue ("Width", 2080)
	oColumns.getByIndex (3).setPropertyValue ("Width", 2080)
	oColumns.getByIndex (4).setPropertyValue ("Width", 2080)
	
	nRow = -2
	' Group description
	nRow = nRow + 2
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Sample Description")
	oCell.setPropertyValue ("CellStyle", "Result2")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 5, nRow)
	oCells.merge (True)
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Sample")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 1, nRow)
	oCells.merge (True)
	oCell = oSheet.getCellByPosition (2, nRow)
	oCell.setString ("N")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCell = oSheet.getCellByPosition (3, nRow)
	oCell.setString ("X")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharOverline", com.sun.star.awt.FontUnderline.SINGLE)
	oCell = oSheet.getCellByPosition (4, nRow)
	oCell.setString ("s")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCell = oSheet.getCellByPosition (5, nRow)
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
	oCells = oSheet.getCellRangeByPosition (0, nRow, 1, nRow)
	oCells.merge (True)
	oCell = oSheet.getCellByPosition (2, nRow)
	sFormula = "=COUNT(" & sCellsXData & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatN)
	sCellXN = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (3, nRow)
	sFormula = "=AVERAGE(" & sCellsXData & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellXMean = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (4, nRow)
	sFormula = "=STDEV(" & sCellsXData & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellXS = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (5, nRow)
	sFormula = "=" & sCellXS & "/SQRT(" & sCellXN & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	
	' The second group
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	sFormula = "=" & sCellYLabel
	oCell.setFormula (sFormula)
	oCells = oSheet.getCellRangeByPosition (0, nRow, 1, nRow)
	oCells.merge (True)
	oCell = oSheet.getCellByPosition (2, nRow)
	sFormula = "=COUNT(" & sCellsYData & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatN)
	sCellYN = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (3, nRow)
	sFormula = "=AVERAGE(" & sCellsYData & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellYMean = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (4, nRow)
	sFormula = "=STDEV(" & sCellsYData & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellYS = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (5, nRow)
	sFormula = "=" & sCellYS & "/SQRT(" & sCellYN & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	
	' The difference between the two groups
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	sFormula = "=""(""&" & sCellXLabel & "&""-""&" & sCellYLabel & "&"")"""
	oCell.setFormula (sFormula)
	oCells = oSheet.getCellRangeByPosition (0, nRow, 1, nRow)
	oCells.merge (True)
	oCell = oSheet.getCellByPosition (2, nRow)
	sFormula = "=" & sCellXN
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatN)
	sCellN = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (3, nRow)
	sFormula = "=" & sCellXMean & "-" & sCellYMean
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	oCell = oSheet.getCellByPosition (4, nRow)
	sFormula = "=SQRT(" & sCellXS & "*" & sCellXS & "-2*SUMPRODUCT(" & sCellsXData & ";" & sCellsYData & ")/(" & sCellN & "-1)+2*" & sCellXMean & "*" & sCellYMean & "*" & sCellN & "/(" & sCellN & "-1)+" & sCellYS & "*" & sCellYS & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellXYS = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (5, nRow)
	sFormula = "=" & sCellXYS & "/SQRT(" & sCellN & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	
	' Draws the table borders.
	oCells = oSheet.getCellByPosition (0, nRow - 3)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellRangeByPosition (1, nRow - 3, 5, nRow - 3)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellRangeByPosition (0, nRow - 2, 0, nRow - 1)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells = oSheet.getCellByPosition (0, nRow)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	oCells = oSheet.getCellRangeByPosition (1, nRow, 5, nRow)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	
	' Correlation
	nRow = nRow + 2
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Pearson’s Correlation")
	oCell.setPropertyValue ("CellStyle", "Result2")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 3, nRow)
	oCells.merge (True)
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("X1")
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
	oCell = oSheet.getCellByPosition (1, nRow)
	oCell.setString ("X2")
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
	oCell = oSheet.getCellByPosition (2, nRow)
	oCell.setString ("r")
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
	
	' The test result.
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	sFormula = "=" & sCellXLabel
	oCell.setFormula (sFormula)
	oCell = oSheet.getCellByPosition (1, nRow)
	sFormula = "=" & sCellYLabel
	oCell.setFormula (sFormula)
	oCell = oSheet.getCellByPosition (2, nRow)
	sFormula = "=CORREL(" & sCellsXData & ";" & sCellsYData & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellR = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (3, nRow)
	sFormula = "=TDIST(ABS(" & sCellR & ")*SQRT((" & sCellN & "-2)/(1-" & sCellR & "*" & sCellR & "))" & ";" & sCellN & "-2;2)"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatP)
	
	' The foot notes of the test.
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Note: *: p<.05, **: p<.01" & Chr (10) & _
		"H0: ρ=0 (the populations of the two samples are irrelavent)." & Chr (10) & _
		"H1: ρ≠0 (the populations of the two samples are relevant) if the probability (p) is small enough.")
	oCell.setPropertyValue ("IsTextWrapped", True)
	oCells = oSheet.getCellRangeByPosition (0, nRow, 3, nRow)
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
	nPos = InStr (sNotes, "ρ")
	Do While nPos <> 0
		oCursor.gotoStart (False)
		oCursor.goRight (nPos - 1, False)
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
		nPos = InStr (nPos + 1, sNotes, "p<")
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
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellByPosition (1, nRow - 2)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellRangeByPosition (2, nRow - 2, 3, nRow - 2)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellByPosition (0, nRow - 1)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	oCells = oSheet.getCellByPosition (1, nRow - 1)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	oCells = oSheet.getCellRangeByPosition (2, nRow - 1, 3, nRow - 1)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	
	' Paired-samples T-test
	nRow = nRow + 2
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Paired-Samples T-test")
	oCell.setPropertyValue ("CellStyle", "Result2")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 4, nRow)
	oCells.merge (True)
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("X1")
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
	oCell = oSheet.getCellByPosition (1, nRow)
	oCell.setString ("X2")
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
	oCell = oSheet.getCellByPosition (2, nRow)
	oCell.setString ("t")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCell = oSheet.getCellByPosition (3, nRow)
	oCell.setString ("df")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCell = oSheet.getCellByPosition (4, nRow)
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
	sFormula = "=" & sCellXLabel
	oCell.setFormula (sFormula)
	oCell = oSheet.getCellByPosition (1, nRow)
	sFormula = "=" & sCellYLabel
	oCell.setFormula (sFormula)
	oCell = oSheet.getCellByPosition (2, nRow)
	sFormula = "=(" & sCellXMean & "-" & sCellYMean & ")/SQRT((" & sCellXS & "*" & sCellXS & "+" & sCellYS & "*" & sCellYS & "-2*" & sCellR & "*" & sCellXS & "*" & sCellYS & ")/" & sCellN & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	oCell = oSheet.getCellByPosition (3, nRow)
	sFormula = "=" & sCellN & "-1"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatN)
	oCell = oSheet.getCellByPosition (4, nRow)
	sFormula = "=TTEST(" & sCellsXData & ";" & sCellsYData & ";2;1)"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatP)
	
	' The foot notes of the test.
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Note: *: p<.05, **: p<.01" & Chr (10) & _
		"H0: μ1=μ2 (the populations of the two samples have the same means)." & Chr (10) & _
		"H1: μ1≠μ2 (the populations of the two samples have different means) if the probability (p) is small enough.")
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
	oCells = oSheet.getCellByPosition (0, nRow - 2)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellByPosition (1, nRow - 2)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellRangeByPosition (2, nRow - 2, 4, nRow - 2)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellByPosition (0, nRow - 1)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	oCells = oSheet.getCellByPosition (1, nRow - 1)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	oCells = oSheet.getCellRangeByPosition (2, nRow - 1, 4, nRow - 1)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
End Sub
