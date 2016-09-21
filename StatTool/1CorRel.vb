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

' 1CorRel: The macros to for generating the report of the Pearson’s correlation coefficient
'   by imacat <imacat@mail.imacat.idv.tw>, 2016-08-10

Option Explicit

' subRunCorrelation: Runs the Pearson’s correlation coefficient.
Sub subRunCorrelation As Object
	Dim oRange As Object
	Dim oSheets As Object, sSheetName As String
	Dim oSheet As Object, mRanges As Object
	Dim sExisted As String, nResult As Integer
	
	DialogLibraries.loadLibrary "StatTool"
	
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
	If oSheets.hasByName (sSheetName & "_correl") Then
		sExisted = "Spreadsheet """ & sSheetName & "_correl"" exists.  Overwrite?"
		nResult = MsgBox (sExisted, MB_YESNO + MB_DEFBUTTON2 + MB_ICONQUESTION)
		If nResult = IDNO Then
			Exit Sub
		End If
		' Drops the existing report
		oSheets.removeByname (sSheetName & "_correl")
	End If
	
	' Reports the paired T-test.
	subReportCorrelation (ThisComponent, mRanges (0), mRanges (1))
	oSheet = oSheets.getByName (sSheetName & "_correl")
	
	' Makes the report sheet active.
	ThisComponent.getCurrentController.setActiveSheet (oSheet)
End Sub

' subReportCorrelation: Reports the Pearson’s correlation coefficient
Sub subReportCorrelation (oDoc As Object, oDataXRange As Object, oDataYRange As Object)
	Dim oSheets As Object, sSheetName As String, oTmpSheet As Object
	Dim mNames () As String, nI As Integer, nSheetIndex As Integer
	Dim oSheet As Object, oColumns As Object, nRow As Integer
	Dim oCell As Object, oCells As Object, oCursor As Object
	Dim nN As Long, sFormula As String
	Dim sNotes As String, nPos As Integer
	Dim nFormatN As Integer, nFormatF As Integer, nFormatP As Integer
	Dim aBorderSingle As New com.sun.star.table.BorderLine
	Dim aBorderDouble As New com.sun.star.table.BorderLine
	Dim sCellXLabel As String, sCellsXData As String
	Dim sCellYLabel As String, sCellsYData As String
	Dim sCellN As String, sCellR As String, sCellP As String
	
	oSheets = oDoc.getSheets
	sSheetName = oDataXRange.getSpreadsheet.getName
	mNames = oSheets.getElementNames
	For nI = 0 To UBound (mNames)
		If mNames (nI) = sSheetName Then
			nSheetIndex = nI
		End If
	Next nI
	oSheets.insertNewByName (sSheetName & "_correl", nSheetIndex + 1)
	oSheet = oSheets.getByName (sSheetName & "_correl")
	
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
	
	' Correlation
	nRow = nRow + 2
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Pearson’s Correlation")
	oCell.setPropertyValue ("CellStyle", "Result2")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 4, nRow)
	oCells.merge (True)
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("X")
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCell = oSheet.getCellByPosition (1, nRow)
	oCell.setString ("Y")
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
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
	oCell.setString ("r")
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
	sFormula = "=COUNT(" & sCellsXData & ")"
	oCell.setFormula (sFormula)
	sCellN = fnGetLocalRangeName (oCell)
	oCell.setPropertyValue ("NumberFormat", nFormatN)
	oCell = oSheet.getCellByPosition (3, nRow)
	sFormula = "=CORREL(" & sCellsXData & ";" & sCellsYData & ")"
	oCell.setFormula (sFormula)
	sCellR = fnGetLocalRangeName (oCell)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	oCell = oSheet.getCellByPosition (4, nRow)
	sFormula = "=TDIST(ABS(" & sCellR & "*SQRT((" & sCellN & "-2)/(1-" & sCellR & "*" & sCellR & ")));" & sCellN & "-2;2)"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatP)
	sCellP = fnGetLocalRangeName (oCell)
	
	' The foot notes of the test.
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Note: *: p<.05, **: p<.01" & Chr (10) & _
		"H0: ρ=0 (the populations of the two groups are irrelavent)." & Chr (10) & _
		"H1: ρ≠0 (the populations of the two groups are relevant) if the probability (p) is small enough.")
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
	nPos = InStr (sNotes, "ρ")
	Do While nPos <> 0
		oCursor.gotoStart (False)
		oCursor.goRight (nPos - 1, False)
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
		nPos = InStr (nPos + 1, sNotes, "ρ")
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
	
	' Adds an X-Y diagram.
	subAddChart (oSheet, nRow, oDataXRange, oDataYRange)
	
	' Adds the linear regression line when there is a linear relation
	If oSheet.getCellRangeByName (sCellP).getValue < 0.05 Then
		' Refresh this sheet and draws the chart in advance.
		oTmpSheet = oDoc.getCurrentController.getActiveSheet
		oDoc.getCurrentController.setActiveSheet (oSheet)
		oDoc.getCurrentController.setActiveSheet (oTmpSheet)
		subAddLinearRegression(oSheet, oDataXRange, oDataYRange)
	End If
End Sub

' subAddChart: Adds a chart for the data
Sub subAddChart (oSheet As Object, nRow As Integer, oDataXRange As Object, oDataYRange As Object)
	Dim nI As Integer, nY As Long
	Dim oCharts As Object, oChart As Object
	Dim oChartDoc As Object, oDiagram As Object
	Dim aPos As New com.sun.star.awt.Rectangle
	Dim mAddrs (1) As New com.sun.star.table.CellRangeAddress
	Dim sTitle As String
	Dim oProvider As Object, oData As Object
	Dim sRange As String, mData () As Object
	
	' Finds the Y position to place the chart.
	nY = 0
	For nI = 0 To nRow + 1
		nY = nY + oSheet.getRows.getByIndex (nI).getPropertyValue ("Height")
	Next nI
	
	' Adds the chart
	With aPos
		.X = 0
		.Y = nY
		.Width = 10000
		.Height = 10000
	End With
	mAddrs (0) = oDataXRange.getRangeAddress
	mAddrs (1) = oDataYRange.getRangeAddress
	oCharts = oSheet.getCharts
	oCharts.addNewByName (oSheet.getName, aPos, mAddrs, True, False)
	oChart = oCharts.getByName (oSheet.getName)
	oChartDoc = oChart.getEmbeddedObject
	
	oDiagram = oChartDoc.createInstance ( _
		"com.sun.star.chart.XYDiagram")
	oDiagram.setPropertyValue ("Lines", False)
	oDiagram.setPropertyValue ("HasXAxisGrid", False)
	oDiagram.setPropertyValue ("HasYAxisGrid", False)
	sTitle = oDataXRange.getCellByPosition (0, 0).getString
	oDiagram.getXAxisTitle.setPropertyValue ("String", sTitle)
	sTitle = oDataYRange.getCellByPosition (0, 0).getString
	oDiagram.getYAxisTitle.setPropertyValue ("String", sTitle)
	'oDiagram.getXAxis.setPropertyValue ("Min", 0)
	'oDiagram.getYAxis.setPropertyValue ("Min", 0)
	With aPos
		.X = 1500
		.Y = 1000
		.Width = 7500
		.Height = 7500
	End With
	oDiagram.setDiagramPositionExcludingAxes (aPos)
	oChartDoc.setDiagram (oDiagram)
	
	' Sets the data sequences for the X-axis and Y-axis
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

' subAddLinearRegression: Adds the linear regression line
Sub subAddLinearRegression (oSheet As Object, oDataXRange As Object, oDataYRange As Object)
	Dim oChart As Object, oChartDoc As Object
	Dim oDrawPage As Object, oChartPageShape As Object
	Dim oDiagramSetShape As Object, oDiagramShape As Object
	Dim aDiagramSize As New com.sun.star.awt.Size
	Dim aDiagramPos As New com.sun.star.awt.Point
	Dim oDiagram As Object
	Dim oXAxis As Object, fXMin As Double, fXMax As Double
	Dim oYAxis As Object, fYMin As Double, fYMax As Double
	Dim oShape As Object, mDataX As Variant, mDataY As Variant
	Dim nI As Long, fSumX As Double, fSumY As Double
	Dim fSumXY As Double, fSumX2 As Double, nN As Long
	Dim fA As Double, fB As Double
	Dim fX0 As Double, fY0 As Double, fX1 As Double, fY1 As Double
	Dim aSize As New com.sun.star.awt.Size
	Dim aPos As New com.sun.star.awt.Point
	Dim aDash As New com.sun.star.drawing.LineDash
	
	oChartDoc = oSheet.getCharts.getByIndex (0).getEmbeddedObject
	
	oChartPageShape = oChartDoc.getDrawPage.getByIndex (0)
	oDiagramSetShape = oChartPageShape.getByIndex (1)
	oDiagramShape = oDiagramSetShape.getByIndex (0)
	aDiagramSize = oDiagramShape.getSize
	aDiagramPos = oDiagramShape.getPosition
	
	oDiagram = oChartDoc.getDiagram
	oXAxis = oDiagram.getXAxis
	fXMin = oXAxis.getPropertyValue ("Min")
	fXMax = oXAxis.getPropertyValue ("Max")
	oYAxis = oDiagram.getYAxis
	fYMin = oYAxis.getPropertyValue ("Min")
	fYMax = oYAxis.getPropertyValue ("Max")
	
	mDataX = oDataXRange.getCellRangeByPosition (0, 1, 0, oDataXRange.getRows.getCount - 1).getDataArray
	mDataY = oDataYRange.getCellRangeByPosition (0, 1, 0, oDataYRange.getRows.getCount - 1).getDataArray
	nN = UBound (mDataX) + 1
	fSumX = 0
	fSumY = 0
	fSumXY = 0
	fSumX2 = 0
	For nI = 0 To UBound (mDataX)
		fSumX = fSumX + mDataX (nI) (0)
		fSumY = fSumY + mDataY (nI) (0)
		fSumXY = fSumXY + mDataX (nI) (0) * mDataY (nI) (0)
		fSumX2 = fSumX2 + mDataX (nI) (0) * mDataX (nI) (0)
	Next nI
	fB = (fSumXY - fSumX * fSumY / nN) / (fSumX2 - fSumX * fSumX / nN)
	fA = (fSumY / nN) - fB * (fSumX / nN)
	fX0 = fXMin
	fY0 = fB * fX0 + fA
	If fY0 < fYMin Then
		fY0 = fYMin
		fX0 = (fY0 - fA) / fB
	End If
	If fY0 > fYMax Then
		fY0 = fYMax
		fX0 = (fY0 - fA) / fB
	End If
	fX1 = fXMax
	fY1 = fB * fX1 + fA
	If fY1 < fYMin Then
		fY1 = fYMin
		fX1 = (fY1 - fA) / fB
	End If
	If fY1 > fYMax Then
		fY1 = fYMax
		fX1 = (fY1 - fA) / fB
	End If
	
	' Adds the linear regression line.
	oShape = oChartDoc.createInstance ("com.sun.star.drawing.LineShape")
	With aSize
		.Width = aDiagramSize.Width * (fX1 - fX0) / (fXMax - fXMin)
		.Height = -aDiagramSize.Height * (fY1 - fY0) / (fYMax - fYMin)
	End With
	oShape.setSize (aSize)
	With aPos
		.X = aDiagramPos.X + aDiagramSize.Width * fX0 / (fXMax - fXMin)
		.Y = aDiagramPos.Y + aDiagramSize.Height - aDiagramSize.Height * fY0 / (fYMax - fYMin)
	End With
	oShape.setPosition (aPos)
	oShape.setPropertyValue ("LineStyle", com.sun.star.drawing.LineStyle.DASH)
	With aDash
		.Style = com.sun.star.drawing.DashStyle.RECT
		.Dots = 1
		.DotLen = 197
		.Dashes = 0
		.DashLen = 0
		.Distance = 120
	End With
	oShape.setPropertyValue ("LineDash", aDash)
	oShape.setPropertyValue ("LineWidth", 100)
	oShape.setPropertyValue ("LineColor", RGB (255, 0, 0))
	
	'oSheet.getDrawPage.add (oShape)
	oChartDoc.getDrawPage.add (oShape)
End Sub
