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

' 5Chi2GoF: The macros to for generating the report of Chi-square goodness of fit
'   by imacat <imacat@mail.imacat.idv.tw>, 2016-09-05

Option Explicit

' subRunChi2GoodnessOfFit: Runs the chi-square goodness of fit.
Sub subRunChi2GoodnessOfFit As Object
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
		"&12.Dlg2SpecData.txtPrompt1.Label5Chi2GoF", _
		"&13.Dlg2SpecData.txtPrompt2.Label5Chi2GoF")
	If IsNull (mRanges) Then
		Exit Sub
	End If
	
	' Checks the existing report
	oSheets = ThisComponent.getSheets
	sSheetName = oRange.getSpreadsheet.getName
	If oSheets.hasByName (sSheetName & "_chi2") Then
		sExisted = "Spreadsheet """ & sSheetName & "_chi2"" exists.  Overwrite?"
		nResult = MsgBox (sExisted, MB_YESNO + MB_DEFBUTTON2 + MB_ICONQUESTION)
		If nResult = IDNO Then
			Exit Sub
		End If
		' Drops the existing report
		oSheets.removeByname (sSheetName & "_chi2")
	End If
	
	' Reports the chi-square goodness of fit
	subReportChi2GoodnessOfFit (ThisComponent, mRanges (0), mRanges (1))
	oSheet = oSheets.getByName (sSheetName & "_chi2")
	
	' Makes the report sheet active.
	ThisComponent.getCurrentController.setActiveSheet (oSheet)
End Sub

' subReportChi2GoodnessOfFit: Reports the chi-square goodness of fit
Sub subReportChi2GoodnessOfFit (oDoc As Object, oColumnColumn As Object, oRowColumn As Object)
	Dim oSheets As Object, sSheetName As String
	Dim nI As Integer, nJ As Integer, nJ1 As Integer, nJ2 As Integer
	Dim mNames () As String, nSheetIndex As Integer
	Dim oSheet As Object, oColumns As Object, nRow As Integer, nStartRow As Integer
	Dim oCell As Object, oCells As Object, oCursor As Object
	Dim sFormula As String
	Dim sNotes As String, nPos As Integer
	Dim nFormatN As Integer, nFormatF As Integer, nFormatP As Integer, nFormatPct As Integer
	Dim aBorderSingle As New com.sun.star.table.BorderLine
	Dim aBorderDouble As New com.sun.star.table.BorderLine
	
	Dim sCellsJData As String, sCellsIData As String
	Dim sLabel As String, sLabelsColumn As String, sLabelsRow As String
	Dim nGroups As Integer, nEvents As Integer
	Dim mCellLabelColomn () As String, mCellLabelRow () As String
	Dim mCellNJ () As String, mCellNI () As String, mCellPI () As String
	Dim mCellFrequency (0, 0) As String, mCellProportion (0, 0) As String
	Dim sCellN As String
	
	Dim sCell As String, sCells As String
	Dim sCellsRow As String, sCellsColumn As String
	Dim sCellChi2 As String, sCellDF As String
	Dim sSE2 AS String, nTotalColumns As Integer
	
	oSheets = oDoc.getSheets
	sSheetName = oColumnColumn.getSpreadsheet.getName
	mNames = oSheets.getElementNames
	For nI = 0 To UBound (mNames)
		If mNames (nI) = sSheetName Then
			nSheetIndex = nI
		End If
	Next nI
	oSheets.insertNewByName (sSheetName & "_chi2", nSheetIndex + 1)
	oSheet = oSheets.getByName (sSheetName & "_chi2")
	
	sCellsJData = fnGetRangeName (oColumnColumn.getCellRangeByPosition (0, 1, 0, oColumnColumn.getRows.getCount - 1))
	sCellsIData = fnGetRangeName (oRowColumn.getCellRangeByPosition (0, 1, 0, oRowColumn.getRows.getCount - 1))
	
	' Counts the number of groups and events
	sLabelsColumn = " "
	sLabelsRow = " "
	nGroups = 0
	nEvents = 0
	For nRow = 1 To oColumnColumn.getRows.getCount - 1
		sLabel = oColumnColumn.getCellByPosition (0, nRow).getString
		If InStr (sLabelsColumn, " " & sLabel & " ") = 0 Then
			sLabelsColumn = sLabelsColumn & sLabel & " "
			nGroups = nGroups + 1
		End If
		sLabel = oRowColumn.getCellByPosition (0, nRow).getString
		If InStr (sLabelsRow, " " & sLabel & " ") = 0 Then
			sLabelsRow = sLabelsRow & sLabel & " "
			nEvents = nEvents + 1
		End If
	Next nRow
	
	ReDim mCellLabelColomn (nGroups - 1) As String, mCellLabelRow (nEvents - 1) As String
	ReDim mCellNJ (nGroups - 1) As String, mCellNI (nEvents - 1) As String
	ReDim mCellPI (nEvents - 1) As String
	ReDim mCellFrequency (nGroups - 1, nEvents - 1) As String
	ReDim mCellProportion (nGroups - 1, nEvents - 1) As String
	
	' Collects the group and event labels
	sLabelsColumn = " "
	sLabelsRow = " "
	nJ = 0
	nI = 0
	For nRow = 1 To oColumnColumn.getRows.getCount - 1
		oCell = oColumnColumn.getCellByPosition (0, nRow)
		sLabel = oCell.getString
		If InStr (sLabelsColumn, " " & sLabel & " ") = 0 Then
			sLabelsColumn = sLabelsColumn & sLabel & " "
			mCellLabelColomn (nJ) = fnGetRangeName (oCell)
			nJ = nJ + 1
		End If
		oCell = oRowColumn.getCellByPosition (0, nRow)
		sLabel = oCell.getString
		If InStr (sLabelsRow, " " & sLabel & " ") = 0 Then
			sLabelsRow = sLabelsRow & sLabel & " "
			mCellLabelRow (nI) = fnGetRangeName (oCell)
			nI = nI + 1
		End If
	Next nRow
	
	' Obtains the format parameters for the report.
	nFormatN = fnQueryFormat (oDoc, "#,##0")
	nFormatF = fnQueryFormat (oDoc, "#,###.000")
	nFormatP = fnQueryFormat (oDoc, "[<0.01]#.000""**"";[<0.05]#.000""*"";#.000")
	nFormatPct = fnQueryFormat (oDoc, "0.0%")
	
	aBorderSingle.OuterLineWidth = 2
	aBorderDouble.OuterLineWidth = 2
	aBorderDouble.InnerLineWidth = 2
	aBorderDouble.LineDistance = 2
	
	' Sets the column widths of the report.
	nTotalColumns = nGroups + 2
	If nEvents = 2 Then
		If nTotalColumns < 5 Then
			nTotalColumns = 5
		End If
	Else
		If nTotalColumns < 6 Then
			nTotalColumns = 6
		End If
	End If
	oColumns = oSheet.getColumns
	For nJ = 0 To nTotalColumns - 1
		oColumns.getByIndex (nJ).setPropertyValue ("Width", 3060)
	Next nJ
	
	nRow = -2
	
	' Group description
	nRow = nRow + 2
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Crosstabulation")
	oCell.setPropertyValue ("CellStyle", "Result2")
	oCells = oSheet.getCellRangeByPosition (0, nRow, nGroups + 1, nRow)
	oCells.merge (True)
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Event")
	For nJ = 0 To nGroups - 1
		oCell = oSheet.getCellByPosition (nJ + 1, nRow)
		sFormula = "=" & mCellLabelColomn (nJ)
		oCell.setFormula (sFormula)
		mCellLabelColomn (nJ) = fnGetLocalRangeName (oCell)
	Next nJ
	oCell = oSheet.getCellByPosition (nGroups + 1, nRow)
	oCell.setString ("Total")
	
	' Shows each event
	nRow = nRow - 1
	For nI = 0 To nEvents - 1
		nRow = nRow + 2
		oCell = oSheet.getCellByPosition (0, nRow)
		sFormula = "=" & mCellLabelRow (nI)
		oCell.setFormula (sFormula)
		oCell.setPropertyValue ("VertJustify", com.sun.star.table.CellVertJustify.TOP)
		mCellLabelRow (nI) = fnGetLocalRangeName (oCell)
		oCells = oSheet.getCellRangeByPosition (0, nRow, 0, nRow + 1)
		oCells.merge (True)
		For nJ = 0 To nGroups - 1
			oCell = oSheet.getCellByPosition (1 + nJ, nRow)
		    sFormula = "=COUNTIFS(" & sCellsJData & ";" & mCellLabelColomn (nJ) & ";" & sCellsIData & ";" & mCellLabelRow (nI) & ")"
		    oCell.setFormula (sFormula)
			oCell.setPropertyValue ("NumberFormat", nFormatN)
			mCellFrequency (nJ, nI) = fnGetLocalRangeName (oCell)
		Next nJ
		oCell = oSheet.getCellByPosition (1 + nGroups, nRow)
		sCells = fnGetLocalRangeName (oSheet.getCellRangeByPosition (1, nRow, nGroups, nRow))
		sFormula = "=SUM(" & sCells & ")"
		oCell.setFormula (sFormula)
		oCell.setPropertyValue ("NumberFormat", nFormatN)
		mCellNI (nI) = fnGetLocalRangeName (oCell)
	Next nI
	' Shows the total
	nRow = nRow + 2
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Total")
	oCell.setPropertyValue ("VertJustify", com.sun.star.table.CellVertJustify.TOP)
	oCells = oSheet.getCellRangeByPosition (0, nRow, 0, nRow + 1)
	oCells.merge (True)
	For nJ = 0 To nGroups - 1
		oCell = oSheet.getCellByPosition (1 + nJ, nRow)
		sFormula = ""
		For nI = 0 To nEvents - 1
		    sFormula = sFormula & "+" & mCellFrequency (nJ, nI)
		Next nI
		sFormula = "=" & Right (sFormula, Len (sFormula) - 1)
		oCell.setFormula (sFormula)
		oCell.setPropertyValue ("NumberFormat", nFormatN)
		mCellNJ (nJ) = fnGetLocalRangeName (oCell)
	Next nJ
	oCell = oSheet.getCellByPosition (1 + nGroups, nRow)
	sCells = fnGetLocalRangeName (oSheet.getCellRangeByPosition (1, nRow, nGroups, nRow))
	sFormula = ""
	For nI = 0 To nEvents - 1
	    sFormula = sFormula & "+" & mCellNI (nI)
	Next nI
	sFormula = "=" & Right (sFormula, Len (sFormula) - 1)
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatN)
	sCellN = fnGetLocalRangeName (oCell)
	' Shows the proportions
	nRow = nRow - nEvents * 2 - 1
	For nI = 0 To nEvents - 1
	    nRow = nRow + 2
        For nJ = 0 To nGroups - 1
			oCell = oSheet.getCellByPosition (1 + nJ, nRow)
		    sFormula = "=" & mCellFrequency (nJ, nI) & "/" & mCellNJ (nJ)
		    oCell.setFormula (sFormula)
			oCell.setPropertyValue ("NumberFormat", nFormatPct)
			mCellProportion (nJ, nI) = fnGetLocalRangeName (oCell)
		Next nJ
		oCell = oSheet.getCellByPosition (1 + nGroups, nRow)
	    sFormula = "=" & mCellNI (nI) & "/" & sCellN
	    oCell.setFormula (sFormula)
		oCell.setPropertyValue ("NumberFormat", nFormatPct)
		mCellPI (nI) = fnGetLocalRangeName (oCell)
	Next nI
	' Shows the total
	nRow = nRow + 2
	For nJ = 0 To nGroups - 1
		oCell = oSheet.getCellByPosition (1 + nJ, nRow)
		sFormula = ""
		For nI = 0 To nEvents - 1
		    sFormula = sFormula & "+" & mCellProportion (nJ, nI)
		Next nI
		sFormula = "=" & Right (sFormula, Len (sFormula) - 1)
		oCell.setFormula (sFormula)
		oCell.setPropertyValue ("NumberFormat", nFormatPct)
	Next nJ
	oCell = oSheet.getCellByPosition (1 + nGroups, nRow)
	sCells = fnGetLocalRangeName (oSheet.getCellRangeByPosition (1, nRow, nGroups, nRow))
	sFormula = ""
	For nI = 0 To nEvents - 1
	    sFormula = sFormula & "+" & mCellPI (nI)
	Next nI
	sFormula = "=" & Right (sFormula, Len (sFormula) - 1)
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatPct)
	
	oCells = oSheet.getCellRangeByPosition (1, nRow - nEvents * 2 - 1, nGroups, nRow - nEvents * 2 - 1)
	sCellsRow = fnGetLocalRangeName (oCells)
	oCells = oSheet.getCellRangeByPosition (1, nRow - nEvents * 2 - 1, 1, nRow - 2)
	sCellsColumn = fnGetLocalRangeName (oCells)
	
	' Draws the table borders.
	oCells = oSheet.getCellByPosition (0, nRow - nEvents * 2 - 2)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellRangeByPosition (1, nRow  - nEvents * 2 - 2, nGroups + 1, nRow - nEvents * 2 - 2)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellRangeByPosition (0, nRow - nEvents * 2 - 1, 0, nRow - 2)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells = oSheet.getCellByPosition (0, nRow - 1)
	oCells.setPropertyValue ("TopBorder", aBorderSingle)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	oCells = oSheet.getCellRangeByPosition (1, nRow - 1, nGroups + 1, nRow - 1)
	oCells.setPropertyValue ("TopBorder", aBorderSingle)
	oCells = oSheet.getCellRangeByPosition (1, nRow, nGroups + 1, nRow)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	
	' The Chi-square test
	nRow = nRow + 2
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Chi-Square Test")
	oCell.setPropertyValue ("CellStyle", "Result2")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 3, nRow)
	oCells.merge (True)
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Test")
	oCell = oSheet.getCellByPosition (1, nRow)
	oCell.setString ("χ2")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.collapseToEnd
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharEscapement", 33)
	oCursor.setPropertyValue ("CharEscapementHeight", 58)
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
	
	' The test result.
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Pearson’s Chi-Square")
	oCell = oSheet.getCellByPosition (1, nRow)
	sFormula = ""
	For nI = 0 To nEvents - 1
	    For nJ = 0 To nGroups - 1
	        sFormula = sFormula & "+POWER(" & mCellFrequency (nJ, nI) & ";2)/(" & mCellNI (nI) & "*" & mCellNJ (nJ) & ")"
	    Next nJ
	Next nI
	sFormula = "=" & sCellN & "*(" & Right (sFormula, Len (sFormula) - 1) & "-1)"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellChi2 = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (2, nRow)
	sFormula = "=(COUNT(" & sCellsRow & ")-1)*(COUNT(" & sCellsColumn & ")/2-1)"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatN)
	sCellDF = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (3, nRow)
	sFormula = "=CHIDIST(" & sCellChi2 & ";" & sCellDF & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatP)
	
	' The foot notes of the test.
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Note: *: p<.05, **: p<.01" & Chr (10) & _
		"H0: P1=P2=…PN=P (the proportions of the events in each group are the same)." & Chr (10) & _
		"H1: The above is false (the proportions of the events in each group are different) if the probability (p) is small enough.")
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
	nPos = InStr (1, sNotes, "P", 0)
	Do While nPos <> 0
		oCursor.gotoStart (False)
		oCursor.goRight (nPos - 1, False)
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.collapseToEnd
		oCursor.goRight (1, True)
		If oCursor.getString <> " " Then
			oCursor.setPropertyValue ("CharEscapement", -33)
			oCursor.setPropertyValue ("CharEscapementHeight", 58)
		End If
		nPos = InStr (nPos + 1, sNotes, "P", 0)
	Loop
	nPos = InStr (sNotes, "PN")
	oCursor.gotoStart (False)
	oCursor.goRight (nPos, False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
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
	oCells = oSheet.getCellRangeByPosition (1, nRow - 2, 3, nRow - 2)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellByPosition (0, nRow - 1)
	oCells.setPropertyValue ("TopBorder", aBorderSingle)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	oCells = oSheet.getCellRangeByPosition (1, nRow - 1, 3, nRow - 1)
	oCells.setPropertyValue ("TopBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	
	' The posteriori comparison
	nRow = nRow + 2
	If nEvents = 2 Then
		oCell = oSheet.getCellByPosition (0, nRow)
		oCell.setString ("Posteriori Comparison")
		oCell.setPropertyValue ("CellStyle", "Result2")
		oCells = oSheet.getCellRangeByPosition (0, nRow, 4, nRow)
		oCells.merge (True)
		nRow = nRow + 1
		oCell = oSheet.getCellByPosition (0, nRow)
		oCell.setString ("j1")
		oCursor = oCell.createTextCursor
		oCursor.gotoStart (False)
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.collapseToEnd
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharEscapement", -33)
		oCursor.setPropertyValue ("CharEscapementHeight", 58)
		oCell = oSheet.getCellByPosition (1, nRow)
		oCell.setString ("j2")
		oCursor = oCell.createTextCursor
		oCursor.gotoStart (False)
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.collapseToEnd
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharEscapement", -33)
		oCursor.setPropertyValue ("CharEscapementHeight", 58)
		oCell = oSheet.getCellByPosition (2, nRow)
		oCell.setString ("Pj1-Pj2")
		oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
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
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharEscapement", -52)
		oCursor.setPropertyValue ("CharEscapementHeight", 34)
		oCursor.collapseToEnd
		oCursor.goRight (1, False)
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
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharEscapement", -52)
		oCursor.setPropertyValue ("CharEscapementHeight", 34)
		oCell = oSheet.getCellByPosition (3, nRow)
		oCell.setString ("χ2")
		oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
		oCursor = oCell.createTextCursor
		oCursor.gotoStart (False)
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.collapseToEnd
		oCursor.gotoEnd (True)
		oCursor.setPropertyValue ("CharEscapement", 33)
		oCursor.setPropertyValue ("CharEscapementHeight", 58)
		oCell = oSheet.getCellByPosition (4, nRow)
		oCell.setString ("p")
		oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
		oCursor = oCell.createTextCursor
		oCursor.gotoStart (False)
		oCursor.gotoEnd (True)
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	Else
		oCell = oSheet.getCellByPosition (0, nRow)
		oCell.setString ("Posteriori Comparison")
		oCell.setPropertyValue ("CellStyle", "Result2")
		oCells = oSheet.getCellRangeByPosition (0, nRow, 5, nRow)
		oCells.merge (True)
		nRow = nRow + 1
		oCell = oSheet.getCellByPosition (0, nRow)
		oCell.setString ("Event")
		oCell = oSheet.getCellByPosition (1, nRow)
		oCell.setString ("j1")
		oCursor = oCell.createTextCursor
		oCursor.gotoStart (False)
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.collapseToEnd
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharEscapement", -33)
		oCursor.setPropertyValue ("CharEscapementHeight", 58)
		oCell = oSheet.getCellByPosition (2, nRow)
		oCell.setString ("j2")
		oCursor = oCell.createTextCursor
		oCursor.gotoStart (False)
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.collapseToEnd
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharEscapement", -33)
		oCursor.setPropertyValue ("CharEscapementHeight", 58)
		oCell = oSheet.getCellByPosition (3, nRow)
		oCell.setString ("Pj1-Pj2")
		oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
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
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharEscapement", -52)
		oCursor.setPropertyValue ("CharEscapementHeight", 34)
		oCursor.collapseToEnd
		oCursor.goRight (1, False)
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
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharEscapement", -52)
		oCursor.setPropertyValue ("CharEscapementHeight", 34)
		oCell = oSheet.getCellByPosition (4, nRow)
		oCell.setString ("χ2")
		oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
		oCursor = oCell.createTextCursor
		oCursor.gotoStart (False)
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.collapseToEnd
		oCursor.gotoEnd (True)
		oCursor.setPropertyValue ("CharEscapement", 33)
		oCursor.setPropertyValue ("CharEscapementHeight", 58)
		oCell = oSheet.getCellByPosition (5, nRow)
		oCell.setString ("p")
		oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
		oCursor = oCell.createTextCursor
		oCursor.gotoStart (False)
		oCursor.gotoEnd (True)
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	End If
	
	' The comparison between groups
	nRow = nRow + 1
	If nEvents = 2 Then
		For nJ1 = 0 To nGroups - 1
			oCell = oSheet.getCellByPosition (0, nRow)
			sFormula = "=" & mCellLabelColomn (nJ1)
			oCell.setFormula (sFormula)
			oCell.setPropertyValue ("VertJustify", com.sun.star.table.CellVertJustify.TOP)
			oCells = oSheet.getCellRangeByPosition (0, nRow, 0, nRow + (nGroups - 1) - 1)
			oCells.merge (True)
			For nJ2 = 0 To nGroups - 1
				If nJ1 <> nJ2 Then
					oCell = oSheet.getCellByPosition (1, nRow)
					sFormula = "=" & mCellLabelColomn (nJ2)
					oCell.setFormula (sFormula)
					oCell = oSheet.getCellByPosition (2, nRow)
					sFormula = "=" & mCellProportion (nJ1, 0) & "-" & mCellProportion (nJ2, 0)
					oCell.setFormula (sFormula)
					oCell.setPropertyValue ("NumberFormat", nFormatPct)
					sCell = fnGetLocalRangeName (oCell)
					sSE2 = "(" & mCellProportion (nJ1, 0) & "*(1-" & mCellProportion (nJ1, 0) & "))/" & mCellNJ (nJ1) & _
						"+(" & mCellProportion (nJ2, 0) & "*(1-" & mCellProportion (nJ2, 0) & "))/" & mCellNJ (nJ2)
					oCell = oSheet.getCellByPosition (3, nRow)
					sFormula = "=IF(" & sSE2 & "=0;""(N/A)"";POWER(" & sCell & ";2)/(" & sSE2 & "))"
					oCell.setFormula (sFormula)
					oCell.setPropertyValue ("NumberFormat", nFormatF)
					oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
					sCellChi2 = fnGetLocalRangeName (oCell)
					oCell = oSheet.getCellByPosition (4, nRow)
					sFormula = "=IF(" & sSE2 & "=0;""(N/A)"";CHIDIST(" & sCellChi2 & ";" & sCellDF & "))"
					oCell.setFormula (sFormula)
					oCell.setPropertyValue ("NumberFormat", nFormatP)
					oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
					nRow = nRow + 1
				End If
			Next nJ2
		Next nJ1
	Else
		For nI = 0 To nEvents - 1
			oCell = oSheet.getCellByPosition (0, nRow)
			sFormula = "=" & mCellLabelRow (nI)
			oCell.setFormula (sFormula)
			oCell.setPropertyValue ("VertJustify", com.sun.star.table.CellVertJustify.TOP)
			oCells = oSheet.getCellRangeByPosition (0, nRow, 0, nRow + nGroups * (nGroups - 1) - 1)
			oCells.merge (True)
			For nJ1 = 0 To nGroups - 1
				oCell = oSheet.getCellByPosition (1, nRow)
				sFormula = "=" & mCellLabelColomn (nJ1)
				oCell.setFormula (sFormula)
				oCell.setPropertyValue ("VertJustify", com.sun.star.table.CellVertJustify.TOP)
				oCells = oSheet.getCellRangeByPosition (1, nRow, 1, nRow + (nGroups - 1) - 1)
				oCells.merge (True)
				For nJ2 = 0 To nGroups - 1
					If nJ1 <> nJ2 Then
						oCell = oSheet.getCellByPosition (2, nRow)
						sFormula = "=" & mCellLabelColomn (nJ2)
						oCell.setFormula (sFormula)
						oCell = oSheet.getCellByPosition (3, nRow)
						sFormula = "=" & mCellProportion (nJ1, nI) & "-" & mCellProportion (nJ2, nI)
						oCell.setFormula (sFormula)
						oCell.setPropertyValue ("NumberFormat", nFormatPct)
						sCell = fnGetLocalRangeName (oCell)
						sSE2 = "(" & mCellProportion (nJ1, nI) & "*(1-" & mCellProportion (nJ1, nI) & "))/" & mCellNJ (nJ1) & _
							"+(" & mCellProportion (nJ2, nI) & "*(1-" & mCellProportion (nJ2, nI) & "))/" & mCellNJ (nJ2)
						oCell = oSheet.getCellByPosition (4, nRow)
						sFormula = "=IF(" & sSE2 & "=0;""(N/A)"";POWER(" & sCell & ";2)/(" & sSE2 & "))"
						oCell.setFormula (sFormula)
						oCell.setPropertyValue ("NumberFormat", nFormatF)
						oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
						sCellChi2 = fnGetLocalRangeName (oCell)
						oCell = oSheet.getCellByPosition (5, nRow)
						sFormula = "=IF(" & sSE2 & "=0;""(N/A)"";CHIDIST(" & sCellChi2 & ";" & sCellDF & "))"
						oCell.setFormula (sFormula)
						oCell.setPropertyValue ("NumberFormat", nFormatP)
						oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
						nRow = nRow + 1
					End If
				Next nJ2
			Next nJ1
		Next nI
	End If
	nRow = nRow - 1
	
	' The foot notes of the test.
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Note: *: p<.05, **: p<.01" & Chr (10) & _
		"H0: Pj1=Pj2 (the proportions of the event in the two groups are the same)." & Chr (10) & _
		"H1: Pj1≠Pj2 (the proportions of the event in the two groups are different) if the probability (p) is small enough.")
	oCell.setPropertyValue ("IsTextWrapped", True)
	If nEvents = 2 Then
		oCells = oSheet.getCellRangeByPosition (0, nRow, 4, nRow)
	Else
		oCells = oSheet.getCellRangeByPosition (0, nRow, 5, nRow)
	End If
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
	nPos = InStr (sNotes, "Pj")
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
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.collapseToEnd
		oCursor.goRight (1, True)
		oCursor.setPropertyValue ("CharEscapement", -52)
		oCursor.setPropertyValue ("CharEscapementHeight", 34)
		nPos = InStr (nPos + 1, sNotes, "Pj")
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
	If nEvents = 2 Then
		nStartRow = nRow - nGroups * (nGroups - 1) - 1
		oCells = oSheet.getCellByPosition (0, nStartRow)
		oCells.setPropertyValue ("TopBorder", aBorderDouble)
		oCells.setPropertyValue ("BottomBorder", aBorderSingle)
		oCells = oSheet.getCellByPosition (1, nStartRow)
		oCells.setPropertyValue ("TopBorder", aBorderDouble)
		oCells.setPropertyValue ("RightBorder", aBorderSingle)
		oCells.setPropertyValue ("BottomBorder", aBorderSingle)
		oCells = oSheet.getCellRangeByPosition (2, nStartRow, 4, nStartRow)
		oCells.setPropertyValue ("TopBorder", aBorderDouble)
		oCells.setPropertyValue ("BottomBorder", aBorderSingle)
		oCells = oSheet.getCellRangeByPosition (1, nStartRow + 1, 1, nRow - 2)
		oCells.setPropertyValue ("RightBorder", aBorderSingle)
		oCells = oSheet.getCellByPosition (0, nRow - 1 - (nGroups - 1) + 1)
		oCells.setPropertyValue ("BottomBorder", aBorderDouble)
		oCells = oSheet.getCellByPosition (1, nRow - 1)
		oCells.setPropertyValue ("RightBorder", aBorderSingle)
		oCells.setPropertyValue ("BottomBorder", aBorderDouble)
		oCells = oSheet.getCellRangeByPosition (2, nRow - 1, 4, nRow - 1)
		oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	Else
		nStartRow = nRow - nEvents * nGroups * (nGroups - 1) - 1
		oCells = oSheet.getCellRangeByPosition (0, nStartRow, 1, nStartRow)
		oCells.setPropertyValue ("TopBorder", aBorderDouble)
		oCells.setPropertyValue ("BottomBorder", aBorderSingle)
		oCells = oSheet.getCellByPosition (2, nStartRow)
		oCells.setPropertyValue ("TopBorder", aBorderDouble)
		oCells.setPropertyValue ("RightBorder", aBorderSingle)
		oCells.setPropertyValue ("BottomBorder", aBorderSingle)
		oCells = oSheet.getCellRangeByPosition (3, nStartRow, 5, nStartRow)
		oCells.setPropertyValue ("TopBorder", aBorderDouble)
		oCells.setPropertyValue ("BottomBorder", aBorderSingle)
		oCells = oSheet.getCellRangeByPosition (2, nStartRow + 1, 2, nRow - 2)
		oCells.setPropertyValue ("RightBorder", aBorderSingle)
		oCells = oSheet.getCellByPosition (0, nRow - 1 - nGroups * (nGroups - 1) + 1)
		oCells.setPropertyValue ("BottomBorder", aBorderDouble)
		oCells = oSheet.getCellByPosition (1, nRow - 1 - (nGroups - 1) + 1)
		oCells.setPropertyValue ("BottomBorder", aBorderDouble)
		oCells = oSheet.getCellByPosition (2, nRow - 1)
		oCells.setPropertyValue ("RightBorder", aBorderSingle)
		oCells.setPropertyValue ("BottomBorder", aBorderDouble)
		oCells = oSheet.getCellRangeByPosition (3, nRow - 1, 5, nRow - 1)
		oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	End If
End Sub
