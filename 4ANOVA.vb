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

' 4ANOVA: The macros to for generating the report of ANOVA (Analyze of Variances)
'   by imacat <imacat@mail.imacat.idv.tw>, 2016-08-31

Option Explicit

' subRunANOVA: Runs the ANOVA (Analyze of Variances).
Sub subRunANOVA As Object
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
		"&10.Dlg2SpecData.txtPrompt1.Label3ITTest", _
		"&11.Dlg2SpecData.txtPrompt2.Label3ITTest")
	If IsNull (mRanges) Then
		Exit Sub
	End If
	
	' Checks the existing report
	oSheets = ThisComponent.getSheets
	sSheetName = oRange.getSpreadsheet.getName
	sExisted = ""
	If oSheets.hasByName (sSheetName & "_anova") Then
		sExisted = sExisted & ", """ & sSheetName & "_anova"""
	End If
	If oSheets.hasByName (sSheetName & "_anovatmp") Then
		sExisted = sExisted & ", """ & sSheetName & "_anovatmp"""
	End If
	If sExisted <> "" Then
		sExisted = Right (sExisted, Len (sExisted) - 2)
		If InStr (sExisted, ",") > 0 Then
			sExisted = "Spreadsheets " & sExisted & " exist.  Overwrite?"
		Else
			sExisted = "Spreadsheet " & sExisted & " exists.  Overwrite?"
		End If
		nResult = MsgBox (sExisted, MB_YESNO + MB_DEFBUTTON2 + MB_ICONQUESTION)
		If nResult = IDNO Then
			Exit Sub
		End If
		' Drops the existing report
		If oSheets.hasByName (sSheetName & "_anova") Then
			oSheets.removeByname (sSheetName & "_anova")
		End If
		If oSheets.hasByName (sSheetName & "_anovatmp") Then
			oSheets.removeByname (sSheetName & "_anovatmp")
		End If
	End If
	
	' Reports the ANOVA (Analyze of Variances)
	subReportANOVA (ThisComponent, mRanges (0), mRanges (1))
	
	' Makes the report sheet active.
	oSheet = oSheets.getByName (sSheetName & "_anova")
	ThisComponent.getCurrentController.setActiveSheet (oSheet)
End Sub

' subTestANOVA: Tests the ANOVA (Analyze of Variances) report
Sub subTestANOVA
	Dim oDoc As Object, oSheets As Object, sSheetName As String
	Dim oSheet As Object, oLabelColumn As Object, oScoreColumn As Object
	
	oDoc = fnFindStatsTestDocument
	If IsNull (oDoc) Then
		MsgBox "Cannot find statstest.ods in the opened documents."
		Exit Sub
	End If
	
	sSheetName = "anova"
	oSheets = oDoc.getSheets
	If Not oSheets.hasByName (sSheetName) Then
		MsgBox "Data sheet """ & sSheetName & """ not found"
		Exit Sub
	End If
	If oSheets.hasByName (sSheetName & "_anova") Then
		oSheets.removeByName (sSheetName & "_anova")
	End If
	If oSheets.hasByName (sSheetName & "_anovatmp") Then
		oSheets.removeByName (sSheetName & "_anovatmp")
	End If
	oSheet = oSheets.getByName (sSheetName)
	oLabelColumn = oSheet.getCellRangeByName ("A13:A35")
	oScoreColumn = oSheet.getCellRangeByName ("B13:B35")
	subReportANOVA (oDoc, oLabelColumn, oScoreColumn)
End Sub

' subReportANOVA: Reports the ANOVA (Analyze of Variances)
Sub subReportANOVA (oDoc As Object, oLabelColumn As Object, oScoreColumn As Object)
	Dim oSheets As Object, sSheetName As String
	Dim nI As Integer, nJ As Integer
	Dim mNames () As String, nSheetIndex As Integer
	Dim oSheet As Object, oColumns As Object, nRow As Integer, nStartRow As Integer
	Dim oCell As Object, oCells As Object, oCursor As Object, oTempDataRange As Object
	Dim nN As Integer, sFormula As String
	Dim sNotes As String, nPos As Integer
	Dim nFormatN As Integer, nFormatF As Integer, nFormatP As Integer
	Dim aBorderSingle As New com.sun.star.table.BorderLine
	Dim aBorderDouble As New com.sun.star.table.BorderLine
	
	Dim nGroups As Integer
	Dim mCellLabel () As String, mCellsData () As String
	Dim mCellN () As String, mCellMean () As String, mCellS () As String
	Dim sCellsData As String
	Dim sCellF As String, sCellDFB As String, sCellDFW As String
	Dim sCellsN As String, sCellN As String, sCellS As String
	Dim sCellSSB As String, sCellSSW As String, sCellSST As String
	Dim sCellMSB As String, sCellMSW As String
	Dim sCellMeanDiff As String
	
	oSheets = oDoc.getSheets
	sSheetName = oLabelColumn.getSpreadsheet.getName
	mNames = oSheets.getElementNames
	For nI = 0 To UBound (mNames)
		If mNames (nI) = sSheetName Then
			nSheetIndex = nI
		End If
	Next nI
	
	oSheets.insertNewByName (sSheetName & "_anovatmp", nSheetIndex + 1)
	oSheet = oSheets.getByName (sSheetName & "_anovatmp")
	oTempDataRange = fnCollectANOVAData (oSheet, oLabelColumn, oScoreColumn)
	nGroups = oTempDataRange.getColumns.getCount / 3
	
	oSheets.insertNewByName (sSheetName & "_anova", nSheetIndex + 1)
	oSheet = oSheets.getByName (sSheetName & "_anova")
	
	ReDim mCellLabel (nGroups - 1) As String, mCellsData (nGroups - 1) As String
	ReDim mCellN (nGroups - 1) As String, mCellMean (nGroups - 1) As String
	ReDim mCellS (nGroups - 1) As String
	
	For nI = 0 To nGroups - 1
		mCellLabel (nI) = fnGetRangeName (oTempDataRange.getCellByPosition (nI, 0))
		nN = oTempDataRange.getCellByPosition (nI, oTempDataRange.getRows.getCount - 3).getValue
		oCells = oTempDataRange.getCellRangeByPosition (nI, 1, nI, nN)
		mCellsData (nI) = fnGetRangeName (oCells)
	Next nI
	
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
	oColumns.getByIndex (5).setPropertyValue ("Width", 2080)
	oColumns.getByIndex (6).setPropertyValue ("Width", 2080)
	
	nRow = -2
	
	' Group description
	nRow = nRow + 2
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Group Description")
	oCell.setPropertyValue ("CellStyle", "Result2")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 5, nRow)
	oCells.merge (True)
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Group")
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
	
	' Show each group
	sCellsData = ""
	For nI = 0 To nGroups - 1
		nRow = nRow + 1
		oCell = oSheet.getCellByPosition (0, nRow)
		sFormula = "=" & mCellLabel (nI)
		oCell.setFormula (sFormula)
		mCellLabel (nI) = fnGetLocalRangeName (oCell)
		oCells = oSheet.getCellRangeByPosition (0, nRow, 1, nRow)
		oCells.merge (True)
		oCell = oSheet.getCellByPosition (2, nRow)
		sFormula = "=COUNT(" & mCellsData (nI) & ")"
		oCell.setFormula (sFormula)
		oCell.setPropertyValue ("NumberFormat", nFormatN)
		mCellN (nI) = fnGetLocalRangeName (oCell)
		oCell = oSheet.getCellByPosition (3, nRow)
		sFormula = "=AVERAGE(" & mCellsData (nI) & ")"
		oCell.setFormula (sFormula)
		oCell.setPropertyValue ("NumberFormat", nFormatF)
		mCellMean (nI) = fnGetLocalRangeName (oCell)
		oCell = oSheet.getCellByPosition (4, nRow)
		sFormula = "=STDEV(" & mCellsData (nI) & ")"
		oCell.setFormula (sFormula)
		oCell.setPropertyValue ("NumberFormat", nFormatF)
		mCellS (nI) = fnGetLocalRangeName (oCell)
		oCell = oSheet.getCellByPosition (5, nRow)
		sFormula = "=" & mCellS (nI) & "/SQRT(" & mCellN (nI) & ")"
		oCell.setFormula (sFormula)
		oCell.setPropertyValue ("NumberFormat", nFormatF)
		sCellsData = sCellsData & ";" & mCellsData (nI)
	Next nI
	sCellsData = Right (sCellsData, Len (sCellsData) - 1)
	' Shows the total
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Total")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 1, nRow)
	oCells.merge (True)
	oCell = oSheet.getCellByPosition (2, nRow)
	sFormula = "=COUNT(" & sCellsData & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatN)
	sCellN = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (3, nRow)
	sFormula = "=AVERAGE(" & sCellsData & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	oCell = oSheet.getCellByPosition (4, nRow)
	sFormula = "=STDEV(" & sCellsData & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellS = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (5, nRow)
	sFormula = "=" & sCellS & "/SQRT(" & sCellN & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	
	' Draws the table borders.
	oCells = oSheet.getCellByPosition (0, nRow - nGroups - 1)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellRangeByPosition (1, nRow  - nGroups - 1, 5, nRow - nGroups - 1)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellRangeByPosition (0, nRow  - nGroups, 0, nRow - 1)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells = oSheet.getCellByPosition (0, nRow)
	oCells.setPropertyValue ("TopBorder", aBorderSingle)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	oCells = oSheet.getCellRangeByPosition (1, nRow, 5, nRow)
	oCells.setPropertyValue ("TopBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	
	' Levene's test for homogeneity of variances
	nRow = nRow + 2
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Test for Homogeneity of Variances")
	oCell.setPropertyValue ("CellStyle", "Result2")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 5, nRow)
	oCells.merge (True)
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Test")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 1, nRow)
	oCells.merge (True)
	oCell = oSheet.getCellByPosition (2, nRow)
	oCell.setString ("F")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCell = oSheet.getCellByPosition (3, nRow)
	oCell.setString ("dfb")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.goRight (2, True)
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
	oCell = oSheet.getCellByPosition (4, nRow)
	oCell.setString ("dfw")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.goRight (2, True)
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
	oCell = oSheet.getCellByPosition (5, nRow)
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
	oCells = oSheet.getCellRangeByPosition (0, nRow, 1, nRow)
	oCells.merge (True)
	oCell = oSheet.getCellByPosition (2, nRow)
	sFormula = "=" & fnGetLeveneTest (oTempDataRange)
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellF = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (3, nRow)
	sCellsN = fnGetRangeName (oTempDataRange.getCellRangeByPosition (0, oTempDataRange.getRows.getCount - 3, nGroups - 1, oTempDataRange.getRows.getCount - 3))
	sFormula = "=COUNT(" & sCellsN & ")-1"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatN)
	sCellDFB = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (4, nRow)
	sCellN = fnGetRangeName (oTempDataRange.getCellByPosition (nGroups * 2, oTempDataRange.getRows.getCount - 3))
	sFormula = "=" & sCellN & "-COUNT(" & sCellsN & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatN)
	sCellDFW = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (5, nRow)
	sFormula = "=FDIST(" & sCellF & ";" & sCellDFB & ";" & sCellDFW & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatP)
	
	' The foot notes of the test.
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Note: *: p<.05, **: p<.01" & Chr (10) & _
		"H0: σ1=σ2=…σN (the populations of the groups have the same variances)." & Chr (10) & _
		"H1: ANOVA does not apply (the populations of the groups does not have the same variances) if the probability (p) is small enough.")
	oCell.setPropertyValue ("IsTextWrapped", True)
	oCells = oSheet.getCellRangeByPosition (0, nRow, 5, nRow)
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
	oCursor.goRight (nPos - 1, False)
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
	nPos = InStr (sNotes, "σN")
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
	oCells = oSheet.getCellRangeByPosition (0, nRow, 4, nRow)
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
	oCells = oSheet.getCellRangeByPosition (1, nRow - 2, 5, nRow - 2)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellByPosition (0, nRow - 1)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	oCells = oSheet.getCellRangeByPosition (1, nRow - 1, 5, nRow - 1)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	
	' The ANOVA (analysis of variances)
	nRow = nRow + 2
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("One-way ANOVA (Analysis of Variances)")
	oCell.setPropertyValue ("CellStyle", "Result2")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 6, nRow)
	oCells.merge (True)
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Source of Variation")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 1, nRow)
	oCells.merge (True)
	oCell = oSheet.getCellByPosition (2, nRow)
	oCell.setString ("SS")
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
	oCell.setString ("MS")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCell = oSheet.getCellByPosition (5, nRow)
	oCell.setString ("F")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCell = oSheet.getCellByPosition (6, nRow)
	oCell.setString ("p")
	oCell.setPropertyValue ("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT)
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	
	' Between groups
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Between Groups")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 1, nRow)
	oCells.merge (True)
	oCell = oSheet.getCellByPosition (2, nRow)
	sFormula = "="
	For nI = 0 To nGroups - 1
		sFormula = sFormula & "POWER(SUM(" & mCellsData (nI) & ");2)/" & mCellN (nI) & "+"
	Next nI
	sFormula = Left (sFormula, Len (sFormula) - 1) & "-POWER(SUM(" & sCellsData & ");2)/" & sCellN
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellSSB = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (3, nRow)
	sFormula = "=" & sCellDFB
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatN)
	sCellDFB = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (4, nRow)
	sFormula = "=" & sCellSSB & "/" & sCellDFB
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellMSB = fnGetLocalRangeName (oCell)
	
	' Within groups
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Within Groups")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 1, nRow)
	oCells.merge (True)
	oCell = oSheet.getCellByPosition (2, nRow)
	sFormula = "="
	For nI = 0 To nGroups - 1
		sFormula = sFormula & "(SUMPRODUCT(" & mCellsData (nI) & ";" & mCellsData (nI) & ")-POWER(SUM(" & mCellsData (nI) & ");2)/" & mCellN (nI) & ")+"
	Next nI
	sFormula = Left (sFormula, Len (sFormula) - 1)
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellSSW = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (3, nRow)
	sFormula = "=" & sCellDFW
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatN)
	sCellDFW = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (4, nRow)
	sFormula = "=" & sCellSSW & "/" & sCellDFW
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellMSW = fnGetLocalRangeName (oCell)
	
	nRow = nRow - 1
	oCell = oSheet.getCellByPosition (5, nRow)
	sFormula = "=" & sCellMSB & "/" & sCellMSW
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	sCellF = fnGetLocalRangeName (oCell)
	oCell = oSheet.getCellByPosition (6, nRow)
	sFormula = "=FDIST(" & sCellF & ";" & sCellDFB & ";" & sCellDFW & ")"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatP)
	nRow = nRow + 1
	
	' Total
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Total")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 1, nRow)
	oCells.merge (True)
	oCell = oSheet.getCellByPosition (2, nRow)
	sFormula = "=" & sCellSSB & "+" & sCellSSW
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatF)
	oCell = oSheet.getCellByPosition (3, nRow)
	sFormula = "=" & sCellDFB & "+" & sCellDFW
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatN)
	
	' The foot notes of the test.
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Note: *: p<.05, **: p<.01" & Chr (10) & _
		"H0: μ1=μ2=…μN (the populations of the groups have the same means)." & Chr (10) & _
		"H1: The above is false (the populations of the groups does not have the same means) if the probability (p) is small enough.")
	oCell.setPropertyValue ("IsTextWrapped", True)
	oCells = oSheet.getCellRangeByPosition (0, nRow, 6, nRow)
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
	nPos = InStr (sNotes, "μN")
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
	oCells = oSheet.getCellByPosition (0, nRow - 4)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellRangeByPosition (1, nRow - 4, 6, nRow - 4)
	oCells.setPropertyValue ("TopBorder", aBorderDouble)
	oCells.setPropertyValue ("BottomBorder", aBorderSingle)
	oCells = oSheet.getCellRangeByPosition (0, nRow - 3, 0, nRow - 2)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells = oSheet.getCellByPosition (0, nRow - 1)
	oCells.setPropertyValue ("TopBorder", aBorderSingle)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	oCells = oSheet.getCellRangeByPosition (1, nRow - 1, 6, nRow - 1)
	oCells.setPropertyValue ("TopBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	
	' The post-hoc test between groups with Scheffé's method
	nRow = nRow + 2
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Post-Hoc Test Between Groups with Scheffé's Method")
	oCell.setPropertyValue ("CellStyle", "Result2")
	oCells = oSheet.getCellRangeByPosition (0, nRow, 4, nRow)
	oCells.merge (True)
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Xi")
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.collapseToEnd
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharEscapement", -33)
	oCursor.setPropertyValue ("CharEscapementHeight", 58)
	oCell = oSheet.getCellByPosition (1, nRow)
	oCell.setString ("Xj")
	oCursor = oCell.createTextCursor
	oCursor.gotoStart (False)
	oCursor.goRight (1, True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.collapseToEnd
	oCursor.gotoEnd (True)
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharEscapement", -33)
	oCursor.setPropertyValue ("CharEscapementHeight", 58)
	oCell = oSheet.getCellByPosition (2, nRow)
	oCell.setString ("Xi-Xj")
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
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
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
	oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
	oCursor.setPropertyValue ("CharEscapement", -33)
	oCursor.setPropertyValue ("CharEscapementHeight", 58)
	oCell = oSheet.getCellByPosition (3, nRow)
	oCell.setString ("F")
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
	
	' The tests between groups
	nRow = nRow + 1
	For nI = 0 To nGroups - 1
		oCell = oSheet.getCellByPosition (0, nRow)
		sFormula = "=" & mCellLabel (nI)
		oCell.setFormula (sFormula)
		For nJ = 0 To nGroups - 1
			If nI <> nJ Then
				oCell = oSheet.getCellByPosition (1, nRow)
				sFormula = "=" & mCellLabel (nJ)
				oCell.setFormula (sFormula)
				oCell = oSheet.getCellByPosition (2, nRow)
				sFormula = "=" & mCellMean (nI) & "-" & mCellMean (nJ)
				oCell.setFormula (sFormula)
				oCell.setPropertyValue ("NumberFormat", nFormatF)
				sCellMeanDiff = fnGetLocalRangeName (oCell)
				oCell = oSheet.getCellByPosition (3, nRow)
				sFormula = "=POWER(" & sCellMeanDiff & ";2)/(" & sCellMSW & "*(1/" & mCellN (nI) & "+1/" & mCellN (nJ) & "))"
				oCell.setFormula (sFormula)
				oCell.setPropertyValue ("NumberFormat", nFormatF)
				sCellF = fnGetLocalRangeName (oCell)
				oCell = oSheet.getCellByPosition (4, nRow)
				sFormula = "=FDIST(" & sCellF & "/" & sCellDFB & ";" & sCellDFB & ";" & sCellDFW & ")"
				oCell.setFormula (sFormula)
				oCell.setPropertyValue ("NumberFormat", nFormatP)
				nRow = nRow + 1
			End If
		Next nJ
	Next nI
	nRow = nRow - 1
	
	' The foot notes of the test.
	nRow = nRow + 1
	oCell = oSheet.getCellByPosition (0, nRow)
	oCell.setString ("Note: *: p<.05, **: p<.01" & Chr (10) & _
		"H0: μi=μj (the populations of the two groups have the same means)." & Chr (10) & _
		"H1: μi≠μj (the populations of the two groups have different means) if the probability (p) is small enough.")
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
		oCursor.setPropertyValue ("CharPosture", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureAsian", com.sun.star.awt.FontSlant.ITALIC)
		oCursor.setPropertyValue ("CharPostureComplex", com.sun.star.awt.FontSlant.ITALIC)
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
	nStartRow = nRow - (nGroups * (nGroups - 1)) - 1
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
	oCells = oSheet.getCellByPosition (0, nRow - 1)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	oCells = oSheet.getCellByPosition (1, nRow - 1)
	oCells.setPropertyValue ("RightBorder", aBorderSingle)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
	oCells = oSheet.getCellRangeByPosition (2, nRow - 1, 4, nRow - 1)
	oCells.setPropertyValue ("BottomBorder", aBorderDouble)
End Sub

' fnCollectANOVAData: Collects the data for the ANOVA (Analyze of Variances).
Function fnCollectANOVAData (oReportSheet As Object, oLabelColumn As Object, oScoreColumn As Object) As Object
	Dim nRow As Integer, nColumn As Integer, nI As Integer
	Dim nNRow As Integer, sCellZMean As String, sCellsN As String
	Dim oCell As Object, oCells As Object, oCursor As Object
	Dim sCell As String, sLabel As String, sFormula As String
	Dim nGroups As Integer, sLabels As String
	Dim mLabels () As String, mCellLabel () As String
	Dim mCellsData () As String, mCellMean () As String
	Dim mN () As Integer, mCellsZData () As String
	Dim mCellZMean () As String
	
	sLabels = " "
	nGroups = 0
	For nRow = 1 To oLabelColumn.getRows.getCount - 1
		sLabel = oLabelColumn.getCellByPosition (0, nRow).getString
		If InStr (sLabels, " " & sLabel & " ") = 0 Then
			sLabels = sLabels & sLabel & " "
			nGroups = nGroups + 1
		End If
	Next nRow
	
	ReDim mLabels (nGroups - 1) As String, mCellLabel (nGroups - 1) As String
	ReDim mCellsData (nGroups - 1) As String, mCellMean (nGroups - 1) As String
	ReDim mN (nGroups - 1) As Integer, mCellsZData (nGroups - 1) As String
	ReDim mCellZMean (nGroups - 1) As String
	
	sLabels = " "
	nGroups = 0
	For nRow = 1 To oLabelColumn.getRows.getCount - 1
		oCell = oLabelColumn.getCellByPosition (0, nRow)
		sLabel = oCell.getString
		If InStr (sLabels, " " & sLabel & " ") = 0 Then
			sLabels = sLabels & sLabel & " "
			mLabels (nGroups) = sLabel
			mCellLabel (nGroups) = fnGetRangeName (oCell)
			nGroups = nGroups + 1
		End If
	Next nRow
	
	' The data labels
	For nI = 0 To nGroups - 1
		oCell = oReportSheet.getCellByPosition (nI, 0)
		sFormula = "=" & mCellLabel (nI)
		oCell.setFormula (sFormula)
	Next nI
	
	' The data
	For nI = 0 To nGroups - 1
		mN (nI) = 0
	Next nI
	For nRow = 1 To oLabelColumn.getRows.getCount - 1
		sLabel = oLabelColumn.getCellByPosition (0, nRow).getString
		For nI = 0 To nGroups - 1
			If sLabel = mLabels (nI) Then
				nColumn = nI
				nI = nGroups - 1
			End If
		Next nI
		mN (nColumn) = mN (nColumn) + 1
		sFormula = "=" & fnGetRangeName (oScoreColumn.getCellByPosition (0, nRow))
		oCell = oReportSheet.getCellByPosition (nColumn, mN (nColumn))
		oCell.setFormula (sFormula)
	Next nRow
	
	' Collects the data
	For nI = 0 To nGroups - 1
		mCellsData (nI) = fnGetLocalRangeName (oReportSheet.getCellRangeByPosition (nI, 1, nI, mN (nI)))
	Next nI
	nNRow = 0
	For nI = 0 To nGroups - 1
		If nNRow < mN (nI) Then
			nNRow = mN (nI)
		End If
	Next nI
	nNRow = nNRow + 1
	For nI = 0 To nGroups - 1
		oCell = oReportSheet.getCellByPosition (nI, nNRow)
		sFormula = "=COUNT(" & mCellsData (nI) & ")"
		oCell.setFormula (sFormula)
		oCell = oReportSheet.getCellByPosition (nI, nNRow + 1)
		sFormula = "=AVERAGE(" & mCellsData (nI) & ")"
		oCell.setFormula (sFormula)
		mCellMean (nI) = fnGetLocalRangeName (oCell)
	Next nI
	oCells = oReportSheet.getCellRangeByPosition (0, nNRow, nGroups - 1, nNRow)
	sCellsN = fnGetLocalRangeName (oCells)
	
	' Calculates the Z values
	For nI = 0 To nGroups - 1
		sCell = fnGetLocalRangeName (oReportSheet.getCellByPosition (nI, 0))
		sFormula = "=""Z""&" & sCell
		oCell = oReportSheet.getCellByPosition (nGroups + nI, 0)
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
		For nRow = 1 To mN (nI)
			sCell = fnGetLocalRangeName (oReportSheet.getCellByPosition (nI, nRow))
			sFormula = "=ABS(" & sCell & "-" & mCellMean (nI) & ")"
			oCell = oReportSheet.getCellByPosition (nGroups + nI, nRow)
			oCell.setFormula (sFormula)
		Next nRow
		mCellsZData (nI) = fnGetLocalRangeName (oReportSheet.getCellRangeByPosition (nGroups + nI, 1, nGroups + nI, mN (nI)))
		oCell = oReportSheet.getCellByPosition (nGroups + nI, nNRow + 1)
		sFormula = "=AVERAGE(" & mCellsZData (nI) & ")"
		oCell.setFormula (sFormula)
		mCellZMean (nI) = fnGetLocalRangeName (oCell)
	Next nI
	
	' Calculates the total average
	oCell = oReportSheet.getCellByPosition (nGroups * 2, nNRow)
	sFormula = "=SUM(" & sCellsN & ")"
	oCell.setFormula (sFormula)
	oCell = oReportSheet.getCellByPosition (nGroups * 2, nNRow + 1)
	sFormula = ""
	For nI = 0 To nGroups - 1
		sFormula = sFormula & ";" & mCellsZData (nI)
	Next nI
	sFormula = "=AVERAGE(" & Right (sFormula, Len (sFormula) - 1)
	oCell.setFormula (sFormula)
	sCellZMean = fnGetLocalRangeName (oCell)
	
	' Calculates the difference of the Z values to their means
	For nI = 0 To nGroups - 1
		sCell = fnGetLocalRangeName (oReportSheet.getCellByPosition (nI, 0))
		sFormula = "=""dZ""&" & sCell
		oCell = oReportSheet.getCellByPosition (nGroups * 2 + nI, 0)
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
		For nRow = 1 To mN (nI)
			sCell = fnGetLocalRangeName (oReportSheet.getCellByPosition (nGroups + nI, nRow))
			sFormula = "=" & sCell & "-" & mCellZMean (nI)
			oCell = oReportSheet.getCellByPosition (nGroups * 2 + nI, nRow)
			oCell.setFormula (sFormula)
		Next nRow
	Next nI
	
	' Calculates the difference of the Z means to the total mean
	For nI = 0 To nGroups - 1
		sCell = fnGetLocalRangeName (oReportSheet.getCellByPosition (nGroups + nI, nNRow + 1))
		sFormula = "=" & sCell & "-" & sCellZMean
		oCell = oReportSheet.getCellByPosition (nGroups + nI, nNRow + 2)
		oCell.setFormula (sFormula)
	Next nI
	
	fnCollectANOVAData = oReportSheet.getCellRangeByPosition (0, 0, nGroups * 3 - 1, nNRow + 2)
End Function
