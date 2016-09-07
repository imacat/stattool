' _1CorRel: The macros to for generating the report of the Pearson’s correlation coefficient
'   by imacat <imacat@mail.imacat.idv.tw>, 2016-08-10

Option Explicit

' subRunCorrelation: Runs the Pearson’s correlation coefficient.
Sub subRunCorrelation As Object
	Dim oRange As Object
	Dim mLabels () As String, nI As Integer, mSelected (0) As Integer
	Dim oDialogModel As Object, oDialog As Object, nResult As Integer
	Dim oTextModel As Object, oListModel1 As Object, oListModel2 As Object
	Dim oButtonModel As Object
	Dim nColumn As Integer, oRange1 As Object, oRange2 As Object
	Dim oSheets As Object, sSheetName As String, sExisted As String
	Dim oSheet As Object
	
	' Asks the user for the data range
	oRange = fnAskDataRange (ThisComponent)
	If IsNull (oRange) Then
		Exit Sub
	End If
	ReDim mLabels (oRange.getColumns.getCount - 1) As String
	For nI = 0 To oRange.getColumns.getCount - 1
		mLabels (nI) = oRange.getCellByPosition (nI, 0).getString
	Next nI
	
	' Creates a dialog
	oDialogModel = CreateUnoService ( _
		"com.sun.star.awt.UnoControlDialogModel")
	oDialogModel.setPropertyValue ("PositionX", 200)
	oDialogModel.setPropertyValue ("PositionY", 200)
	oDialogModel.setPropertyValue ("Height", 80)
	oDialogModel.setPropertyValue ("Width", 95)
	oDialogModel.setPropertyValue ("Title", "Step 2/2: Specify the data")
	
	' Adds the prompt.
	oTextModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlFixedTextModel")
	oTextModel.setPropertyValue ("PositionX", 5)
	oTextModel.setPropertyValue ("PositionY", 5)
	oTextModel.setPropertyValue ("Height", 10)
	oTextModel.setPropertyValue ("Width", 85)
	oTextModel.setPropertyValue ("Label", "First score column:")
	oTextModel.setPropertyValue ("MultiLine", True)
	oTextModel.setPropertyValue ("TabIndex", 1)
	oDialogModel.insertByName ("txtPromptGroup", oTextModel)
	
	' Adds the drop down list
	oListModel1 = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlListBoxModel")
	oListModel1.setPropertyValue ("PositionX", 5)
	oListModel1.setPropertyValue ("PositionY", 15)
	oListModel1.setPropertyValue ("Height", 10)
	oListModel1.setPropertyValue ("Width", 85)
	oListModel1.setPropertyValue ("Dropdown", True)
	oListModel1.setPropertyValue ("StringItemList", mLabels)
	mSelected (0) = 0
	oListModel1.setPropertyValue ("SelectedItems", mSelected)
	oDialogModel.insertByName ("lstGroup", oListModel1)
	
	' Adds the prompt.
	oTextModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlFixedTextModel")
	oTextModel.setPropertyValue ("PositionX", 5)
	oTextModel.setPropertyValue ("PositionY", 30)
	oTextModel.setPropertyValue ("Height", 10)
	oTextModel.setPropertyValue ("Width", 85)
	oTextModel.setPropertyValue ("Label", "Second score column:")
	oTextModel.setPropertyValue ("MultiLine", True)
	oTextModel.setPropertyValue ("TabIndex", 1)
	oDialogModel.insertByName ("txtPromptScore", oTextModel)
	
	' Adds the drop down list
	oListModel2 = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlListBoxModel")
	oListModel2.setPropertyValue ("PositionX", 5)
	oListModel2.setPropertyValue ("PositionY", 40)
	oListModel2.setPropertyValue ("Height", 10)
	oListModel2.setPropertyValue ("Width", 85)
	oListModel2.setPropertyValue ("Dropdown", True)
	oListModel2.setPropertyValue ("StringItemList", mLabels)
	mSelected (0) = 1
	oListModel2.setPropertyValue ("SelectedItems", mSelected)
	oDialogModel.insertByName ("lstScore", oListModel2)
	
	' Adds the buttons.
	oButtonModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlButtonModel")
	oButtonModel.setPropertyValue ("PositionX", 5)
	oButtonModel.setPropertyValue ("PositionY", 60)
	oButtonModel.setPropertyValue ("Height", 15)
	oButtonModel.setPropertyValue ("Width", 40)
	oButtonModel.setPropertyValue ("PushButtonType", _
		com.sun.star.awt.PushButtonType.CANCEL)
	oDialogModel.insertByName ("btnClose", oButtonModel)
	
	oButtonModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlButtonModel")
	oButtonModel.setPropertyValue ("PositionX", 50)
	oButtonModel.setPropertyValue ("PositionY", 60)
	oButtonModel.setPropertyValue ("Height", 15)
	oButtonModel.setPropertyValue ("Width", 40)
	oButtonModel.setPropertyValue ("PushButtonType", _
		com.sun.star.awt.PushButtonType.OK)
	oButtonModel.setPropertyValue ("DefaultButton", True)
	oDialogModel.insertByName ("btnOK", oButtonModel)
	
	' Adds the dialog model to the control and runs it.
	oDialog = CreateUnoService ("com.sun.star.awt.UnoControlDialog")
	oDialog.setModel (oDialogModel)
	oDialog.setVisible (True)
	nResult = oDialog.execute
	oDialog.dispose
	
	' Cancelled
	If nResult = 0 Then
		Exit Sub
	End If
	
	nColumn = oListModel1.getPropertyValue ("SelectedItems") (0)
	oRange1 = oRange.getCellRangeByPosition ( _
		nColumn, 0, nColumn, oRange.getRows.getCount - 1)
	nColumn = oListModel2.getPropertyValue ("SelectedItems") (0)
	oRange2 = oRange.getCellRangeByPosition ( _
		nColumn, 0, nColumn, oRange.getRows.getCount - 1)
	
	' Checks the existing report
	oSheets = ThisComponent.getSheets
	sSheetName = oRange1.getSpreadsheet.getName
	sExisted = ""
	If oSheets.hasByName (sSheetName & "_correl") Then
		sExisted = sExisted & ", """ & sSheetName & "_correl"""
	End If
	If sExisted <> "" Then
		sExisted = Right (sExisted, Len (sExisted) - 2)
		If InStr (sExisted, ",") > 0 Then
			sExisted = "Spreadsheets " & sExisted & " exist.  Overwrite?"
		Else
			sExisted = "Spreadsheet " & sExisted & " exists.  Overwrite?"
		End If
		nResult = MsgBox(sExisted, MB_YESNO + MB_DEFBUTTON2 + MB_ICONQUESTION)
		If nResult = IDNO Then
			Exit Sub
		End If
		
		' Drops the existing report
		If oSheets.hasByName (sSheetName & "_correl") Then
			oSheets.removeByname (sSheetName & "_correl")
		End If
	End If
	
	' Reports the paired T-test.
	subReportCorrelation (ThisComponent, oRange1, oRange2)
	
	' Makes the report sheet active.
	oSheet = oSheets.getByName (sSheetName & "_correl")
	ThisComponent.getCurrentController.setActiveSheet (oSheet)
End Sub

' subTestCorrelation: Tests the Pearson’s correlation coefficient report
Sub subTestCorrelation
	Dim oDoc As Object, oSheets As Object, sSheetName As String
	Dim oSheet As Object, oXRange As Object, oYRange As Object
	
	oDoc = fnFindStatsTestDocument
	If IsNull (oDoc) Then
		MsgBox "Cannot find statstest.ods in the opened documents."
		Exit Sub
	End If
	
	sSheetName = "correl"
	oSheets = oDoc.getSheets
	If Not oSheets.hasByName (sSheetName) Then
		MsgBox "Data sheet """ & sSheetName & """ not found"
		Exit Sub
	End If
	If oSheets.hasByName (sSheetName & "_correl") Then
		oSheets.removeByName (sSheetName & "_correl")
	End If
	oSheet = oSheets.getByName (sSheetName)
	oXRange = oSheet.getCellRangeByName ("B3:B13")
	oYRange = oSheet.getCellRangeByName ("C3:C13")
	subReportCorrelation (oDoc, oXRange, oYRange)
End Sub

' subReportCorrelation: Reports the Pearson’s correlation coefficient
Sub subReportCorrelation (oDoc As Object, oDataXRange As Object, oDataYRange As Object)
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
	Dim sCellYLabel As String, sCellsYData As String
	Dim sCellN As String, sCellR As String
	
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
	sFormula = "=TDIST(" & sCellR & "*SQRT((" & sCellN & "-2)/(1-" & sCellR & "*" & sCellR & "))" & ";" & sCellN & "-2;2)"
	oCell.setFormula (sFormula)
	oCell.setPropertyValue ("NumberFormat", nFormatP)
	
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
End Sub
