' _0Main: The main module for the statistics macros
'   by imacat <imacat@mail.imacat.idv.tw>, 2016-08-10

Option Explicit

' subMain: The main program
Sub subMain
	BasicLibraries.loadLibrary "XrayTool"
	
	subRunCorrelation
	'subRunPairedTTest
	'subRunIndependentTTest
	'subRunAnova
	'subRunChi2GoodnessOfFit
	'subTestCorrelation
	'subTestPairedTTest
	'subTestIndependentTTest
	'subTestANOVA
	'subTestChi2GoodnessOfFit
	
End Sub

' fnCheckRangeName: Checks the range name and returns the range when
'                   found, or null when not found.
Function fnCheckRangeName (oDoc As Object, sRangeName As String) As Object
	On Error Goto ErrorHandler
	Dim oController As Object, oSheet As Object
	Dim nPos As Integer, sSheetName As String, oRange As Object
	
	oController = oDoc.getCurrentController
	nPos = InStr (sRangeName, ".")
	If nPos = 0 Then
		oSheet = oController.getActiveSheet
	Else
		sSheetName = Left (sRangeName, nPos - 1)
		If Left (sSheetName, 1) = "$" Then
			sSheetName = Right (sSheetName, Len (sSheetName) - 1)
		End If
		oSheet = oDoc.getSheets.getByName (sSheetName)
	End If
	fnCheckRangeName = oSheet.getCellRangeByName (sRangeName)
	
	ErrorHandler:
End Function

' fnQueryFormat: Returns the index of the number format, and creates
'                the number format if required.
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
	Dim oEnum As Object, oDoc As Object, sFile As String
	
	sFile = "/statstest.ods"
	oEnum = StarDesktop.getComponents.createEnumeration
	Do While oEnum.hasMoreElements
		oDoc = oEnum.nextElement
		If oDoc.supportsService ("com.sun.star.document.OfficeDocument") Then
			If Right (oDoc.getLocation, Len (sFile)) = sFile Then
				fnFindStatsTestDocument = oDoc
				Exit Function
			End If
		End If
	Loop
End Function

' fnAskDataRange: Asks the user for the data range, or null when
'                 the user cancelled
Function fnAskDataRange (oDoc As Object) As Object
	Dim oRange As Object, sPrompt As String, sCellsData As String
	
	oRange = fnFindActiveDataRange (oDoc)
	If IsNull (oRange) Then
		sCellsData = ""
	Else
		sCellsData = oRange.getPropertyValue ("AbsoluteName")
	End If
	sPrompt = "Cells with the data:"
	
	' Loop until we get good answer
	Do While sPrompt <> ""
		sCellsData = InputBox (sPrompt, "Step 1/2: Select the data range", sCellsData)
		
		' Cancelled
		If sCellsData = "" Then
			Exit Function
		End If
		
		oRange = fnCheckRangeName (oDoc, sCellsData)
		If IsNull (oRange) Then
			sPrompt = "The range """ & sCellsData & """ does not exist."
		Else
			If oRange.getRows.getCount < 2 Or oRange.getColumns.getCount < 2 Then
				sPrompt = "The range """ & sCellsData & """ is too small (at least 2×2)."
			Else
				sPrompt = ""
				oDoc.getCurrentController.select (oRange)
				fnAskDataRange = oRange
				Exit Function
			End If
		End If
	Loop
End Function

' fnAskDataRange2: Asks the user for the data range, or null when
'                 the user cancelled
Function fnAskDataRange2 (oDoc As Object) As Object
	Dim oRange As Object
	Dim oDialogModel As Object, oDialog As Object, nResult As Integer
	Dim oTextModel As Object, oEditModel As Object
	Dim oButtonModel As Object
	Dim sPrompt As String, sCellsData As String
	
	oRange = fnFindActiveDataRange (oDoc)
	If IsNull (oRange) Then
		sCellsData = ""
	Else
		sCellsData = oRange.getPropertyValue ("AbsoluteName")
	End If
	sPrompt = "Cells with the data:"
	
	' Loop until we finds good data
	Do While sPrompt <> ""
		' Creates a dialog
		oDialogModel = CreateUnoService ( _
			"com.sun.star.awt.UnoControlDialogModel")
		oDialogModel.setPropertyValue ("PositionX", 200)
		oDialogModel.setPropertyValue ("PositionY", 200)
		oDialogModel.setPropertyValue ("Height", 65)
		oDialogModel.setPropertyValue ("Width", 95)
		oDialogModel.setPropertyValue ("Title", "Step 1/2: Select the data range")
		
		' Adds the prompt.
		oTextModel = oDialogModel.createInstance ( _
			"com.sun.star.awt.UnoControlFixedTextModel")
		oTextModel.setPropertyValue ("PositionX", 5)
		oTextModel.setPropertyValue ("PositionY", 5)
		oTextModel.setPropertyValue ("Height", 15)
		oTextModel.setPropertyValue ("Width", 85)
		oTextModel.setPropertyValue ("Label", sPrompt)
		oTextModel.setPropertyValue ("MultiLine", True)
		oTextModel.setPropertyValue ("TabIndex", 1)
		oDialogModel.insertByName ("txtPrompt", oTextModel)
		
		' Adds the text input.
		oEditModel = oDialogModel.createInstance ( _
			"com.sun.star.awt.UnoControlEditModel")
		oEditModel.setPropertyValue ("PositionX", 5)
		oEditModel.setPropertyValue ("PositionY", 25)
		oEditModel.setPropertyValue ("Height", 15)
		oEditModel.setPropertyValue ("Width", 85)
		oEditModel.setPropertyValue ("Text", sCellsData)
		oDialogModel.insertByName ("edtCellsData", oEditModel)
		
		' Adds the buttons.
		oButtonModel = oDialogModel.createInstance ( _
			"com.sun.star.awt.UnoControlButtonModel")
		oButtonModel.setPropertyValue ("PositionX", 5)
		oButtonModel.setPropertyValue ("PositionY", 45)
		oButtonModel.setPropertyValue ("Height", 15)
		oButtonModel.setPropertyValue ("Width", 40)
		oButtonModel.setPropertyValue ("PushButtonType", _
			com.sun.star.awt.PushButtonType.CANCEL)
		oDialogModel.insertByName ("btnClose", oButtonModel)
		
		oButtonModel = oDialogModel.createInstance ( _
			"com.sun.star.awt.UnoControlButtonModel")
		oButtonModel.setPropertyValue ("PositionX", 50)
		oButtonModel.setPropertyValue ("PositionY", 45)
		oButtonModel.setPropertyValue ("Height", 15)
		oButtonModel.setPropertyValue ("Width", 40)
		oButtonModel.setPropertyValue ("PushButtonType", _
			com.sun.star.awt.PushButtonType.OK)
		oDialogModel.insertByName ("btnOK", oButtonModel)
		
		' Adds the dialog model to the control and runs it.
		oDialog = CreateUnoService ("com.sun.star.awt.UnoControlDialog")
		oDialog.setModel (oDialogModel)
		oDialog.setVisible (True)
		nResult = oDialog.execute
		oDialog.dispose
		
		' Cancelled
		If nResult = 0 Then
			Exit Function
		End If
		
		sCellsData = oEditModel.getPropertyValue ("Text")
		If sCellsData = "" Then
			sPrompt = "Cells with the data:"
		Else
			oRange = fnCheckRangeName (oDoc, sCellsData)
			If IsNull (oRange) Then
				sPrompt = "The range """ & sCellsData & """ does not exist."
			Else
				If oRange.getRows.getCount < 2 Or oRange.getColumns.getCount < 2 Then
					sPrompt = "The range """ & sCellsData & """ is too small (at least 2×2)."
				Else
					sPrompt = ""
					oDoc.getCurrentController.select (oRange)
					fnAskDataRange = oRange
					Exit Function
				End If
			End If
		End If
	Loop
End Function

' fnFindActiveDataRange: Finds the selected data range.
Function fnFindActiveDataRange (oDoc)
	Dim oSelection As Object, nI As Integer
	Dim oRanges As Object, oRange As Object
	Dim aCellAddress As New com.sun.star.table.CellAddress
	Dim aRangeAddress As New com.sun.star.table.CellRangeAddress
	
	oSelection = oDoc.getCurrentSelection
	
	' Some data ranges are already selected.
	If Not oSelection.supportsService ("com.sun.star.sheet.SheetCell") Then
		' Takes the first selection in multiple selections
		If oSelection.supportsService ("com.sun.star.sheet.SheetCellRanges") Then
			fnFindActiveDataRange = oSelection.getByIndex (0)
		' The only selection
		Else
			fnFindActiveDataRange = oSelection
		End If
		Exit Function
	End If
	
	' Finds the data range containing the single active cell
	aCellAddress = oSelection.getCellAddress
	oRanges = oSelection.getSpreadsheet.queryContentCells ( _
		com.sun.star.sheet.CellFlags.VALUE _
		+ com.sun.star.sheet.CellFlags.DATETIME _
		+ com.sun.star.sheet.CellFlags.STRING _
		+ com.sun.star.sheet.CellFlags.FORMULA)
	For nI = 0 To oRanges.getCount - 1
		oRange = oRanges.getByIndex (nI)
		aRangeAddress = oRange.getRangeAddress
		If 		aRangeAddress.StartRow <= aCellAddress.Row _
				And aRangeAddress.EndRow >= aCellAddress.Row _
				And aRangeAddress.StartColumn <= aCellAddress.Column _
				And aRangeAddress.EndColumn >= aCellAddress.Column Then
			oDoc.getCurrentController.select (oRange)
			fnFindActiveDataRange = oRange
			Exit Function
		End If
	Next nI
	' Not in a data cell range
End Function
