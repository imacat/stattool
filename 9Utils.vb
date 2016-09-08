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

' 9Utils: The utility macros.
'   by imacat <imacat@mail.imacat.idv.tw>, 2016-08-10

Option Explicit

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

' fnSpecifyData: Specifies the data
Function fnSpecifyData (oRange As Object, sPrompt1 As String, sPrompt2 As String) As Object
	Dim mLabels (oRange.getColumns.getCount - 1) As String
	Dim nI As Integer, mSelected (0) As Integer
	Dim oDialog As Object, oTextModel As Object
	Dim oListModel1 As object, oListModel2 As Object
	Dim nResult As Integer, nColumn As Integer, mRanges (1) As Object
	
	For nI = 0 To oRange.getColumns.getCount - 1
		mLabels (nI) = oRange.getCellByPosition (nI, 0).getString
	Next nI
	
	' Runs the dialog
	oDialog = CreateUnoDialog (DialogLibraries.Stats.Dlg2SpecData)
	oTextModel = oDialog.getControl ("txtPrompt1").getModel
	oTextModel.setPropertyValue ("Label", sPrompt1)
	oListModel1 = oDialog.getControl ("lstData1").getModel
	oListModel1.setPropertyValue ("StringItemList", mLabels)
	mSelected (0) = 0
	oListModel1.setPropertyValue ("SelectedItems", mSelected)
	oTextModel = oDialog.getControl ("txtPrompt2").getModel
	oTextModel.setPropertyValue ("Label", sPrompt2)
	oListModel2 = oDialog.getControl ("lstData2").getModel
	oListModel2.setPropertyValue ("StringItemList", mLabels)
	mSelected (0) = 1
	oListModel2.setPropertyValue ("SelectedItems", mSelected)
	
	nResult = oDialog.execute
	oDialog.dispose
	
	' Cancelled
	If nResult = 0 Then
		Exit Function
	End If
	
	nColumn = oListModel1.getPropertyValue ("SelectedItems") (0)
	mRanges (0) = oRange.getCellRangeByPosition ( _
		nColumn, 0, nColumn, oRange.getRows.getCount - 1)
	nColumn = oListModel2.getPropertyValue ("SelectedItems") (0)
	mRanges (1) = oRange.getCellRangeByPosition ( _
		nColumn, 0, nColumn, oRange.getRows.getCount - 1)
	fnSpecifyData = mRanges
End Function

' fnAskDataRange: Asks the user for the data range, or null when
'                 the user cancelled
Function fnAskDataRange (oDoc As Object) As Object
	Dim oRange As Object
	Dim oDialog As Object, nResult As Integer
	Dim oTextModel As Object, oEditModel As Object
	Dim sPrompt As String, sCellsData As String
	
	oRange = fnFindActiveDataRange (oDoc)
	If IsNull (oRange) Then
		sCellsData = ""
	Else
		sCellsData = oRange.getPropertyValue ("AbsoluteName")
	End If
	sPrompt = "&27.Dlg1AskRange.txtPrompt.Label"
	
	' Loop until we finds good data
	Do While sPrompt <> ""
		' Runs the dialog
		oDialog = CreateUnoDialog (DialogLibraries.Stats.Dlg1AskRange)
		oTextModel = oDialog.getControl ("txtPrompt").getModel
		oTextModel.setPropertyValue ("Label", sPrompt)
		oEditModel = oDialog.getControl ("edtCellsData").getModel
		oEditModel.setPropertyValue ("Text", sCellsData)
		
		nResult = oDialog.execute
		oDialog.dispose
		
		' Cancelled
		If nResult = 0 Then
			Exit Function
		End If
		
		sCellsData = oEditModel.getPropertyValue ("Text")
		If sCellsData = "" Then
			sPrompt = "&27.Dlg1AskRange.txtPrompt.Label"
		Else
			oRange = fnCheckRangeName (oDoc, sCellsData)
			If IsNull (oRange) Then
				sPrompt = "&35.Dlg1AskRange.txtPrompt.LabelNotExists"
			Else
				If oRange.getRows.getCount < 2 Or oRange.getColumns.getCount < 2 Then
					sPrompt = "&36.Dlg1AskRange.txtPrompt.LabelTooSmall"
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
