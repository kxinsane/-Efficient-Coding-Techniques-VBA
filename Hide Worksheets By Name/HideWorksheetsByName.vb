Option Explicit
'==========================================================================================================
' ## HideWorksheetsByName
'    hides all Worksheets in a Workbook by Names
'==========================================================================================================
Public Sub HideWorksheetsByName()

	Dim wsObject As Worksheet

	For Each wsObject In ThisWorkbook.Worksheets
		Select Case UCase(wsObject.Name)
			' modify Sheets Names below 
			Case "Sheet2", "Sheet3", "*Other*"
				wsObject.Visible = True
			Case Else
				wsObject.Visible = False
		End Select
	Next wsObject

End Sub