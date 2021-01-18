Option Explicit
'==========================================================================================================
' ## ToggleColumns
'    Hides single Column or Range of adjoining Columns without selecting Worksheet
'
' // Example of calling
'	 ToggleColumns wsWorksheetName:=Sheet1, iColIndex:=2, iNoColumnsToHide:=4, boolToggle:=True
'	 ToggleColumns Sheets("Sheet1"), 3, 4, True
'
' // Inline method
'	 Worksheets("Sheet1").Columns(2).Resize(, 4).EntireColumn.Hidden = True
'
' // Parameters
'	 wsWorksheetName:= Worksheet Name (e.g. "Arkusz1", "Sheet1")
'	 iColIndex:= Starting Column index to hide
'	 iNoColumnsToHide:= Number of Columns to hide from starting Column
'	 boolToggle:= True or False (Hide or Show Columns)
'==========================================================================================================
Private Sub ToggleColumns(ByVal wsWorksheetName as Worksheet, ByVal iColIndex as Integer, ByVal iNoColumnsToHide as Integer, ByVal boolToggle as Boolean)
	
	With wsWorksheetName
		.Columns(iColIndex).Resize(, iNoColumnsToHide).EntireColumn.Hidden = boolToggle
	End With
	
End Sub