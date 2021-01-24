Option Explicit
'==========================================================================================================
' ## ToggleAutoFilter
'	 switches Autofilter On/Off for adjacent Header Range then Select default Cell
'
' // Example of calling
'	 ToggleAutoFilter Sheets("Sheet1"), True, "C1:E1", ActiveCell.Address
'	 ToggleAutoFilter Sheet1, False, "C1:E1", "C1"
'
' // Parameters
' 	 wsSelectedWorksheet:= Sheet1, Sheets("Sheet1"), ActiveSheet
'	 boolToggleState:= True, switch On, False, switch Off
'	 strHeaderRange:= "C1:E1"
'	 strDefaultCell:= "C1"
'
'==========================================================================================================
Public Function ToggleAutoFilter(ByVal wsSelectedWorksheet As Worksheet, ByVal boolToggleState as Boolean, ByVal strHeaderRange as String, ByVal strDefaultCell as String)

	Application.ScreenUpdating = False
	
	wsSelectedWorksheet.Select
	If ActiveSheet.AutoFilterMode = True And boolToggleState = False Or ActiveSheet.AutoFilterMode = False And boolToggleState = True Then
		wsSelectedWorksheet.Range(strHeaderRange).Select
		Selection.AutoFilter
		wsSelectedWorksheet.Range(strDefaultCell).Select
	End If
	
	Application.ScreenUpdating = True

End Function