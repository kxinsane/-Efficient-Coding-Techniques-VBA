Option Explicit
'==========================================================================================================
' ## ColourAllFormulaCells
'    will colour any Cells with Formula on specified Worksheet, function returns Count of coloured cells 
'
' // Example of calling
'    Call/MsgBox ColourAllFormulaCells(5, ActiveSheet)
'
' // Parameters
'    ColourIndex := 5
'	 wbObject := ActiveSheet, Sheet1, Sheets("Sheet1"), [Sheet1], ThisWorkbook.ActiveSheet, ActiveWorkbook.ActiveSheet, Sheets(CStr([Defined_Sheet_Name]))
'==========================================================================================================

Public Function ColourAllFormulaCells(ByVal intColourIndex As Integer, ByVal wbObject As Worksheet) as Long
	
	wbObject.UsedRange.SpecialCells(xlCellTypeFormulas).Interior.ColorIndex = intColourIndex
	
	ColourAllFormulaCells = wbObject.UsedRange.SpecialCells(xlCellTypeFormulas).Count

End Function