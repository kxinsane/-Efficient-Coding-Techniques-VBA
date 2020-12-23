Option Explicit
'==========================================================================================================
' ## ScrollToPosition
'    scroll to position and Select the entire Column 
'
' // Example of calling
'    Call ScrollToPosition Range("Z1"), Columns("AA:AA")
'
' // Parameters
'    rngPosition := ActiveSheet.Range("B2"), Range("B2"), Sheet1.Range("B2")
'	 rngColumns := ActiveSheet.Columns("B:B"), Columns("B:B"), Sheet1.Columns("B:B")
'==========================================================================================================

Public Sub ScrollToPosition(ByVal rngPosition As Range, ByVal rngColumns As Range)

	Application.GoTo Reference:=Range("A1")
	Application.GoTo Reference:=rngPosition
	rngColumns.Select

End Sub