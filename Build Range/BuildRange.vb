Option Explicit
'==========================================================================================================
' ## BuildRange
'    Builds Range for single Cell Reference X Columns by Y Rows, Range Object can be set or selected
'
' // Example of calling
'    Dim rngExample As Range
'    Set rngExample = BuildRange("A1", 1, 2)
'    If Not rngExample Is Nothing Then rngExample.Select
'
' // Parameters
'    strStartCell := start Cell (e.g. "A1")
'    iColumns := number of Columns (e.g. 1)
'    lRows := number of Rows (e.g. 2)
'
' // Return
'    Range Object identifying Range that can be set or selected
'    Nothing for Range that cannot be set
'==========================================================================================================
Public Function BuildRange(ByVal strStartCell As String, ByVal iColumns As Integer, ByVal lRows As Long) As Range

	Dim rng as Range
		
	On Error Resume Next
        
        If iColumns <= 0 Or lRow <= 0 Then Exit Function
		
	Set rng = Range(strStartCell, Range(Left(Cells(1, _
		Int(iColumns + Range(strStartCell).Column - 1)).Address(1, 0), _
		InStr(1, Cells(1, _
		Int(iColumns + Range(strStartCell).Column)).Address(1, 0), "$") - 1) & _
		Range(strStartCell).Row + lRows - 1)).Cells

End Function
