Option Explicit
'==========================================================================================================
' ## SetRange
'    search Worksheet for Criteria
'    return Range object which is Start Cell x Last Row Cell x NumberOfColumns wide or Nothing with the option to use a Defined Name
'
' // Example of calling
'    Dim rngExample As Range
'    Const strHeader as String = "Header"
'    '[1] Set rngExample = SetRange(Sheet1, strHeader, 1) := find Header, return single Column Range by x Rows or Nothing
'    '[2] Set rngExample = SetRange(Sheets("Sheet1"), strHeader, 3, True) := find Defined Name, return Range it is 3 Columns wide by x Rows or Nothing
'
' // Parameters
'    wsName := Sheet1, ActiveSheet, Sheets("Sheet1"), [Sheet1], ThisWorkbook.ActiveSheet, ActiveWorkbook.ActiveSheet, Sheets(CStr([Defined_Sheet_Name]))
'    strCriteria := Header
'    iNumberOfCols := 3 (wide)
'    boolDefinedName := True / False
'==========================================================================================================
Public Function SetRange( _
    ByVal wsName As Worksheet, _
    ByVal strCriteria As String, _
    Optional iNumberOfCols As Integer = 1, _
    Optional boolDefinedName As Boolean = False) As Range
    
        Dim objName As Name
        Dim rngCell As Range
    
        With wsName
            If boolDefinedName Then
                For Each objName In ThisWorkbook.Names
                    If objName.Name Like strCriteria Then
                        Set rngCell = .Range(strCriteria): Exit For
                    End If
                Next objName
            Else
                Set rngCell = .Cells.Find(strCriteria, .Cells(1, 1), xlValues, xlWhole, xlByRows, xlNext, True)
                If Not rngCell Is Nothing Then
                    Set SetRange = .Range(rngCell.Address & ":" & Split(.Cells(, rngCell.Column + iNumberOfCols - 1).Address, "$")(1) & _
                        .Range(Split(.Cells(, rngCell.Column).Address, "$")(1) & .Rows.Count).End(xlUp).Row).Cells
                End If
            End If
        End With
End Function
