Sub FindLastRow()

    Dim wsName as Worksheet
    Dim lLastRow as Long

    Set wsName = ActiveSheet

    ' By Find Function - in my opinion the best one and the most reliable
    lLastRow = wsName.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row

    ' Using: SpecialCells Function
    lLastRow = wsName.Cells.SpecialCells(xlCellTypeLastCell).Row

    ' Using: Like CTRL + SHIFT + END
    lLastRow = wsName.Cells(wsName.Rows.Count, "A").End(xlUp).Row

    ' Using: UsedRange
    wsName.UsedRange 'Refresh UsedRange
    lLastRow = wsName.UsedRange.Rows(wsName.UsedRange.Rows.Count).Row

    'Using: Table Range
    lLastRow = wsName.ListObjects("TableName").Range.Rows.Count

    'Using: Named Range
    lLastRow = wsName.Range("RangeName").Rows.Count

    ' Using: Like CTRL + SHIFT + DOWN, Range should be first cell in data set
    lLastRow = wsName.Range("A1").CurrentRegion.Rows.Count

End Sub