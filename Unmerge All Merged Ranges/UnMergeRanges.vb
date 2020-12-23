Option Explicit
'==========================================================================================================
' ## UnMergeRanges
'    unmerge all merged Ranges for ActiveWorkbook, Worksheet UsedRange
'
' // Example of calling
'    Call UnMergeRanges(ActiveWorkbook.ActiveSheet.UsedRange)
'
' // Parameters
'    rngUsedRange := ActiveWorkbook.ActiveSheet.UsedRange
'==========================================================================================================

Public Sub UnMergeRanges(ByVal rngUsedRange As Range)

    Dim rngMerged As Range
    Dim rngCell As Range
    
	' iterate UsedRange and identify Cells containing merged areas
    For Each rngCell In rngUsedRange
        If rngCell.MergeCells Then
            Set rngMerged = rngCell.MergeArea
            rngMerged.Cells(1, 1).UnMerge
        End If
    Next rngCell

	' autofit Columns for unmerged Cells
    ActiveWorkbook.ActiveSheet.Cells.EntireColumn.AutoFit

End Sub
