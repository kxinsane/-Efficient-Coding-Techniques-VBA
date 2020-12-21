Option Explicit
'==========================================================================================================
' ## SortWorksheets
'    sorts visible Worksheets for Workbook: Ascending or Descending
'
' // Example of calling
'    Call SortWorksheets(ActiveWorkbook, xlAscending)
'    Call SortWorksheets(ActiveWorkbook, xlDescending)
'    Call SortWorksheets(Workbooks("Some Workbook Name.xls"), xlDescending)
'
' // Parameters
'    wbObject      := ActiveWorkbook, Workbooks("Some Workbook Name.xls")
'    SortDirection := xlAscending (1), xlDescending (2)
'==========================================================================================================

Public Sub SortWorksheets(ByVal wbObject As Workbook, ByVal SortDirection As XlSortOrder)

    Dim iCurrentIndex As Integer
    Dim iMovingIndex As Integer

    For iCurrentIndex = 1 To wbObject.Sheets.Count
        For iMovingIndex = 1 To wbObject.Sheets.Count - 1
            If SortDirection = xlAscending Then
                If UCase$(wbObject.Sheets(iMovingIndex).Name) > UCase$(wbObject.Sheets(iMovingIndex + 1).Name) Then
                    wbObject.Sheets(iMovingIndex).Move After:=wbObject.Sheets(iMovingIndex + 1)
                End If
            ElseIf SortDirection = xlDescending Then
                If UCase$(wbObject.Sheets(iMovingIndex).Name) < UCase$(wbObject.Sheets(iMovingIndex + 1).Name) Then
                    wbObject.Sheets(iMovingIndex).Move After:=wbObject.Sheets(iMovingIndex + 1)
                End If
            End If
        Next iMovingIndex
    Next iCurrentIndex

End Sub