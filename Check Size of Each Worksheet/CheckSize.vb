Option Explicit

'==========================================================================================================
' ## CheckSizeEachWorksheet
'    Check the Size of Each Worksheet of ActiveWorkbook
'==========================================================================================================
Public Sub CheckSizeEachWorksheet()

    Dim wsWorksheet As Worksheet
    Dim strFullFile As String
    Dim sWBName As String

    sWBName = "TempFileToErase.xls"
    strFullFile = ThisWorkbook.Path & Application.PathSeparator & sWBName

    For Each wsWorksheet In Application.ActiveWorkbook.Worksheets
        wsWorksheet.Copy
        
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs strFullFile
        ActiveWorkbook.Close SaveChanges:=False
        Application.DisplayAlerts = True
        
        MsgBox wsWorksheet.Name & ": " & VBA.FileLen(strFullFile) & " bytes"
        
        Kill strFullFile
    Next wsWorksheet

End Sub
