Option Explicit

Public Function CountUnique(rngData As Range) As Long

	On Error Resume Next 

	' turn off screen refreshing
    Application.ScreenUpdating = False
    
    Dim lCounter As Long
    Dim vntData As Variant
    
    Dim objDict As Object: Set objDict = CreateObject("Scripting.Dictionary") 'New Dictionary
    
    ' vbBinaryCompare:=0, binary comparison, vbTextCompare:=1, textual comparison
    objDict.CompareMode = 0
    
	' pick up Range Data to Variant
    vntData = rngData
    
    For lCounter = 1 To UBound(vntData)
        If vntData(lCounter, 1) <> "" And Not objDict.Exists(vntData(lCounter, 1)) Then
            objDict.Add vntData(lCounter, 1), 1
        End If
    Next lCounter
    
	' return Count
    CountUnique = objDict.Count
    
    Set objDict = Nothing
    Erase vntData
    
	' turn on screen refreshing
    Application.ScreenUpdating = True

End Function
