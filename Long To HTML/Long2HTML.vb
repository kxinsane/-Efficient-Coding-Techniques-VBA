Option Explicit
'==========================================================================================================
' ## ColorDec2ColorHex
'    Converts Long Colour to HTML String
'
' // Example of calling
'    LongColor2HTML(12345678)   'return: #4E61BC
'    LongColor2HTML(9253632)    'return: #00338D
'==========================================================================================================
Public Function LongColor2HTML(ByVal lColour As Long) As String
    
    Dim lRed, lGreen, lBlue As Long
    
    lRed = (lColour Mod 256)
    lGreen = (lColour \ 256) Mod 256
    lBlue = (lColour \ 65536) Mod 256
    LongColor2HTML = _
        "#" & Application.WorksheetFunction.Dec2Hex(lRed, 2) & Application.WorksheetFunction.Dec2Hex(lGreen, 2) & Application.WorksheetFunction.Dec2Hex(lBlue, 2)

End Function