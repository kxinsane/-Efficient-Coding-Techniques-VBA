Option Explicit
'==========================================================================================================
' ## BrightOrDarkFont
'    Specifies whether White Font should be used 
'
' // Example of calling
'    BrightOrDarkFont(RGB(0, 0, 0), vbBlack)        'return: 16777215 (VBA White Font Color in Decimal)
'    BrightOrDarkFont(RGB(255, 255, 255), vbWhite)  'return: 0 (Black Font Color in Decimal)
'==========================================================================================================
Public Function BrightOrDarkFont(ByVal lRGB As Long, ByVal lFontColour As Long) As Long
    
    Dim lRed, lGreen, lBlue As Long
    
    lRed = (lRGB And &HFF)
    lGreen = (lRGB And &HFF00&) / 256
    lBlue = (lRGB And &HFF0000) / 65536
    
    If lRed * 0.3 + lGreen * 0.59 + lBlue * 0.11 > 128 Then BrightOrDarkFont = lFontColour Else BrightOrDarkFont = vbWhite
End Function