Option Explicit
'==========================================================================================================
' ## ColorDec2ColorHex
'    converts Decimal Color Code (0 to 16777215) to Hexadecimal HTML color code (#000000 to #FFFFFF) 
'
' // Example of calling
'    ColorDec2ColorHex(vbRed)               'return: #FF0000
'    ColorDec2ColorHex(RGB(0, 51, 141))     'return: #00338D
'    ColorDec2ColorHex(12345678)            'return: #4E61BC
'    ColorDec2ColorHex(16777215)            'return: #FFFFFF
'    ColorDec2ColorHex(0)                   'return: #000000
'==========================================================================================================
Public Function ColorDec2ColorHex(ByVal lColorInDec as Long)

    Dim lColour As Long
    
    If lColorInDec > 16777215 Then
        lColour = 16777215
    ElseIf lColorInDec < 0 Then
        lColour = 0
    Else
        lColour = lColorInDec
    End If

    ColorDec2ColorHex = "#" & Right$("00" & Hex$((lColour Mod 256)), 2) & _
        Right$("00" & Hex$((lColour \ 256) Mod 256), 2) & _
        Right$("00" & Hex$(lColour \ 65536), 2)

End Function