Option Explicit
'==========================================================================================================
' ## IsValidFileName
'    checking Valid Filename using Regular Expression
'
' // Example of calling
'    IsValidFileName("filename.txt"), IsValidFileName("filename_without_extension")
'
' // Parameters
'    strFileName:= Filename with or without extension
'==========================================================================================================
Public Function IsValidFileName(ByVal strFileName As String)
    Dim objRegExp As Object: Set objRegExp = CreateObject("VBScript.RegExp")
    
    ' Pattern by Rich Jenkins, it will checks that a string is valid on:
	'	Windows (NTFS), Mac (HFS+) and most Linux distros as a file/folder name as well as part of a URI without encoding.
    With objRegExp
        .Pattern = "^(?!.{256,})(?!(aux|clock\$|con|nul|prn|com[1-9]|lpt[1-9])(?:$|\.))[^ ][ \.\w-$()+=[\];#@~,&amp;']+[^\. ]$"
        .Global = True
        .IgnoreCase = True
    End With
    
    IsValidFileName = objRegExp.Test(strFileName)
End Function