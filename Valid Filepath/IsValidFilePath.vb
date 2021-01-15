Option Explicit
'==========================================================================================================
' ## IsValidFilePath
'    checking Valid FilePath using Regular Expression
'
' // Example of calling
'    IsValidFilePath("C:/test/hello.html"), IsValidFilePath("//10.1.1.107/test/hello")
'
' // Parameters
'    strFileName:= FilePath
'
' #### Full Regex Patern Meaning:
'	^   (?# <- Start at the beginning of the line #)
'    		(?# check opening drive or path delimiter -> #)
'        		(?: (?# "C:", "C:\", "C:..\", "C:.\" -> #)
'               	(?:[A-Z]:(?:\.{1,2}[\/\\]|[\/\\])?)
'            		| (?# or "\", "..\", ".\", "\\" -> #)
'               	(?:[\/\\]{1,2}|\.{1,2}[\/\\])
'        		)?
'    	(?# check for invalid patterns in the string -> #)
'        	(?:([^\/\\. ]|[^\/. \\][\/. \\][^\/. \\]|[\/\\]$)*
'    		[^\x00-\x1A|*?\s+,;"'`:<.>=[\]] (?# <- Ensure that the last character is valid #)
'	$   (?# <- End of the line #)
'==========================================================================================================
Public Function IsValidFilePath(ByVal strFileName As String)
    Dim objRegExp As Object: Set objRegExp = CreateObject("VBScript.RegExp")
    
    With objRegExp
        .Pattern = "^(?:(?:[A-Z]:(?:\.{1,2}[\/\\]|[\/\\])?)|(?:[\/\\]{1,2}|\.{1,2}[\/\\]))?(?:[^\x00-\x1A|*?\v\r\n\f+\/,;'" & chr(34) & "`\\:<>=[\]]+[\/\\]?)+$"
        .Global = True
        .IgnoreCase = True
    End With
    
    IsValidFilePath = objRegExp.Test(strFileName)
End Function