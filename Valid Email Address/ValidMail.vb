Option Explicit

Public Function IsMailValid(ByVal strEmailAddress As String) As Boolean

    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")

    ' regular expression pattern match
   objRegExp.IgnoreCase = True
   objRegExp.Pattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"

   ' return True for valid email or False for invalid email
   IsMailValid = objRegExp.Test(strEmailAddress)

End Function