Option Explicit
'==========================================================================================================
' ## ExcelWorksheetToPDF
' 	 Write Excel data to PDF File with Optional Document Properties
'
' // Example of calling
' 	 ExcelWorksheetToPDF Sheets("Sheet1"), _
'    		"C:\Users\KX\Desktop\test.pdf", _
'       	xlQualityStandard, _
'       	True, _
'       	False, _
'       	False, _
'       	"Document Title", _
'       	"John Doe", _
'       	"Report Subject", _
'       	"Excel PDF"
'
' // Parameters
'	 wsName:= Worksheet Name
'	 strFilepath:= Filepath and Filename
'	 xlQualityFormat:= Quality of spreadsheets (xlQualityStandard or xlQualityMinimum / 0 or 1)
'	 boolIncludeDocumentProperties:= Include Document Properties (True or False)
'	 boolIgnorePrintAreAs:= Ignores any print areAs set when publishing (True or False)
'	 boolOpenAfterPublish:= Displays the file in the viewer after it is published (True or False)
'	 strDocumentTitle:= Optional, Document Title
'	 strAuthor:= Optional, Author of Document
'	 strSubject:= Optional, Report Subject of Document
'	 strKeyWords:= Optional, Key Words of Document
'==========================================================================================================
Private Function ExcelWorksheetToPDF(ByVal wsName As Worksheet, _
	ByVal strFilepath As String, _
	ByVal xlQualityFormat As XlFixedFormatQuality, _
	ByVal boolIncludeDocumentProperties As Boolean, _
	ByVal boolIgnorePrintAreAs As Boolean, _
	ByVal boolOpenAfterPublish As Boolean, _
	Optional ByVal strDocumentTitle As String, _
	Optional ByVal strAuthor As String, _
	Optional ByVal strSubject As String, _
	Optional ByVal strKeyWords As String)

	' // write Document Properties if required
	On Error Resume Next
	If boolIncludeDocumentProperties Then
		ActiveWorkbook.BuiltinDocumentProperties("Title").Value = strDocumentTitle
		ActiveWorkbook.BuiltinDocumentProperties("Author").Value = strAuthor
		ActiveWorkbook.BuiltinDocumentProperties("Subject").Value = strSubject
		ActiveWorkbook.BuiltinDocumentProperties("Keywords").Value = strKeyWords
	End If
	On Error GoTo 0

	' // export Worksheet As PDF File
	wsName.ExportAsFixedFormat _
	Type:=xlTypePDF, _
	Filename:=strFilepath, _
	Quality:=xlQualityFormat, _
	IncludeDocProperties:=boolIncludeDocumentProperties, _
	IgnorePrintAreAs:=boolIgnorePrintAreAs, _
	OpenAfterPublish:=boolOpenAfterPublish

End Function
