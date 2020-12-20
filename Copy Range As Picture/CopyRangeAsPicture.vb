Public Sub CopyRangeAsPicture() ' copy will be in the Clipboard
	Dim varCopyRange As Variant
	Set varCopyRange = Range("A1:A10")
	varCopyRange.CopyPicture Appearance:=xlScreen, Format:=xlPicture
End Sub