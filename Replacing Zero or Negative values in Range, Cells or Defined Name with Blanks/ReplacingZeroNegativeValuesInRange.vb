	' Columns, default replace zero ("0") with blanks
	Columns("A:D").Replace What:="0", Replacement:=""

	' Search for containing zero - will replace 0 and also replace 20, 200, 202 etc.
	Columns("A:D").Replace What:="*0*", Replacement:=""

	' Columns, default replace negatives with zero (0) with uses a Minus sign
	Columns("I:I").Replace What:="-*", Replacement:="0"

	' Look at whole Cells, Match Case and Replace.
	Columns("A:D").Replace What:="0", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

	' Rows, look at whole Cells
	Rows("1:5").Replace What:="0", Replacement:="", LookAt:=xlWhole

	' Range, simple, replace all zeros (0) with Blanks
	Range("A1:B4").Replace What:="0", Replacement:=""

	' Defined Name or Named Range
	[MyRange].Replace What:="0", Replacement:="", LookAt:=xlWhole

	' Short Notation
	[A1:D5].Replace What:="0", Replacement:=""

	' Using With for the ActiveSheet in Workbook
	With ThisWorkbook.ActiveSheet.Cells
		.Replace What:="0", Replacement:=""
	End With

	' Using With for Columns
	With Columns("A:D").Cells
		.Replace What:="0", Replacement:=""
	End With

	' Replacing Errors
	.Replace What:="#N/A", Replacement:="", LookAt:=xlWhole