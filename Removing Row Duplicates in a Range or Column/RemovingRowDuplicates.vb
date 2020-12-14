    ' using an Offset method to exclude Headers, Range must be contiguous because blanks will stop the process
    Range("A1").CurrentRegion.Offset(1, 0).RemoveDuplicates (Array(1))
    
	' using short notation and include or exclude Headers - it will cover Entire Column A
    [A:A].RemoveDuplicates Columns:=1, Header:=xlNo
    [A:A].RemoveDuplicates Columns:=1, Header:=xlYes

	' will cover Entire Column A with include Header
    Columns(1).RemoveDuplicates Columns:=1
    [A:A].RemoveDuplicates Columns:=1

	' Range, using Short Notation and include or exclude Headers - it will cover the Entire Column A
	' it will remove duplicates of any Rows that are the same for the first 4 Columns, not duplicates anywhere in the Range
    ActiveSheet.Range("A:D").RemoveDuplicates Columns:=Array(1, 2, 3, 4), Header:=xlYes