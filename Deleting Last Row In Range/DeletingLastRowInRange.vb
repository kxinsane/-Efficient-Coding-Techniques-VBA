' delete the last Row in Column 1 ("A")
With Activesheet
	.Cells(Rows.Count, 1).End(xlUp).EntireRow.Delete
End with

' delete the last Row in Column 2 ("B")
Me.Cells(Rows.Count, 2).End(xlUp).EntireRow.Delete
