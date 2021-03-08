Option Explicit
'==========================================================================================================
' ## InsertCheckboxes
'    Insert Checkboxes in Selection Range with Linked Cells
'
' // Example of calling
'    InsertCheckboxes(2, True): Linked Cells will be 2 columns offset with 3D Shading Checkboxe
'    InsertCheckboxes: Linked Cells will be in the same Cell with Checkbox, without 3D Shading
'
' // Parameters
' lLinkedCellOffset:= Offset to Linked Cell (Long Value)
' boolShading:= True/False (Boolean Value) to Display 3D Shading Checkboxes

Public Sub InsertCheckboxes(Optional ByVal lLinkedCellOffset As Long, Optional boolShading as Boolean)

    Dim rngCell As Range
    Dim cb As Checkbox

    For Each rngCell in Selection
        Set cb = Activesheet.Checkboxes.Add(rngCell.Left, rngCell.Top, rngCell.Width, rngCell.Height)
    
        With cb
            .Caption = ""
            .Value = xlOff
            .LinkedCell = rngCell.Offset(0,lLinkedCellOffset).Address
            .Display3DShading = boolShading
        End With
    
    Next rngCell

End Sub