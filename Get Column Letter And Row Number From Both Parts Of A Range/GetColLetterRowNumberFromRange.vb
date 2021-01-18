' Select a Range for the Examples
Range("D2:L23").Select

' identify Column from the 1st part of the Range following the ":"
Debug.Print Split(Split(Selection.Address, ":")(0), "$")(1)

' identify Row from the first part of the Range following the ":"
Debug.Print Split(Split(Selection.Address, ":")(0), "$")(2)

' identify Column from the 2nd part of the Range following the ":"
Debug.Print Split(Split(Selection.Address, ":")(1), "$")(1)

' identify Row from the 2nd part of the Range following the ":"
Debug.Print Split(Split(Selection.Address, ":")(1), "$")(2)
