'To convert the hectare equivalent of an acre value
Sub sendAcre()
    Dim acre As Double
    acre = InputBox("Enter the value in acre")
    MsgBox ("The hectare equivalent of " & acre & " is " & getHectare(acre))
End Sub
Function getHectare(acre)
    getHectare = acre * 0.404686
End Function
