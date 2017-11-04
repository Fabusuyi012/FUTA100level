'To determine whether an Integer is positive, negative or zero
Sub posNeg()
    Dim number As Integer
    number = InputBox("Enter an INTEGER")
    If number = 0 Then
        MsgBox ("Number is zero")
    End If
    If number < 0 Then
        MsgBox ("Negative number")
    End If
    If number > 0 Then
        MsgBox ("Positive number")
    End If
End Sub
