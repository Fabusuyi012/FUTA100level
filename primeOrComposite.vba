Sub doStuffs()
    'To determine whether a number is a prime number or a composite number
    Dim number, usable, count As Integer
    number = InputBox("Enter the number")
    If number = 1 Then
        MsgBox ("The number is neither a prime number nor a composite number")
    End If
    For i = 1 To number
        If number Mod i = 0 Then
            count = count + 1
            End If
    Next i
    If count = 2 Then
        MsgBox ("The number is a prime number")
    End If
    If count > 2 Then
        MsgBox ("The number is a composite number")
    End If
End Sub



