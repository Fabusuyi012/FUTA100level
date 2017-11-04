Sub solveExp()
    'to solve the expression a^x + b^y = c^z
    Dim A, B, C, x, y, z As Integer
    Dim lhs, rhs As Double
    A = InputBox("Enter the value of A")
    B = InputBox("Enter the value of B")
    C = InputBox("Enter the value of C")
    x = InputBox("Enter the value of x")
    y = InputBox("Enter the value of y")
    z = InputBox("Enter the value of z")
    If x <= 2 Or y <= 2 Or z <= 2 Then
        MsgBox ("x, y and/or z must be greater than 2")
    End If
    If x > 2 And y > 2 And z > 2 Then
        lhs = (A ^ x) + (B ^ y)
        rhs = C ^ z
        If lhs = rhs Then
            MsgBox ("(" & A & "^" & x & ")+(" & B & "^" & y & ")=" & C & "^" & z)
        End If
        If lhs <> rhs Then
            MsgBox ("(" & A & "^" & x & ")+(" & B & "^" & y & ") is not equal to " & C & "^" & z)
        End If
    End If
End Sub
