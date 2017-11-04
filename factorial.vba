'To find the factorial of a number
Sub getFactorial()
    Dim number As Integer
    number = InputBox("Enter the number whose factoria you want to find")
    MsgBox factorial(number)
End Sub
Function factorial(number)
    Dim i, sumProduct As Integer
    sumProduct = 1
    If number = 0 Then
        factorial = 1
    Else
        For i = 1 To number
            sumProduct = sumProduct * i
        Next i
        factorial = sumProduct
    End If
End Function
