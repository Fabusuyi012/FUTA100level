'To print the nth number in the fibonacci sequence
Sub fibonacciNth()
    Dim answer, fibonacci, s1, s2, s3, nth, n As Integer
    n = InputBox("Enter the value of n")
    s1 = 0
    s2 = 1
    s3 = 1
    For i = 4 To n
        nth = s2 + s3
        s2 = s3
        s3 = nth
    Next i
    MsgBox ("The " & n & "th number in the fibonacci sequence is " & nth)
End Sub
