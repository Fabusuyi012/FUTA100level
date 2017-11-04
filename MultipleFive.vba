'To print product and sum of multiples of 5 between 1 and 20
Sub MultipleFive()
    Dim x, sum, product As Integer
    sum = 0
    product = 1
    For x = 1 To 20
        If (x Mod 5 = 0) Then
            sum = sum + x
            product = product * x
        End If
    Next x
    MsgBox ("Their sum is " & sum & " while their product is " & product)
    
End Sub
