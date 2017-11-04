'To print the first 15 odd numbers between 1 and 50
Sub odd15()
    Dim x, sum, count As Integer
    count = 0
    For x = 1 To 50
        If (x Mod 2 <> 0) Then
            count = count + 1
            MsgBox x
            If count = 15 Then
                x = 51
            End If
        End If
    Next x
End Sub

