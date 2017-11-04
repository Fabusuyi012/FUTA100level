'To find the sum and product of the LCM and HCF of two given numbers
Sub sumProduct()
    Dim number1, number2, A, B, lowerLimit, higherLimit, hcf, lcm, count, i, remainder As Integer
    number1 = InputBox("Enter the first number")
    number2 = InputBox("Enter the second number")
    A = number1
    B = number2
    'Start of HCF calculation
    Do While (number1 > 0 Or number2 > 0)
        remainder = number1 Mod number2
        number1 = number2
        If remainder = 0 Then
            hcf = number1
            Exit Do
        End If
        number2 = remainder
        hcf = number1
    Loop
    'End of HCF calculation
    'Start of LCM calculation
    higherLimit = A * B
    count = 0
    If (A > B) Then
        lowerLimit = A
    Else
        lowerLimit = B
    End If
    For i = lowerLimit To higherLimit
        If (i Mod A = 0 And i Mod B = 0) Then
            lcm = i
            count = count + 1
            If count = 1 Then
                i = higherLimit + 1
            End If
        End If
    Next i
    'End of LCM calculation
    MsgBox ("The sum and product of the LCM and HCF of " & A & " and " & B & " is " & hcf + lcm & " and " & hcf * lcm & " respectively")
End Sub


