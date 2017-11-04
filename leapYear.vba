'To determine if a year is a leap year
Sub getLeapYear()
    Dim year As Integer
    year = InputBox("Enter the year")
    If (testLeapYear(year) = 1) Then
        MsgBox (year & " is a leap year")
    End If
    If (testLeapYear(year) = 0) Then
        MsgBox (year & " is not a leap year")
    End If
End Sub
Function testLeapYear(year) As Integer
    If (year Mod 4 = 0) And (year Mod 100 <> 0) Or (year Mod 400 = 0) Then
        testLeapYear = 1 '1 denotes being a leap year
    Else
        testLeapYear = 0 '0 denotes not being a leap year
    End If
End Function
