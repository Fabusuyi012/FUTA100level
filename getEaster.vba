'To get easter date of a given year between 1982 and 2048
Sub getEaster()
    Dim year, a, b, c, d, e, sum, day, easterDay As Integer
    year = InputBox("Enter a year between 1982 and 2048")
    If year < 1982 Or year > 2048 Then
        MsgBox ("Year not in proper range")
    Else
        a = year Mod 19
        b = year Mod 4
        c = year Mod 7
        d = ((19 * a) + 24) Mod 30
        e = ((2 * b) + (4 * c) + (6 * d) + 5) Mod 7
        sum = d + e
        day = 22 + sum
        If day <= 31 Then
            easterDay = day
            MsgBox ("Easter in year " & year & " falls on March " & easterDay & ", " & year)
        Else
            easterDay = day - 31
            MsgBox ("Easter in year " & year & " falls on April " & easterDay & ", " & year)
        End If
    End If
End Sub
