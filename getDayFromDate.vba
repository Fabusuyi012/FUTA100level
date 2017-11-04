'To give corresponding day based on date input
Sub getDayNum()
    Dim inputDate As Date
    Dim dayNum, Day, month, prevMonth, year, solution As Integer
    inputDate = InputBox("Enter the date in the form DD/MM/YYYY")
    MsgBox (CStr(inputDate))
    Day = DatePart("d", CStr(inputDate))
    month = DatePart("m", CStr(inputDate))
    year = DatePart("yyyy", CStr(inputDate))
    If testLeapYear(year) = 1 Then
        If month > 2 Then
            prevMonth = month - 1
            solution = ((31 * (prevMonth)) + Day) - (((4 * month) + 23) / 10)
            dayNum = Math.Round(solution) + 1
            MsgBox ("The day number is " & dayNum)
        End If
        If month <= 2 Then
            prevMonth = month - 1
            If month = 1 Then
                prevMonth = 12
            End If
            dayNum = (31 * (prevMonth)) + Day + 1
            MsgBox ("The day number is " & dayNum)
        End If
    Else
        If month > 2 Then
            prevMonth = month - 1
            solution = ((31 * (prevMonth)) + Day) - (((4 * month) + 23) / 10)
            dayNum = Math.Round(solution)
            MsgBox ("The day number is " & dayNum)
        End If
        If month <= 2 Then
            prevMonth = month - 1
            If month = 1 Then
                prevMonth = 12
            End If
            dayNum = (31 * (prevMonth)) + Day
            MsgBox ("The day number is " & dayNum)
        End If
    End If
End Sub
Function testLeapYear(year) As Integer
    If (year Mod 4 = 0) And (year Mod 100 <> 0) Or (year Mod 400 = 0) Then
        testLeapYear = 1 '1 denotes being a leap year
    Else
        testLeapYear = 0 '0 denotes not being a leap year
    End If
End Function
