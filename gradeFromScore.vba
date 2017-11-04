'To determine grade corresponding to user score input
Sub getGrades()
    Dim score As Integer
    Dim grade As String
    score = InputBox("Enter the score")
    If score < 0 Or score > 100 Then
        MsgBox ("You entered an invalid score")
    End If
    If score > 0 And score <= 100 Then
        If score >= 70 Then
            grade = "A"
        End If
        If score >= 60 And score < 70 Then
            grade = "B"
        End If
        If score >= 50 And score < 60 Then
            grade = "C"
        End If
        If score >= 45 And score < 50 Then
            grade = "D"
        End If
        If score < 45 Then
            grade = "F"
        End If
            MsgBox ("Your corresponding grade is " & grade)
    End If
End Sub

