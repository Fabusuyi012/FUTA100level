'To determine whether a number is a palindrome
Sub palindrome()
    Dim word, rev As String
    Dim number As Long
    number = InputBox("Enter a number")
    word = CStr(number)
    rev = StrReverse(word)
    If word = rev Then
        MsgBox (number & " is a palindrome")
    End If
    If word <> rev Then
        MsgBox (number & " is not a palindrome")
    End If
End Sub
