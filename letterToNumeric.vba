'To convert a given name to the sum of the numeric equivalent of its letters
Function toCharArray(ByRef sIn As String) As String()
    toCharArray = Split(StrConv(sIn, vbUnicode), Chr(0))
End Function
Sub test()
    Dim arr() As String
    Dim answer, sum As Integer
    Dim name As String
    name = InputBox("Enter name without space")
    arr = toCharArray(name)
    sum = 0
    For i = LBound(arr) To UBound(arr) - 1
        answer = Asc(UCase(arr(i))) - 64
        sum = sum + answer
    Next i
    MsgBox (sum)
End Sub


