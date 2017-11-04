'To find the first n odd numbers
Sub oddNumbers()
    Dim n, count, i As Integer
    Dim numArray() As Integer
    Dim arr() As String
    n = InputBox("Enter the value of n to calculate the first n odd numbers")
    count = 0
    ReDim numArray(n - 1) As Integer
    For i = 1 To 2 * n
        If i Mod 2 <> 0 Then
            numArray(count) = i
            count = count + 1
        End If
    Next i
    ReDim arr(UBound(numArray))
    For i = LBound(numArray) To UBound(numArray)
        arr(i) = CStr(numArray(i))
    Next i
    MsgBox Join(arr, vbCrLf)
End Sub
