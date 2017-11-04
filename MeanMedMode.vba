'To determine mean, median and mode of a set of numbers
Sub getStuffs()
    Dim arr() As String
    Dim numArr() As Integer
    Dim i As Integer
    Dim numWord, tempTxt1, tempTxt2 As String
    Dim x, y, k As Long
    numWord = InputBox("Enter numbers seperated by a comma to make a set in the form 1,2,3,4 to make up the set [1,2,3,4]")
    arr = Split(numWord, ",")
    For x = LBound(arr) To UBound(arr)
        For y = x To UBound(arr)
            If UCase(arr(y)) < UCase(arr(x)) Then
                tempTxt1 = arr(x)
                tempTxt2 = arr(y)
                arr(x) = tempTxt2
                arr(y) = tempTxt1
            End If
        Next y
    Next x
    ReDim numArr(UBound(arr))
    For i = LBound(arr) To UBound(arr)
        numArr(i) = CInt(arr(i))
    Next i
    MsgBox ("The mean is " & getMean(numArr))
    MsgBox ("The median is " & getMedian(numArr))
    MsgBox ("The mode is " & getMode(numArr))
End Sub
Function getMean(arr() As Integer)
    Dim setSum, setTotal, i As Integer
    setTotal = UBound(arr) + 1
    setSum = WorksheetFunction.sum(arr)
    getMean = setSum / setTotal
End Function
Function getMedian(arr() As Integer)
    Dim n As Integer
    n = UBound(arr) + 1
    If n Mod 2 = 0 Then
        getMedian = arr(((n / 2) + 1) - 1) + arr((n / 2) - 1)
    Else
        getMedian = arr(((n + 1) / 2) - 1)
    End If
End Function
Function getMode(arr() As Integer)
        getMode = WorksheetFunction.mode(arr)
End Function




