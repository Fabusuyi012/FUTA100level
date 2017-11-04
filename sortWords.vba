'To sort a set of given text alphabetically
Sub sortABC()
    Dim x, y, i As Long
    Dim tempTxt1 As String
    numWord = InputBox("Enter words seperated by a comma and no space e.g Name1,Name2,Name5 ")
    myArray = Split(numWord, ",")
    For x = LBound(myArray) To UBound(myArray)
        For y = x To UBound(myArray)
            If UCase(myArray(y)) < UCase(myArray(x)) Then
                tempTxt1 = myArray(x)
                tempTxt2 = myArray(y)
                myArray(x) = tempTxt2
                myArray(y) = tempTxt1
            End If
        Next y
    Next x
        MsgBox Join(myArray, vbCrLf)
End Sub

