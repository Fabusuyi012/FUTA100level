'To sort a given set of text alphabetically
Sub sortABC()
    Dim x, y, i As Long
    Dim tempTxt1 As String
    Dim tempTxt2 As String
    myArray = Array("Myname", "Yourname", "Ournames", "FUTA", "Lastname") 'array that contains the words to be sorted and whose contents can be modified
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
    For i = LBound(myArray) To UBound(myArray)
        MsgBox (myArray(i))
    Next i
End Sub

