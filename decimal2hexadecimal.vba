'To convert decimal value into its hexadecimal equivalent
Sub dec2Hexa()
    Dim number, output As Integer
    Dim usable As String
    number = InputBox("Enter the number's decimal value")
    usable = CStr(number)
    MsgBox Evaluate("DEC2HEX(" & usable & ")")
End Sub
