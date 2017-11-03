'This code was initially written at extramiledata.com and modified by Fabusuyi Ayodeji
'This code permits user to enter a number which is in turn split into a set of numbers. Upon spliting,
'repeated numbers are checked for. E.g 224546 is split into the set [2,2,3,5,4,6], hence, 2 and 4 are the duplicate numbers
'If there are no duplicates, the word "No Duplicate" is printed
'feel free to modify this code
Option Explicit
Sub test()
    Dim arr() As String
    Dim number, i, j, temp As Integer
    Dim name As String
    number = InputBox("Enter number")
    name = CStr(number)
    arr = toCharArray(name)
    MsgBox (DuplicatesInArray(arr))
End Sub
Function toCharArray(ByRef sIn As String) As String()
    toCharArray = Split(StrConv(sIn, vbUnicode), Chr(0))
End Function

Public Function DuplicatesInArray(ArrayOfValues) As String
' This function checks to see if there are duplicate values in the
' ArrayOfValues argument, which is an array.  If there are, it returns
' an unsorted, comma+space separated list of the duplicated values.
' If there are no duplicates, it returns a blank string, "".  The
' function ignores Nulls.

' DuplicatesInArray() Version 1.0.0
' Copyright © 2009 Extra Mile Data, www.extramiledata.com.
' For questions or issues, please contact support@extramiledata.com.
' Use (at your own risk) and modify freely as long as proper credit is given.

On Error GoTo Err_DuplicatesInArray

    Dim intUB As Integer
    Dim intElem As Integer
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim varValue
    Dim varLoop
    Dim strResults As String
    
    ' Get the upper bound of the array.
    intUB = UBound(ArrayOfValues)
    ' Initialize the variable that holds the results.
    strResults = ""
    
    ' Loop through the array of values, examining each value.
    For intElem = 0 To intUB
        ' Initialize the count of occurrences.
        intCount = 0
        ' Get the value that we're working with.
        varValue = ArrayOfValues(intElem)
        ' If the value is not Null, then continue.  We're ignoring
        ' Null values.
        If Not IsNull(varValue) Then
            ' Now that we have the value that we are checking,
            ' loop through the array and compare the value with all
            ' the other values.
            For intLoop = 0 To intUB
                ' Get the next value in the array.
                varLoop = ArrayOfValues(intLoop)
                ' We are ignoring Nulls, but if it is not null, and
                ' it matches the value that we are checking for, then
                ' increment the counter.
                If Not IsNull(varLoop) Then
                    If varLoop = varValue Then
                        intCount = intCount + 1
                    End If
                End If
            Next intLoop
            ' We would expect a count of 1, the value itself.  If the
            ' count is greater than 1, then there is a duplicate.  If
            ' we have not already listed the duplicate, then add it
            ' to the results.
            If intCount > 1 Then
                If InStr(strResults, varValue & ", ") = 0 Then
                    strResults = strResults & varValue & ", "
                End If
            End If
        End If
    Next intElem

    ' If there were some duplicates, then strip off the last
    ' comma+space and pass back the results.  If there were no
    ' duplicates, then pass back a blank string.
    If Len(strResults) > 0 Then
        DuplicatesInArray = Left(strResults, Len(strResults) - 2)
    Else
        DuplicatesInArray = "No Duplicate"
    End If

Exit_DuplicatesInArray:
    On Error Resume Next
    Exit Function
    
Err_DuplicatesInArray:
    MsgBox Err.number & " " & Err.Description, vbCritical, "DuplicatesInArray()"
    DuplicatesInArray = ""
    Resume Exit_DuplicatesInArray
End Function

