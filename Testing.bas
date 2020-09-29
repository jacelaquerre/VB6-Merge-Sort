Attribute VB_Name = "Testing"
Option Explicit

Public Sub Main()
    Dim i&
    Dim arr(8) As Long
  
    arr(0) = -4
    arr(1) = 2
    arr(2) = 3
    arr(3) = 1
    arr(4) = 5
    arr(4) = 9
    arr(5) = 0
    arr(6) = 99
    arr(7) = 15
    arr(8) = 93
    
    Dim result$
    result$ = LongArrToString(arr, ", ", 0)
    Debug.Print (result$)
    
    ReDim marrayA(UBound(arr)) As ArraySortElement
    For i& = 0 To UBound(arr)
        marrayA(i&).Index = i&
        marrayA(i&).sort_str = CStr(arr(i&))
        Next i&
    Call MergeSort(marrayA, 0, (UBound(arr)), 0, 2)
    
    result$ = ElementArrToString(marrayA, ", ", 0)
    Debug.Print (result$)
    
    ReDim marrayD(UBound(arr)) As ArraySortElement
    For i& = 0 To UBound(arr)
        marrayD(i&).Index = i&
marrayD(i&).sort_str = CStr(arr(i&))
        Next i&
    Call MergeSort(marrayD, 0, (UBound(arr)), 1, 2)

    result$ = ElementArrToString(marrayD, ", ", 0)
    Debug.Print (result$)
    
End Sub

Public Sub ErrDump(tl&, msg$)
    'Standard error handler. Display the error to the user
    'tl&    - Required by the library. In this program, it does nothing.
    'msg$   - Error message to be displayed and sent.
    
    Call MsgBox("An error has occurred. Details below:" & vbCrLf & msg$, vbOKOnly + vbCritical)
    Debug.Assert False
    Close
    End
End Sub

Public Function LongArrToString(arr() As Long, sep$, array_start&) As String
    ' arr is the incoming array of Longs
    ' sep is what sperates the elements of the array
    ' Examples could be ",", " ", or "-"
' array_start& = 0 or 1, depending on how we are handling this array
    Dim i&
    Dim result$ ' stores the array as a string to be returned
    
    result$ = ""
    For i& = array_start& To UBound(arr)
        If (i& = array_start&) Then
                result$ = arr(i&)
            Else
                result$ = result$ & sep$ & arr(i&)
        End If
        Next i&
     LongArrToString = result$
End Function

Public Function StrArrToString(arr() As String, sep$, array_start&) As String
    ' arr is the incoming array of Strings
    ' sep is what sperates the elements of the array
    ' Examples could be ",", " ", or "-"
    ' array_start& = 0 or 1, depending on how we are is handling this array
    Dim i&
    Dim result$ ' stores the array as a string to be returned
    
    result$ = ""
    For i& = array_start& To UBound(arr)
        If (i& = array_start&) Then
                result$ = arr(i&)
            Else
                result$ = result$ & sep$ & arr(i&)
        End If
        Next i&
     StrArrToString = result$
End Function

Public Function ElementArrToString(arr() As ArraySortElement, sep$, array_start&) As String
    ' arr is the incoming array of ArraySortElement's
    ' sep is what sperates the elements of the array
    ' Examples could be ",", " ", or "-"
    ' array_start& = 0 or 1, depending on how we are is handling this array
    Dim i&
    Dim result$ ' stores the array as a string to be returned
    
    result$ = ""
    For i& = array_start& To UBound(arr)
        If (i& = array_start&) Then
                result$ = arr(i&).sort_str
            Else
                result$ = result$ & sep$ & arr(i&).sort_str
        End If
        Next
     ElementArrToString = result$
End Function
