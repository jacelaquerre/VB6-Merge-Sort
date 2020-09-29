Attribute VB_Name = "mergeSort"
' Jace Laquerre
' Merge Sort in Visual Basic 6.0
Option Explicit

Public Type ArraySortElement
    Index As Long
    sort_str As String
    End Type
    
Public Sub mergeSort(ByRef ase() As ArraySortElement, array_start&, array_end&, sort_dir&, flg&)
    ' Merge sort has a best and worst case runtime complexity of O(n Log(n))
    ' Merge sort will work better for larger sets of data, if the data set is small,
    ' (less than about 1750 elements)
    ' Then a sorting algoritm with a better best case time complexity should be used
    '
    ' array_start& = 0 or 1, depending on how we are is handling this array
    ' array_end& the last index of the arrray so call UBound(ase) in function call for this param
    '
    ' sort_dir& = 0, ascending sort
    ' sort_dir$ = 1, decending sort
    '
    ' flg&   1 = case insensitive
    
    Dim middle&         ' the middle index of the array
    Dim i&
    On Error GoTo MergeSort_Error
    
    If (flg And 1) = 1 Then ' Apply ucase to sort_strs to for case insensitive sorts
        For i& = array_start& To UBound(ase)
            ase(i&).sort_str = UCase$(ase(i&).sort_str)
            Next i&
        End If
        
    If (array_end& > array_start&) Then
        ' Recursively sort the two halves of the list.
        middle& = (array_start& + array_end&) \ 2
        mergeSort ase, array_start&, middle&, sort_dir&, flg&
        mergeSort ase, (middle& + 1), array_end&, sort_dir&, flg&

        ' Merge the results.
        Merge ase, array_start&, middle&, array_end&, sort_dir&, flg&
        End If

    Exit Sub
MergeSort_Error:
    Call ErrDump(1, "MergeSort, Line: " & Information.Erl & " of Module ArraySort")
    Resume
End Sub

Public Sub Merge(ByRef ase() As ArraySortElement, start&, middle&, last&, sort_dir&, flg&)
    ' Part of MergeSort
    Dim i&, j&, k&
    Dim n1&, n2&
    Dim tleft() As ArraySortElement
    Dim tright() As ArraySortElement
    On Error GoTo Merge_Error
    
    ' reDim temp arrays
    n1& = middle& - start& + 1
    n2& = last& - middle&
    ReDim tleft(n1&)
    ReDim tright(n2&)
    
    ' Copy data into temp arrays
    Do While (i& < n1&)
        tleft(i&) = ase(start& + i&)
        i& = i& + 1
        Loop
    Do While (j& < n2&)
        tright(j&) = ase(middle& + 1 + j&)
        j& = j& + 1
        Loop
          
    i& = 0
    j& = 0
    k& = start&
    If (sort_dir& = 0) Then ' ASC
            Do While ((i& < n1) And (j < n2))
                If ((flg And 2) = 2) Then
                        If (Val(tleft(i&).sort_str) <= Val(tright(j&).sort_str)) Then
                                ase(k&) = tleft(i&)
                                i& = i& + 1
                            Else
                                ase(k&) = tright(j&)
                                j& = j& + 1
                            End If
                            k& = k& + 1
                    Else
                        If (tleft(i&).sort_str <= tright(j&).sort_str) Then
                                ase(k&) = tleft(i&)
                                i& = i& + 1
                            Else
                                ase(k&) = tright(j&)
                                j& = j& + 1
                            End If
                            k& = k& + 1
                    End If
                Loop
        Else ' DESC
            Do While ((i& < n1) And (j < n2))
                If ((flg And 2) = 2) Then
                        If (Val(tleft(i&).sort_str) >= Val(tright(j&).sort_str)) Then
                                ase(k&) = tleft(i&)
                                i& = i& + 1
                            Else
                                ase(k&) = tright(j&)
                                j& = j& + 1
                            End If
                            k& = k& + 1
                    Else
                        If (tleft(i&).sort_str >= tright(j&).sort_str) Then
                                ase(k&) = tleft(i&)
                                i& = i& + 1
                            Else
                                ase(k&) = tright(j&)
                                j& = j& + 1
                            End If
                            k& = k& + 1
                    End If
                Loop
        End If
        
    Do While (i& < n1&)
        ase(k&) = tleft(i&)
        i& = i& + 1
        k& = k& + 1
        Loop
        
    Do While (j& < n2&)
        ase(k&) = tright(j&)
        j& = j& + 1
        k& = k& + 1
        Loop
    Exit Sub
    
Merge_Error:
    Call ErrDump(1, "Merge, Line: " & Information.Erl & " of Module ArraySort")
    Resume
End Sub
