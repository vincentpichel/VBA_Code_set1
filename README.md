# VBA_Code_set1
Check a range of variables for a combination which is equal to a specified total

Sub Test_AllSumsForTotalFromSet()
    Dim numberSet, total As Long, result As Collection

    'Will need to modify numberset to a specified column or set of cells
    numberSet = Array(65536, 131072, 262144, 524288, 104576, 2097152)
    'numberSet = Range("I2:I7").value
    total = Range("G2")

    Set result = GetAllSumsForTotalFromSet(total, numberSet)

    Debug.Print "Possible sums: " & result.Count

    PrintResult result
End Sub

Function GetAllSumsForTotalFromSet(total As Long, ByRef numberSet As Variant) As Collection
    Set GetAllSumsForTotalFromSet = New Collection
    Dim partialSolution(1 To 1) As Long

    Set GetAllSumsForTotalFromSet = AllSumsForTotalFromSet(total, numberSet, UBound(numberSet), partialSolution)
End Function

Function AllSumsForTotalFromSet(total As Long, ByRef numberSet As Variant, numberSetIndex As Long, ByRef partialSolution() As Long) As Collection
    Dim index As Variant, number As Long, result As Collection

    Set AllSumsForTotalFromSet = New Collection

    'break if numberSetIndex is too small
    If numberSetIndex < LBound(numberSet) Then Exit Function

    For index = numberSetIndex To LBound(numberSet) Step -1
        number = numberSet(index)

        If number <= total Then
            'append the number to the partial solution
            partialSolution(UBound(partialSolution)) = number

            If number = total Then
                AllSumsForTotalFromSet.Add partialSolution

            Else
                'For the following line, set "index - 1" to "index" should you need to find sums using the same values more than once
                Set result = AllSumsForTotalFromSet(total - number, numberSet, index - 1, CopyAndReDimPlus1(partialSolution))
                AppendCollection AllSumsForTotalFromSet, result
            End If
        End If
    Next index
End Function



'copy the passed array and increase the copy's size by 1
Function CopyAndReDimPlus1(ByVal sourceArray As Variant) As Long()
    Dim i As Long, destArray() As Long
    ReDim destArray(LBound(sourceArray) To UBound(sourceArray) + 1)

    For i = LBound(sourceArray) To UBound(sourceArray)
        destArray(i) = sourceArray(i)
    Next i

    CopyAndReDimPlus1 = destArray
End Function

'append sourceCollection to destCollection
Sub AppendCollection(ByRef destCollection As Collection, ByRef sourceCollection As Collection)
    Dim e
    For Each e In sourceCollection
        destCollection.Add e
    Next e
End Sub

Sub PrintResult(ByRef result As Collection)
    Dim r, a

    For Each r In result
        For Each a In r
            Debug.Print a;
        Next
        Debug.Print
    Next
End Sub

Sub SetArray()

Dim smsarray As Variant
Dim i As Long
        
With Worksheets("Data")
    smsarray = .Range("B17", .Range("B" & Rows.Count).End(xlUp)).Value

        
        For i = LBound(smsarray) To UBound(smsarray)
            If smsarray(i, 1) <> vbNullString Then
                MsgBox smsarray(i, 1)
            End If
        Next i
    End With
    
    
End Sub
