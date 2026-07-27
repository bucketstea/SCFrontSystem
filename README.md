Private Function RangeToArray1D(ByVal targetRange As Range) As Variant

    Dim source As Variant
    Dim result() As Variant
    Dim i As Long

    source = targetRange.Value

    ' 1セルだけの場合
    If targetRange.Cells.CountLarge = 1 Then
        RangeToArray1D = Array(source)
        Exit Function
    End If

    ReDim result(1 To UBound(source, 1))

    For i = 1 To UBound(source, 1)
        result(i) = source(i, 1)
    Next i

    RangeToArray1D = result

End Function
