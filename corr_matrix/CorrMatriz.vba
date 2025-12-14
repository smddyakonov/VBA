'function to create a correlation matrix given the data
Function CorrMatriz(Mat_data As Variant)
    Dim i As Integer
    Dim j As Integer
    Dim corr As Variant
    Dim M1 As Variant
    Dim M2 As Variant
    ReDim corr(1 To Mat_data.Columns.Count, 1 To Mat_data.Columns.Count)

    ReDim M1(1 To Mat_data.Rows.Count, 1 To 1)
    ReDim M2(1 To Mat_data.Rows.Count, 1 To 1)

    For i = 1 To Mat_data.Columns.Count
        M1 = ExtraeMatriz(Mat_data, i)
        For j = 1 To Mat_data.Columns.Count
            M2 = ExtraeMatriz(Mat_data, j)
            corr(i, j) = Application.Correl(M1, M2)
        Next j
    Next i
    CorrMatriz = corr
End Function

' function to extract one column
Function ExtraeMatriz(Matriz As Variant, columna As Integer)
    Dim i As Integer
    Dim data_final As Variant
    ReDim data_final(1 To Matriz.Rows.Count, 1)
    
    For i = 1 To Matriz.Rows.Count
    data_final(i, 1) = Matriz(i, columna)
    Next i
    
    ExtraeMatriz = data_final
End Function
