Option Base 1
    Function MCorrelation(rango As Range) As Variant
        Dim x As Variant
        Dim y As Variant
        Dim s As Integer
        Dim t As Integer
        Dim c() As Variant
        
        ReDim c(rango.Columns.Count, rango.Columns.Count)
        
        For i = 1 To rango.Columns.Count Step 1
            For j = 1 To i Step 1
                c(i, j) = Application.Correl(Application.Index(rango, , i), Application.Index(rango, , j))
            Next j
        Next i
        
        MCorrelation = c
        
    End Function
