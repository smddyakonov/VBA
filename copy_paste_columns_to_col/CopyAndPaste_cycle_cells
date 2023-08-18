Sub CopyAndPaste_cycle_cells()
    
    Dim dataRange As Range 'определяем диапазон данных
    Dim lastRow As Long 'определяем последнюю строку в столбце
    Dim i As Long 'счетчик
    Dim start_cell As String 'стартовая ячейка
    Dim finish_cell As String 'финишная ячейка
    Dim start_row As Long 'стартовая строка
    Dim finish_row As Long 'финишная строка
    Dim finish_col As String 'стартовый столбец
    Dim start_col As String 'финишный столбец

    
    'замените на нужный диапазон
    start_cell = "A1"
    finish_cell = "C3"
    start_col = Range(start_cell).Column
    start_row = Range(start_cell).Row
    finish_col = Range(finish_cell).Column
    finish_row = Range(finish_cell).Row
    
    For i = start_col To finish_col
    
        Set dataRange = Range(Cells(start_row, i), Cells(finish_row, i))

        lastRow = Cells(Rows.Count, "A").End(xlUp).Row 'последняя заполненная строка в столбце "A"
        lastCell = "A" + CStr(lastRow + 1) 'ячейка, с которой будет заполненение в столбце "A"
    
        Range(lastCell).Value = "#" 'вставляем символ "#" разделитель
        
        dataRange.Copy 'копируем данные в буфер обмена
    
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row 'последняя заполненная строка в столбце "A"
    
        Range("A" & lastRow + 1).PasteSpecial xlPasteValues 'вставляем данные в столбец "A", начиная со следующей строки после последней заполненной
        
    Next i
    
End Sub
