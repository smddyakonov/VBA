Sub Mcorrel()
    'В MS Excel должен быть включен пакет "Анализ данные"
    Dim data_str As String
    Dim out_cell_str As String
    Dim groub_str As String
    Dim metki As Boolean
    
    data_str = "A2:D21" 'данные
    out_cell_str = "F18" 'выходная ячейка
    groub_str = "К" '"К" - колонки, кирилица для русской версии MS Excel, "С" - строки, кирилица для русской версии MS Excel
    metki = False '"False" - без меток, True - с метками

    Application.Run "ATPVBAEN.XLAM!Mcorrel", ActiveSheet.Range(data_str), ActiveSheet.Range(out_cell_str), groub_str, metki
    
End Sub
