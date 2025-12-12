Option Explicit

' long-таблица: модификации по СТРОКАМ, параметры по СТОЛБЦАМ
Sub WideToLong_FromInputRanges_RC()
    Dim srcWs As Worksheet
    Dim wb As Workbook
    Dim rngMods As Range      ' диапазон с МОДИФИКАЦИЯМИ (столбец, много строк)
    Dim rngParams As Range    ' диапазон с НАИМЕНОВАНИЯМИ ПАРАМЕТРОВ (строка, много столбцов)
    Dim rngValues As Range    ' диапазон со ЗНАЧЕНИЯМИ (матрица: строки = модификации, столбцы = параметры)
    
    Dim wsOut As Worksheet
    Dim outName As String
    Dim r As Long, c As Long
    Dim outRow As Long
    Dim modName As String
    Dim paramName As String
    Dim valText As String
    
    ' --- выбираем диапазоны ---
    On Error Resume Next
    
    Set rngMods = Application.InputBox( _
        Prompt:="Выделите диапазон с МОДИФИКАЦИЯМИ (один столбец, по СТРОКАМ).", _
        Title:="Диапазон модификаций (по строкам)", _
        Type:=8)
    If rngMods Is Nothing Then Exit Sub
    
    ' ВАЖНО: после выбора rngMods определяем лист и книгу ИЗ ЭТОГО диапазона
    Set srcWs = rngMods.Worksheet
    Set wb = srcWs.Parent
    
    Set rngParams = Application.InputBox( _
        Prompt:="Выделите диапазон с НАИМЕНОВАНИЯМИ ПАРАМЕТРОВ (одна строка, по СТОЛБЦАМ).", _
        Title:="Диапазон параметров (по столбцам)", _
        Type:=8)
    If rngParams Is Nothing Then Exit Sub
    
    Set rngValues = Application.InputBox( _
        Prompt:="Выделите диапазон со ЗНАЧЕНИЯМИ (матрица: строки = модификации, столбцы = параметры).", _
        Title:="Диапазон значений", _
        Type:=8)
    If rngValues Is Nothing Then Exit Sub
    
    On Error GoTo 0
    
    ' --- проверки размеров ---
    ' число модификаций (строк) = число строк матрицы
    If rngMods.Rows.Count <> rngValues.Rows.Count Then
        MsgBox "Количество модификаций (строк) и строк в матрице значений не совпадает!", vbCritical
        Exit Sub
    End If
    
    ' число параметров (столбцов) = число столбцов матрицы
    If rngParams.Columns.Count <> rngValues.Columns.Count Then
        MsgBox "Количество параметров (столбцов) и столбцов в матрице значений не совпадает!", vbCritical
        Exit Sub
    End If
    
    ' --- создаём лист <имя_текущего_листа>_long В ЭТОЙ ЖЕ КНИГЕ ---
    outName = srcWs.Name & "_long"
    If Len(outName) > 31 Then
        outName = Left(outName, 31)
    End If
    
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(outName).Delete   ' если лист уже был в этой книге — удалим
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsOut = wb.Worksheets.Add(After:=srcWs)
    wsOut.Name = outName
    
    ' --- заголовки ---
    wsOut.Cells(1, 1).Value = "Modification"
    wsOut.Cells(1, 2).Value = "ParamName"
    wsOut.Cells(1, 3).Value = "Value"
    
    outRow = 2
    
    ' --- основной цикл ---
    ' r — по модификациям (строки), c — по параметрам (столбцы)
    For r = 1 To rngMods.Rows.Count
        modName = CStr(rngMods.Cells(r, 1).Value)
        
        For c = 1 To rngParams.Columns.Count
            paramName = CStr(rngParams.Cells(1, c).Value)
            valText = CStr(rngValues.Cells(r, c).Value)
            
            If Len(Trim(modName)) > 0 Or Len(Trim(paramName)) > 0 Or Len(Trim(valText)) > 0 Then
                wsOut.Cells(outRow, 1).Value = modName
                wsOut.Cells(outRow, 2).Value = paramName
                wsOut.Cells(outRow, 3).Value = valText
                outRow = outRow + 1
            End If
        Next c
    Next r
    
    MsgBox "Готово! Результат на листе: " & wsOut.Name & " в книге " & wb.Name, vbInformation
End Sub
