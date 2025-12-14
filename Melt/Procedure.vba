Sub MeltData(rng As Range, id_vars As Variant, value_vars As Variant, ByRef output() As Variant)
    Dim ws As Worksheet
    Set ws = rng.Worksheet

    Dim rowCount As Long
    Dim colCount As Long
    Dim idCount As Long
    Dim valueCount As Long

    ' Определяем количество строк и столбцов в исходном диапазоне
    rowCount = rng.Rows.Count
    colCount = rng.Columns.Count

    ' Определяем количество идентификаторов и значений
    idCount = UBound(id_vars) - LBound(id_vars) + 1
    valueCount = UBound(value_vars) - LBound(value_vars) + 1

    ' Создаем выходной массив для результата
    ReDim output(1 To rowCount * valueCount, 1 To idCount + 2)

    Dim i As Long, j As Long, outputRow As Long
    outputRow = 1

    ' Заголовки для выходного массива
    For i = LBound(id_vars) To UBound(id_vars)
        output(1, i + 1) = rng.Cells(1, id_vars(i)).Value ' Копируем заголовки идентификаторов
    Next i
    output(1, idCount + 1) = "variable" ' Заголовок для переменной
    output(1, idCount + 2) = "value" ' Заголовок для значения

    ' Заполняем выходной массив данными
    For i = 2 To rowCount ' Начинаем с 2, чтобы пропустить заголовок
        For j = LBound(value_vars) To UBound(value_vars)
            outputRow = outputRow + 1 ' Переходим к следующей строке в выходном массиве
            For k = LBound(id_vars) To UBound(id_vars)
                output(outputRow, k + 1) = rng.Cells(i, id_vars(k)).Value ' Копируем идентификаторы
            Next k
            output(outputRow, idCount + 1) = rng.Cells(1, value_vars(j)).Value ' Имя переменной
            output(outputRow, idCount + 2) = rng.Cells(i, value_vars(j)).Value ' Значение переменной
        Next j
    Next i
End Sub

 Параметр output: Мы добавили параметр ByRef output() As Variant, который будет использоваться для передачи выходного массива. Он должен быть объявлен заранее в вызывающем коде.

2. Удаление оператора Return: Вместо возврата массива через оператор Return, мы просто заполняем переданный массив output.

▎Пример вызова процедуры:

Чтобы вызвать эту процедуру и получить результат, вы можете использовать следующий код:

Sub TestMeltData()
    Dim result() As Variant
    Dim idVars As Variant
    Dim valueVars As Variant

    ' Пример идентификаторов и значений (измените под свои данные)
    idVars = Array(1, 2) ' Индексы столбцов с идентификаторами
    valueVars = Array(3, 4) ' Индексы столбцов со значениями

    ' Вызываем процедуру
    Call MeltData(Sheet1.Range("A1:D10"), idVars, valueVars, result)

    ' Выводим результат на новый лист (например)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Range("A1").Resize(UBound(result, 1), UBound(result, 2)).Value = result
End Sub


Этот код создает новый лист и выводит заполненный массив result на него. Не забудьте изменить диапазон и индексы в соответствии с вашими данными!
