Function MeltData(rng As Range, id_vars As Variant, value_vars As Variant) As Variant
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
    Dim output() As Variant
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

    ' Возвращаем заполненный выходной массив как результат функции
    MeltData = output
End Function

Комментарии к коду:

- Определение переменных: Мы определяем переменные для работы с диапазоном, такие как количество строк и столбцов.
- Выходной массив: Создается выходной массив, который будет содержать преобразованные данные.
- Заполнение заголовков: В первой строке выходного массива добавляются заголовки для идентификаторов, переменных и значений.
- Циклы для заполнения данных: Мы проходим по всем строкам исходного диапазона и заполняем выходной массив соответствующими значениями.
- Возврат результата: В конце функция возвращает заполненный массив.

Dim result As Variant
result = MeltData(Sheet1.Range("A1:C5"), Array(1), Array(2, 3))

Как использовать эту функцию:

1. Откройте Excel и нажмите Alt + F11, чтобы открыть редактор VBA.
2. Вставьте новый модуль: Insert > Module.
3. Скопируйте и вставьте приведенный выше код в модуль.
4. Вернитесь в Excel и используйте функцию в виде массива.

▎Пример вызова функции:
Если у вас есть данные в диапазоне A1:C5, где:
- Столбец A — идентификатор,
- Столбец B и C — значения
