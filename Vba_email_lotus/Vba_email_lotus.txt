Sub SendEmailsFromSelection_Solver()

    Dim objNotes As Object
    Dim objNotesDB As Object
    Dim objNotesMailDoc As Object
    Dim objAttachment As Object
    Dim i As Long
    Dim ws As Worksheet
    Dim selectedRange As Range

    ' Проверка на выделение диапазона
    On Error Resume Next
    Set selectedRange = Selection
    On Error GoTo 0

    If selectedRange.Columns(selectedRange.Columns.Count).Column <> 7 Then
        MsgBox "Выделите 7 колонок: subj, msg, sendTo, copyTo, blindCopyTo,
pth_file, отметка"
        Exit Sub
    End If

    ' Установка ссылки на текущий лист
    Set ws = ActiveSheet

    ' Создание объекта Notes
    Set objNotes = CreateObject("Notes.NotesSession")
    Set objNotesDB = objNotes.GetDatabase("", "")
    Call objNotesDB.OPENMAIL

    ' Перебор строк в выделенном диапазоне
    For Each Cell In selectedRange.Rows
        i = Cell.Row

        ' Получение значений из каждой колонки
        Dim subj As String
        Dim msg As String
        Dim sendTo As String
        Dim copyTo As String
        Dim blindCopyTo As String
        Dim pth_file As String
        Dim startTime As Date
        Dim endTime As Date
        Dim executionTime As String

        Dim privetstvie As String
        Dim podpis As String
        Dim obrachenie As String
        'Dim blagodarnost As String


        ' Проверка содержимого ячейки в седьмой колонке перед созданием
документа
        If InStr(1, ws.Cells(i, 7).Value, "Отправлено на репликацию") > 0
Then
            MsgBox "Внимание! Письмо" & " " & ws.Cells(i, 1).Value & " " &
ws.Cells(i, 7).Value
        Else

            subj = ws.Cells(i, 1).Value ' Тема письма
            If subj = "" Then
                MsgBox "Внимание! Макрос завершен с ошибкой: письмо без
темы, проверьте столбец subj"
                Exit Sub
            End If

            msg = ws.Cells(i, 2).Value ' Текст письма
            If msg = "" Then
                MsgBox "Внимание! Макрос завершен с ошибкой: письмо без
сообщения, проверьте столбец msg"
                Exit Sub
            End If

            sendTo = ws.Cells(i, 3).Value ' Адрес отправки
            If sendTo = "" Then
                sendTo = InputBox(ws.Cells(i, 1).Value & " " & "введите
e-mail:")
                ws.Cells(i, 3).Value = sendTo
            End If

            copyTo = ws.Cells(i, 4).Value ' Адрес для копии
            blindCopyTo = ws.Cells(i, 5).Value ' Адрес для скрытой копии
            pth_file = ws.Cells(i, 6).Value ' Путь к файлу

            privetstvie = ws.Range("D1").Value
            podpis = ws.Range("D2").Value
            blagodarnost = ws.Range("D3").Value
            obrachenie = ws.Cells(i, 15).Value

            ' Создание нового документа
            Set objNotesMailDoc = objNotesDB.CreateDocument

            ' Заполнение параметров письма
            With objNotesMailDoc
                .DeliveryReport = "a" 'Запрос уведомления о доставке
адресату
                .ReturnReceipt = "1" 'Запрос уведомления о прочтении письма
адресатом
                .Subject = subj 'Тема

                'Сообщение
                .Body = privetstvie & ", " & obrachenie & "!" & vbCrLf &
vbCrLf _
                                                    & msg & vbCrLf & vbCrLf
_
                                                    & podpis

                .SaveMessageOnSend = True 'Сохранять или нет в папке
"Отправленные"
                '.SignOnSend = True 'цифровая подписывать
                .Importance = "2" 'важность док-та(Высокая = 1, Обычная =
2, Низкая = 3)
                '.EncryptOnSend = True 'шифровать
                .Form = "Memo"
                .sendTo = sendTo
                .copyTo = copyTo
                .blindCopyTo = blindCopyTo

                ' Добавление вложения, если указан путь к файлу
                If pth_file <> "" Then
                    Set objAttachment = .CreateRichTextItem("Attachment")
                    Call objAttachment.EmbedObject(1454, "", pth_file)
                End If

                ' Запуск таймера
                startTime = Now()

                ' Отправка письма
                Call .Send(False)

                ' Остановка таймера и рассчет времени выполнения
                'endTime = Now()
                'executionTime = Format(endTime - startTime, "hh:mm:ss")

                ' Запись времени выполнения в колонку 7
                ws.Cells(i, 7).Value = "Отправлено на репликацию" + " " +
Format(startTime, "dd.mm.yyyy hh:mm:ss")
            End With

            ' Освобождение ресурсов
            Set objNotesMailDoc = Nothing
            Set objAttachment = Nothing

            'MsgBox "Макрос отработал без ошибок!"

        End If

    Next Cell

    ' Освобождение ресурсов Lotus Notes
    Set objNotes = Nothing
    Set objNotesDB = Nothing



End Sub
