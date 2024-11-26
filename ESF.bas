Attribute VB_Name = "Module1"
Sub ExtractESFNumbers()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim startPos As Long
    Dim endPos As Long
    Dim extractedText As String
    
    ' Установите ссылку на активный лист
    Set ws = ActiveSheet
    ' Укажите диапазон столбца D (номер столбца 4)
    Set rng = ws.Range("D1:D" & ws.Cells(ws.Rows.Count, 4).End(xlUp).Row)
    
    ' Обрабатываем каждую ячейку
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            ' Ищем начало текста "ESF-"
            startPos = InStr(cell.Value, "ESF-")
            
            If startPos > 0 Then
                ' Ищем конец текста (первую закрывающую скобку после номера)
                endPos = InStr(startPos, cell.Value, ")")
                If endPos > 0 Then
                    ' Извлекаем текст между "ESF-" и закрывающей скобкой
                    extractedText = Mid(cell.Value, startPos, endPos - startPos)
                Else
                    ' Если закрывающая скобка не найдена, берем текст до конца строки
                    extractedText = Mid(cell.Value, startPos)
                End If
                
                ' Заменяем содержимое ячейки на найденный текст
                cell.Value = extractedText
            Else
                ' Если "ESF-" не найден, очищаем ячейку
                cell.Value = ""
            End If
        End If
    Next cell

    MsgBox "Обработка завершена!", vbInformation
End Sub

