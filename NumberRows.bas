Sub NumberRows(sheetIndex As Integer, col As String, colR As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Установите лист, на котором вы хотите выполнить операцию
    Set ws = ThisWorkbook.Sheets(sheetIndex)

    ' Найдите последнюю заполненную строку в столбце colR
    lastRow = ws.Cells(ws.Rows.Count, colR).End(xlUp).Row

    ' Начнем нумерацию с col строки 3
    For i = 3 To lastRow + 1
        ' Если текущая ячейка в столбце B заполнена или это следующая строка после последней заполненной
        If i <= lastRow Or ws.Range(colR & (i - 1)).Value <> "" Then
            ws.Range(col & i).Value = i - 2 ' Нумерация начинается с 1 в col строки 3
        End If
    Next i
End Sub
