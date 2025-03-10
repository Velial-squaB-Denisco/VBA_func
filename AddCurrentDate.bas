Sub AddCurrentDate(sheetIndex As Integer, col As String)
    Dim ws As Worksheet
    Dim lastRow As Long

    ' Установить лист по имени
    Set ws = ThisWorkbook.Sheets(sheetIndex)

    ' Найти последнюю заполненную строку в указанной колонке
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    ' Вставить текущую дату в следующую пустую строку
    ws.Cells(lastRow + 1, col).Value = Date
End Sub

Sub main()

    Call AddCurrentDate(1, "B") 'B_Date
    
End Sub