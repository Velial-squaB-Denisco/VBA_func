Sub CheckLastRow(sheetIndex As Integer, col1 As String, col2 As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Integer
    Dim isEmpty As Boolean
    Dim emptyColumns As String
    Dim colNum1 As Integer
    Dim colNum2 As Integer

    ' Преобразование букв столбцов в числа
    colNum1 = Range(col1 & "1").Column
    colNum2 = Range(col2 & "1").Column

    ' Установите лист, на котором выполняется проверка
    Set ws = ThisWorkbook.Sheets(sheetIndex)

    ' Найти последнюю строку в столбце col1
    lastRow = ws.Cells(ws.Rows.Count, colNum1).End(xlUp).Row
    If lastRow < 3 Then lastRow = 3

    ' Проверка строки на наличие пустых ячеек
    isEmpty = False
    emptyColumns = ""

    For i = colNum1 To colNum2 ' Проверяем столбцы от col1 до col2
        If ws.Cells(lastRow, i).Value = "" Then
            isEmpty = True
            emptyColumns = emptyColumns & Chr(64 + i) & " "
        End If
    Next i

    ' Если все ячейки заполнены, закрасить строку зелёным
    If Not isEmpty Then
        ws.Range(ws.Cells(lastRow, colNum1), ws.Cells(lastRow, colNum2)).Interior.Color = RGB(0, 255, 0)
    Else
        ' Если есть пустые ячейки, закрасить строку жёлтым и вывести сообщение о столбцах с пустыми ячейками
        ws.Range(ws.Cells(lastRow, colNum1), ws.Cells(lastRow, colNum2)).Interior.Color = RGB(255, 255, 0)
        MsgBox "Пустые ячейки в столбцах: " & emptyColumns
    End If
End Sub

Sub main()

    Call CheckLastRow(1, "A", "O")
    
End Sub
