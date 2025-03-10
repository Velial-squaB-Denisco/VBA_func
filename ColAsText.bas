Sub ColAsText(sheetIndex As Integer, col As String, startRow As Long)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetIndex) ' Укажите название вашего листа
    
    With ws.Range(col & startRow & ":" & col & ws.Rows.Count)
        .NumberFormat = "@" ' Задаем формат как текст
    End With
    
End Sub

Sub main()

    Call ColAsText(1, "K", 3)
    
End Sub