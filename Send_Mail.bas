Sub SendEmail(mail As String, emailBody As String)
    Dim OutlookApp As Object
    Dim OutlookMail As Object

    ' Создание объекта приложения Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    ' Создание нового элемента почты
    Set OutlookMail = OutlookApp.CreateItem(0)

    With OutlookMail
        .To = mail ' Установка адреса получателя
        .CC = "" ' Установка адреса CC, если необходимо
        .BCC = "" ' Установка адреса BCC, если необходимо
        .Subject = "Запрос RC Lab" ' Установка темы письма
        .Body = emailBody ' Установка тела письма
        ' .Attachments.Add "C:\path\to\file.txt" ' Добавление вложения, если необходимо
        .Send ' Отправка письма
    End With

    ' Очистка объектов
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub
