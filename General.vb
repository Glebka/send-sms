Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Module General
    Dim WithEvents wb As New WebBrowser
    Dim exitflag As Boolean = False
    Dim action As Integer
    Dim doc As mshtml.HTMLDocument
    Dim m As Match
    Dim cUrl As String
    Dim sessionId As String
    Private Sub ShowHelp()
        Console.WriteLine()
        Console.WriteLine("Console SMS sender v " + Application.ProductVersion)
        Console.WriteLine("Автор: Романов Глеб <rgewebppc@gmail.com>")
        Console.WriteLine()
        Console.WriteLine("Эта программа позволяет отправлять SMS на мобильные телефоны абонентов" + vbNewLine + " МТС Украина, используя услугу MTC Messenger.")
        Console.WriteLine("Информацию о MTS Messenger смотрите на сайте mts.com.ua")
        Console.WriteLine()
        Console.WriteLine("sms.exe -u пользователь -p пароль -num номер_телефона [/q] [/t] [""сообщение""]")
        Console.WriteLine()
        Console.WriteLine("номер_телефона - мобильный номер (пример: 0504957799)")
        Console.WriteLine("/q - добавить в контакт-лист, отправить запрос на подтверждение")
        Console.WriteLine("/t - отправить текстовое сообщение")
        Console.WriteLine()
        Console.WriteLine("ВАЖНО: Для использования этой программы необходима регистрация в MTS Messenger.")
        Console.WriteLine("http://www.mts.com.ua/ukr/mts_messenger_register.php")
        Console.WriteLine()
        Console.WriteLine("ВАЖНО: После регисрации в контакт-лист MTS Messenger следует добавить " + vbNewLine + "номер абонента, которому будут отправляться сообщения. " + vbNewLine + "Можно добавлять несколько абонентов. " + vbNewLine + "Добавить абонента в контакт-лист сайта можно используя ключ /q")
        Console.WriteLine()
        Console.WriteLine("ВАЖНО: Абонент должен отправить пустое SMS на номер 10780, а затем " + vbNewLine + "отправить SMS с любым текстом на номер 10780, первым словом " + vbNewLine + "которого будет ник (логин), который Вы указали при регистрации.")
        Console.WriteLine()
        Console.WriteLine("ВАЖНО: Если эти шаги не будут выполнены, то программа не сможет " + vbNewLine + "отправлять SMS абоненту.")
        Console.WriteLine()
        Console.WriteLine("ВАЖНО: Количество запросов на подтверждение НЕ ОГРАНИЧЕНО, количество " + vbNewLine + "SMS ОГРАНИЧЕНО - 3 SMS в сутки без ответа абонента.")
        Console.WriteLine()
        Console.WriteLine("ВАЖНО: Чтобы получить еще 3 SMS, абонент должен отправить SMS " + vbNewLine + "с любым текстом на номер 10780, первым словом которого будет ник (логин), " + vbNewLine + "который Вы указали при регистрации.")
        Console.WriteLine()
        Console.WriteLine("ВАЖНО: Стоимость отправки SMS на номер 10780 - стандартна для пакета абонента.")
        Console.WriteLine()
        Console.WriteLine("Текстовое сообщение не должно содержать двойных кавычек, " + vbNewLine + "его длина не должна превышать 147 символов.")
        Console.WriteLine("Желательно, чтобы сообщение было написано латинскими символами.")
    End Sub
    Sub Main()
        Dim regex As Regex
        Dim pattern As String = "\-u\s(?<1>\S{1,})\s-p\s(?<2>\S{1,})\s-num\s(?<3>[0-9]{10})\s(?<4>\S{2})\s?""?(?<5>[^""]*)""?" '"(?:\-user)\s(?<user>[a-zA-Z_-\.0-9]{3,})\s(?<pass>[\S]{1,})\s(?<number>[0-9]{10,10})\s(?<command>\\q|\\m)\s""(?<message>[\S]*)"""
        wb.ScriptErrorsSuppressed = True
        Select Case Command()
            Case "-help", "-h", "?", "help"
                ShowHelp()
            Case Else
                regex = New Regex(pattern, RegexOptions.IgnoreCase)
                Dim matches As MatchCollection = regex.Matches(Command)
                If matches.Count = 0 Then
                    Console.WriteLine()
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Ошибка в синтаксисе команды.")
                    Console.WriteLine("Введите sms.exe help для получения справки.")
                    Exit Sub
                    Application.Exit()
                End If
                m = matches.Item(0)
                If m.Groups("1").Value.Length < 3 Or m.Groups("2").Value.Length < 3 Then
                    Console.WriteLine()
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Имя пользователя и пароль должны иметь длину не менее 3 символа.")
                    Console.WriteLine("Введите sms.exe help для получения справки.")
                    Exit Sub
                    Application.Exit()
                End If
                If m.Groups("4").Value <> "/q" And m.Groups("4").Value <> "/t" Then
                    Console.WriteLine()
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Ошибка в синтаксисе команды.")
                    Console.WriteLine("Введите sms.exe help для получения справки.")
                    ShowHelp()
                    Exit Sub
                    Application.Exit()
                End If
                If m.Groups("5").Value.Length = 0 And m.Groups("4").Value = "/t" Then
                    Console.WriteLine()
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Запрещено отправлять пустые сообщения.")
                    Console.WriteLine("Введите sms.exe help для получения справки.")
                    Exit Sub
                    Application.Exit()
                ElseIf m.Groups("5").Value.Length >= 147 Then
                    Console.WriteLine()
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Слишком длинное сообщение.")
                    Console.WriteLine("Введите sms.exe help для получения справки.")
                    Exit Sub
                    Application.Exit()
                End If
                action = 1
                wb.Navigate("http://www.mts.com.ua/ukr/mts_messenger_register.php")
                Do Until exitflag
                    Application.DoEvents()
                Loop
        End Select
        Application.Exit()
        'Console.ReadKey()
        ' While wb.ReadyState <> WebBrowserReadyState.Complete
        'Application.DoEvents()
        'End While
    End Sub

    Private Function SendRequest(ByVal number As String) As Boolean
        Try
            doc.getElementsByName("number").item(0).value = number
            For Each el As mshtml.IHTMLFormElement In doc.getElementsByTagName("form")
                If el.action.Contains("mts_messenger.php") Then
                    el.submit()
                End If
            Next
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Sub SendMessage(Optional ByVal act As Integer = 1)
        Select Case act
            Case 1
                'Console.WriteLine("Щас перейдем вот сюда:")
                'Console.WriteLine("http://www.mts.com.ua/ukr/mts_messenger_chat.php?node=input&u=38" + m.Groups(3).Value + "&PHPSESSID=" + sessionId)
                wb.Navigate("http://www.mts.com.ua/ukr/mts_messenger_chat.php?node=input&u=38" + m.Groups(3).Value + "&PHPSESSID=" + sessionId)
                action = 3
            Case 2
                action = 3
                'Console.WriteLine(wb.Url.ToString)
                'Console.WriteLine("Мы должны находится в инпут...")
                If doc.url.Contains("input") Then
                    Try
                        'Console.WriteLine("Да, это так... Пишем сообщение, отправляем... Не факт, что дойдет :)")
                        doc.getElementsByName("message").item(0).value = m.Groups(5).Value
                        action = 7
                        doc.getElementsByTagName("form").item(0).submit()
                        Console.WriteLine("[" + DateAndTime.Now.ToString + "] Сообщеие отправлено...")
                        exitflag = True
                    Catch ex As Exception
                        Console.WriteLine("[" + DateAndTime.Now.ToString + "] Ошибка: " + ex.Message)
                        Console.WriteLine("Не удалось отправить сообщение.")
                        exitflag = True
                    End Try
                Else
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Сообщеие не отправлено.")
                    Console.WriteLine(doc.url)
                End If
            Case 3
                action = 5
                wb.Navigate("http://www.mts.com.ua/ukr/mts_messenger_chat.php?u=38" + m.Groups(3).Value)
            Case 4
                Try
                    Dim frames As mshtml.FramesCollection = doc.frames
                    Dim frame As mshtml.IHTMLWindow4
                    For i = 0 To frames.length
                        frame = frames.item(i)
                        If frame.name = "main" Then
                            Dim fmain As mshtml.HTMLDocument = frame.document
                            action = 6
                            Do While fmain.readyState = "loading"
                                Debug.Print(fmain.readyState)
                                Application.DoEvents()
                            Loop
                            Dim elements As mshtml.HTMLElementCollection = fmain.getElementsByTagName("div")
                            Dim lastEl As mshtml.IHTMLElement
                            For Each el In elements
                                lastEl = el
                            Next

                            If lastEl.innerText.Contains("Увага:") Then
                                Console.WriteLine("[" + DateAndTime.Now.ToString + "] Сообщеие не отправлено. Текст ошибки:")
                                Console.WriteLine(lastEl.innerText)
                            Else
                                Console.WriteLine("[" + DateAndTime.Now.ToString + "] Сообщеие успешно отправлено")
                                Console.WriteLine(lastEl.innerText)
                            End If
                            exitflag = True
                            Exit For
                        End If
                    Next
                Catch ex As Exception
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Ошибка: " + ex.Message)
                    Console.WriteLine("Невозможно проверить отправку сообщения.")
                    exitflag = True
                End Try
                
        End Select
    End Sub

    Private Sub wb_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles wb.DocumentCompleted
        If e.Url.ToString.Contains("node=iframe") Then
            Exit Sub
        End If
        Select Case action
            Case 1 ' загрузка страницы входа
                If wb.Document.Url.ToString.Contains("res:") Then
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Ошибка при загрузке http://www.mts.com.ua/ukr/mts_messenger_register.php: ")
                    Console.WriteLine(wb.Document.GetElementsByTagName("html").Item(0).InnerText)
                    exitflag = True
                ElseIf wb.Document.Url.ToString = "http://www.mts.com.ua/ukr/mts_messenger_register.php" Then
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Авторизация...")
                    Try
                        doc = wb.Document.DomDocument
                        doc.getElementsByName("login").item(0).value = m.Groups("1").Value
                        doc.getElementsByName("password").item(0).value = m.Groups("2").Value
                        action = 2
                        doc.getElementsByName("submit").item(0).click()
                    Catch ex As Exception
                        Console.WriteLine("[" + DateAndTime.Now.ToString + "] Ошибка: на странице " + wb.Document.Url.ToString + " не найдены поля для ввода логина и (или) пароля.")
                        exitflag = True
                        Exit Sub
                    End Try
                Else
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Ошибка: перенаправление на " + wb.Document.Url.ToString)
                    exitflag = True
                End If
            Case 2 ' вход в МТС
                '&err_login=1
                If wb.Document.Url.ToString.Contains("res:") Then
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Ошибка при загрузке http://www.mts.com.ua/ukr/mts_messenger_register.php: ")
                    Console.WriteLine(wb.Document.GetElementsByTagName("html").Item(0).InnerText)
                    exitflag = True
                ElseIf wb.Document.Url.ToString.Contains("&err_login=1") Then
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Ошибка: неправильный логин и(или) пароль.")
                    exitflag = True
                Else
                    cUrl = wb.Document.Url.ToString
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Авторизация прошла успешно.")
                    Dim ar() As String = Split(wb.Document.Url.ToString, "?PHPSESSID=")
                    sessionId = ar(1)
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] PHPSESSID:" + sessionId)
                    'Console.WriteLine(sessionId)
                    If m.Groups("4").Value = "/q" Then
                        Console.WriteLine("[" + DateAndTime.Now.ToString + "] Отправка запроса на подтверждение...")
                        If SendRequest(m.Groups("3").Value) Then
                            Console.WriteLine("[" + DateAndTime.Now.ToString + "] Запрос отправлен.")
                            exitflag = True
                            action = 0
                        Else
                            Console.WriteLine("[" + DateAndTime.Now.ToString + "] Ошибка: отправка запроса не удалась.")
                        End If
                    ElseIf m.Groups("4").Value = "/t" Then
                        Console.WriteLine("[" + DateAndTime.Now.ToString + "] Отправка текстового сообщения (этап 1).")
                        SendMessage()
                    End If
                End If
            Case 3
                If wb.Document.Url.ToString.Contains("res:") Then
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Ошибка при отправке сообщения.")
                    Console.WriteLine(wb.Document.GetElementsByTagName("html").Item(0).InnerText)
                    exitflag = True
                Else
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Отправка текстового сообщения (этап 2).")
                    SendMessage(2)
                End If
            Case 4
                If wb.Document.Url.ToString.Contains("res:") Then
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Ошибка при получении подтверждения об отправке.")
                    Console.WriteLine(wb.Document.GetElementsByTagName("html").Item(0).InnerText)
                    exitflag = True
                Else
                    SendMessage(3)
                End If
            Case 5
                If wb.Document.Url.ToString.Contains("res:") Then
                    Console.WriteLine("[" + DateAndTime.Now.ToString + "] Ошибка при получении подтверждения об отправке.")
                    Console.WriteLine(wb.Document.GetElementsByTagName("html").Item(0).InnerText)
                    exitflag = True
                Else
                    SendMessage(4)
                End If
            Case Else
                'Console.WriteLine(wb.Document.Url.ToString)
                'Console.WriteLine(wb.Document.GetElementsByTagName("html").Item(0).InnerText)
                exitflag = True
        End Select
    End Sub
End Module
