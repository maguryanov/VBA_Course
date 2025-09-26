Attribute VB_Name = "m_01_18_Errors"
Option Explicit

'Нет обработки ошибок
Sub Errors_NoHandling_01()
    Dim intValue As Integer
    intValue = 0
    Debug.Print 1 / intValue 'Потенциальная ошибка
    Debug.Print "Действия после потенциальной ошибки"
End Sub


'Последовательность выполнения при отсутствии ошибки On Error GoTo
Sub Errors_GoTo_01()
    On Error GoTo ErrorHandler
    Dim intDividor As Integer
    intDividor = 100
        Debug.Print 1 / intDividor 'Потенциальная ошибка
    Debug.Print "Действия после потенциальной ошибки"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Последовательность выполнения при наличии ошибки On Error GoTo
Sub Errors_GoTo_02()
    On Error GoTo ErrorHandler
    Dim intDividor As Integer
    intDividor = 0
    Debug.Print 1 / intDividor 'Потенциальная ошибка
    Debug.Print "Действия после потенциальной ошибки"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



'Отмена обработки ошибок. On Error GoTo 0
Sub Errors_GoTo_0_01()
    On Error GoTo ErrorHandler
    Dim intValue As Integer
    intValue = 1
    Debug.Print "Обработка ошибок выполняется"
    Debug.Print 1 / intValue 'Потенциальная ошибка
    Debug.Print "Действия после потенциальной ошибки"
    On Error GoTo 0 'Отмена обработки ошибок
    Debug.Print "Обработка ошибок НЕ выполняется"
    Debug.Print 1 / (intValue - 1) 'Потенциальная ошибка
    Debug.Print "Действия после потенциальной ошибки"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Продолжение после ошибки. On Error Resume Next
Sub Errors_Resume_01()
    On Error Resume Next
    Dim intValue As Integer
    intValue = 0
    Debug.Print 1 / intValue 'Потенциальная ошибка
    Debug.Print "Действия после потенциальной ошибки"
    Exit Sub
ErrorHandler: 'Не нужен в таком случае
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Обработка после ошибки. On Error Resume Next
Sub Errors_Resume_02()
    On Error Resume Next
    Dim intValue As Integer
    intValue = 0
    Debug.Print 1 / intValue 'Потенциальная ошибка
    If Err.Number = 11 Then Debug.Print "Деление на ноль"
    Debug.Print 1 / intValue 'Потенциальная ошибка
    If Err.Number = 11 Then Debug.Print "Деление на ноль"
    Debug.Print "Действия после потенциальных ошибок"
End Sub



'Безусловное выполнение действий Finalization. Отсутствие ошибки
Sub File_Always_01()
    On Error GoTo ErrorHandler
    Dim intValue As Integer
    Open "D:\VBA\ANSI.txt" For Input As #1
    Debug.Print "Чтение из файла"
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'Безусловное выполнение действий Finalization. Наличие ошибки
Sub File_Always_02()
    On Error GoTo ErrorHandler
    Dim intValue As Integer
    Open "D:\VBA\ANSI.txt" For Input As #1
    Debug.Print "Чтение из файла"
    intValue = "100$"
    Debug.Print "Если бы закрытие файла было сдесь?"
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Получение информации об ошибке
Sub Errors_Info_01()
    On Error GoTo ErrorHandler
    Dim intValue As Integer: intValue = 0
    Debug.Print 1 / intValue 'Потенциальная ошибка
    Debug.Print "Действия после потенциальной ошибки"
    Exit Sub
ErrorHandler:
    Debug.Print "Err:", Err
    Debug.Print "Number:", Err.Number
    Debug.Print "Description:", Err.Description
    Debug.Print "HelpContext:", Err.HelpContext
    Debug.Print "HelpFile:", Err.HelpFile
    Debug.Print "LastDllError:", Err.LastDllError
    Debug.Print "Source:", Err.Source
End Sub

'Очистка информации об ошибке. Err.Clear
Sub Errors_Info_02()
    On Error GoTo ErrorHandler
    Dim intValue As Integer: intValue = 0
    Debug.Print 1 / intValue 'Потенциальная ошибка
    Debug.Print "Действия после потенциальной ошибки"
    Exit Sub
ErrorHandler:
    Err.Clear
    Debug.Print "Err:", Err
    Debug.Print "Number:", Err.Number
    Debug.Print "Description:", Err.Description
    Debug.Print "HelpContext:", Err.HelpContext
    Debug.Print "HelpFile:", Err.HelpFile
    Debug.Print "LastDllError:", Err.LastDllError
    Debug.Print "Source:", Err.Source
End Sub



'Частые ошибки. 11 / Division by zero
Sub Errors_Frequent_01()
    On Error GoTo ErrorHandler
    Debug.Print 1 / 0
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Частые ошибки. 6 / Overflow
Sub Errors_Frequent_02()
    On Error GoTo ErrorHandler
    Dim intValue As Integer
    intValue = 100500
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Частые ошибки. 9 / Subscript out of range
Sub Errors_Frequent_07()
    On Error GoTo ErrorHandler
    Dim aValues(3) As Integer
    aValues(4) = 100
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Частые ошибки. 13 / Type mismatch
Sub Errors_Frequent_03()
    On Error GoTo ErrorHandler
    Dim intValue As Integer
    intValue = "100$"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Частые ошибки. 55 / File already open
Sub Errors_Frequent_04()
    On Error GoTo ErrorHandler
    Open "D:\VBA\ANSI.txt" For Input As #1
    Debug.Print "Чтение из файла"
    Open "D:\VBA\ANSI.txt" For Append As #1
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'Частые ошибки. 53 / File not found
Sub Errors_Frequent_05()
    On Error GoTo ErrorHandler
    Open "D:\VBA\Несуществующий файл.txt" For Input As #1
    Debug.Print "Чтение из файла"
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Частые ошибки. 54 / Bad file mode
Sub Errors_Frequent_08()
    On Error GoTo ErrorHandler
    Open "D:\VBA\Files\Данные.txt" For Input As #1
    Write #1, "Попытка записи данных"
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'Частые ошибки. 75 / Path/File access error
Sub Errors_Frequent_06()
    On Error GoTo ErrorHandler
    Open "D:\VBA\Нет разрешений.txt" For Input As #1
    Debug.Print "Чтение из файла"
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub



'Генерирование собственной ошибки. Err.Raise
Sub Errors_Raise_01()
    On Error GoTo ErrorHandler
    Dim intProductQty As Integer
    intProductQty = -100
    If intProductQty < 0 Then Err.Raise _
        50001, "Errors_Raise_06", "Значение не может быть отрицательным"
    Debug.Print "Дальнейшие действия"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description & " / " & Err.Source
End Sub


'Генерирование собственной ошибки. Очистка Err. Явная и неявная
Sub Errors_Raise_02()
    On Error Resume Next
    Dim intProductQty As Integer
    intProductQty = -100
    Debug.Print 1 / 0
    If intProductQty < 0 Then
        'Err.Clear
        Err.Raise 50001
    End If
    Debug.Print Err.Number & " / " & Err.Description & " / " & Err.Source
End Sub


'0–512 зарезервирован для системных ошибок
Sub Errors_Raise_03()
    On Error GoTo ErrorHandler
    Dim intProductQty As Integer
    intProductQty = -100
    If intProductQty < 0 Then Err.Raise _
        6, "Errors_Raise_06", "Значение не может быть отрицательным"
    Debug.Print "Дальнейшие действия"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description & " / " & Err.Source
End Sub

'Оператор Error не рекомендуется. Для обратной совместимости
Sub Errors_Raise_04()
    On Error GoTo ErrorHandler
    Dim intProductQty As Integer
    intProductQty = -100
    If intProductQty < 0 Then Error 50001
    If intProductQty < 0 Then Error 6
    Debug.Print "Дальнейшие действия"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description & " / " & Err.Source
End Sub

'Ошибка произошла в классе. Она обработана
Sub Errors_Raise_05()
    On Error GoTo ErrorHandler
    Dim varResult As Variant
    Dim oVideoLesson As New VideoLesson
    'varResult = oVideoLesson.ErrorHandlingDemo(0)
    varResult = oVideoLesson.ErrorHandlingDemo("100$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Ошибка произошла в классе. Она обработана и сгенерирована
Sub Errors_Raise_06()
    On Error GoTo ErrorHandler
    Dim varResult As Variant
    Dim oVideoLesson As New VideoLesson
    'varResult = oVideoLesson.ErrorHandlingDemo(0)
    varResult = oVideoLesson.ErrorHandlingRaiseDemo("100$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Обработка ошибок при отладке. Опция Error Trapping. Break on all
Sub Errors_Testing_01()
    On Error GoTo ErrorHandler
    Dim dblDivider As Double
    Debug.Print 1 / dblDivider
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Обработка ошибок при тестировании. Опция Error Trapping. Break In Сlass Module
Sub Errors_Testing_02_1()
    On Error GoTo ErrorHandler
    Dim varResult As Variant
    Dim oVideoLesson As New VideoLesson
    'varResult = oVideoLesson.ErrorHandlingDemo(0)
    varResult = oVideoLesson.ErrorHandlingRaiseDemo("100$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Обработка ошибок при тестировании. Опция Error Trapping. Break In class Module
Sub Errors_Testing_02_2()
    On Error GoTo ErrorHandler
        Err.Raise 50000, , "Для тестировани Break In class Module"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Обработка ошибок при тестировании. Опция Error Trapping. Break on Unhandled
Sub Errors_Testing_03_01()
    On Error GoTo 0
    Debug.Print 1 / 0
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Максимальное и минимальное значение типа Integer ' % - Integer
Sub d_20_Integer()
On Error GoTo ErrorHandler
    Dim intVar As Integer
    Do
        intVar = intVar + 1
    Loop
    Debug.Print intVar
    Exit Sub
ErrorHandler:
    Debug.Print intVar
    Debug.Print Err.Number & " / " & Err.Description
End Sub

