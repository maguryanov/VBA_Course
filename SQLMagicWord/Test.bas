Attribute VB_Name = "Test"
Option Explicit

Private Sub DisplayArithmeticProgression()
    On Error GoTo ErrorHandler
    ' Объявление переменных с венгерской нотацией
    Dim dblFirstElement As Double  ' Первый элемент прогрессии
    Dim dblStep As Double          ' Шаг прогрессии
    Dim strResult As String        ' Строка для формирования результата
    Dim intCounter As Integer      ' Счетчик цикла
    
    ' Запрос данных у пользователя
    dblFirstElement = InputBox("Введите первый элемент прогрессии:", "Арифметическая прогрессия")
    dblStep = InputBox("Введите шаг прогрессии:", "Арифметическая прогрессия")
    
    ' Проверка на корректный ввод
    If Not IsNumeric(dblFirstElement) Or Not IsNumeric(dblStep) Then
        MsgBox "Ошибка: необходимо ввести числовые значения!", vbCritical
        Exit Sub
    End If
    
    ' Преобразование введенных строк в числа
    dblFirstElement = CDbl(dblFirstElement)
    dblStep = CDbl(dblStep)
    
    ' Формирование результата
    strResult = "Первые 3 члена арифметической прогрессии:" & vbCrLf & vbCrLf
    
    For intCounter = 0 To 2
        strResult = strResult & "Член " & (intCounter + 1) & ": " & _
                   Format(dblFirstElement + intCounter * dblStep, "0.#####") & vbCrLf
    Next intCounter
    
    ' Вывод результата
    MsgBox strResult, vbInformation, "Результат"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

