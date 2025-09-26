Attribute VB_Name = "m_01_19_Debug"
Option Explicit
Private mintVar As Integer
Private mdtVar As Date
Public gstrVar As String

'1.  Синтаксические ошибки. Опция Auto syntax check
Private Sub d01_DisplayArithmeticProgression()
    ' Объявление переменных
    'Dim dblFirstElement Double  ' Первый элемент прогрессии
End Sub
'2.  Ошибки компиляции.
Private Sub d02_DisplayArithmeticProgression()
    ' Объявление переменных
    Dim FirstElement As String  ' Ввод данных
    'Dim FirstElement As Double  ' Преобразованное значение
    ' Получение данных для наглядности пропущено
    ' Преобразование введенных строк в числа
    FirstElement = CDbl("100.5")
    'Step = CDbl("2")
End Sub

'3.  Ошибки времени выполнения
Private Sub d03_DisplayArithmeticProgression()
    ' Объявление переменных
    Dim dblFirstElement As Double  ' Первый элемент прогрессии
    Dim dblStep As Double          ' Шаг прогрессии
    Dim strResult As String        ' Строка для формирования результата
    Dim intCounter As Integer      ' Счетчик цикла
    ' Получение данных для наглядности пропущено
    ' Преобразование введенных строк в числа
    dblFirstElement = CDbl("1,5")
    dblStep = CDbl("2")
    ' Формирование результата
    For intCounter = 0 To 2
        strResult = strResult & "Элемент " & (intCounter + 1) & ": " & _
                   Format(dblFirstElement + intCounter, "0.#####") & vbCrLf
    Next intCounter
    ' Вывод результата
    Debug.Print strResult
    Exit Sub
End Sub

'4.  Логические ошибки
Private Sub d04_DisplayArithmeticProgression()
    On Error GoTo ErrorHandler
    Dim dblFirstElement As Double, dblStep As Double
    Dim strFirstElement As String, strStep As String
    Dim strResult As String, intCounter As Integer
    ' Получение данных для наглядности пропущено
    strFirstElement = "2": strStep = "0,5"
    ' Проверка на корректный ввод
    If IsNumeric(strFirstElement) And IsNumeric(strStep) Then
        dblFirstElement = strFirstElement: dblStep = strStep
        Else: MsgBox "Необходимо ввести числовые значения!", vbCritical: Exit Sub
    End If
    For intCounter = 0 To 2
        strResult = strResult & "Элемент " & (intCounter + 1) & ": " & _
                   Format(dblFirstElement + intCounter, "0.#####") & vbCrLf
    Next intCounter
    Debug.Print strResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Immediate Window (Ctrl+G)
Private Sub d05_ImmediateWindow()
    Debug.Print "Демонстрация работы с Immediate Window"
'print 2 * 2
'? 2 * 2
'? Choose(10, 1, 2)
'? strVar
'dim intVar as Integer 'Нельзя
'intVar  =  100
'? intVar
'MkDir "d:\Vba\FromImmediateWindow"
'Help по F1
End Sub



' Режим прерывания и точка останова
Private Sub d06_Breakpoint()
' Просмотр через Immediate window, Watch, Tips, locals
' Настройки Auto Datа Tips
' Изменение значений переменных
    On Error GoTo ErrorHandler
    Dim intCounter As Integer, intResult As Integer
    Static intAttempt As Integer: intAttempt = intAttempt + 1
    For intCounter = 1 To 10
        intResult = intResult + intCounter
    Next intCounter
    Debug.Print "Успешное завершение #" & intAttempt
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


' Stop Безусловная остановка
Private Sub d07_Stop()
    On Error GoTo ErrorHandler
    Dim intCounter As Integer, intResult As Integer
    For intCounter = 1 To 10
        intResult = intResult + intCounter
        Stop 'Безусловная остановка
    Next intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


' Stop Остановка по условию
Private Sub d08_Stop()
    On Error GoTo ErrorHandler
    Dim intCounter As Integer, intResult As Integer
    For intCounter = 1 To 10
        If intCounter = 5 Then Stop 'Остановка по условию
        intResult = intResult + intCounter
    Next intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

' StepInto - F8
Private Sub d09_StepInto()
    On Error GoTo ErrorHandler
    Dim intCounter As Integer, intResult As Integer
    Dim intDivider As Integer: intDivider = 0
    For intCounter = 1 To 10
        intResult = intResult + intCounter / intDivider
    Next intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Ищем причину ошибок
Private Sub d10_DisplayArithmeticProgression_Demo()
    On Error GoTo ErrorHandler
    Dim dblFirstElement As Double, dblStep As Double
    Dim strFirstElement As String, strStep As String
    Dim strResult As String, intCounter As Integer
    ' Получение данных для наглядности пропущено
    strFirstElement = "2": strStep = "0,5"
    ' Проверка на корректный ввод
    If IsNumeric(strFirstElement) And IsNumeric(strStep) Then
        dblFirstElement = strFirstElement: dblStep = strStep
        Else: MsgBox "Необходимо ввести числовые значения!", vbCritical: Exit Sub
    End If
    For intCounter = 0 To 2
        strResult = strResult & "Элемент " & (intCounter + 1) & ": " & _
                   Format(dblFirstElement + intCounter, "0.#####") & vbCrLf
    Next intCounter
    Debug.Print strResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Step Over - пропустить процедуру (Shift+F8)
Private Sub d11_DisplayArithmeticProgression_Step_Over()
    On Error GoTo ErrorHandler
    Dim dblFirstElement As Double, dblStep As Double
    Dim strFirstElement As String, strStep As String
    Dim strResult As String, intCounter As Integer
    ' Получение данных для наглядности пропущено
    strFirstElement = "2": strStep = "0,5"
    ' Проверка на корректный ввод для наглядности пропущена
    dblFirstElement = strFirstElement: dblStep = strStep
    For intCounter = 0 To 2
        strResult = strResult & "Элемент " & (intCounter + 1) & ": " & _
                   Format(d11_ElementN(dblFirstElement, intCounter, dblStep), "0.00") & vbCrLf
    Next intCounter
    Debug.Print strResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
'Функция вычисления N элемента арифметической последовательности
Private Function d11_ElementN _
    (ByVal FirstElement As Double, ByVal Number As Long, ByVal Step As Double) As Double
    d11_ElementN = FirstElement + (Number - 1) * Step
End Function


'Step Out (Ctrl+Shift+F8)
Private Sub d12_DisplayArithmeticProgression_Step_Out()
    On Error GoTo ErrorHandler
    Dim dblFirstElement As Double, dblStep As Double
    Dim strFirstElement As String, strStep As String
    ' Получение данных для наглядности пропущено
    strFirstElement = "2": strStep = "0,5"
    ' Проверка на корректный ввод для наглядности пропущена
    dblFirstElement = CDbl(strFirstElement)
    dblStep = CDbl(strStep)
    Debug.Print d12_ArithmeticProgressionText(dblFirstElement, dblStep)
    Debug.Print "Успешное завершение"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Функция получения строки для вывода арифметической последовательности
Private Function d12_ArithmeticProgressionText _
    (ByVal FirstElement As Double, ByVal Step As Double) As String
    Dim intCounter As Integer
    Dim strResult As String
    For intCounter = 1 To 3
        strResult = strResult & "Элемент " & intCounter & ": " & _
                   Format(d11_ElementN(FirstElement, intCounter, Step), "0.00") & vbCrLf
    Next intCounter
    d12_ArithmeticProgressionText = strResult
End Function

'Call Stack-Стек вызовов. Вызов окна (Ctrl+L), Locals
Sub d13_CallStack()
    d12_DisplayArithmeticProgression_Step_Out
End Sub

'Run To Cursor (Ctrl+F8) Переход к позиции курсора
Private Sub d14_RunToCursor()
    On Error GoTo ErrorHandler
    Dim dblFirstElement As Double, dblStep As Double
    Dim strFirstElement As String, strStep As String
    Dim strResult As String, intCounter As Integer
    ' Получение данных для наглядности пропущено
    strFirstElement = "2": strStep = "0,5"
    ' Проверка на корректный ввод
    If IsNumeric(strFirstElement) And IsNumeric(strStep) Then
        dblFirstElement = strFirstElement: dblStep = strStep
        Else: MsgBox "Необходимо ввести числовые значения!", vbCritical: Exit Sub
    End If
    For intCounter = 0 To 2
        strResult = strResult & "Элемент " & (intCounter + 1) & ": " & _
                   Format(dblFirstElement + intCounter, "0.#####") & vbCrLf
    Next intCounter
    Debug.Print strResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Навигация по программе. "Перетаскивание стрелочки"
Private Sub d15_ExecutionBySteps()
    Debug.Print "Шаг первый"
    Debug.Print "Шаг второй"
    Debug.Print "Шаг третий"
    Debug.Print "Шаг четвёртый"
    Debug.Print "Шаг пятый"
    Debug.Print "Шаг шестой"
    Debug.Print "Шаг седьмой"
End Sub


'Пропуск части кода. Комментарии и Set Next Statement
Private Sub d17_SetNextStatement()
    Debug.Print "Шаг первый"
    Debug.Print "Шаг второй"
    Debug.Print "Шаг третий"
    Debug.Print "Шаг четвёртый"
    Debug.Print "Шаг пятый"
    Debug.Print "Шаг шестой"
    Debug.Print "Шаг седьмой"
End Sub


'Пропуск части кода. Show Next Statement
Private Sub d17_ShowNextStatement()
    Debug.Print "Демонстрация работы Show Next Statement"
End Sub


'Watch Window - Окно контрольных значений. Quick Watch (Shift+F9)
Private Sub d18_QuickWatch()
    On Error GoTo ErrorHandler
    Dim dblFirstElement As Double, dblStep As Double
    Dim strFirstElement As String, strStep As String
    Dim strResult As String, intCounter As Integer
    ' Получение данных и проверка на корректный ввод для наглядности пропущены
    dblFirstElement = 2: dblStep = 0.5
    For intCounter = 0 To 2
        strResult = strResult & "Элемент " & (intCounter + 1) & ": " & _
                   Format(dblFirstElement + intCounter * dblStep, "0.00") & vbCrLf
    Next intCounter
    Debug.Print strResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Watch Window. Определение контекста
Private Sub d19_WatchContext()
    Dim strVar As String
    strVar = "One"
    Debug.Print "Демонстрация контекста"
'    Переменные определённые в модуле
'    Private mintVar As Integer
'    Private mdtVar As Date
'    Public gstrVar As String
End Sub

'Watch Window. Вывод значения выражения
Private Sub d20_ExpressionWatch()
    On Error GoTo ErrorHandler
    Dim intCounter As Integer, intResult As Integer
    For intCounter = 1 To 10
        intResult = intResult + intCounter
    Next intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Watch Window. Остановка по условию
Private Sub d21_ConditionalWatch()
    On Error GoTo ErrorHandler
    Dim intCounter As Integer, intResult As Integer
    For intCounter = 1 To 10
        intResult = intResult + intCounter
    Next intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Watch Window. Остановка по изменению значения
Private Sub d22_ChangeWatch()
    On Error GoTo ErrorHandler
    Dim intCounter As Integer, intResult As Integer
    For intCounter = 1 To 10
        intResult = intResult + intCounter
    Next intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Debug.Assert
Private Sub d24_Debug_Assert()
    On Error GoTo ErrorHandler
    
    Dim lngQty As Long 'Значение количества не должно быть отрицательным
    Dim intCounter As Integer
    'Моделируем ситуацию что количество может стать отрицательным
    For intCounter = 100 To -100 Step -1
        lngQty = intCounter
        Debug.Assert lngQty >= 0
    Next intCounter

ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Debug.Print
Private Sub d24_Debug_Print()
    On Error GoTo ErrorHandler
    
    Debug.Print "Демо", "работы "; "Debug.Print",
    Debug.Print "продолжение";
    Debug.Print " строки"
    Debug.Print "Демонстрация"; Spc(5); "использования Spc"
    Debug.Print "1234567890"; Tab(12); "использование Tab"

ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Особенности отладки при использовании обработки ошибок
Private Sub d25_ErrorHandling()
    On Error GoTo ErrorHandler
    Dim intResult As Integer
    Err.Raise 50000, , "Неверные данные"
    intResult = intResult + 10000
    intResult = intResult + 10000
    intResult = intResult + 10000
    intResult = intResult + 10000
    intResult = intResult + 10000
        
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Опция Error Trapping. Break In Сlass Module
Sub d26_BreakInСlassModule()
    On Error GoTo ErrorHandler
    Dim varResult As Variant
    Dim oVideoLesson As New VideoLesson
    'varResult = oVideoLesson.ErrorHandlingDemo(0)
    varResult = oVideoLesson.ErrorHandlingRaiseDemo("100$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub





