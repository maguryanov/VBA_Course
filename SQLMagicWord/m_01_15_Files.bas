Attribute VB_Name = "m_01_15_Files"
Option Explicit

'Запись текстовых файлов. Добавление
Sub File_Append_01()
    On Error GoTo ErrorHandler
    Static intAttempt As Integer
    intAttempt = intAttempt + 1
    Open "D:\VBA\Тестирование файлов.txt" For Append As #1
    Print #1, "Выгрузка из программы на VBA (Append) попытка " & intAttempt
    Debug.Print "Успешная запись в файл попытка " & intAttempt
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Запись текстовых файлов. Перезапись
Sub File_Output_01()
    On Error GoTo ErrorHandler
    Static intAttempt As Integer: intAttempt = intAttempt + 1
    Open "D:\VBA\Тестирование файлов.txt" For Output As #1
    Print #1, "Выгрузка из программы на VBA попытка " & intAttempt
    Debug.Print "Успешная запись в файл попытка " & intAttempt
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Функция FreeFile. Аргументы 0 и 1
Sub File_FreeFile_01()
    On Error GoTo ErrorHandler
    Debug.Print "Номер файла при FreeFile(0): " & FreeFile(0)
    Debug.Print "Номер файла пре FreeFile(1): " & FreeFile(1)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Функция FreeFile. Практический пример
Sub File_FreeFile_02()
    On Error GoTo ErrorHandler
    Static intAttempt As Integer
    Dim intFile As Integer
    intAttempt = intAttempt + 1
    intFile = FreeFile
    Open "D:\VBA\Тестирование файлов.txt" For Output As #intFile
    Print #intFile, "Выгрузка из программы на VBA, попытка " & intAttempt
    Debug.Print "Успешная запись в файл попытка: " & intAttempt
Finalization:
    Debug.Print "Доступный свободный файл до закрытия: " & FreeFile
    Close #intFile
    Debug.Print "Доступный свободный файл после закрытия: " & FreeFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Оператор Print. Работа с типами данных
Sub File_DataTypes_Print_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Static intAttempt As Integer: intAttempt = intAttempt + 1
    Open "D:\VBA\Тестирование Print.txt" For Output As #intFile
    Print #intFile, "Запись данных Print, попытка"; intAttempt
    Print #intFile, "Integer", "Double", "Boolean", "Null", "Date"
    Print #intFile, 100, 100.5, True, Null, Now
    Debug.Print "Успешная запись в файл, попытка: " & intAttempt
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'Оператор Write. Работа с типами данных
Sub File_DataTypes_Write_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Static intAttempt As Integer: intAttempt = intAttempt + 1
    Open "D:\VBA\Тестирование Write.txt" For Output As #intFile
    Write #intFile, "Запись данных Write, попытка"; intAttempt
    Write #intFile, "Integer", "Double", "Boolean", "Null", "Date"
    Write #intFile, 100, 100.5, True, Null, Now
    Debug.Print "Успешная запись в файл, попытка: " & intAttempt
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'Оператор Write. Запись данных
Sub File_DataTypes_Write_02()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Static intAttempt As Integer: intAttempt = intAttempt + 1
    Open "D:\VBA\Записанные данные.txt" For Output As #intFile
    Write #intFile, 100, 100.5, True, Null, Now
    Write #intFile, 200, 100.6, False, Empty, Now
    Write #intFile, 300, 100.7, True, Null, Now
    Write #intFile, 400, 100.8, True, Null, Now
    Debug.Print "Успешная запись в файл, попытка: " & intAttempt
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Чтение записанных Write данных
Sub File_DataTypes_Input_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim intVar As Integer
    Dim dblVar As Double
    Dim boolVar As Boolean
    Dim varVar As Variant
    Dim dtVar As Date
    Open "D:\VBA\Записанные данные.txt" For Input As #intFile
    Input #intFile, intVar, dblVar, boolVar, varVar, dtVar
    Debug.Print intVar, dblVar, boolVar, varVar, dtVar
    Input #intFile, intVar, dblVar, boolVar, varVar, dtVar
    Debug.Print intVar, dblVar, boolVar, varVar, dtVar
    Debug.Print TypeName(varVar)
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'Сохранение данных объекта Person
Sub WritePeople()
    Dim oPeople As Collection
    Dim oPerson As Person
    Set oPerson = New Person
    oPerson.LastName = "Гурьянов"
    oPerson.FirstName = "Михаил"
    oPerson.Gender = "М"
    oPerson.BirthDate = #2/13/1972#
    oPerson.AddToFile ("D:\VBA\People.txt")
    Set oPeople = New Collection
    oPeople.Add oPerson
    Set oPerson = New Person
    oPerson.LastName = "Ларина"
    oPerson.FirstName = "Татьяна"
    oPerson.Gender = "f"
    oPerson.BirthDateIsUnknown = True
    oPerson.AddToFile ("D:\VBA\People.txt")

End Sub
'Чтение данных объекта в Person коллекцию
Sub File_DataTypes_Input_02()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim oPeople As New Collection
    Dim oPerson As Person
    Dim strFirstName As String, strLastName As String, strGender As String, _
                dtBirthDate As Date, BirthDateIsUnknown As Boolean
    Open "D:\VBA\People.txt" For Input As #intFile
    Do Until EOF(intFile)
         Input #intFile, strFirstName, strLastName, strGender, _
                dtBirthDate, BirthDateIsUnknown
         Set oPerson = New Person
         oPerson.FirstName = strFirstName
         oPerson.LastName = strLastName
         oPerson.Gender = strGender
         oPerson.BirthDate = dtBirthDate
         oPerson.BirthDateIsUnknown = BirthDateIsUnknown
         oPeople.Add oPerson, strLastName & strFirstName
    Loop
    PrintPeopleCollection oPeople
Finalization:
    Set oPerson = Nothing
    Set oPeople = Nothing
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Проверка чтения текста одним Input
Sub File_DataTypes_Input_04()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim strText As String
    Open "D:\VBA\Просто текст.txt" For Input As #intFile
    Input #intFile, strText
    MsgBox strText
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'Чтение одной строки Line Input и использование цикла
Sub File_DataTypes_Input_06()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim strText As String
        Open "D:\VBA\Просто текст.txt" For Input As #intFile
        Do Until EOF(intFile)
            Line Input #intFile, strText
            MsgBox strText
        Loop
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'Если файла нет
Sub File_DataTypes_Input_07()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim strText As String
        Open "D:\VBA\Ошибочное название.txt" For Input As #intFile
        Do Until EOF(intFile)
            Line Input #intFile, strText
            MsgBox strText
        Loop
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Команды форматированного вывода Print
Sub File_Commands_Print()
    On Error GoTo ErrorHandler
    Static intAttempt As Integer
    intAttempt = intAttempt + 1
    Open "D:\VBA\Тестирование файлов.txt" For Output As #1
    Print #1, "попытка " & intAttempt
    Print #1, "Запятая попытка ", intAttempt
    Print #1, "Точка с запятой попытка "; intAttempt
    Print #1, "Spc(2) попытка "; Spc(2); intAttempt
    Print #1, "Space(2) попытка " & Space(2) & intAttempt
    Print #1, "String(String(2, vbTab)) попытка " & String(2, vbTab) & intAttempt
    Print #1, "Tab попытка "; Tab; intAttempt
    Print #1, "Tab(1) попытка "; Tab(1); intAttempt
    Print #1, "Tab(2) попытка "; Tab(2); intAttempt
    Print #1, "Tab(10) попытка "; Tab(10); intAttempt
    Print #1, "Tab(11) попытка "; Tab(11); intAttempt
    Print #1, "Пример продолжения записи ";
    Print #1, "в ту же самую строку"
    
    Debug.Print "Успешная запись в файл попытка " & intAttempt
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'Команды форматированного вывода Write
Sub File_Commands_Write()
    On Error GoTo ErrorHandler
    Static intAttempt As Integer
    intAttempt = intAttempt + 1
    Open "D:\VBA\Тестирование файлов.txt" For Output As #1
    Write #1, "попытка " & intAttempt
    Write #1, "Запятая попытка ", intAttempt
    Write #1, "Точка с запятой попытка "; intAttempt
    Write #1, "Spc(2) попытка "; Spc(2); intAttempt
    Write #1, "Tab попытка "; Tab; intAttempt
    Write #1, "Tab(1) попытка "; Tab(1); intAttempt
    Write #1, "Tab(2) попытка "; Tab(2); intAttempt
    Write #1, "Tab(10) попытка "; Tab(10); intAttempt
    Write #1, "Tab(11) попытка "; Tab(11); intAttempt
    Write #1, "Пример продолжения записи ";
    Write #1, "в ту же самую строку"
    Debug.Print "Успешная запись в файл попытка " & intAttempt
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Блокировки. Модификатор Shared
Sub File_Lock_Shared_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Open "D:\VBA\Запись данных.txt" For Output Shared As #intFile
    Write #intFile, "Запись Word Shared", 200
    Debug.Print "Успешная запись данных Word"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Блокировки. Модификатор Lock Read
Sub File_Procedure_Lock_Read_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Open "D:\VBA\Запись данных.txt" For Output Lock Read As #intFile
    Write #intFile, "Запись Word Lock Read", 200
    Debug.Print "Успешная запись данных Word"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Блокировки. Модификатор Lock Write
Sub File_Procedure_Lock_Write_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Open "D:\VBA\Запись данных.txt" For Output Lock Write As #intFile
    Write #intFile, "Данные моей программы", 200
    Debug.Print "Успешная запись данных Word"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'Блокировки. Модификатор Lock Read Write
Sub File_Procedure_Lock_Read_Write_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Open "D:\VBA\Запись данных.txt" For Output Lock Read Write As #intFile
    Write #intFile, "Данные моей программы", 200
    Debug.Print "Успешная запись данных Word"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'Файлы произвольного доступа. Put
Sub File_Random_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Open "D:\VBA\Random.txt" For Random As #intFile Len = 4
    Put #intFile, 1, 100
    Put #intFile, 2, 200
    Put #intFile, 3, 100500
    Debug.Print "Успешная запись данных"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Файлы произвольного доступа. Get
Sub File_Random_02()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim lngVar As Long
    Open "D:\VBA\Random.txt" For Random As #intFile Len = 4
    Get #intFile, 3, lngVar
    Debug.Print "Успешное чтение данных:", lngVar
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'Файлы произвольного доступа. Нельзя Write, Print
Sub File_Random_03()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Open "D:\VBA\Запись данных.txt" For Random As #intFile
    Print #intFile, "Данные моей программы", 200
    Debug.Print "Успешная запись данных"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'Файлы произвольного доступа. Добавление новой записи
Sub File_Random_04()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Open "D:\VBA\Random.txt" For Random As #intFile Len = 4
    Put #intFile, 10, 100
    Debug.Print "Успешная запись данных"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Файлы произвольного доступа. Последовательное добавление новой записи
Sub File_Random_05()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim byteRecordLenth As Byte
    byteRecordLenth = 4
    Open "D:\VBA\Random.txt" For Random As #intFile Len = byteRecordLenth
    Put #intFile, LOF(intFile) / byteRecordLenth + 1, 100
    Debug.Print "Успешная запись данных"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Random доступен на чтение и запись. Read Write по умолчанию
Sub File_Random_06()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim lngVar As Long
    Open "D:\VBA\Random.txt" For Random Access Read Write As #intFile Len = 4
    Put #intFile, 3, 200500
    Debug.Print "Успешная запись данных"
    Get #intFile, 3, lngVar
    Debug.Print "Успешное чтение данных:"; lngVar
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'Открытие файла произвольного доступа только на чтение
Sub File_Random_07()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim lngVar As Long
    Open "D:\VBA\Random.txt" For Random Access Read As #intFile Len = 4
    'Put #intFile, 3, 200500
    'Debug.Print "Успешная запись данных"
    Get #intFile, 3, lngVar
    Debug.Print "Успешное чтение данных:"; lngVar
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'Открытие файла произвольного доступа только на запись
Sub File_Random_08()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim lngVar As Long
    Open "D:\VBA\Random.txt" For Random Access Write As #intFile Len = 4
    Put #intFile, 3, 300500
    Debug.Print "Успешная запись данных"
    'Get #intFile, 3, lngVar
    'Debug.Print "Успешное чтение данных:"; lngVar
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Файловый указатель. Оператор Seek, функция Seek и Loc
Sub File_Random_09()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim lngVar As Long
    Open "D:\VBA\Файловый указатель.txt" For Random As #intFile Len = 8
    Put #intFile, , "Первый"
    Debug.Print "Seek после Первый:"; Seek(intFile)
    Put #intFile, , "Второй"
    Debug.Print "Seek после Второй:"; Seek(intFile)
    Debug.Print "Loc после Второй: "; Loc(intFile)
    Seek intFile, 1
    Debug.Print "Seek после Seek intFile, 1:"; Seek(intFile)
    Debug.Print "Loc после Seek intFile, 1:"; Loc(intFile)
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'Бинарные файлы. Запись
Sub File_Binary_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim strFirstName As String * 20: strFirstName = "Михаил"
    Dim strLastName As String * 20: strLastName = "Гурьянов"
    Dim strGender As String * 1: strGender = "M"
    Open "D:\VBA\Person.bin" For Binary As #intFile
    Put #intFile, , strFirstName
    Put #intFile, , strLastName
    Put #intFile, , strGender
    Put #intFile, , #2/13/1972#
    Put #intFile, , True
    Debug.Print "Успешная запись данных"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Бинарные файлы. Чтение
Sub File_Binary_02()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim strFirstName As String * 20
    Dim strLastName As String * 20
    Dim strGender As String * 1
    Dim dtBirthDate As Date
    Dim boolRFCitizen As Boolean
    Open "D:\VBA\Person.bin" For Binary As #intFile
    Get #intFile, , strFirstName
    Get #intFile, , strLastName
    Get #intFile, , strGender
    Get #intFile, , dtBirthDate
    Get #intFile, , boolRFCitizen
    Debug.Print strFirstName; strLastName; strGender, dtBirthDate; boolRFCitizen
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

