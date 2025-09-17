VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "People"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Класс: People
' Описание: Коллекция объектов Person

Private colPeople As Collection

' Инициализация коллекции
Private Sub Class_Initialize()
    Set colPeople = New Collection
End Sub

' Освобождение ресурсов
Private Sub Class_Terminate()
    Set colPeople = Nothing
End Sub

' Добавление объекта Person в коллекцию
Public Sub Add(objPerson As Person)
    On Error Resume Next
    ' Проверяем, существует ли уже объект с таким ID
    Dim temp As Person
    Set temp = Nothing
    Set temp = colPeople(CStr(objPerson.ID))
    
    If Not temp Is Nothing Then
        ' Если объект с таким ID уже существует, генерируем ошибку
        Err.Raise 457, "People.Add", "Объект с ID " & objPerson.ID & " уже существует в коллекции"
    Else
        ' Добавляем объект в коллекцию с ключом, равным его ID
        colPeople.Add objPerson, CStr(objPerson.ID)
    End If
    On Error GoTo 0
End Sub

' Удаление объекта из коллекции по индексу
Public Sub RemoveByIndex(intIndex As Integer)
    If intIndex > 0 And intIndex <= colPeople.Count Then
        colPeople.Remove intIndex
    Else
        Err.Raise 9, "People.RemoveByIndex", "Индекс вне диапазона"
    End If
End Sub

' Удаление объекта из коллекции по ID
Public Sub RemoveByID(lngID As Long)
    On Error Resume Next
    colPeople.Remove CStr(lngID)
    If Err.Number <> 0 Then
        Err.Raise 5, "People.RemoveByID", "Объект с ID " & lngID & " не найден в коллекции"
    End If
    On Error GoTo 0
End Sub

' Получение количества элементов в коллекции
Public Property Get Count() As Long
    Count = colPeople.Count
End Property

' Получение элемента по индексу
Public Function ItemByIndex(intIndex As Integer) As Person
    If intIndex > 0 And intIndex <= colPeople.Count Then
        Set ItemByIndex = colPeople(intIndex)
    Else
        Set ItemByIndex = Nothing
        Err.Raise 9, "People.ItemByIndex", "Индекс вне диапазона"
    End If
End Function

' Получение элемента по ID
Public Function ItemByID(lngID As Long) As Person
    On Error Resume Next
    Set ItemByID = colPeople(CStr(lngID))
    If Err.Number <> 0 Then
        Set ItemByID = Nothing
        Err.Raise 5, "People.ItemByID", "Объект с ID " & lngID & " не найден в коллекции"
    End If
    On Error GoTo 0
End Function

' Проверка существования элемента по ID
Public Function Exists(lngID As Long) As Boolean
    On Error Resume Next
    Dim temp As Person
    Set temp = colPeople(CStr(lngID))
    Exists = (Err.Number = 0)
    On Error GoTo 0
End Function

' Очистка коллекции
Public Sub Clear()
    Set colPeople = New Collection
End Sub

' Получение коллекции для использования в For Each
Public Function GetCollection() As Collection
    Set GetCollection = colPeople
End Function

' Поиск объектов по критерию
Public Function Find(strCriteria As String, Optional strField As String = "LastName") As People
    Dim objResult As People
    Set objResult = New People
    Dim objPerson As Person
    Dim varValue As Variant
    
    For Each objPerson In colPeople
        ' Получаем значение поля для проверки
        Select Case LCase(strField)
            Case "id": varValue = objPerson.ID
            Case "lastname": varValue = objPerson.LastName
            Case "firstname": varValue = objPerson.FirstName
            Case "patronymic": varValue = objPerson.Patronymic
            Case "phone": varValue = objPerson.Phone
            Case "email": varValue = objPerson.Email
            Case "city": varValue = objPerson.City
            Case "address": varValue = objPerson.Address
            Case "birthdate": varValue = objPerson.BirthDate
            Case "gender": varValue = objPerson.Gender
            Case "passportnumber": varValue = objPerson.PassportNumber
            Case "passportseries": varValue = objPerson.PassportSeries
            Case "passportissuedby": varValue = objPerson.PassportIssuedBy
            Case Else: varValue = Null
        End Select
        
        ' Проверяем соответствие критерию
        If Not IsNull(varValue) Then
            If TypeName(varValue) = "String" Then
                ' Для строк проверяем частичное совпадение
                If InStr(1, CStr(varValue), strCriteria, vbTextCompare) > 0 Then
                    objResult.Add objPerson
                End If
            Else
                ' Для других типов проверяем точное совпадение
                If varValue = strCriteria Then
                    objResult.Add objPerson
                End If
            End If
        End If
    Next objPerson
    
    Set Find = objResult
End Function


' Функция для загрузки данных из файла в коллекцию People
Public Sub LoadFromFile(strFilePath As String)
    On Error GoTo ErrorHandler
    Dim intFileNumber As Integer
    Dim strLine As String
    Dim arrValues() As String
    Dim objPerson As Person
    Dim lngRow As Long
    
    ' Проверяем существование файла
    If Dir(strFilePath) = "" Then
        Err.Raise 50002, "LoadFromFile", "Файл не найден"
        Exit Sub
    End If
    
    ' Открываем файл для чтения
    intFileNumber = FreeFile()
    Open strFilePath For Input As intFileNumber
    
    lngRow = 0
    Do Until EOF(intFileNumber)
        Line Input #intFileNumber, strLine
        lngRow = lngRow + 1
        
        ' Пропускаем заголовок (первую строку)
        If lngRow = 1 Then
            GoTo EndOfLoop
        End If
        
        ' Разбиваем строку на значения
        arrValues = Split(strLine, ",")
        
        ' Проверяем, что в строке достаточно значений
        If UBound(arrValues) < 12 Then
            Debug.Print "Пропущена строка " & lngRow & ": недостаточно значений"
            GoTo EndOfLoop
        End If
        
        ' Создаем новый объект Person
        Set objPerson = New Person
        
        ' Заполняем свойства объекта
        With objPerson
            ' Первое значение - ID (Long)
            If IsNumeric(Trim(arrValues(0))) Then
                .ID = CLng(Trim(arrValues(0)))
            Else
                .ID = 0
            End If
            
            ' Второе значение - Фамилия (String)
            .LastName = ParseString(arrValues(1))
            
            ' Третье значение - Имя (String)
            .FirstName = ParseString(arrValues(2))
            
            ' Четвертое значение - Отчество (String)
            .Patronymic = ParseString(arrValues(3))
            
            ' Пятое значение - Телефон (String)
            .Phone = ParseString(arrValues(4))
            
            ' Шестое значение - Email (String)
            .Email = ParseString(arrValues(5))
            
            ' Седьмое значение - Город (String)
            .City = ParseString(arrValues(6))
            
            ' Восьмое значение - Адрес (String)
            .Address = ParseString(arrValues(7))
            
            ' Девятое значение - Дата рождения (Date)
            .BirthDate = ParseDate(arrValues(8))
            
            ' Десятое значение - Пол (String)
            .Gender = ParseString(arrValues(9))
            
            ' Одиннадцатое значение - Номер паспорта (String)
            .PassportNumber = ParseString(arrValues(10))
            
            ' Двенадцатое значение - Серия паспорта (String)
            .PassportSeries = ParseString(arrValues(11))
            
            ' Тринадцатое значение - Кем выдан паспорт (String)
            .PassportIssuedBy = ParseString(arrValues(12))
        End With
        
        ' Добавляем объект в коллекцию
        colPeople.Add objPerson
EndOfLoop:
    Loop
Finalization:
    ' Закрываем файл
    Close intFileNumber
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

' Вспомогательная функция для парсинга строковых значений
Private Function ParseString(strValue As String) As String
    ' Удаляем лишние пробелы и кавычки
    strValue = Trim(strValue)
    
    ' Обрабатываем строки в кавычках
    If Left(strValue, 1) = """" And Right(strValue, 1) = """" Then
        strValue = Mid(strValue, 2, Len(strValue) - 2)
        ' Восстанавливаем экранированные кавычки
        strValue = Replace(strValue, """""", """")
    End If
    
    ' Обрабатываем значение Null
    If strValue = "Null" Then
        ParseString = ""
    Else
        ParseString = strValue
    End If
End Function

' Вспомогательная функция для парсинга дат
Private Function ParseDate(strValue As String) As Date
    ' Обрабатываем даты в формате #yyyy-mm-dd#
    If Left(strValue, 1) = "#" And Right(strValue, 1) = "#" Then
        Dim strDate As String
        strDate = Mid(strValue, 2, Len(strValue) - 2)
        
        ' Пытаемся преобразовать в дату
        On Error Resume Next
        ParseDate = CDate(strDate)
        On Error GoTo 0
    Else
        ' Возвращаем дату по умолчанию при ошибке
        ParseDate = #1/1/1900#
    End If
End Function


' Функция для загрузки данных из файла в коллекцию People
Public Sub LoadFromFile2(strFilePath As String)
    On Error GoTo ErrorHandler
    Dim intFileNumber As Integer
    Dim strLine As String
    Dim arrValues() As String
    Dim objPerson As Person
    Dim lngRow As Long
    Dim lngID As Integer, strLastName As String, strFirstName As String
    Dim strPatronymic As String, strPhone As String, strEmail As String
    Dim strCity As String, strAddress As String, dtBirthDate As Date
    Dim strGender As String, strPassportNumber As String
    Dim strPassportSeries As String, strPassportIssuedBy As String
    
    
    ' Проверяем существование файла
    If Dir(strFilePath) = "" Then
        Err.Raise 50002, "LoadFromFile", "Файл не найден"
        Exit Sub
    End If
    
    ' Открываем файл для чтения
    intFileNumber = FreeFile()
    Open strFilePath For Input As intFileNumber
    
    lngRow = 0
    Do Until EOF(intFileNumber)
        lngRow = lngRow + 1
                        
        ' Заполняем свойства объекта
        Input #intFileNumber, lngID, strLastName, strFirstName, _
            strPatronymic, strPhone, strEmail, strCity, _
            strAddress, dtBirthDate, strGender, _
            strPassportNumber, strPassportSeries, strPassportIssuedBy
        
        ' Пропускаем заголовок (первую строку)
        If lngRow = 1 Then GoTo EndOfLoop

        ' Создаем новый объект Person
        Set objPerson = New Person
                
        objPerson.ID = lngID: objPerson.LastName = strLastName
        objPerson.FirstName = strFirstName: objPerson.Patronymic = strPatronymic
        objPerson.Phone = strPhone: objPerson.Email = strEmail
        objPerson.City = strCity:  objPerson.Address = strAddress
        objPerson.BirthDate = dtBirthDate: objPerson.Gender = strGender
        objPerson.PassportNumber = strPassportNumber
        objPerson.PassportSeries = strPassportSeries
        objPerson.PassportIssuedBy = strPassportIssuedBy
                               
        ' Добавляем объект в коллекцию
        colPeople.Add objPerson
EndOfLoop:
    Loop
Finalization:
    ' Закрываем файл
    Close intFileNumber
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


