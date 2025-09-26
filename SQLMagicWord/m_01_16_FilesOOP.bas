Attribute VB_Name = "m_01_16_FilesOOP"
Option Explicit
'Чтение всего файла
Sub FileReading_01()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Dim strContent As String
    Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\ANSI.txt", ForReading)
    strContent = oTextStream.ReadAll()
    MsgBox strContent
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Чтение отдельной строки
Sub FileReading_02()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Dim strContent As String
    Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\ANSI.txt", ForReading)
    strContent = oTextStream.ReadLine()
    MsgBox strContent
    strContent = oTextStream.ReadLine()
    MsgBox strContent
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'Прочитать (Read) и пропустить (Skip) заданное количество символов
Sub FileReading_03()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Dim strContent As String
    Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\ANSI.txt", ForReading)
    oTextStream.Skip 7
    strContent = oTextStream.Read(5)
    Debug.Print strContent
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Пропустить строку в файле
Sub FileReading_04()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Dim strContent As String
    Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\ANSI.txt", ForReading)
    oTextStream.SkipLine
    strContent = oTextStream.ReadLine()
    Debug.Print strContent
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Создать файл, перезаписать, если существует
Sub FileWriting_01()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Static intAttemp As Integer: intAttemp = intAttemp + 1
    Set oTextStream = oFileSystem.CreateTextFile _
        ("D:\VBA\Тестирование FileSystemObject.txt", ForWriting)
    oTextStream.WriteLine ("Тестирование ForWriting, попытка: " & intAttemp)
    Debug.Print "Успешная запись файла, попытка: " & intAttemp; ""
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Открыть файл, дописать
Sub FileWriting_02()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Static intAttemp As Integer: intAttemp = intAttemp + 1
    Set oTextStream = oFileSystem.OpenTextFile _
        ("D:\VBA\Тестирование FileSystemObject.txt", ForAppending)
    oTextStream.WriteLine ("Тестирование ForAppending, попытка: " & intAttemp)
    Debug.Print "Успешная запись файла, попытка: " & intAttemp; ""
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

' Создаем новый файл, если он не существует
Sub FileWriting_03()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject, oTextStream As TextStream
    Dim strFilePath As String: strFilePath = "D:\VBA\Новый файл.txt"
    If Not oFileSystem.FileExists(strFilePath) Then
' Второй параметр - True, если файл можно перезаписать, и False в противном случае
        Set oTextStream = oFileSystem.CreateTextFile(strFilePath, False)
        oTextStream.Close
        Debug.Print "Новый файл создан"
    Else
        Debug.Print "Файл существует"
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Запись многострочного файла одной строкой
Sub FileWriting_04()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Static intAttemp As Integer: intAttemp = intAttemp + 1
    Set oTextStream = oFileSystem.CreateTextFile _
        ("D:\VBA\Многострочный файл.txt", ForWriting)
    oTextStream.WriteLine _
        ("Первая строка" & vbCrLf & "Вторая строка" & vbCrLf & "Третья строка")
    Debug.Print "Успешная запись файла, попытка: " & intAttemp; ""
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Запись без перевода строки (продолжаем строку)
Sub FileWriting_05()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Static intAttemp As Integer: intAttemp = intAttemp + 1
    Set oTextStream = oFileSystem.CreateTextFile _
        ("D:\VBA\Многострочный файл.txt", ForWriting)
    oTextStream.Write ("Первые данные ")
    oTextStream.Write ("Вторые данные")
    Debug.Print "Успешная запись файла, попытка: " & intAttemp; ""
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Вставка пустых строк
Sub FileWriting_06()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Static intAttemp As Integer: intAttemp = intAttemp + 1
    Set oTextStream = oFileSystem.CreateTextFile _
        ("D:\VBA\Многострочный файл.txt", ForWriting)
    oTextStream.WriteLine ("Первая строка")
    oTextStream.WriteBlankLines (5)
    oTextStream.WriteLine ("Последняя строка")
    Debug.Print "Успешная запись файла, попытка: " & intAttemp; ""
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Файловый указатель
Sub FileNavigation_01()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject, oTextStream As TextStream
    Dim strContent As String
    Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\ANSI.txt", ForReading)
    strContent = oTextStream.ReadLine()
    Debug.Print "------ После первого ReadLine ------"
    Debug.Print "AtEndOfLine:", , oTextStream.AtEndOfLine
    Debug.Print "AtEndOfStream:", oTextStream.AtEndOfStream
    Debug.Print "Column:", , oTextStream.Column
    Debug.Print "Line:", , oTextStream.Line
    strContent = oTextStream.ReadLine()
    Debug.Print "------ После второго ReadLine ------"
    Debug.Print "AtEndOfLine:", , oTextStream.AtEndOfLine
    Debug.Print "AtEndOfStream:", oTextStream.AtEndOfStream
    Debug.Print "Column:", , oTextStream.Column
    Debug.Print "Line:", , oTextStream.Line
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Файловый указатель Skip, Read
Sub FileNavigation_02()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject, oTextStream As TextStream
    Dim strContent As String
    Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\ANSI.txt", ForReading)
    oTextStream.Skip (7)
    Debug.Print "------ После Skip(7) ------"
    Debug.Print "AtEndOfLine:", , oTextStream.AtEndOfLine
    Debug.Print "AtEndOfStream:", oTextStream.AtEndOfStream
    Debug.Print "Column:", , oTextStream.Column
    Debug.Print "Line:", , oTextStream.Line
    strContent = oTextStream.Read(6)
    Debug.Print "------ После Read(6) ------"
    Debug.Print "AtEndOfLine:", , oTextStream.AtEndOfLine
    Debug.Print "AtEndOfStream:", oTextStream.AtEndOfStream
    Debug.Print "Column:", , oTextStream.Column
    Debug.Print "Line:", , oTextStream.Line
    Debug.Print "strContent:", strContent
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Работа с UTF-16,  Работа с UTF-8 не поддерживается
Sub FileReading_10()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oFile As File
    Dim oTextStream As TextStream
    Dim strContent As String
    Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\ANSI.txt", ForReading, , False)
    Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\UTF-16.txt", ForReading, , True)
    'Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\UTF-8.txt", ForReading, , True)
    strContent = oTextStream.ReadAll()
    MsgBox strContent
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Чтение UTF-8. Читаем всё
Sub ADODB_Reading_01()
    On Error GoTo ErrorHandler
    Dim oStream As New ADODB.stream
    With oStream
        .Type = 2 ' Текстовый тип
        'HKEY_CLASSES_ROOT\MIME\Database\Charset
        .Charset = "utf-8"
        .Open
        .LoadFromFile "D:\VBA\UTF-8.txt"
        Debug.Print .ReadText
        .Close
    End With
Finalization:
    Set oStream = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Чтение UTF-8. Читаем выборочно
Sub ADODB_Reading_02()
    On Error GoTo ErrorHandler
    Dim oStream As New ADODB.stream
    With oStream
        .Type = 2 ' Текстовый тип
        .Charset = "utf-8"
        .Open
        .LoadFromFile "D:\VBA\UTF-8.txt"
        .SkipLine
        Debug.Print .ReadText(3)
        .Close
    End With
Finalization:
    Set oStream = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'Запись UTF-8
Sub ADODB_Writing_01()
    On Error GoTo ErrorHandler
    Dim oStream As New ADODB.stream
    Dim strContent As String: strContent = "Код, написанный на VBA, компилируется в промежуточный Microsoft P-код (байт-код)"
    With oStream
        .Type = 2 ' Текстовый тип
        .Charset = "utf-8"
        .Open
        .WriteText strContent
        .SaveToFile "D:\VBA\UTF-8 Запись.txt", adSaveCreateOverWrite
        .Close
    End With
    Debug.Print "Запись успешно проведена"
Finalization:
    Set oStream = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Запись cp866
Sub ADODB_Writing_02()
    On Error GoTo ErrorHandler
    Dim oStream As New ADODB.stream
    Dim strContent As String: strContent = "Код, написанный на VBA, компилируется в промежуточный Microsoft P-код (байт-код)"
    With oStream
        .Type = 2 ' Текстовый тип
        'HKEY_CLASSES_ROOT\MIME\Database\Charset
        .Charset = "cp866"
        .Open
        .WriteText strContent
        .SaveToFile "D:\VBA\cp866 Запись.txt", adSaveCreateOverWrite
        .Close
    End With
    Debug.Print "Запись успешно проведена"
Finalization:
    Set oStream = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'Чтение cp866
Sub ADODB_Readingcp866_01()
    On Error GoTo ErrorHandler
    Dim oStream As New ADODB.stream
    With oStream
        .Type = adTypeText
        'HKEY_CLASSES_ROOT\MIME\Database\Charset
        .Charset = "cp866"
        .Open
        .LoadFromFile "D:\VBA\cp866 Запись.txt"
        Debug.Print .ReadText
        .Close
    End With
Finalization:
    Set oStream = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
