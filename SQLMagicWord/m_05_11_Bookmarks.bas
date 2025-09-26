Attribute VB_Name = "m_05_11_Bookmarks"
Option Explicit

'Создаём копию файла "Заявление на отпуск" для экспериментов
Private Sub d_01_CopyFileForExperiment()
    On Error GoTo ErrorHandler
    Dim strSourceFilename As String, strTargetFilename As String
    strSourceFilename = "D:\VBA\Для чтения Word\Заявление на отпуск.docx"
    strTargetFilename = "D:\VBA\Word\Заявление.docx"
    Call CopyFile(SourceFilename:=strSourceFilename, TargetFilename:=strTargetFilename)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Просмотр закладок
Private Sub d_02_ShowBookmarks()
    On Error GoTo ErrorHandler
    Dim docStatement As Document, objBookmarkItem As Bookmark
    Set docStatement = Documents.Open("D:\VBA\Word\Заявление.docx")
    For Each objBookmarkItem In docStatement.Bookmarks
        Debug.Print objBookmarkItem.Name
    Next objBookmarkItem
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Проверка наличия закладки
Private Sub d_03_Exists()
    On Error GoTo ErrorHandler
    Dim docStatement As Document, objBookmarkItem As Bookmark
    Set docStatement = Documents.Open("D:\VBA\Word\Заявление.docx")
    If docStatement.Bookmarks.Exists("Сотрудник") Then
        Debug.Print "Закладка существует"
    Else
        Debug.Print "Закладка не существует"
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Форматирование закладок
Private Sub d_04_FormatBookmarks()
    On Error GoTo ErrorHandler
    Dim docStatement As Document, objBookmarkItem As Bookmark
    Set docStatement = Documents.Open("D:\VBA\Word\Заявление.docx")
    For Each objBookmarkItem In docStatement.Bookmarks
        objBookmarkItem.Range.Font.Bold = False
    Next objBookmarkItem
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



'Присваивание значений тексту закладок
Private Sub d_05_SetBookmarkText()
    On Error GoTo ErrorHandler
    Dim docStatement As Document
    Dim strEmployee As String, dtStart As Date, dtEnd As Date
    strEmployee = "менеджера О. И. Бендера"
    dtStart = #9/15/2025#
    dtEnd = #9/26/2025#
    Set docStatement = Documents.Add("D:\VBA\Для чтения Word\Заявление на отпуск.docx")
    'Set docStatement = Documents.Add("D:\VBA\Для чтения Word\Простой текст.docx")
    docStatement.Bookmarks("Сотрудник").Range.Text = strEmployee
    docStatement.Bookmarks("ДатаНачала").Range.Text = Format(dtStart, "«dd» mmmm yyyy")
    docStatement.Bookmarks("ДатаКонца").Range.Text = Format(dtEnd, "«dd» mmmm yyyy")
    Exit Sub
ErrorHandler:
    MsgBox "Oшибка при формировании документа: " & Err.Number & vbCrLf _
        & Err.Description & vbCrLf & "Проверьте документ и обратитесь в техподдержку"
End Sub

'Добавление закладок из выделения
Private Sub d_06_AddBookmark()
    On Error GoTo ErrorHandler
    Dim docStatement As Document
    Dim strEmployee As String, dtStart As Date, dtEnd As Date
    Set docStatement = Documents.Open("D:\VBA\Word\Заявление.docx")
    docStatement.Bookmarks.Add "ИзВыделения"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Удаление закладок
Private Sub d_07_DeleteBookmark()
    On Error GoTo ErrorHandler
    Dim docStatement As Document
    Dim strEmployee As String, dtStart As Date, dtEnd As Date
    Set docStatement = Documents.Open("D:\VBA\Word\Заявление.docx")
    docStatement.Bookmarks("ИзВыделения").Delete
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Добавление закладки из диапазона
Private Sub d_08_AddBookmarkFromRange()
    On Error GoTo ErrorHandler
    Dim docStatement As Document
    Dim rngPositionOfBoss As Range
    Set docStatement = Documents.Open("D:\VBA\Word\Заявление.docx")
    Set rngPositionOfBoss = docStatement.Range(0, docStatement.Range.Words(2).End - 1)
    docStatement.Bookmarks.Add "ДолжностьПервогоЛица", rngPositionOfBoss
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Присваивание значений тексту закладок c сохранением закладок
Private Sub d_09_BookmarksSaving()
    On Error GoTo ErrorHandler
    Dim docStatement As Document, rngItem As Range
    Dim strEmployee As String, dtStart As Date, dtEnd As Date
    strEmployee = "менеджера О. И. Бендера"
    dtStart = #9/15/2025#
    dtEnd = #9/26/2025#
    Set docStatement = Documents.Add("D:\VBA\Для чтения Word\Заявление на отпуск.docx")
    
    If docStatement.Bookmarks.Exists("Сотрудник") Then
        Set rngItem = docStatement.Bookmarks("Сотрудник").Range
        docStatement.Bookmarks("Сотрудник").Range.Text = strEmployee
        docStatement.Bookmarks.Add "Сотрудник", rngItem
    End If
    
    If docStatement.Bookmarks.Exists("ДатаНачала") Then
        Set rngItem = docStatement.Bookmarks("ДатаНачала").Range
        docStatement.Bookmarks("ДатаНачала").Range.Text = Format(dtStart, "«dd» mmmm yyyy")
        docStatement.Bookmarks.Add "ДатаНачала", rngItem
    End If
    
    If docStatement.Bookmarks.Exists("ДатаКонца") Then
        Set rngItem = docStatement.Bookmarks("ДатаКонца").Range
        docStatement.Bookmarks("ДатаКонца").Range.Text = Format(dtEnd, "«dd» mmmm yyyy")
        docStatement.Bookmarks.Add "ДатаКонца", rngItem
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



