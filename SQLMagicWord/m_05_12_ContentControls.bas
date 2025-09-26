Attribute VB_Name = "m_05_12_ContentControls"
Option Explicit

'Создаём копию файла "Заявление на отпуск" для экспериментов
Private Sub d_01_CopyFileForExperiment()
    On Error GoTo ErrorHandler
    Dim strSourceFilename As String, strTargetFilename As String
    strSourceFilename = "D:\VBA\Для чтения Word\Заявление на отпуск Content Controls.docx"
    strTargetFilename = "D:\VBA\Word\Заявление.docx"
    Call CopyFile(SourceFilename:=strSourceFilename, TargetFilename:=strTargetFilename)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Просмотр. Ни Tag ни Title не являются уникальныму идентификаторами! Свойства Name нет
Private Sub d_02_ShowContentControls()
    On Error GoTo ErrorHandler
    Dim docApplication As Document, objCCItem As ContentControl
    Dim intCCIndex As Integer
    Set docApplication = Documents.Open("D:\VBA\Word\Заявление.docx")
    For intCCIndex = 1 To docApplication.ContentControls.Count
        Debug.Print intCCIndex,
        Debug.Print docApplication.ContentControls(intCCIndex).Tag,
        Debug.Print docApplication.ContentControls(intCCIndex).Title,
        Debug.Print docApplication.ContentControls(intCCIndex).Type
        'wdContentControlRichText 0 / wdContentControlText 1
    Next intCCIndex
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Обращение к ContentControl. Элементy управления содержимым
Private Sub d_03_GetContentControl()
    On Error GoTo ErrorHandler
    Dim docApplication As Document
    Set docApplication = Documents.Open("D:\VBA\Word\Заявление.docx")
    Debug.Print docApplication.ContentControls.Item(1).Tag
    Debug.Print docApplication.ContentControls.Item(1).Title
    Debug.Print docApplication.ContentControls(2).Title
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Форматирование ContentControl. Элемента управления содержимым
Private Sub d_04_FormatContentControls()
    On Error GoTo ErrorHandler
    Dim docApplication As Document, objCCItem As ContentControl
    Set docApplication = Documents.Open("D:\VBA\Word\Заявление.docx")
    For Each objCCItem In docApplication.ContentControls
        objCCItem.Range.Font.Bold = False
        objCCItem.Range.HighlightColorIndex = wdYellow
    Next objCCItem
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Присваивание значений ContentControl. Элементy управления содержимым
Private Sub d_05_SetContentControlText()
    On Error GoTo ErrorHandler
    Dim docApplication As Document, objCCItem As ContentControl
    Dim strEmployee As String, dtStart As Date, dtEnd As Date
    strEmployee = "менеджера О. И. Бендера"
    dtStart = #9/15/2025#
    dtEnd = #9/26/2025#
    Set docApplication = Documents.Open("D:\VBA\Word\Заявление.docx")
    'Set docApplication = Documents.Add("D:\VBA\Для чтения Word\Заявление на отпуск Content Controls.docx")
    For Each objCCItem In docApplication.ContentControls
        Select Case objCCItem.Tag
            Case "EmployeeTag": objCCItem.Range.Text = strEmployee
            Case "StartTag": objCCItem.Range.Text = Format(dtStart, "«dd» mmmm yyyy")
            Case "EndTag": objCCItem.Range.Text = Format(dtEnd, "«dd» mmmm yyyy")
        End Select
    Next objCCItem
    Exit Sub
ErrorHandler:
    MsgBox "Oшибка при формировании документа: " & Err.Number & vbCrLf _
        & Err.Description & vbCrLf & "Проверьте документ и обратитесь в техподдержку"
End Sub


'Окончательный вариант формирования заявления
Private Sub d_06_CreateApplication()
    On Error GoTo ErrorHandler
    Dim docApplication As Document, objCCItem As ContentControl
    Dim strEmployee As String, dtStart As Date, dtEnd As Date
    strEmployee = "менеджера О. И. Бендера"
    dtStart = #9/15/2025#
    dtEnd = #9/26/2025#
    Set docApplication = Documents.Add("D:\VBA\Для чтения Word\Заявление на отпуск Content Controls.docx")
    For Each objCCItem In docApplication.ContentControls
        Select Case objCCItem.Tag
            Case "EmployeeTag": objCCItem.Range.Text = strEmployee
            Case "StartTag": objCCItem.Range.Text = Format(dtStart, "«dd» mmmm yyyy")
            Case "EndTag": objCCItem.Range.Text = Format(dtEnd, "«dd» mmmm yyyy")
        End Select
    Next objCCItem
    Exit Sub
ErrorHandler:
    MsgBox "Oшибка при формировании документа: " & Err.Number & vbCrLf _
        & Err.Description & vbCrLf & "Проверьте документ и обратитесь в техподдержку"
End Sub



'Добавление ContentControl из выделения. Элемент управления содержимым.
Private Sub d_07_AddContentControl()
    On Error GoTo ErrorHandler
    Dim docApplication As Document, ccItem As ContentControl
    Set docApplication = Documents.Open("D:\VBA\Word\Заявление.docx")
    Set ccItem = docApplication.ContentControls.Add(Type:=wdContentControlRichText)
    ccItem.Tag = "Selection"
    ccItem.Title = "Selection"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Удаление ContentControl. Элемент управления содержимым.
Private Sub d_08_DeleteContentControl()
    On Error GoTo ErrorHandler
    Dim docApplication As Document
    Set docApplication = Documents.Open("D:\VBA\Word\Заявление.docx")
    'docApplication.ContentControls(2).LockContentControl = False
    docApplication.ContentControls(2).Delete
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Добавление ContentControl из диапазона. Элемент управления содержимым.
Private Sub d_09_AddContentControlFromRange()
    On Error GoTo ErrorHandler
    Dim docApplication As Document
    Dim rngPositionOfBoss As Range
    Dim ccItem As ContentControl
    Set docApplication = Documents.Open("D:\VBA\Word\Заявление.docx")
    Set rngPositionOfBoss = docApplication.Range(0, docApplication.Range.Words(2).End - 1)
    Set ccItem = docApplication.ContentControls.Add(Range:=rngPositionOfBoss)
    ccItem.Title = "PosOfBoss"
    ccItem.Tag = "PosOfBoss"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
