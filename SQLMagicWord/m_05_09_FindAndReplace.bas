Attribute VB_Name = "m_05_09_FindAndReplace"
Option Explicit

'Создаём копию файла "Visual Basic" для экспериментов
Private Sub d_01_CopyFileForExperiment()
    On Error GoTo ErrorHandler
    Dim strSourceFilename As String, strTargetFilename As String
    strSourceFilename = "D:\VBA\Для чтения Word\Простой текст.docx"
    strTargetFilename = "D:\VBA\Word\Visual Basic.docx"
    Call CopyFile(SourceFilename:=strSourceFilename, TargetFilename:=strTargetFilename)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Поиск во всём тексте. Наличие текста
Private Sub d_02_FindWholeDocument()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    With objDocument.Content.Find
        .ClearFormatting          ' без учёта форматирования
        .Text = "visual basic"    ' что ищем
        .MatchCase = False        ' не учитывать регистр
        .MatchWholeWord = True    ' искать только целое слово
        .Execute
        If .Found Then
            Debug.Print "Текст найден"
        Else
            Debug.Print "Текст не найден"
        End If
    End With
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



'Поиск во всём тексте. Нахождение текста
Private Sub d_03_FindAll()
    On Error GoTo ErrorHandler
    Dim objDocument As Document, rngItem As Range, intCounter As Integer
    Dim rngSentence
    Set objDocument = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    Set rngItem = objDocument.Content
    intCounter = 0
    With rngItem.Find
        .ClearFormatting          ' без учёта форматирования
        .Text = "visual basic"    ' что ищем
        .MatchCase = False        ' не учитывать регистр
        .MatchWholeWord = True    ' искать только целое слово
        Do While .Execute 'True , если операция поиска выполнена успешно
            rngItem.Font.Bold = True ' выделим найденное
            Set rngSentence = rngItem.Duplicate
            rngSentence.Expand Unit:=wdSentence
            rngSentence.HighlightColorIndex = wdYellow ' выделим предложение в котором искомый текст
            rngSentence.Font.Color = vbBlue
            rngItem.Collapse wdCollapseEnd ' смещаем курсор дальше
            intCounter = intCounter + 1
        Loop
    End With
    Debug.Print "Найдено:"; intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Замена текста
Private Sub d_04_ReplaceAll()
    On Error GoTo ErrorHandler
    Dim objDocument As Document, rngItem As Range, intCounter As Integer
    Dim rngSentence
    Set objDocument = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    With objDocument.Content.Find
        .Text = "VBA"
        .Replacement.Text = "Visual Basic for Applications"
        .MatchCase = False
        .Execute Replace:=wdReplaceAll
    End With
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Поиск по форматированию. Поиск только в заголовках первого уровня
Private Sub d_05_FindByFormat()
    On Error GoTo ErrorHandler
    Dim objDocument As Document, rngItem As Range, intCounter As Integer
    Dim rngSentence
    Set objDocument = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    Set rngItem = objDocument.Content
    intCounter = 0
    With rngItem.Find
        .ClearFormatting          ' очищаем форматирование
        .Style = "Заголовок 1"
        .Text = "visual basic"    ' что ищем
        .MatchCase = False        ' не учитывать регистр
        .MatchWholeWord = True    ' искать только целое слово
        Do While .Execute 'True , если операция поиска выполнена успешно
            rngItem.Font.Bold = True ' выделим найденное
            rngItem.Font.Italic = True ' выделим найденное
            rngItem.Collapse wdCollapseEnd ' смещаем курсор дальше
            intCounter = intCounter + 1
        Loop
    End With
    Debug.Print "Найдено:"; intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



'Поиск в выделении
Private Sub d_06_FindInSelection()
    On Error GoTo ErrorHandler
    Dim rngItem As Range, intCounter As Integer, lngEnd As Long
    Set rngItem = Selection.Range
    lngEnd = Selection.Range.End
    intCounter = 0
    With rngItem.Find
        .ClearFormatting          ' без учёта форматирования
        .Text = "visual basic"    ' что ищем
        .MatchCase = False        ' не учитывать регистр
        .MatchWholeWord = True    ' искать только целое слово
        .Wrap = wdFindStop
        Do While .Execute And rngItem.End < lngEnd
            rngItem.Font.Bold = True ' выделим найденное
            rngItem.HighlightColorIndex = wdYellow
            rngItem.Collapse wdCollapseEnd ' смещаем курсор дальше
            intCounter = intCounter + 1
        Loop
    End With
    Debug.Print "Найдено:"; intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
