Attribute VB_Name = "m_05_05_Selection"
Option Explicit
'Что в Selection?
Private Sub d_01_SelectionContent()
    On Error GoTo ErrorHandler
    Select Case Selection.Type
        Case wdNoSelection: Debug.Print "wdNoSelection"
        Case wdSelectionBlock: Debug.Print "wdSelectionBlock"
        Case wdSelectionColumn: Debug.Print "wdSelectionColumn"
        Case wdSelectionInlineShape: Debug.Print "wdSelectionInlineShape"
        Case wdSelectionIP: Debug.Print "wdSelectionIP"
        Case wdSelectionNormal: Debug.Print "wdSelectionNormal"
        Case wdSelectionRow: Debug.Print "wdSelectionRow"
        Case wdSelectionShape: Debug.Print "wdSelectionShape"
    End Select
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Создаём копию файла для экспериментов
Private Sub d_03_CopyFile()
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = wdAlertsNone   ' отключаем предупреждения
    Documents.Add("D:\VBA\Для чтения Word\Демонстрации.docx"). _
        SaveAs2 "Microsoft Visual Basic.docx"
Finalization:
    Application.DisplayAlerts = wdAlertsAll   ' возвращаем обратно
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'Свойства Selection 1
Private Sub d_03_SelectionProperties_1()
    On Error GoTo ErrorHandler
    Selection.Font.AllCaps = True
    Selection.Font.Bold = True
    Selection.Font.Name = "Courier New"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Свойства Selection 2
Private Sub d_04_SelectionProperties_2()
    On Error GoTo ErrorHandler
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Borders.Shadow = True
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Методы Selection. DetectLanguage
Private Sub d_05_DetectLanguage()
    On Error GoTo ErrorHandler
    Selection.LanguageDetected = False
    Selection.DetectLanguage
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Методы Selection
Private Sub d_06_Methods()
    On Error GoTo ErrorHandler
    Selection.Copy
    Selection.Cut
    Selection.Paste
    Selection.CopyFormat
    Selection.PasteFormat
    Selection.CopyAsPicture
    Selection.PasteSpecial DataType:=wdPasteMetafilePicture
    Selection.InsertAfter "Вставка после выделения"
    Selection.InsertBefore "Вставка перед выделением"
    Selection.InsertDateTime
    Selection.InsertBreak
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Методы Selection Sorting
Private Sub d_07_Sorting()
    On Error GoTo ErrorHandler
    Selection.Sort SortOrder:=wdSortOrderAscending
    Selection.Sort SortOrder:=wdSortOrderDescending
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Устанавляваем английский язык для слов латиницей
Sub d08_SetEnglishLanguage()
    On Error GoTo ErrorHandler
    Const strCharactersLat As String = " abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim rngSelection As Range, rngWord As Range, rngChar As Range
    Dim boolIsLat As Boolean, intPos As Integer
    Set rngSelection = Selection.Range
    For Each rngWord In rngSelection.Words
        boolIsLat = True
        For Each rngChar In rngWord.Characters
            intPos = InStr(strCharactersLat, rngChar.Text)
            If intPos = 0 Then
                boolIsLat = False
                Exit For
            End If
        Next rngChar
        If boolIsLat = True Then
            rngWord.LanguageID = wdEnglishUS
            rngWord.NoProofing = False
        End If
    Next rngWord
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка: " & Err.Number & " / " & Err.Description
End Sub
