Attribute VB_Name = "m_07_03_Dialogs"
Option Explicit

'Выбор папки
Private Sub d_01_FileDialogFolderPicker()
    On Error Resume Next
    Dim objFileDialog As FileDialog
    Set objFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With objFileDialog
        .Title = "Выберите папку"
        .InitialFileName = "d:\vba" ' Начальная папка
        If .Show = -1 Then ' -1 - Выбрана кнопка действия, 0 - Отмена
            Debug.Print .SelectedItems(1)
        End If
    End With
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка: " & Err.Description, vbCritical
End Sub


'Выбор файлов
Private Sub d_02_FileDialogFilePicker()
    On Error Resume Next
    Dim objFileDialog As FileDialog, varFileName As Variant
    Set objFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    With objFileDialog
        .Title = "Выберите файлы. Можно несколько, нажимая Shift и Ctrl"
        .InitialFileName = "d:\vba\*.*" ' Начальная папка
        .AllowMultiSelect = True
        If .Show = -1 Then '-1 - Выбрана кнопка действия, 0 - Отмена
            For Each varFileName In .SelectedItems
                Debug.Print varFileName
            Next varFileName
        End If
    End With
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка: " & Err.Description, vbCritical
End Sub


'Открытие файлов
Private Sub d_03_FileDialogOpen()
    On Error Resume Next
    Dim objFileDialog As FileDialog, varFileName As Variant
    Set objFileDialog = Application.FileDialog(msoFileDialogOpen)
    With objFileDialog
        .Title = "Откройте файлы. Можно несколько, нажимая Shift и Ctrl"
        .InitialFileName = "D:\VBA\Для чтения Word\*.*" ' Начальные файлы
        .Filters.Add "Файлы Word", "*.docx; *.docm", 1 'Добавляем фильтры
        .FilterIndex = 1 'Фильтр по умолчанию
        .AllowMultiSelect = True
        If .Show = -1 Then '-1 - Выбрана кнопка действия, 0 - Отмена
            For Each varFileName In .SelectedItems
                Debug.Print varFileName
            Next varFileName
        End If
        If .SelectedItems.Count > 3 Then
            MsgBox "Количество файлов превышает максимальное количество - три"
        Else
            .Execute
        End If
    End With
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка: " & Err.Description, vbCritical
End Sub


'Сохранение файлов
Private Sub d_04_FileDialogSaveAs()
    On Error Resume Next
    Dim objFileDialog As FileDialog, varFileName As Variant
    Set objFileDialog = Application.FileDialog(msoFileDialogSaveAs)
    With objFileDialog
        .Title = "Сохранить файл как"
        .InitialFileName = "d:\vba\Word\SaveAs.docx" ' Начальный файл
        .ButtonName = "Сохранить как"
        If .Show = -1 Then '-1 - Выбрана кнопка действия, 0 - Отмена
            Debug.Print .SelectedItems(1)
        End If
        .Execute
    End With
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка: " & Err.Description, vbCritical
End Sub
