Attribute VB_Name = "m_07_01_UI"
Option Explicit

'Получение ссылки на запущенный Outlook
Private Sub d_01_CommandBar()
    
    Dim objDocsCommandBar As CommandBar
    Set objDocsCommandBar = CommandBars.Add("Documents", msoBarFloating)
    objDocsCommandBar.Visible = True
    
    Exit Sub
ErrorHandler:
    Debug.Print "Oшибка:"; Err.Number & " / " & Err.Description
End Sub

'Создание панели инструментов
Private Sub d_02_CreateAdvancedToolbar()
    Dim objToolbar As CommandBar
    Dim objControl As Object
    Dim strBarName As String
    
    strBarName = "РасширеннаяПанель"
    
    ' Удаляем старую панель
    On Error Resume Next
    Application.CommandBars(strBarName).Delete
    On Error GoTo 0
    
    ' Создаем панель
    Set objToolbar = Application.CommandBars.Add( _
        Name:=strBarName, _
        Position:=msoBarFloating, _
        MenuBar:=False, _
        Temporary:=True)
    
    ' Добавляем разные типы элементов
    
    ' 1. Кнопка
    Set objControl = objToolbar.Controls.Add(Type:=msoControlButton)
    With objControl
        .Caption = "Сохранить"
        .FaceId = 3
        .OnAction = "SaveDocument"
        .TooltipText = "Сохранить документ (Ctrl+S)"
    End With
    
    ' 2. Разделитель
    Set objControl = objToolbar.Controls.Add(Type:=msoControlButton)
    objControl.BeginGroup = True
    
    ' 3. Выпадающий список
    Set objControl = objToolbar.Controls.Add(Type:=msoControlDropdown)
    With objControl
        .Caption = "Стили текста"
        .AddItem "Обычный", 1
        .AddItem "Заголовок 1", 2
        .AddItem "Заголовок 2", 3
        .OnAction = "ApplyStyleMacro"
    End With
    
    ' 4. Еще кнопка
    Set objControl = objToolbar.Controls.Add(Type:=msoControlButton)
    With objControl
        .Caption = "Печать"
        .FaceId = 4
        .OnAction = "PrintDocument"
        .TooltipText = "Печать документа (Ctrl+P)"
    End With
    ' Настраиваем панель
    With objToolbar
        .Visible = True
        .Top = 100
        .Left = 200
        .Width = 200
        .Protection = msoBarNoMove ' Запрещаем перемещение
    End With
        
CleanExit:
    Set objControl = Nothing
    Set objToolbar = Nothing
End Sub

'Удаление панели инструментов
Private Sub d_03_DeleteAdvancedToolbar()
    Dim strBarName As String
    
    strBarName = "РасширеннаяПанель"
    
    ' Удаляем старую панель
    On Error Resume Next
    Application.CommandBars(strBarName).Delete
    On Error GoTo 0
    
End Sub

'Создание контекстного меню
Sub d_04_CreateContextMenu()
    Dim objContextMenu As CommandBar
    Dim objMenuItem As CommandBarButton
    Dim strMenuName As String
    
    strMenuName = "МоеКонтекстноеМеню"
    
    ' Удаляем существующее меню
    On Error Resume Next
    Application.CommandBars(strMenuName).Delete
    On Error GoTo 0
    
    ' Создаем контекстное меню
    Set objContextMenu = Application.CommandBars.Add( _
        Name:=strMenuName, _
        Position:=msoBarPopup, _
        MenuBar:=False, _
        Temporary:=True)
    
    ' Добавляем пункты меню
    Set objMenuItem = objContextMenu.Controls.Add(Type:=msoControlButton)
    With objMenuItem
        .Caption = "Форматировать ячейки"
        .BeginGroup = True
        .OnAction = "FormatCellsMacro"
    End With
    
    Set objMenuItem = objContextMenu.Controls.Add(Type:=msoControlButton)
    With objMenuItem
        .Caption = "Вставить специально"
        .OnAction = "SpecialPasteMacro"
    End With
    
    Set objMenuItem = objContextMenu.Controls.Add(Type:=msoControlButton)
    With objMenuItem
        .Caption = "Быстрый анализ"
        .BeginGroup = True
        .OnAction = "QuickAnalysisMacro"
    End With
    
    ' Показываем меню в позиции курсора
    objContextMenu.ShowPopup
    
    MsgBox "Контекстное меню создано!", vbInformation
    
CleanExit:
    Set objMenuItem = Nothing
    Set objContextMenu = Nothing
End Sub

'Создание меню
Sub CreateMenuInMenuBar()
    Dim objMenuBar As CommandBar
    Dim objNewMenu As CommandBarPopup
    Dim objMenuItem As CommandBarButton
    Dim strMenuName As String
    
    strMenuName = "Мои Инструменты"
    
    On Error Resume Next
    ' Удаляем старое меню если существует
    Application.CommandBars(2).Controls(strMenuName).Delete
    On Error GoTo 0
    
    ' Получаем главную строку меню
    Set objMenuBar = Application.CommandBars(2)
    
    ' Добавляем новое меню в конец строки меню
    Set objNewMenu = objMenuBar.Controls.Add( _
        Type:=msoControlPopup, _
        Temporary:=True)
    
    With objNewMenu
        .Caption = "&" & strMenuName ' & для горячей клавиши Alt+М
        
        ' Добавляем пункты меню
        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        With objMenuItem
            .Caption = "Экспорт в PDF"
            .OnAction = "ExportToPDFMacro"
            .FaceId = 109
        End With
        
        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        With objMenuItem
            .Caption = "Импорт данных"
            .OnAction = "ImportDataMacro"
            .FaceId = 23
        End With
        
        ' Добавляем разделитель
        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.BeginGroup = True
        
        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        With objMenuItem
            .Caption = "Настройки"
            .OnAction = "SettingsMacro"
            .FaceId = 487
        End With
    End With
    
    ' Обновляем строку меню
    objMenuBar.Visible = True
    
    MsgBox "Меню '" & strMenuName & "' добавлено в строку меню!", vbInformation
    
CleanExit:
    Set objMenuItem = Nothing
    Set objNewMenu = Nothing
    Set objMenuBar = Nothing
End Sub


' Удаление меню
Private Sub DeleteMenuInMenuBar()
    On Error GoTo ErrorHandler
    Dim strMenuName As String
    strMenuName = "Мои Инструменты"
    Application.CommandBars(1).Controls(strMenuName).Delete
    Exit Sub
ErrorHandler:
    Debug.Print "Oшибка:"; Err.Number & " / " & Err.Description
End Sub



