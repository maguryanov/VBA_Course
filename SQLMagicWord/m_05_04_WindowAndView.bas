Attribute VB_Name = "m_05_04_WindowAndView"
Option Explicit

'Перебор окон для документа
Sub d_01_WindowsList()
On Error GoTo ErrorHandler
    Dim oDoc As Document, oItem As Window
    Set oDoc = Documents.Open("D:\VBA\Для чтения Word\Демонстрации.docx")
    For Each oItem In oDoc.Windows
        Debug.Print oItem.Caption
    Next oItem
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Перебор всех окон
Sub d_02_AllWindowsList()
On Error GoTo ErrorHandler
    Dim oItem As Window
    For Each oItem In Application.Windows
        Debug.Print oItem.Caption
    Next oItem
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Обращение к окну
Sub d_03_GetWindow()
On Error GoTo ErrorHandler
    Dim oWindow As Window
    Set oWindow = Windows("Демонстрации.docx:2")
    Debug.Print oWindow.Document.fullName
    Debug.Print oWindow.Caption
    Set oWindow = Documents("Демонстрации.docx").Windows(3)
    Debug.Print oWindow.Document.fullName
    Debug.Print oWindow.Caption
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Изменение заголовка и размеров окна
Sub d_04_ChangeWindow()
On Error GoTo ErrorHandler
    Dim oWindow As Window
    Set oWindow = Documents.Open("D:\VBA\Для чтения Word\Демонстрации.docx").Windows(3)
    oWindow.Caption = "Пример изменения заголовка окна"
    oWindow.WindowState = wdWindowStateNormal
    oWindow.Left = 0
    oWindow.Top = 0
    oWindow.Width = 700
    oWindow.Height = 800
    oWindow.DisplayRulers = False
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Включение/выключение линейки
Sub d_05_ChangeWindow()
On Error GoTo ErrorHandler
    Dim oWindow As Window
    Set oWindow = Documents.Open("D:\VBA\Для чтения Word\Демонстрации.docx").Windows(3)
    oWindow.DisplayRulers = Not oWindow.DisplayRulers
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Переключение режима просмотра
Sub d_06_ChangeView()
On Error GoTo ErrorHandler
    Dim oWindow As Window, byteSwitch As Byte
    Set oWindow = Documents.Open("D:\VBA\Для чтения Word\Демонстрации.docx").Windows(3)
    byteSwitch = 0
    Select Case byteSwitch
        Case 0: oWindow.View.Type = wdNormalView
        Case 1: oWindow.View.Type = wdWebView
        Case 2: oWindow.View.Type = wdPrintPreview
        Case 3: oWindow.View.Type = wdPrintView
        Case 4: oWindow.View.Type = wdReadingView
    End Select
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Включение/выключение полного экрана
Sub d_07_ChangeView()
On Error GoTo ErrorHandler
    Dim oWindow As Window, byteSwitch As Byte
    Set oWindow = Documents.Open("D:\VBA\Для чтения Word\Демонстрации.docx").Windows(3)
    oWindow.View.FullScreen = Not oWindow.View.FullScreen
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Включение/выключение ShowAll
Sub d_09_ShowAll()
On Error GoTo ErrorHandler
    Dim oWindow As Window, byteSwitch As Byte
    Set oWindow = Documents.Open("D:\VBA\Для чтения Word\Демонстрации.docx").Windows(3)
    oWindow.View.ShowAll = Not oWindow.View.ShowAll
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Включение/выключение ShowTabs
Sub d_10_ShowTabs()
On Error GoTo ErrorHandler
    Dim oWindow As Window, byteSwitch As Byte
    Set oWindow = Documents.Open("D:\VBA\Для чтения Word\Демонстрации.docx").Windows(3)
    oWindow.View.ShowTabs = Not oWindow.View.ShowTabs
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Включение/выключение ShowComments
Sub d_11_ShowComments()
On Error GoTo ErrorHandler
    Dim oWindow As Window, byteSwitch As Byte
    Set oWindow = Documents.Open("D:\VBA\Для чтения Word\Демонстрации.docx").Windows(3)
    oWindow.View.ShowComments = Not oWindow.View.ShowComments
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Включение/выключение ShowBookmarks
Sub d_12_ShowBookmarks()
On Error GoTo ErrorHandler
    Dim oWindow As Window, byteSwitch As Byte
    Set oWindow = Documents.Open("D:\VBA\Для чтения Word\Демонстрации.docx").Windows(3)
    oWindow.View.ShowBookmarks = Not oWindow.View.ShowBookmarks
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Включение/выключение TableGridlines
Sub d_13_TableGridlines()
On Error GoTo ErrorHandler
    Dim oWindow As Window, byteSwitch As Byte
    Set oWindow = Documents.Open("D:\VBA\Для чтения Word\Демонстрации.docx").Windows(3)
    oWindow.View.TableGridlines = Not oWindow.View.TableGridlines
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Установка значений по-умолчанию для всех окон
Sub d_14_SetDafault()
On Error GoTo ErrorHandler
    Dim oWindow As Window, oDocument As Document
    For Each oDocument In Documents
        For Each oWindow In oDocument.Windows
            oWindow.DisplayRulers = True
            oWindow.View.Type = wdPrintView
            oWindow.View.ShowAll = False
            oWindow.View.ShowTabs = False
            oWindow.View.ShowComments = False
            oWindow.View.ShowBookmarks = False
            oWindow.View.TableGridlines = True
        Next oWindow
    Next oDocument
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


