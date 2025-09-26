Attribute VB_Name = "m_Export"
Option Explicit

Sub ExportAllModules_Word()
    Dim objComponent As Object
    Dim strFolder As String
    Dim strFile As String
    Dim strProjectName As String
    
    ' Имя проекта (Project Name в свойствах VBAProject)
    strProjectName = Application.VBE.ActiveVBProject.Name & "Word"
    
    ' Папка рядом с документом Word
    If ActiveDocument.Path = "" Then
        MsgBox "Сначала сохраните документ на диск.", vbExclamation
        Exit Sub
    End If
    
    strFolder = ActiveDocument.Path & "\" & strProjectName
    
    ' Создаём папку, если её нет
    On Error Resume Next
    MkDir strFolder
    On Error GoTo 0
    
    ' Перебор всех компонентов проекта
    For Each objComponent In Application.VBE.ActiveVBProject.VBComponents
        Select Case objComponent.Type
            Case 1 ' Стандартный модуль
                strFile = strFolder & "\" & objComponent.Name & ".bas"
            Case 2 ' Класс
                strFile = strFolder & "\" & objComponent.Name & ".cls"
            Case 3 ' Форма
                strFile = strFolder & "\" & objComponent.Name & ".frm"
            Case 100 ' Документ (ThisDocument)
                strFile = strFolder & "\" & objComponent.Name & ".cls"
            Case Else
                strFile = strFolder & "\" & objComponent.Name & ".txt"
        End Select
        
        ' Экспортируем модуль
        objComponent.Export strFile
    Next
    
    MsgBox "Все модули выгружены в папку: " & vbCrLf & strFolder, vbInformation
End Sub

