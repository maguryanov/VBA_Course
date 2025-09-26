Attribute VB_Name = "m_03_01_Forms"
Option Explicit

Private Sub FormShow_01()
On Error GoTo ErrorHandler
    FormFirst.Show
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub FormUnload_01()
On Error GoTo ErrorHandler
    Unload FormFirst
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Свойства форм
'Форматирование форм

Private Sub FormLoad_01()
On Error GoTo ErrorHandler
    Load FormFirst
    FormFirst.Caption = "Форма загружена но не отображается"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub Hide_01()
On Error GoTo ErrorHandler
    FormFirst.Hide
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

