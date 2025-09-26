Attribute VB_Name = "m_05_03_Parameters"
Option Explicit

'��������� ���������� � ���������� ����������
Private Sub d_01_ApplicationInfo()
    On Error GoTo ErrorHandler
    Debug.Print "UserName: "; Application.UserName
    Debug.Print "UserInitials: "; Application.UserInitials = "�. �."
    Debug.Print "DefaultSaveFormat: "; Application.DefaultSaveFormat
    Debug.Print "PrintComments: "; Options.PrintComments
    Debug.Print "PrintHiddenText: "; Options.PrintHiddenText
    Debug.Print "AllowDragAndDrop: "; Options.AllowDragAndDrop
    Debug.Print "SaveInterval: "; Options.SaveInterval
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'��������� �������� ����������
Private Sub d_02_SetOptions()
    On Error GoTo ErrorHandler
    Application.UserName = "������ ��������"
    Application.UserInitials = "�. �."
    Application.DefaultSaveFormat = "MacroEnabledDocument"
    Application.DefaultSaveFormat = ""
    Options.SaveInterval = 5
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'����� ������, ����� Show
Private Sub d_03_FontDialogShow()
    On Error GoTo ErrorHandler
    Dim oFontDialog As Dialog
    Set oFontDialog = Dialogs(wdDialogFormatFont)
    oFontDialog.Show
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'����� ������, ����� Display
Private Sub d_04_FontDialogDisplay()
    On Error GoTo ErrorHandler
    Dim oFontDialog As Dialog
    Set oFontDialog = Dialogs(wdDialogFormatFont)
    oFontDialog.Display
    Debug.Print oFontDialog.Name
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'����� ������, ����� Display + Execute
Private Sub d_05_FontDialogDisplayExecute()
    On Error GoTo ErrorHandler
    Dim oFontDialog As Dialog
    Set oFontDialog = Dialogs(wdDialogFormatFont)
    oFontDialog.Display
    oFontDialog.Execute
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������ ������������� Display
Sub d_06_DisplayRightIndent()
On Error GoTo ErrorHandler
    Dim dlgParagraph As Dialog
    Set dlgParagraph = Dialogs(wdDialogFormatParagraph)
    dlgParagraph.Display
    Debug.Print "����� ������ " & dlgParagraph.RightIndent
    Debug.Print "������ ������ " & dlgParagraph.LeftIndent
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
