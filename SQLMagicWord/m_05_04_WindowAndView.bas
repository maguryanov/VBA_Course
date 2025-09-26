Attribute VB_Name = "m_05_04_WindowAndView"
Option Explicit

'������� ���� ��� ���������
Sub d_01_WindowsList()
On Error GoTo ErrorHandler
    Dim oDoc As Document, oItem As Window
    Set oDoc = Documents.Open("D:\VBA\��� ������ Word\������������.docx")
    For Each oItem In oDoc.Windows
        Debug.Print oItem.Caption
    Next oItem
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'������� ���� ����
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


'��������� � ����
Sub d_03_GetWindow()
On Error GoTo ErrorHandler
    Dim oWindow As Window
    Set oWindow = Windows("������������.docx:2")
    Debug.Print oWindow.Document.fullName
    Debug.Print oWindow.Caption
    Set oWindow = Documents("������������.docx").Windows(3)
    Debug.Print oWindow.Document.fullName
    Debug.Print oWindow.Caption
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'��������� ��������� � �������� ����
Sub d_04_ChangeWindow()
On Error GoTo ErrorHandler
    Dim oWindow As Window
    Set oWindow = Documents.Open("D:\VBA\��� ������ Word\������������.docx").Windows(3)
    oWindow.Caption = "������ ��������� ��������� ����"
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

'���������/���������� �������
Sub d_05_ChangeWindow()
On Error GoTo ErrorHandler
    Dim oWindow As Window
    Set oWindow = Documents.Open("D:\VBA\��� ������ Word\������������.docx").Windows(3)
    oWindow.DisplayRulers = Not oWindow.DisplayRulers
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'������������ ������ ���������
Sub d_06_ChangeView()
On Error GoTo ErrorHandler
    Dim oWindow As Window, byteSwitch As Byte
    Set oWindow = Documents.Open("D:\VBA\��� ������ Word\������������.docx").Windows(3)
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

'���������/���������� ������� ������
Sub d_07_ChangeView()
On Error GoTo ErrorHandler
    Dim oWindow As Window, byteSwitch As Byte
    Set oWindow = Documents.Open("D:\VBA\��� ������ Word\������������.docx").Windows(3)
    oWindow.View.FullScreen = Not oWindow.View.FullScreen
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'���������/���������� ShowAll
Sub d_09_ShowAll()
On Error GoTo ErrorHandler
    Dim oWindow As Window, byteSwitch As Byte
    Set oWindow = Documents.Open("D:\VBA\��� ������ Word\������������.docx").Windows(3)
    oWindow.View.ShowAll = Not oWindow.View.ShowAll
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'���������/���������� ShowTabs
Sub d_10_ShowTabs()
On Error GoTo ErrorHandler
    Dim oWindow As Window, byteSwitch As Byte
    Set oWindow = Documents.Open("D:\VBA\��� ������ Word\������������.docx").Windows(3)
    oWindow.View.ShowTabs = Not oWindow.View.ShowTabs
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'���������/���������� ShowComments
Sub d_11_ShowComments()
On Error GoTo ErrorHandler
    Dim oWindow As Window, byteSwitch As Byte
    Set oWindow = Documents.Open("D:\VBA\��� ������ Word\������������.docx").Windows(3)
    oWindow.View.ShowComments = Not oWindow.View.ShowComments
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'���������/���������� ShowBookmarks
Sub d_12_ShowBookmarks()
On Error GoTo ErrorHandler
    Dim oWindow As Window, byteSwitch As Byte
    Set oWindow = Documents.Open("D:\VBA\��� ������ Word\������������.docx").Windows(3)
    oWindow.View.ShowBookmarks = Not oWindow.View.ShowBookmarks
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'���������/���������� TableGridlines
Sub d_13_TableGridlines()
On Error GoTo ErrorHandler
    Dim oWindow As Window, byteSwitch As Byte
    Set oWindow = Documents.Open("D:\VBA\��� ������ Word\������������.docx").Windows(3)
    oWindow.View.TableGridlines = Not oWindow.View.TableGridlines
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'��������� �������� ��-��������� ��� ���� ����
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


