Attribute VB_Name = "m_07_01_UI"
Option Explicit

'��������� ������ �� ���������� Outlook
Private Sub d_01_CommandBar()
    
    Dim objDocsCommandBar As CommandBar
    Set objDocsCommandBar = CommandBars.Add("Documents", msoBarFloating)
    objDocsCommandBar.Visible = True
    
    Exit Sub
ErrorHandler:
    Debug.Print "O�����:"; Err.Number & " / " & Err.Description
End Sub

'�������� ������ ������������
Private Sub d_02_CreateAdvancedToolbar()
    Dim objToolbar As CommandBar
    Dim objControl As Object
    Dim strBarName As String
    
    strBarName = "�����������������"
    
    ' ������� ������ ������
    On Error Resume Next
    Application.CommandBars(strBarName).Delete
    On Error GoTo 0
    
    ' ������� ������
    Set objToolbar = Application.CommandBars.Add( _
        Name:=strBarName, _
        Position:=msoBarFloating, _
        MenuBar:=False, _
        Temporary:=True)
    
    ' ��������� ������ ���� ���������
    
    ' 1. ������
    Set objControl = objToolbar.Controls.Add(Type:=msoControlButton)
    With objControl
        .Caption = "���������"
        .FaceId = 3
        .OnAction = "SaveDocument"
        .TooltipText = "��������� �������� (Ctrl+S)"
    End With
    
    ' 2. �����������
    Set objControl = objToolbar.Controls.Add(Type:=msoControlButton)
    objControl.BeginGroup = True
    
    ' 3. ���������� ������
    Set objControl = objToolbar.Controls.Add(Type:=msoControlDropdown)
    With objControl
        .Caption = "����� ������"
        .AddItem "�������", 1
        .AddItem "��������� 1", 2
        .AddItem "��������� 2", 3
        .OnAction = "ApplyStyleMacro"
    End With
    
    ' 4. ��� ������
    Set objControl = objToolbar.Controls.Add(Type:=msoControlButton)
    With objControl
        .Caption = "������"
        .FaceId = 4
        .OnAction = "PrintDocument"
        .TooltipText = "������ ��������� (Ctrl+P)"
    End With
    ' ����������� ������
    With objToolbar
        .Visible = True
        .Top = 100
        .Left = 200
        .Width = 200
        .Protection = msoBarNoMove ' ��������� �����������
    End With
        
CleanExit:
    Set objControl = Nothing
    Set objToolbar = Nothing
End Sub

'�������� ������ ������������
Private Sub d_03_DeleteAdvancedToolbar()
    Dim strBarName As String
    
    strBarName = "�����������������"
    
    ' ������� ������ ������
    On Error Resume Next
    Application.CommandBars(strBarName).Delete
    On Error GoTo 0
    
End Sub

'�������� ������������ ����
Sub d_04_CreateContextMenu()
    Dim objContextMenu As CommandBar
    Dim objMenuItem As CommandBarButton
    Dim strMenuName As String
    
    strMenuName = "������������������"
    
    ' ������� ������������ ����
    On Error Resume Next
    Application.CommandBars(strMenuName).Delete
    On Error GoTo 0
    
    ' ������� ����������� ����
    Set objContextMenu = Application.CommandBars.Add( _
        Name:=strMenuName, _
        Position:=msoBarPopup, _
        MenuBar:=False, _
        Temporary:=True)
    
    ' ��������� ������ ����
    Set objMenuItem = objContextMenu.Controls.Add(Type:=msoControlButton)
    With objMenuItem
        .Caption = "������������� ������"
        .BeginGroup = True
        .OnAction = "FormatCellsMacro"
    End With
    
    Set objMenuItem = objContextMenu.Controls.Add(Type:=msoControlButton)
    With objMenuItem
        .Caption = "�������� ����������"
        .OnAction = "SpecialPasteMacro"
    End With
    
    Set objMenuItem = objContextMenu.Controls.Add(Type:=msoControlButton)
    With objMenuItem
        .Caption = "������� ������"
        .BeginGroup = True
        .OnAction = "QuickAnalysisMacro"
    End With
    
    ' ���������� ���� � ������� �������
    objContextMenu.ShowPopup
    
    MsgBox "����������� ���� �������!", vbInformation
    
CleanExit:
    Set objMenuItem = Nothing
    Set objContextMenu = Nothing
End Sub

'�������� ����
Sub CreateMenuInMenuBar()
    Dim objMenuBar As CommandBar
    Dim objNewMenu As CommandBarPopup
    Dim objMenuItem As CommandBarButton
    Dim strMenuName As String
    
    strMenuName = "��� �����������"
    
    On Error Resume Next
    ' ������� ������ ���� ���� ����������
    Application.CommandBars(2).Controls(strMenuName).Delete
    On Error GoTo 0
    
    ' �������� ������� ������ ����
    Set objMenuBar = Application.CommandBars(2)
    
    ' ��������� ����� ���� � ����� ������ ����
    Set objNewMenu = objMenuBar.Controls.Add( _
        Type:=msoControlPopup, _
        Temporary:=True)
    
    With objNewMenu
        .Caption = "&" & strMenuName ' & ��� ������� ������� Alt+�
        
        ' ��������� ������ ����
        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        With objMenuItem
            .Caption = "������� � PDF"
            .OnAction = "ExportToPDFMacro"
            .FaceId = 109
        End With
        
        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        With objMenuItem
            .Caption = "������ ������"
            .OnAction = "ImportDataMacro"
            .FaceId = 23
        End With
        
        ' ��������� �����������
        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.BeginGroup = True
        
        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        With objMenuItem
            .Caption = "���������"
            .OnAction = "SettingsMacro"
            .FaceId = 487
        End With
    End With
    
    ' ��������� ������ ����
    objMenuBar.Visible = True
    
    MsgBox "���� '" & strMenuName & "' ��������� � ������ ����!", vbInformation
    
CleanExit:
    Set objMenuItem = Nothing
    Set objNewMenu = Nothing
    Set objMenuBar = Nothing
End Sub


' �������� ����
Private Sub DeleteMenuInMenuBar()
    On Error GoTo ErrorHandler
    Dim strMenuName As String
    strMenuName = "��� �����������"
    Application.CommandBars(1).Controls(strMenuName).Delete
    Exit Sub
ErrorHandler:
    Debug.Print "O�����:"; Err.Number & " / " & Err.Description
End Sub



