Attribute VB_Name = "m_06_01_OutlookObjects"
Option Explicit
Private mobjOutlook As Outlook.Application


'��������� ������ �� ���������� Outlook
Private Sub d_01_GetObject()
    On Error GoTo ErrorHandler
    Dim objOutlook As Object
    Set objOutlook = GetObject(Class:="Outlook.Application")
    Debug.Print objOutlook.Version
    Exit Sub
ErrorHandler:
    Debug.Print "O�����:"; Err.Number & " / " & Err.Description
End Sub

'������ Outlook � �������������� � ���. ��� ����������. ����� ��� ������ PowerShell
Private Sub d_02_EarlyBinding()
    On Error GoTo ErrorHandler
    Dim objOutlook As Outlook.Application
    Dim objExplorer As Object
    Set objOutlook = New Outlook.Application
    Debug.Print objOutlook.Version
    Exit Sub
ErrorHandler:
    Debug.Print "O�����:"; Err.Number & " / " & Err.Description
End Sub

'������ Outlook � ����������� ����������
Sub d_03_ShowOutlook()
    Dim objOutlook As Object
    Dim objExplorer As Object
    
    ' �������� �������� ���������� Outlook
    On Error Resume Next
    Set objOutlook = GetObject(, "Outlook.Application")
    On Error GoTo 0
    
    ' ���� �� ����� � ��������� ����� ���������
    If objOutlook Is Nothing Then
        Set objOutlook = CreateObject("Outlook.Application")
    End If
    
    ' ���� ������� ���� (Explorer)
    Set objExplorer = objOutlook.ActiveExplorer
    If objExplorer Is Nothing Then
        Set objExplorer = objOutlook.Explorers.Add(objOutlook.Session.GetDefaultFolder(6), 0) ' 6 = olFolderInbox
    End If
    
    ' ������ ���� ������� � ��������
    objExplorer.Display
    objExplorer.Activate
    Debug.Print objOutlook.Version
End Sub



'��������� ������ �� ���������� Outlook � ���������� ������
Private Sub d_04_GetIntoModuleVariable()
    On Error GoTo ErrorHandler
    Set mobjOutlook = GetObject(Class:="Outlook.Application")
    Exit Sub
ErrorHandler:
    Debug.Print "O�����:"; Err.Number & " / " & Err.Description
End Sub


Private Sub d_05_WorkWithModuleVariable()
    On Error GoTo ErrorHandler
    Debug.Print "������: "; mobjOutlook.Version
    Exit Sub
ErrorHandler:
    Debug.Print "O�����:"; Err.Number & " / " & Err.Description
End Sub


