Attribute VB_Name = "m_Export"
Option Explicit

Sub ExportAllModules_Word()
    Dim objComponent As Object
    Dim strFolder As String
    Dim strFile As String
    Dim strProjectName As String
    
    ' ��� ������� (Project Name � ��������� VBAProject)
    strProjectName = Application.VBE.ActiveVBProject.Name & "Word"
    
    ' ����� ����� � ���������� Word
    If ActiveDocument.Path = "" Then
        MsgBox "������� ��������� �������� �� ����.", vbExclamation
        Exit Sub
    End If
    
    strFolder = ActiveDocument.Path & "\" & strProjectName
    
    ' ������ �����, ���� � ���
    On Error Resume Next
    MkDir strFolder
    On Error GoTo 0
    
    ' ������� ���� ����������� �������
    For Each objComponent In Application.VBE.ActiveVBProject.VBComponents
        Select Case objComponent.Type
            Case 1 ' ����������� ������
                strFile = strFolder & "\" & objComponent.Name & ".bas"
            Case 2 ' �����
                strFile = strFolder & "\" & objComponent.Name & ".cls"
            Case 3 ' �����
                strFile = strFolder & "\" & objComponent.Name & ".frm"
            Case 100 ' �������� (ThisDocument)
                strFile = strFolder & "\" & objComponent.Name & ".cls"
            Case Else
                strFile = strFolder & "\" & objComponent.Name & ".txt"
        End Select
        
        ' ������������ ������
        objComponent.Export strFile
    Next
    
    MsgBox "��� ������ ��������� � �����: " & vbCrLf & strFolder, vbInformation
End Sub

