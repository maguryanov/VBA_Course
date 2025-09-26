Attribute VB_Name = "m_05_10_DocVariables"
Option Explicit

'������ ����� ����� "Visual Basic" ��� �������������
Private Sub d_01_CopyFileForExperiment()
    On Error GoTo ErrorHandler
    Dim strSourceFilename As String, strTargetFilename As String
    strSourceFilename = "D:\VBA\��� ������ Word\��������� �����.docx"
    strTargetFilename = "D:\VBA\Word\���� VBA.docx"
    Call CopyFile(SourceFilename:=strSourceFilename, TargetFilename:=strTargetFilename)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub d_02_WhatToStore()
'��� ����� ������� � ���������� ���������?
'���������� ���������
'DocAuthor � ����� ���������
'DocEditor � ��������� ��������
'DocVersion � ������ ���������
'DocStatus � ������ ��������� (Draft, Approved, Final)
'DocCreatedDate � ���� ��������
'DocApprovedBy � ��� ��������

'������ ������� / ��������
'ProjectCode � ��� �������
'ProjectName � �������� �������
'ClientName � ��� ���������
'ContractNumber � ����� ��������
'ContractDate � ���� ��������

'��������� �������� / �������
'Style_FirstLineIndent � ������ ������ ������
'Style_SpaceAfter � �������� ����� ������
'PrintComments � �������� �� ����������� (True/False)
'Lang_Main � �������� ���� ������ (��������, en-US, ru-RU)
'ThemeMode � ���� ���������� (Light/Dark)

'��������� � �����������
'DocGUID � ���������� ������������� ���������
'LastCursorPos � ������� ������� ��� ��������� ��������
'UpdateTOC � ��������� �� ���������� (True/False)
'ProcessedByMacro � ��������� �� �������� ��������
'ExportPath � ���� ��� �������� ���������
End Sub

'�������� ���������� ���������
Private Sub d_03_AddVariables()
    On Error GoTo ErrorHandler
    Dim docVBA As Document
    Set docVBA = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    With docVBA.Variables
        .Item("DocAuthorLastName").Value = "��������"
        .Item("DocAuthorPatronymic").Value = "������"
        .Item("DocAuthorFirstName").Value = "����������"
        .Item("DocVersion").Value = "0.1"
        .Item("DocStatus").Value = "��������"
        .Item("ProjectCode").Value = "VBACourse"
        .Item("PrintComments").Value = "False"
    End With
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'����������� ��� ���������� ���������
Private Sub d_04_ShowVariables()
    On Error GoTo ErrorHandler
    Dim varItem As Variant '��� ������������� for each
    '���������� ��������� ������ �������� ������
    Dim docVBA As Document
    Set docVBA = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    For Each varItem In docVBA.Variables
        Debug.Print varItem.Name & " = " & varItem.Value
    Next varItem
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'��������� � ���������� ���������
Private Sub d_05_GetVariableValue()
    On Error GoTo ErrorHandler
    Dim strItem As String
    Dim docVBA As Document
    Set docVBA = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    Debug.Print "DocStatus="; docVBA.Variables("DocStatus")
    Debug.Print "ProjectCode="; docVBA.Variables("ProjectCode")
    Debug.Print "DocVersion="; docVBA.Variables("DocVersion")
    docVBA.Variables("DocVersion") = "0.2"
    Debug.Print "DocVersion="; docVBA.Variables("DocVersion")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



