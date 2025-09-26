Attribute VB_Name = "m_05_02_Document"
Option Explicit

'���������� ���������
Private Sub d_01_AddDocument()
    On Error GoTo ErrorHandler
    Dim oDoc As Document
    Set oDoc = Documents.Add()
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'���������� ���������
Private Sub d_02_SaveDocument()
    On Error GoTo ErrorHandler
    Dim oDoc As Document
    Set oDoc = Documents.Add()
    oDoc.SaveAs2 ("D:\VBA\Word\������ ���������� ��������.docx")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'�������� ���������
Private Sub d_03_OpenDocument()
    On Error GoTo ErrorHandler
    Dim oDoc As Document
    Set oDoc = Documents.Open("D:\VBA\Word\������ ���������� ��������.docx")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'��������� ������ �� ��������
Private Sub d_04_OpenDocument()
    On Error GoTo ErrorHandler
    Dim oDoc As Document
    Set oDoc = Documents(1)
    Debug.Print oDoc.Name
    Set oDoc = Documents("D:\VBA\Word\������ ���������� ��������.docx")
    Debug.Print oDoc.Name
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'�������� ���������
Private Sub d_05_CloseDocument()
    On Error GoTo ErrorHandler
    Dim oDoc As Document
    Set oDoc = Documents("D:\VBA\Word\������ ���������� ��������.docx")
    oDoc.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������� �������� ����������
Private Sub d_06_ListOfDocuments()
    On Error GoTo ErrorHandler
    Dim oDoc As Document, oItem As Document
    Set oDoc = Documents.Open("D:\VBA\Word\������ ���������� ��������.docx")
    For Each oItem In Documents
        Debug.Print oItem.fullName
    Next oItem
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'��������� ���������
Private Sub d_07_ActivateDocument()
    On Error GoTo ErrorHandler
    Dim oDoc As Document
    Set oDoc = Documents("D:\VBA\Word\������ ���������� ��������.docx")
    oDoc.Activate
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'��������� ���������� � ���������
Private Sub d_08_DocumentInfo()
    On Error GoTo ErrorHandler
    Dim oDoc As Document
    Set oDoc = Documents.Open("D:\VBA\Word\������ ���������� ��������.docx")
    Debug.Print oDoc.Name
    Debug.Print oDoc.fullName
    Debug.Print oDoc.Type 'wdTypeDocument  0   ��������
    Debug.Print oDoc.HasVBProject
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'�������� �� �������� ��������
Public Function IsDocumentOpen(strFileName As String) As Boolean
    Dim objDoc As Document
    
    IsDocumentOpen = False
    For Each objDoc In Application.Documents
        If LCase(objDoc.fullName) = LCase(strFileName) Then
            IsDocumentOpen = True
            Exit Function
        End If
    Next objDoc
End Function


'������ ����� �����
Public Sub CopyFile(ByVal SourceFilename As String, ByVal TargetFilename As String)
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = wdAlertsNone   ' ��������� ��������������
    If IsDocumentOpen(TargetFilename) Then Documents(TargetFilename).Close
    Documents.Add(SourceFilename).SaveAs2 TargetFilename
Finalization:
    Application.DisplayAlerts = wdAlertsAll   ' ���������� �������
    Exit Sub
ErrorHandler:
    Application.DisplayAlerts = wdAlertsAll   ' ���������� �������
    Debug.Print Err.Number & " / " & Err.Description
    Err.Raise Err.Number
    GoTo Finalization
End Sub
