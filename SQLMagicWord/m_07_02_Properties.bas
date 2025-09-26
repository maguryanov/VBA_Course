Attribute VB_Name = "m_07_02_Properties"
Option Explicit

'������ ���������� �������
Private Sub d_01_BuildinPropoperties()
    On Error Resume Next
    Dim objPropertyItem As DocumentProperty
    For Each objPropertyItem In ThisDocument.BuiltInDocumentProperties
        Debug.Print objPropertyItem.Name; " = ";
        Debug.Print objPropertyItem.Value
        If Err.Number <> 0 Then
            Debug.Print "O�����"
            Err.Clear
        End If
    Next objPropertyItem
    Exit Sub
ErrorHandler:
    Debug.Print "O�����:"; Err.Number & " / " & Err.Description
End Sub

'�������� ���������������� �������
Private Sub d_02_CustomPropoperties()
    On Error Resume Next
    Dim objPropertyItem As DocumentProperty
    For Each objPropertyItem In ThisDocument.CustomDocumentProperties
        Debug.Print objPropertyItem.Name; " = ";
        Debug.Print objPropertyItem.Value; "/";
        If Err.Number <> 0 Then
            Debug.Print "O�����"
            Err.Clear
        End If
        Debug.Print objPropertyItem.Type
        If Err.Number <> 0 Then
            Debug.Print "O�����"
            Err.Clear
        End If
    Next objPropertyItem
    Exit Sub
End Sub

'���������� ���������������� �������
Private Sub d_03_AddCustomPropertiesToWordDoc()
    Dim objDocument As Document
    Dim strPropName As String
    
    On Error GoTo ErrorHandler
    
    Set objDocument = ActiveDocument
    
    With objDocument.CustomDocumentProperties
        ' ���������� � ���������
        .Add Name:="������ ���������", _
             LinkToContent:=False, _
             Type:=msoPropertyTypeString, _
             Value:="���������"
        
        .Add Name:="���������� ������� �� ������ ������� �������", _
             LinkToContent:=False, _
             Type:=msoPropertyTypeNumber, _
             Value:=objDocument.Content.ComputeStatistics(wdStatisticPages)
        
        .Add Name:="���� ������� �������", _
             LinkToContent:=False, _
             Type:=msoPropertyTypeDate, _
             Value:=Now
        
        ' ��������� ����������
        .Add Name:="���������� �����", _
             LinkToContent:=False, _
             Type:=msoPropertyTypeString, _
             Value:="DOC-" & Format(Now, "yymmddhhmm")
    End With
    
    
CleanExit:
    Set objDocument = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "������: " & Err.Description, vbCritical
    Resume CleanExit
End Sub


'��������� �� ��������
Private Sub d_04_GetValueOfProperty()
    On Error GoTo ErrorHandler
    Debug.Print ThisDocument.BuiltInDocumentProperties.Item(1).Value
    Debug.Print ThisDocument.CustomDocumentProperties.Item(1).Value
    Exit Sub
ErrorHandler:
    MsgBox "������: " & Err.Description, vbCritical
End Sub


'�������� ���������������� �������
Private Sub d_06_DeleteCustomPropoperties()
    On Error Resume Next
    Dim objPropertyItem As DocumentProperty
    For Each objPropertyItem In ThisDocument.CustomDocumentProperties
        objPropertyItem.Delete
    Next objPropertyItem
    Exit Sub
ErrorHandler:
    MsgBox "������: " & Err.Description, vbCritical
End Sub

