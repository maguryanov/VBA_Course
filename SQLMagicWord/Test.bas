Attribute VB_Name = "Test"
Option Explicit

Private Sub DisplayArithmeticProgression()
    On Error GoTo ErrorHandler
    ' ���������� ���������� � ���������� ��������
    Dim dblFirstElement As Double  ' ������ ������� ����������
    Dim dblStep As Double          ' ��� ����������
    Dim strResult As String        ' ������ ��� ������������ ����������
    Dim intCounter As Integer      ' ������� �����
    
    ' ������ ������ � ������������
    dblFirstElement = InputBox("������� ������ ������� ����������:", "�������������� ����������")
    dblStep = InputBox("������� ��� ����������:", "�������������� ����������")
    
    ' �������� �� ���������� ����
    If Not IsNumeric(dblFirstElement) Or Not IsNumeric(dblStep) Then
        MsgBox "������: ���������� ������ �������� ��������!", vbCritical
        Exit Sub
    End If
    
    ' �������������� ��������� ����� � �����
    dblFirstElement = CDbl(dblFirstElement)
    dblStep = CDbl(dblStep)
    
    ' ������������ ����������
    strResult = "������ 3 ����� �������������� ����������:" & vbCrLf & vbCrLf
    
    For intCounter = 0 To 2
        strResult = strResult & "���� " & (intCounter + 1) & ": " & _
                   Format(dblFirstElement + intCounter * dblStep, "0.#####") & vbCrLf
    Next intCounter
    
    ' ����� ����������
    MsgBox strResult, vbInformation, "���������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

