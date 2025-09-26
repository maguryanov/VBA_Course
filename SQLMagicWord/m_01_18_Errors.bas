Attribute VB_Name = "m_01_18_Errors"
Option Explicit

'��� ��������� ������
Sub Errors_NoHandling_01()
    Dim intValue As Integer
    intValue = 0
    Debug.Print 1 / intValue '������������� ������
    Debug.Print "�������� ����� ������������� ������"
End Sub


'������������������ ���������� ��� ���������� ������ On Error GoTo
Sub Errors_GoTo_01()
    On Error GoTo ErrorHandler
    Dim intDividor As Integer
    intDividor = 100
        Debug.Print 1 / intDividor '������������� ������
    Debug.Print "�������� ����� ������������� ������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'������������������ ���������� ��� ������� ������ On Error GoTo
Sub Errors_GoTo_02()
    On Error GoTo ErrorHandler
    Dim intDividor As Integer
    intDividor = 0
    Debug.Print 1 / intDividor '������������� ������
    Debug.Print "�������� ����� ������������� ������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



'������ ��������� ������. On Error GoTo 0
Sub Errors_GoTo_0_01()
    On Error GoTo ErrorHandler
    Dim intValue As Integer
    intValue = 1
    Debug.Print "��������� ������ �����������"
    Debug.Print 1 / intValue '������������� ������
    Debug.Print "�������� ����� ������������� ������"
    On Error GoTo 0 '������ ��������� ������
    Debug.Print "��������� ������ �� �����������"
    Debug.Print 1 / (intValue - 1) '������������� ������
    Debug.Print "�������� ����� ������������� ������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'����������� ����� ������. On Error Resume Next
Sub Errors_Resume_01()
    On Error Resume Next
    Dim intValue As Integer
    intValue = 0
    Debug.Print 1 / intValue '������������� ������
    Debug.Print "�������� ����� ������������� ������"
    Exit Sub
ErrorHandler: '�� ����� � ����� ������
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'��������� ����� ������. On Error Resume Next
Sub Errors_Resume_02()
    On Error Resume Next
    Dim intValue As Integer
    intValue = 0
    Debug.Print 1 / intValue '������������� ������
    If Err.Number = 11 Then Debug.Print "������� �� ����"
    Debug.Print 1 / intValue '������������� ������
    If Err.Number = 11 Then Debug.Print "������� �� ����"
    Debug.Print "�������� ����� ������������� ������"
End Sub



'����������� ���������� �������� Finalization. ���������� ������
Sub File_Always_01()
    On Error GoTo ErrorHandler
    Dim intValue As Integer
    Open "D:\VBA\ANSI.txt" For Input As #1
    Debug.Print "������ �� �����"
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'����������� ���������� �������� Finalization. ������� ������
Sub File_Always_02()
    On Error GoTo ErrorHandler
    Dim intValue As Integer
    Open "D:\VBA\ANSI.txt" For Input As #1
    Debug.Print "������ �� �����"
    intValue = "100$"
    Debug.Print "���� �� �������� ����� ���� �����?"
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'��������� ���������� �� ������
Sub Errors_Info_01()
    On Error GoTo ErrorHandler
    Dim intValue As Integer: intValue = 0
    Debug.Print 1 / intValue '������������� ������
    Debug.Print "�������� ����� ������������� ������"
    Exit Sub
ErrorHandler:
    Debug.Print "Err:", Err
    Debug.Print "Number:", Err.Number
    Debug.Print "Description:", Err.Description
    Debug.Print "HelpContext:", Err.HelpContext
    Debug.Print "HelpFile:", Err.HelpFile
    Debug.Print "LastDllError:", Err.LastDllError
    Debug.Print "Source:", Err.Source
End Sub

'������� ���������� �� ������. Err.Clear
Sub Errors_Info_02()
    On Error GoTo ErrorHandler
    Dim intValue As Integer: intValue = 0
    Debug.Print 1 / intValue '������������� ������
    Debug.Print "�������� ����� ������������� ������"
    Exit Sub
ErrorHandler:
    Err.Clear
    Debug.Print "Err:", Err
    Debug.Print "Number:", Err.Number
    Debug.Print "Description:", Err.Description
    Debug.Print "HelpContext:", Err.HelpContext
    Debug.Print "HelpFile:", Err.HelpFile
    Debug.Print "LastDllError:", Err.LastDllError
    Debug.Print "Source:", Err.Source
End Sub



'������ ������. 11 / Division by zero
Sub Errors_Frequent_01()
    On Error GoTo ErrorHandler
    Debug.Print 1 / 0
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������ ������. 6 / Overflow
Sub Errors_Frequent_02()
    On Error GoTo ErrorHandler
    Dim intValue As Integer
    intValue = 100500
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������ ������. 9 / Subscript out of range
Sub Errors_Frequent_07()
    On Error GoTo ErrorHandler
    Dim aValues(3) As Integer
    aValues(4) = 100
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������ ������. 13 / Type mismatch
Sub Errors_Frequent_03()
    On Error GoTo ErrorHandler
    Dim intValue As Integer
    intValue = "100$"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������ ������. 55 / File already open
Sub Errors_Frequent_04()
    On Error GoTo ErrorHandler
    Open "D:\VBA\ANSI.txt" For Input As #1
    Debug.Print "������ �� �����"
    Open "D:\VBA\ANSI.txt" For Append As #1
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'������ ������. 53 / File not found
Sub Errors_Frequent_05()
    On Error GoTo ErrorHandler
    Open "D:\VBA\�������������� ����.txt" For Input As #1
    Debug.Print "������ �� �����"
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'������ ������. 54 / Bad file mode
Sub Errors_Frequent_08()
    On Error GoTo ErrorHandler
    Open "D:\VBA\Files\������.txt" For Input As #1
    Write #1, "������� ������ ������"
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'������ ������. 75 / Path/File access error
Sub Errors_Frequent_06()
    On Error GoTo ErrorHandler
    Open "D:\VBA\��� ����������.txt" For Input As #1
    Debug.Print "������ �� �����"
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub



'������������� ����������� ������. Err.Raise
Sub Errors_Raise_01()
    On Error GoTo ErrorHandler
    Dim intProductQty As Integer
    intProductQty = -100
    If intProductQty < 0 Then Err.Raise _
        50001, "Errors_Raise_06", "�������� �� ����� ���� �������������"
    Debug.Print "���������� ��������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description & " / " & Err.Source
End Sub


'������������� ����������� ������. ������� Err. ����� � �������
Sub Errors_Raise_02()
    On Error Resume Next
    Dim intProductQty As Integer
    intProductQty = -100
    Debug.Print 1 / 0
    If intProductQty < 0 Then
        'Err.Clear
        Err.Raise 50001
    End If
    Debug.Print Err.Number & " / " & Err.Description & " / " & Err.Source
End Sub


'0�512 �������������� ��� ��������� ������
Sub Errors_Raise_03()
    On Error GoTo ErrorHandler
    Dim intProductQty As Integer
    intProductQty = -100
    If intProductQty < 0 Then Err.Raise _
        6, "Errors_Raise_06", "�������� �� ����� ���� �������������"
    Debug.Print "���������� ��������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description & " / " & Err.Source
End Sub

'�������� Error �� �������������. ��� �������� �������������
Sub Errors_Raise_04()
    On Error GoTo ErrorHandler
    Dim intProductQty As Integer
    intProductQty = -100
    If intProductQty < 0 Then Error 50001
    If intProductQty < 0 Then Error 6
    Debug.Print "���������� ��������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description & " / " & Err.Source
End Sub

'������ ��������� � ������. ��� ����������
Sub Errors_Raise_05()
    On Error GoTo ErrorHandler
    Dim varResult As Variant
    Dim oVideoLesson As New VideoLesson
    'varResult = oVideoLesson.ErrorHandlingDemo(0)
    varResult = oVideoLesson.ErrorHandlingDemo("100$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'������ ��������� � ������. ��� ���������� � �������������
Sub Errors_Raise_06()
    On Error GoTo ErrorHandler
    Dim varResult As Variant
    Dim oVideoLesson As New VideoLesson
    'varResult = oVideoLesson.ErrorHandlingDemo(0)
    varResult = oVideoLesson.ErrorHandlingRaiseDemo("100$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'��������� ������ ��� �������. ����� Error Trapping. Break on all
Sub Errors_Testing_01()
    On Error GoTo ErrorHandler
    Dim dblDivider As Double
    Debug.Print 1 / dblDivider
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'��������� ������ ��� ������������. ����� Error Trapping. Break In �lass Module
Sub Errors_Testing_02_1()
    On Error GoTo ErrorHandler
    Dim varResult As Variant
    Dim oVideoLesson As New VideoLesson
    'varResult = oVideoLesson.ErrorHandlingDemo(0)
    varResult = oVideoLesson.ErrorHandlingRaiseDemo("100$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'��������� ������ ��� ������������. ����� Error Trapping. Break In class Module
Sub Errors_Testing_02_2()
    On Error GoTo ErrorHandler
        Err.Raise 50000, , "��� ����������� Break In class Module"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'��������� ������ ��� ������������. ����� Error Trapping. Break on Unhandled
Sub Errors_Testing_03_01()
    On Error GoTo 0
    Debug.Print 1 / 0
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������������ � ����������� �������� ���� Integer ' % - Integer
Sub d_20_Integer()
On Error GoTo ErrorHandler
    Dim intVar As Integer
    Do
        intVar = intVar + 1
    Loop
    Debug.Print intVar
    Exit Sub
ErrorHandler:
    Debug.Print intVar
    Debug.Print Err.Number & " / " & Err.Description
End Sub

