Attribute VB_Name = "m_02_05_ConditionalFunctions"
Option Explicit

Function IIF_01(Balance As Currency) As String
On Error GoTo ErrorHandler
    IIF_01 = IIf(Balance > 0, "�������������", "������� ��� �������������")
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function


Function IIF_02(Balance As Currency) As String
On Error GoTo ErrorHandler
    IIF_02 = IIf(Balance > 0, "�������������", _
                IIf(Balance = 0, "�������", "�������������"))
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function

Function Switch_01(Balance As Currency) As String
On Error GoTo ErrorHandler
    Switch_01 = Switch(Balance > 0, "�������������", _
        Balance < 0, "�������������", Balance = 0, "�������")
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function

Function Switch_02(Balance As Currency) As String
On Error GoTo ErrorHandler
    Switch_02 = Switch(Balance > 0, "�������������")
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function

Function Switch_03(Salary As Currency) As String
On Error GoTo ErrorHandler
    Switch_03 = Switch(Salary < 100000, "������", Salary < 200000, "�������", _
                                True, "�������")
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function


Function Switch_04(ColorNumber As Integer) As String
On Error GoTo ErrorHandler
    Switch_04 = Switch(ColorNumber = 1, "�������", ColorNumber = 2, "Ƹ����", _
            ColorNumber = 3, "������", True, "�������������")
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function


Function Choose_01(ColorNumber As Integer) As String
On Error GoTo ErrorHandler
    Choose_01 = Choose(ColorNumber, "�������", "Ƹ����", "������")
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function

Function Choose_02(ColorNumber As Integer) As Variant
On Error GoTo ErrorHandler
    Choose_02 = Choose(ColorNumber, "�������", "Ƹ����", "������")
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function

Function Choose_03(ColorNumber As Integer) As String
On Error GoTo ErrorHandler
    Dim varColor As Variant
    varColor = Choose(ColorNumber, "�������", "Ƹ����", "������")
    Choose_03 = IIf(IsNull(varColor), "�������������", varColor)
    '� Access ���� ������� NZ
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function

Sub DataTypes_04()
On Error GoTo ErrorHandler
    Debug.Print TypeName(IIf(True, 100, "������"))
    Debug.Print TypeName(Switch(False, 100, True, "������"))
    Debug.Print TypeName(Choose(3, "�������", 100, Now))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

