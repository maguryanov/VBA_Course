Attribute VB_Name = "m_01_12_Functions"
Option Explicit

Function Parameters(CourseName As String, _
                 Category As String, _
                  trainer As String, _
                   online As Boolean)
    Debug.Print "�������� �����: " & CourseName
    Debug.Print "���������     : " & Category
    Debug.Print "������        : " & trainer
    Debug.Print "������        : " & online
End Function

Sub TestParameters()
    Debug.Print Parameters("���������������� �� VBA", _
        "Microsoft Office", "������ ��������", True)
End Sub

Function CallDemonstration()
    Debug.Print "� ��� ������������ ������������� CALL"
End Function

Sub TestCallDemonstration()
    Debug.Print CallDemonstration()
End Sub

Function Parameters_01(par1 As String, par2 As String)
    Debug.Print par1 & " " & par2
End Function

Sub TestParameters_01()
    Debug.Print Parameters_01("�����", "��� CALL")
End Sub

Function Parameters_ByRef(par As String)
    Debug.Print "�������� ��� �������� � ���������: " & par
    par = "��������� � ��������� ��������"
End Function

Sub TestParameters_ByRef()
    Dim strVar As String
    strVar = "�������� �� ������ ���������"
    Debug.Print Parameters_ByRef(strVar)
    Debug.Print "�������� ����� ������ ���������: " & strVar
End Sub

Function Parameters_ByRef_02(ByRef par As String)
    Debug.Print "�������� ��� �������� � ���������: " & par
    par = "��������� � ��������� ��������"
End Function

Sub TestParameters_ByRef_02()
    Dim strVar As String
    strVar = "�������� �� ������ ���������"
    Debug.Print Parameters_ByRef_02(strVar)
    Debug.Print "�������� ����� ������ ���������: " & strVar
End Sub


Function Parameters_ByVal(ByVal par As String)
    Debug.Print "�������� ��� �������� � ���������: " & par
    par = "��������� � ��������� ��������"
End Function

Sub TestParameters_ByVal()
    Dim strVar As String
    strVar = "�������� �� ������ ���������"
    Call Parameters_ByVal(strVar)
    Debug.Print "�������� ����� ������ ���������: " & strVar
End Sub


Function OutputValues(Balance As Currency, outResult As String)
    On Error GoTo ErrorHandler
    If Balance > 0 Then
        outResult = "�������������"
    Else
        outResult = "������� ��� �������������"
    End If
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function
Sub TestOutputValues()
    Dim curBalance As Currency
    Dim strResult As String
    curBalance = 100
    Call OutputValues(curBalance, strResult)
    Debug.Print strResult & " ������"
End Sub


Function PrintArray(arr() As Integer)
    On Error GoTo ErrorHandler
    Dim strResult As String
    Dim iCounter As Integer
    Dim iFirstIdx As Integer: iFirstIdx = LBound(arr())
    Dim iLastIdx As Integer: iLastIdx = UBound(arr())
    For iCounter = iFirstIdx To iLastIdx
        strResult = strResult & arr(iCounter)
        If iCounter < iLastIdx Then strResult = strResult & ", "
    Next iCounter
    Debug.Print strResult
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function

Sub TestPrintArray()
    Dim iArray(2 To 5) As Integer
    iArray(2) = 1
    iArray(3) = 11
    iArray(4) = 21
    iArray(5) = 31
    Call PrintArray(iArray())
End Sub


Function ParamArrayDemo(ParamArray arr())
    On Error GoTo ErrorHandler
    Dim strResult As String
    Dim iCounter As Integer
    Dim iLastIdx As Integer: iLastIdx = UBound(arr())
    For iCounter = 0 To iLastIdx
        strResult = strResult & arr(iCounter)
        If iCounter < iLastIdx Then strResult = strResult & ", "
    Next iCounter
    Debug.Print strResult
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function

Sub TestParamArray()
     Call ParamArrayDemo(1, "Two", 3, 4, 5)
End Sub

Function DefaultParameters(CourseName As String, _
                    Optional Category As String, _
                    Optional trainer As String = "������ ��������", _
                    Optional online As Boolean = True)
    Debug.Print "�������� �����: " & CourseName
    Debug.Print "���������     : " & Category
    Debug.Print "������        : " & trainer
    Debug.Print "������        : " & online
End Function

Sub TestDefaultParameters()
    'DefaultParameters ("���������������� �� VBA")
    Call DefaultParameters("���������������� �� VBA", "Microsoft Office", , False)
End Sub

Function ArgumentTypes(CourseName As String, _
                    Optional Category As String, _
                    Optional trainerFirstName As String, _
                    Optional trainerLastName As String, _
                    Optional online As Boolean = True)
    Debug.Print "�������� ����� : " & CourseName
    Debug.Print "���������      : " & Category
    Debug.Print "��� �������    : " & trainerFirstName
    Debug.Print "������� �������: " & trainerLastName
    Debug.Print "������         : " & online
End Function

Sub TestPositionalArguments()
    Call ArgumentTypes("���������������� �� VBA", _
                   "Microsoft Office", _
                   "��������", _
                   "������", _
                    online:=False)
End Sub

Sub TestNamedArguments()
    Call ArgumentTypes(trainerLastName:="��������", _
                 trainerFirstName:="������", _
                       CourseName:="���������������� �� VBA", _
                           online:=False)
End Sub

Sub TestMixedArguments()
    Call ArgumentTypes("���������������� �� VBA", online:=False)
End Sub


Function GetBalanceState(Balance As Currency) As String
    On Error GoTo ErrorHandler
    If Balance > 0 Then
        GetBalanceState = "�������������"
    Else
        GetBalanceState = "������� ��� �������������"
    End If
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function
Sub TestOutputValues_02()
    Dim curBalance As Currency
    curBalance = 100
    Debug.Print GetBalanceState(curBalance) & " ������"
End Sub

Sub TestOutputValues_03()
    Dim curBalance As Currency
    Dim strResult As String
    curBalance = 100
    OutputValues curBalance, strResult
    Debug.Print strResult & " ������"
End Sub
' ��� ������������� ��������
Function WithReturnedType() As String
    WithReturnedType = "�������� ������������� ����"
    WithReturnedType = 100
End Function

Function WithoutReturnedType()
    WithoutReturnedType = "�������� ������������� ����"
    WithoutReturnedType = 100
End Function
Sub TestReturnedType()
    Debug.Print "WithReturnedType", TypeName(WithReturnedType())
    Debug.Print "WithoutReturnedType", TypeName(WithoutReturnedType())
End Sub
' ������� ��� ������������ �������
Function WithoutReturnValue()
End Function

Sub TestWithoutReturnValue()
    Debug.Print "/" & WithoutReturnValue() & "/"
End Sub



