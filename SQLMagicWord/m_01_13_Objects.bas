Attribute VB_Name = "m_01_13_Objects"
Option Explicit

Sub TestClassPerson_01()
    Dim oPerson As New Person
    oPerson.LastName = "��������"
    oPerson.FirstName = "������"
    oPerson.Gender = "�"
    oPerson.BirthDate = Now - 10000
    Debug.Print oPerson.fullName
    'oPerson.FullName = "������ ��������"
    oPerson.PrintForm
End Sub


Sub TestClassPerson_02()
    Dim oPerson As New Person
    oPerson.LastName = "������"
    oPerson.FirstName = "������"
    oPerson.Gender = "�"
    oPerson.BirthDate = #2/13/1972#
    Debug.Print oPerson.fullName
    oPerson.PrintForm
    Debug.Print oPerson.FullYears
End Sub

Sub TestClassPerson_03()
    Dim oPerson As New Person
    oPerson.LastName = "������"
    oPerson.FirstName = "������"
    oPerson.Gender = "�"
    oPerson.BirthDate = #2/1/1991#
    oPerson.PrintForm
    Debug.Print oPerson.GetFullYears(#1/1/2000#)
    Debug.Print oPerson.GetFullYears()
End Sub

Sub TestClassPerson_04()
    Dim oPerson As New Person
    oPerson.LastName = "��������"
    oPerson.FirstName = "������"
    oPerson.Gender = "�"
    Debug.Print oPerson.fullName
    'oPerson.FullName = "������ ��������"
    oPerson.PrintForm
    'Debug.Print oPerson.GetForm()
End Sub

Sub TestClassPerson_05()
    Dim oPerson As New Person
    oPerson.PrintForm
    Debug.Print oPerson.GetFullYears(#1/1/2000#)
    Debug.Print oPerson.GetFullYears()
End Sub

Sub TestClass_Events()
    Dim oPerson As New Person
    Dim oEventProcessor As New EventProcessor
    Set oEventProcessor.oPerson = oPerson
    oPerson.BirthDate = Now + 10
End Sub


Sub TestVariables_PrimitiveVars_01()
    Dim strVar1 As String, strVar2 As String
    Dim oVar1 As Person, oVar2 As Person
    strVar1 = "01"
    strVar2 = strVar1
    strVar2 = "02"
    Debug.Print "strVar1 = " & strVar1
End Sub

Sub TestVariables_ObjectVars_02()
    Dim oVar1 As New Person, oVar2 As New Person
    oVar1.FirstName = "01"
    Set oVar2 = oVar1
    oVar2.FirstName = "02"
    Debug.Print "oVar1.FirstName = " & oVar1.FirstName
End Sub

Sub TestVariables_Is_03()
    Dim oVar1 As New Person, oVar2 As New Person, oVar3 As New Person
    oVar1.FirstName = "01"
    Set oVar2 = oVar1
    oVar3.FirstName = "01"
    Debug.Print oVar1 Is oVar2
    Debug.Print oVar1 Is oVar3
End Sub


'������� ���������� (Late Binding). �������� ����� �� ����� ����������

Sub TestLateBinding_01()
    Dim oApp As Object
    Set oApp = CreateObject("Word.Application")
    Debug.Print oApp.Version
End Sub

'������ ���������� (Early Binding). �������� ����� �� ����� ����������

Sub TestEarlyBinding_03()
    Dim oApp As New Application
    Debug.Print oApp.Version
End Sub

Sub TestEarlyBinding_04()
    Dim oApp As Application
    Set oApp = New Application
    Debug.Print oApp.Version
End Sub

'������ ���������� (Early Binding). �������� ����� �� ����� ����������
' ��� ����� ������� � �������

Sub TestEarlyBinding_01()
    Dim oPerson As New Person
    Debug.Print "������ �� ������!"
    oPerson.FirstName = "������"
End Sub

Sub TestEarlyBinding_02()
    Dim oPerson As Person
    Set oPerson = New Person
    Debug.Print "������ ������!"
    oPerson.FirstName = "������"
End Sub

Sub DeleteObject()
    Dim oPerson As Person
    Set oPerson = New Person
    Debug.Print "������ ������!"
    oPerson.FirstName = "������"
    Set oPerson = Nothing
End Sub
