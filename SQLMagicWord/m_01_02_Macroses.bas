Attribute VB_Name = "m_01_02_Macroses"
Option Explicit

Sub Greeting()
    MsgBox "������, � ������� �������� VBA!"
End Sub

Sub ProgramText()
    Dim a As Integer, b As Integer, c As Integer
    a = 1: b = 2: c = 3
    Debug.Print a, b, c
    Debug.Print "����� ������� ������ "; "����� ������� ������ "; "����� ������� ������ "; "����� ������� ������ "
    Debug.Print "����� ������� ������" _
        ; " ����� ������� ������" _
        ; " ����� ������� ������" _
        ; " ����� ������� ������"
End Sub
