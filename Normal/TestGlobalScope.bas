Attribute VB_Name = "TestGlobalScope"
Option Explicit

Sub GlobalScopeProcedure()
    Debug.Print "� �������� �� ���� �������"
End Sub

Sub SameNameProcedure()
    Debug.Print "� � ������ TestGlobalScope"
End Sub

Sub TestSameNameProcedure()
    SameNameProcedure
End Sub

