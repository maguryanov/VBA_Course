Attribute VB_Name = "TestGlobalScope"
Option Explicit

Sub GlobalScopeProcedure()
    Debug.Print "� �������� �� ���� �������"
End Sub

Private Sub ModuleScopeProcedure()
    Debug.Print "� �������� ������ � ���� ������"
End Sub

Sub TestModuleScopeProcedure()
    ModuleScopeProcedure
End Sub

Sub SameNameProcedure()
    Debug.Print "� � ������ TestGlobalScope"
End Sub

Sub TestSameNameProcedure()
    SameNameProcedure
End Sub
