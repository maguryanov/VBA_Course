Attribute VB_Name = "TestGlobalScope2"
Option Explicit

Private Sub SameNameProcedure()
    Debug.Print "� � ������ TestGlobalScope2"
End Sub

Sub TestSameNameProcedure()
    SameNameProcedure
End Sub

