Attribute VB_Name = "TestGlobalScope2"
Option Explicit

Private Sub SameNameProcedure()
    Debug.Print "Я в модуле TestGlobalScope2"
End Sub

Sub TestSameNameProcedure()
    SameNameProcedure
End Sub

