Attribute VB_Name = "TestGlobalScope"
Option Explicit

Sub GlobalScopeProcedure()
    Debug.Print "Я доступна во всех модулях"
End Sub

Sub SameNameProcedure()
    Debug.Print "Я в модуле TestGlobalScope"
End Sub

Sub TestSameNameProcedure()
    SameNameProcedure
End Sub

