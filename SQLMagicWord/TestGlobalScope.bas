Attribute VB_Name = "TestGlobalScope"
Option Explicit

Sub GlobalScopeProcedure()
    Debug.Print "Я доступна во всех модулях"
End Sub

Private Sub ModuleScopeProcedure()
    Debug.Print "Я доступна только в своём модуле"
End Sub

Sub TestModuleScopeProcedure()
    ModuleScopeProcedure
End Sub

Sub SameNameProcedure()
    Debug.Print "Я в модуле TestGlobalScope"
End Sub

Sub TestSameNameProcedure()
    SameNameProcedure
End Sub
