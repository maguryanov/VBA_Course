Attribute VB_Name = "GlobalVariables"
Option Explicit
Public gstrGlobal As String

Sub GlobalVariable_Test()
    
    gstrGlobal = "Переменная определена глобально"
    Debug.Print "pstrNormalProject = "; gstrGlobal
End Sub
