Attribute VB_Name = "ProjectVariables"
Option Explicit
Option Private Module

Public pstrNormalProject As String

Sub Let_pstrNormalProject()
    
    pstrNormalProject = "Переменная определена в проекте Normal"
    Debug.Print "pstrNormalProject = "; pstrNormalProject
    
End Sub
