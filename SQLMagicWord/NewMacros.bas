Attribute VB_Name = "NewMacros"
Option Explicit

Sub Макрос2()
Attribute Макрос2.VB_ProcData.VB_Invoke_Func = "SQLMagic.NewMacros.Макрос2"
'
' Макрос2 Макрос
'
'
    Application.UserName = "Михаил Гурьянов"
End Sub
Sub Макрос3()
Attribute Макрос3.VB_ProcData.VB_Invoke_Func = "SQLMagic.NewMacros.Макрос3"
'
' Макрос3 Макрос
'
'
    Application.DefaultSaveFormat = "MacroEnabledDocument"
End Sub
Sub Макрос4()
Attribute Макрос4.VB_ProcData.VB_Invoke_Func = "SQLMagic.NewMacros.Макрос4"
'
' Макрос4 Макрос
'
'
    Options.SaveInterval = 5
    Application.DefaultSaveFormat = ""
End Sub
