Attribute VB_Name = "m_02_08_OtherOperators"
Option Explicit
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub AppActivate_01()
On Error GoTo ErrorHandler
    Dim dProgramID As Double
    dProgramID = Shell("notepad.exe", vbNormalFocus)
    Debug.Print dProgramID
    AppActivate dProgramID
    SendKeys "Данные введенные из VBA"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub AppActivate_02()
On Error GoTo ErrorHandler
    Dim oShell As Object
    Dim iCounter As Integer
    Const cDelay As Long = 100
    Set oShell = CreateObject("WScript.Shell")
    oShell.Run "calc.exe"
    oShell.AppActivate "Калькулятор"
    Sleep 2000
    'shell.SendKeys "100500"
    SendKeys "100500"
    For iCounter = 1 To 15
        Sleep cDelay
        SendKeys "{TAB}"
    Next iCounter
    Sleep cDelay
    SendKeys " "
    Sleep cDelay
    SendKeys "100"
    Sleep cDelay
    SendKeys "{ENTER}"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


