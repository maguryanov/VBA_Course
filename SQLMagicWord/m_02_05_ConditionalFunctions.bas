Attribute VB_Name = "m_02_05_ConditionalFunctions"
Option Explicit

Function IIF_01(Balance As Currency) As String
On Error GoTo ErrorHandler
    IIF_01 = IIf(Balance > 0, "Положительный", "Нулевой или отрицательный")
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function


Function IIF_02(Balance As Currency) As String
On Error GoTo ErrorHandler
    IIF_02 = IIf(Balance > 0, "Положительный", _
                IIf(Balance = 0, "Нулевой", "отрицательный"))
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function

Function Switch_01(Balance As Currency) As String
On Error GoTo ErrorHandler
    Switch_01 = Switch(Balance > 0, "Положительный", _
        Balance < 0, "Отрицательный", Balance = 0, "Нулевой")
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function

Function Switch_02(Balance As Currency) As String
On Error GoTo ErrorHandler
    Switch_02 = Switch(Balance > 0, "Положительный")
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function

Function Switch_03(Salary As Currency) As String
On Error GoTo ErrorHandler
    Switch_03 = Switch(Salary < 100000, "Низкая", Salary < 200000, "Средняя", _
                                True, "Высокая")
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function


Function Switch_04(ColorNumber As Integer) As String
On Error GoTo ErrorHandler
    Switch_04 = Switch(ColorNumber = 1, "Красный", ColorNumber = 2, "Жёлтый", _
            ColorNumber = 3, "Зелёный", True, "Неопределённый")
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function


Function Choose_01(ColorNumber As Integer) As String
On Error GoTo ErrorHandler
    Choose_01 = Choose(ColorNumber, "Красный", "Жёлтый", "Зелёный")
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function

Function Choose_02(ColorNumber As Integer) As Variant
On Error GoTo ErrorHandler
    Choose_02 = Choose(ColorNumber, "Красный", "Жёлтый", "Зелёный")
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function

Function Choose_03(ColorNumber As Integer) As String
On Error GoTo ErrorHandler
    Dim varColor As Variant
    varColor = Choose(ColorNumber, "Красный", "Жёлтый", "Зелёный")
    Choose_03 = IIf(IsNull(varColor), "Неопределённый", varColor)
    'В Access есть функция NZ
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function

Sub DataTypes_04()
On Error GoTo ErrorHandler
    Debug.Print TypeName(IIf(True, 100, "Двести"))
    Debug.Print TypeName(Switch(False, 100, True, "Двести"))
    Debug.Print TypeName(Choose(3, "Красный", 100, Now))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

