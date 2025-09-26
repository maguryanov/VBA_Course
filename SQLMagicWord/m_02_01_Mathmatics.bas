Attribute VB_Name = "m_02_01_Mathmatics"
Option Explicit
Private Const PI As Double = 3.14159265358979
Private Const E As Double = 2.718282
Private Const DEG_TO_RAD As Double = PI / 180
Private Const RAD_TO_DEG As Double = 180 / PI
Private Const SEED As Long = 1000

Sub Functions_Abs()
On Error GoTo ErrorHandler
    Debug.Print Abs(-100), TypeName(Abs(-100))
    Debug.Print Abs(-100&), TypeName(Abs(-100&))
    Debug.Print Abs(-100.4343!), TypeName(Abs(-100.4343!))
    Debug.Print Abs(-100.4343), TypeName(Abs(-100.4343))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Тригонометрия, угол в радианах
Sub Functions_Trigonometry()
On Error GoTo ErrorHandler
    Debug.Print "Sin(PI)", Sin(PI), TypeName(Sin(PI))
    Debug.Print "Cos(PI)", Cos(PI), TypeName(Cos(PI))
    Debug.Print "Sin(180)", Sin(DEG_TO_RAD * 180), TypeName(Sin(DEG_TO_RAD * 90))
    Debug.Print "Cos(180)", Cos(DEG_TO_RAD * 180), TypeName(Cos(DEG_TO_RAD * 90))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Экспонента и натуральный логарифм
Sub Functions_ExpLog()
On Error GoTo ErrorHandler
    Debug.Print "Exp(10)", Exp(10), TypeName(Exp(10))
    Debug.Print "E ^ 10", E ^ 10#, TypeName(E ^ 10#)
    Debug.Print "Log(10)", Log(10), TypeName(Log(10))
    'логарифмы по основанию 10
    Debug.Print Log(100) / Log(10#)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

' Int для отрицательных значений возвращает первое отрицательное
' целое число меньше или равно числу
' Преобразует к меньшему целому или оставляет как есть
Sub Functions_FixInt_01()
On Error GoTo ErrorHandler
    Debug.Print Int(1234.56), TypeName(Int(1234.56))
    Debug.Print Int(-1234.56), TypeName(Int(-1234.56))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

' Fix для отрицательных значений возвращает первое отрицательное
' целое число больше или равно числу
' Отбрасывает дробную часть
Sub Functions_FixInt_02()
On Error GoTo ErrorHandler
    Debug.Print Fix(1234.56), TypeName(Fix(1234.56))
    Debug.Print Fix(-1234.56), TypeName(Fix(-1234.56))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_FixInt_03()
On Error GoTo ErrorHandler
    Debug.Print TypeName(Int(1234)), TypeName(Int(1234&))
    Debug.Print TypeName(Int(1234.56!)), TypeName(Int(1234.56))
    Debug.Print TypeName(Int(1234.56@)), TypeName(Int(CDec(1234.56)))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_FixInt_04()
On Error GoTo ErrorHandler
    Debug.Print #4/12/1981 5:34:25 PM#, Format(CDate(0), "dd mmmm yyyy г.")
    Debug.Print Int(#4/12/1981 5:34:25 PM#), TypeName(Int(#4/12/1981 5:34:25 PM#))
    Debug.Print Int(#4/12/1881 5:34:25 PM#), TypeName(Int(#4/12/1881 5:34:25 PM#))
    Debug.Print Fix(#4/12/1981 5:34:25 PM#), TypeName(Fix(#4/12/1981 5:34:25 PM#))
    Debug.Print Fix(#4/12/1881 5:34:25 PM#), TypeName(Fix(#4/12/1881 5:34:25 PM#))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Function RandomFromRange(LowerBound As Long, UpperBound As Long) As Long
    RandomFromRange = Int((UpperBound - LowerBound + 1) * Rnd + LowerBound)
End Function

Function RandomFromRange2(LowerBound As Long, UpperBound As Long) As Long
    Randomize Timer
    RandomFromRange2 = Int((UpperBound - LowerBound + 1) * Rnd + LowerBound)
End Function

Sub Functions_Random_01()
On Error GoTo ErrorHandler
    Randomize Timer
    Dim iCounter As Integer
    For iCounter = 1 To 5
        Debug.Print Rnd()
    Next iCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_Random_02()
On Error GoTo ErrorHandler
    Randomize Timer
    Dim iCounter As Integer
    For iCounter = 1 To 5
        Debug.Print RandomFromRange(100000, 200000)
    Next iCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Rnd(отрицательное число) - Воспроизводимая последовательность (с фиксированным seed)
Sub Functions_Random_03()
On Error GoTo ErrorHandler
    Debug.Print Rnd(-1000)
    Dim iCounter As Integer
    For iCounter = 1 To 5
        Debug.Print RandomFromRange(100000, 200000)
    Next iCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Rnd(0) - повторение предыдущего случайного числа
Sub Functions_Random_04()
On Error GoTo ErrorHandler
    Rnd 0
    Dim iCounter As Integer
    For iCounter = 1 To 5
        Debug.Print Rnd(), Rnd(0)
    Next iCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Rnd(положительное число) - следующее случайное число
Sub Functions_Random_05()
On Error GoTo ErrorHandler
    Const mcSeed As Long = 1000
    Dim iCounter As Integer
    Debug.Print "-----------"
    Debug.Print Rnd(mcSeed)
    For iCounter = 1 To 3
        Debug.Print Rnd()
    Next iCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Functions_Random_06()
On Error GoTo ErrorHandler
    Dim aColor(2) As String
    aColor(0) = "Красный"
    aColor(1) = "Жёлтый"
    aColor(2) = "Зелёный"
    Debug.Print aColor(RandomFromRange2(0, 2))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub




Sub Functions_Sgn()
On Error GoTo ErrorHandler
    Debug.Print Sgn(-100), TypeName(Sgn(-100))
    Debug.Print Sgn(0), TypeName(Sgn(0))
    Debug.Print Sgn(100), TypeName(Sgn(100))
    Debug.Print Sgn(100.123), TypeName(Sgn(100.123))
    Debug.Print Sgn(Now), Sgn(#1/1/1745#)
    Debug.Print Sgn(False), Sgn(True)
    Debug.Print Sgn("One")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_Sqr()
On Error GoTo ErrorHandler
    Debug.Print Sqr(4), TypeName(Sqr(4))
    Debug.Print Sqr(CDec(4)), TypeName(Sqr(CDec(4)))
    Debug.Print 4 ^ (1 / 2), TypeName(4 ^ (1 / 2))
    Debug.Print 8 ^ (1 / 3), TypeName(8 ^ (1 / 3))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

