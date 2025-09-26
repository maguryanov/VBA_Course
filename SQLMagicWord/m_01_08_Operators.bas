Attribute VB_Name = "m_01_08_Operators"
Option Explicit
'Option Compare Binary
Option Compare Text

Sub ImplicitConvertion()
    On Error GoTo ErrorHandler
'    Debug.Print TypeName(1 + 1)
'    Debug.Print 25000 + 25000
'    Debug.Print TypeName(25000 + 25000)
'    Debug.Print 25000& + 25000&, TypeName(25000& + 25000&)
'    Debug.Print 50000 + 50000, TypeName(50000 + 50000)
'    Debug.Print 100500 + 1, TypeName(100500 + 1)
'    Debug.Print 100500# + 1, TypeName(100500# + 1)
'    Debug.Print 100500.12345 + 1, TypeName(100500.12345 + 1)
'    Debug.Print 100500.1234@ + 1, TypeName(100500.1234@ + 1)
'    Debug.Print 100500.1234@ + 1.12345, TypeName(100500.1234@ + 1.12345)
'    Debug.Print "100" + "500"
'    Debug.Print "100" + 500
'    Debug.Print 100000 + "$"
'    Debug.Print 100000 & "$"
'    Debug.Print CStr(100000) + "$", CStr(100000) & "$"
'    Debug.Print "Результат - " + True
'    Debug.Print "Результат - " & True
'    Debug.Print "Дата рождения: " + #2/13/1972#
'    Debug.Print "Дата рождения: " & #2/13/1972#
'    Debug.Print "Сегодня: " & Now
'    Debug.Print "Сегодня: " & Format(Now, "dd.mm.yyyy"), "Сегодня: " _
'        + Format(Now, "dd.mm.yyyy")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub

Sub AssignmentStatement()
    On Error GoTo ErrorHandler
    Dim CurBal As Currency
    CurBal = 100
    Debug.Print "****************************"
    Debug.Print "Остаток:" & CurBal
    CurBal = CurBal + 200
    Debug.Print "****************************"
    Debug.Print "Остаток:" & CurBal
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

' % - Integer
' & - Long
' ! - Single
' # - Double
' @ - Currency'
Sub ArithmeticAddition()
    On Error GoTo ErrorHandler
    Debug.Print "******************************"
    Debug.Print 1 + 1, TypeName(1 + 1)
    Debug.Print 0.5! + 0.7!, TypeName(0.5! + 0.5!)
    'Debug.Print -25000 - 25000, TypeName(25000 + 25000)
    Debug.Print 1 + 500000, TypeName(1 + 500000)
    Debug.Print 500000 + 1, TypeName(500000 + 1)
    Debug.Print 500000.12345 + 1, TypeName(500000.12345 + 1)
    Debug.Print 0.5! + 100500100500.123, TypeName(0.5! + 100500100500.123)
    Debug.Print 100500100500.123 + 0.5@, TypeName(100500100500.123 + 0.5@)
    Debug.Print 123456789012345@ + CDec(0.1234567789), TypeName(0.5@ + CDec(0.5))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


' % - Integer
' & - Long
' ! - Single
' # - Double
' @ - Currency'
Sub ArithmeticSubtraction()
    On Error GoTo ErrorHandler
    Debug.Print "*************Результат*****************"
    Debug.Print 1 - 1, TypeName(1 - 1)
    Debug.Print 0.5! - 0.5!, TypeName(0.5! - 0.5!)
    'Debug.Print 25000 - 25000, TypeName(25000 - 25000)
    Debug.Print 1 - 500000, TypeName(1 - 500000)
    Debug.Print 500000 - 1, TypeName(500000 - 1)
    Debug.Print 500000.12345 - 1, TypeName(500000.12345 - 1)
    Debug.Print 0.5! - 100500100500.123, TypeName(0.5! - 100500100500.123)
    Debug.Print 100500100500.123 - 0.5@, TypeName(100500100500.123 - 0.5@)
    Debug.Print 123456789012345@ - CDec(0.1234567789), TypeName(0.5@ - CDec(0.5))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

' % - Integer
' & - Long
' ! - Single
' # - Double
' @ - Currency'
Sub ArithmeticMultiplication()
    On Error GoTo ErrorHandler
    Debug.Print "******************************"
    Debug.Print 1 * 1, TypeName(1 * 1)
    Debug.Print 0.5! * 0.7!, TypeName(0.5! * 0.5!)
    Debug.Print 0.5@ * 0.7@, TypeName(0.5@ * 0.7@)
    'Debug.Print 25000 * 2, TypeName(25000 * 2)
    Debug.Print 25000 * 2&, TypeName(25000 * 2&)
    Debug.Print 1 * 500000, TypeName(1 * 500000)
    Debug.Print 500000 * 1, TypeName(500000 * 1)
    Debug.Print 500000.12345 * 1, TypeName(500000.12345 * 1)
    Debug.Print 0.5! * 100500100500.123, TypeName(0.5! * 100500100500.123)
    Debug.Print 100500100500.123 * 0.5@, TypeName(100500100500.123 * 0.5@)
    Debug.Print 123456789012345@ * CDec(0.1234567789), TypeName(123456789012345@ * CDec(0.1234567789))
    Debug.Print 12345@ * CDec(0.1234567789), TypeName(12345@ * CDec(0.1234567789))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


' % - Integer
' & - Long
' ! - Single
' # - Double
' @ - Currency'
Sub ArithmeticDivision()
    On Error GoTo ErrorHandler
    Debug.Print "******************************"
    Debug.Print 1 / 1, TypeName(1 / 1)
    Debug.Print 0.5! / 0.7!, TypeName(0.5! / 0.5!)
    Debug.Print 0.5@ / 0.7@, TypeName(0.5@ / 0.7@)
    Debug.Print 1! / 1E-25!, TypeName(1! / 1E-25!)
    Debug.Print 25000 / 2&, TypeName(25000 / 2&)
    Debug.Print 1 / 500000, TypeName(1 / 500000)
    Debug.Print 500000 / 1, TypeName(500000 / 1)
    Debug.Print 500000.12345 / 1, TypeName(500000.12345 / 1)
    Debug.Print 0.5! / 100500100500.123, TypeName(0.5! / 100500100500.123)
    Debug.Print 0.5@ / 100500100500.123, TypeName(0.5@ / 100500100500.123)
    Debug.Print 100500100500.123 / 0.5@, TypeName(100500100500.123 / 0.5@)
    Debug.Print 123456789012345@ / CDec(0.1234567789), TypeName(123456789012345@ / CDec(0.1234567789))
    Debug.Print 12345@ / CDec(0.1234567789), TypeName(12345@ / CDec(0.1234567789))
    Debug.Print 1 / 1 * 1, TypeName(1 / 1 * 1)
    'Debug.Print 1 / 0, TypeName(1 / 0)
    
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

' % - Integer
' & - Long
' ! - Single
' # - Double
' @ - Currency'
Sub IntegerDivision_01()
    On Error GoTo ErrorHandler
    
    Debug.Print "******************************"
    Debug.Print 5 / 2, TypeName(5 / 2)
    Debug.Print 5 \ 2, TypeName(5 \ 2)
    Debug.Print 5 Mod 2, TypeName(5 Mod 2)
    Debug.Print 5& \ 2, TypeName(5& \ 2)
    Debug.Print 5& Mod 2, TypeName(5& Mod 2)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
' % - Integer
' & - Long
' ! - Single
' # - Double
' @ - Currency'
Sub IntegerDivision_02()
    On Error GoTo ErrorHandler
    Debug.Print "******************************"
    Debug.Print 5.5 \ 2.1, TypeName(5.5 \ 2.1)
    Debug.Print 5.5 Mod 2.1, TypeName(5.5 Mod 2.1)
    Debug.Print 5.5@ \ 2.1@, TypeName(5.5@ \ 2.1@)
    Debug.Print 5.5@ Mod 2.1@, TypeName(5.5@ Mod 2.1@)
    Debug.Print CDec(5.5@) \ CDec(2.1@), TypeName(CDec(5.5@) \ CDec(2.1@))
    Debug.Print CDec(5.5@) Mod CDec(2.1@), TypeName(CDec(5.5@) Mod CDec(2.1@))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

' % - Integer
' & - Long
' ! - Single
' # - Double
' @ - Currency'
Sub IntegerDivision_03()
    On Error GoTo ErrorHandler
    Debug.Print "******************************"
    Debug.Print 5.5 \ 2.1, TypeName(5.5 \ 2.1)
    Debug.Print 5.5 Mod 2.1, TypeName(5.5 Mod 2.1)
    Debug.Print 5.5@ \ 2.1@, TypeName(5.5@ \ 2.1@)
    Debug.Print 5.5@ Mod 2.1@, TypeName(5.5@ Mod 2.1@)
    Debug.Print CDec(5.5@) \ CDec(2.1@), TypeName(CDec(5.5@) \ CDec(2.1@))
    Debug.Print CDec(5.5@) Mod CDec(2.1@), TypeName(CDec(5.5@) Mod CDec(2.1@))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

' % - Integer
' & - Long
' ! - Single
' # - Double
' @ - Currency
Sub Power()
    On Error GoTo ErrorHandler
    Debug.Print "******************************"
    Debug.Print 3 ^ 3, TypeName(2 ^ 3)
    Debug.Print 3 ^ 100, TypeName(2 ^ 3)
    Debug.Print 3@ ^ 3@, TypeName(2@ ^ 3@)
    Debug.Print CDec(3) ^ CDec(100), TypeName(2 ^ 3)
    Debug.Print 3 * 3 * 3, TypeName(3 * 3 * 3)
    Debug.Print CDec(3) * CDec(3) * CDec(3), TypeName(CDec(3) * CDec(3) * CDec(3))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub EmptyValue()
    On Error GoTo ErrorHandler
    Dim vVar As Variant
    Debug.Print "******************************"
    Debug.Print vVar, TypeName(vVar)
    Debug.Print vVar + 100, TypeName(vVar + 100)
    Debug.Print vVar - 100, TypeName(vVar - 100)
    Debug.Print vVar * 100, TypeName(vVar * 100)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Concatenation_01()
    On Error GoTo ErrorHandler
    Dim strFirstName As String: strFirstName = "Михаил"
    Dim strLastName As String: strLastName = "Гурьянов"
    Debug.Print "******************************"
    Debug.Print strFirstName & " " & strLastName
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Concatenation_By_Plus()
    On Error GoTo ErrorHandler
    Debug.Print "******************************"
    Debug.Print "100" + "500", TypeName("100" + "500")
    Debug.Print "100" + 500, TypeName("100" + 500)
    Debug.Print 100 + "500", TypeName(100 + "500")
    'Debug.Print 10000 + "$", TypeName(10000 + "$")
    Debug.Print CStr(10000) + "$", TypeName(CStr(10000) + "$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Concatenation_By_Ampersand()
    On Error GoTo ErrorHandler
    Debug.Print "******************************"
    Debug.Print 10000 & "$", TypeName(10000 & "$")
    Debug.Print "100" & "500", TypeName("100" & "500")
    Debug.Print 100 & "500", TypeName(100 & "500")
    Debug.Print "100" & 500, TypeName("100" & 500)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub DateArithmetic_01()
    On Error GoTo ErrorHandler
    Dim dtNow As Date: dtNow = Now
    Debug.Print "******************************"
    Debug.Print #2/13/1972# + 1, TypeName(#2/13/1972# + 1)
    Debug.Print dtNow, dtNow + 1
    Debug.Print #2/13/1972# - 10, TypeName(#2/13/1972# - 10)
    Debug.Print dtNow - #2/2/1972#, TypeName(dtNow - #2/2/1972#)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub ComparisonOperators_01()
    On Error GoTo ErrorHandler
    Dim lValue As Long 'lValue = 0
    Debug.Print "***********************************************"
    Debug.Print lValue = 0, lValue <> 0, TypeName(lValue = 0)
    Debug.Print lValue = 0#, lValue <> 0#, TypeName(lValue = 0#)
    Debug.Print lValue = 1E-16, lValue <> 1E-16, TypeName(lValue = 1E-16)
    Debug.Print lValue = False, lValue <> True, TypeName(lValue = False)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub ComparisonOperators_02()
    On Error GoTo ErrorHandler
    Dim lValue As Long: lValue = 1
    Debug.Print "***********************************************"
    Debug.Print lValue > 0, lValue < 0, lValue >= 0, lValue <= 0, TypeName(lValue <= 0)
    Debug.Print lValue > 0.2, lValue < 0.2, lValue >= 0.2, lValue <= 0.2, TypeName(lValue <= 0.2)
    
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub StringComparisonOperators()
    On Error GoTo ErrorHandler
    Dim strVar As String: strVar = "$"
    Dim strFix As String * 2: strFix = "$"
    Debug.Print "***********************************************"
    Debug.Print strVar = "$", strFix = "$", RTrim(strFix) = "$", Trim(strFix) = "$"
    Debug.Print strVar = "$", strFix = "$", RTrim(strFix) = "$", Trim(strFix) = "$"
    Debug.Print "Муха" > "Слон", "Муха" < "Слон", "Муха" = "Слон", "Муха" <> "Слон"
    Debug.Print "2" > "1000000", CCur("2") > CCur("1000000")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub StringComparisonOperators_Like_01()
    On Error GoTo ErrorHandler
    Dim strVar As String: strVar = "Microsoft SQL Server 2025"
    Debug.Print "***********************************************"
    Debug.Print strVar = "Microsoft"
    Debug.Print strVar Like "Microsoft"
    Debug.Print strVar Like "Microsoft*"
    Debug.Print strVar Like "*SQL*"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub StringComparisonOperators_Like_02()
    On Error GoTo ErrorHandler
    Dim strVar As String: strVar = "Microsoft SQL Server 2025"
    Debug.Print "***********************************************"
    Debug.Print strVar Like "*SQL*"
    Debug.Print strVar Like "*sql*"
    Debug.Print "BE-2908" Like "*E*", "ES-2908" Like "*E*"
    Debug.Print "BE-2908" Like "?E*", "ES-2908" Like "?E*"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub StringComparisonOperators_Like_03()
    On Error GoTo ErrorHandler
    Debug.Print "***********************************************"
    Debug.Print "BE-2908" Like "??-####", "BE-290A" Like "??-####"
    Debug.Print "BE-2908" Like "[acd]*", "CE-290A" Like "[acd]*"
    Debug.Print "BE-2908" Like "[a-c]*", "CE-290A" Like "[a-c]*"
    Debug.Print "BE-2908" Like "[!acd]*", "CE-290A" Like "[!acd]*"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub StringComparisonOperators_Is()
    On Error GoTo ErrorHandler
    Dim oApp1 As Application
    Set oApp1 = Application
    Dim oApp2 As Application
    Set oApp2 = Application
    Debug.Print "***********************************************"
    Debug.Print oApp1 Is Application, oApp1 Is oApp2
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub LogicalOperators_And_01()
    On Error GoTo ErrorHandler
    Dim boolPaid As Boolean: boolPaid = True
    Dim boolReady As Boolean: boolReady = False
    Debug.Print "***********************************************"
    If boolPaid And boolReady Then
        Debug.Print "Можно отгружать"
    Else
        Debug.Print "Нельзя отгружать"
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub LogicalOperators_And_02()
    On Error GoTo ErrorHandler
    Debug.Print "***********************************************"
    Debug.Print "False And False", False And False
    Debug.Print "False And True ", False And True
    Debug.Print "True And False ", True And False
    Debug.Print "True And True  ", True And True
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub LogicalOperators_And_03()
    On Error GoTo ErrorHandler
    Dim sTemp As Single: sTemp = -1
    Debug.Print "***********************************************"
    If sTemp >= 0 And sTemp <= 25 Then
        Debug.Print "Нормальная температура хранения"
    Else
        Debug.Print "Недопустимая температура хранения"
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub LogicalOperators_Or_01()
    On Error GoTo ErrorHandler
    Dim boolPassport As Boolean: boolPassport = True
    Dim boolDriverLicense As Boolean: boolDriverLicense = False
    Debug.Print "***********************************************"
    If boolPassport Or boolDriverLicense Then
        Debug.Print "Личность подтверждена"
    Else
        Debug.Print "Личность НЕ подтверждена"
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub LogicalOperators_Or_02()
    On Error GoTo ErrorHandler
    Debug.Print "***********************************************"
    Debug.Print "False Or False", False Or False
    Debug.Print "False Or True ", False Or True
    Debug.Print "True Or False ", True Or False
    Debug.Print "True Or True  ", True Or True
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
Sub LogicalOperators_Or_03()
    On Error GoTo ErrorHandler
    Dim dtDay As Date: dtDay = #3/22/2025#
    Debug.Print "***********************************************"
    If (dtDay >= #2/12/2025# And dtDay <= #2/22/2025#) _
       Or (dtDay >= #4/12/2025# And dtDay <= #4/22/2025#) Then
        Debug.Print "Отпуск"
    Else
        Debug.Print "Рабочий день"
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub LogicalOperators_Xor()
    On Error GoTo ErrorHandler
    Debug.Print "***********************************************"
    Debug.Print "False Xor False", False Xor False
    Debug.Print "False Xor True ", False Xor True
    Debug.Print "True Xor False ", True Xor False
    Debug.Print "True Xor True  ", True Xor True
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub LogicalOperators_Xor_02()
    On Error GoTo ErrorHandler
    Dim boolPositive_1 As Boolean: boolPositive_1 = True
    Dim boolPositive_2 As Boolean: boolPositive_2 = False
    Debug.Print "***********************************************"
    If boolPositive_1 Xor boolPositive_2 Then
        Debug.Print "Притягиваются"
    Else
        Debug.Print "Отталкиваются"
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub LogicalOperators_Equivalence()
    On Error GoTo ErrorHandler
    Debug.Print "***********************************************"
    Debug.Print "False Eqv False", False Eqv False
    Debug.Print "False Eqv True ", False Eqv True
    Debug.Print "True Eqv False ", True Eqv False
    Debug.Print "True Eqv True  ", True Eqv True
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub LogicalOperators_Implication()
    On Error GoTo ErrorHandler
    Debug.Print "***********************************************"
    Debug.Print "False Imp False", False Imp False
    Debug.Print "False Imp True ", False Imp True
    Debug.Print "True Imp False ", True Imp False
    Debug.Print "True Imp True  ", True Imp True
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub LogicalOperators_Not_01()
    On Error GoTo ErrorHandler
    Dim dtDay As Date: dtDay = #3/22/2025#
    Debug.Print "***********************************************"
    If Not ((dtDay >= #2/12/2025# And dtDay <= #2/22/2025#) _
       Or (dtDay >= #4/12/2025# And dtDay <= #4/22/2025#)) Then
        Debug.Print "Рабочий день"
    Else
        Debug.Print "Отпуск"
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub LogicalOperators__Not_02()
    On Error GoTo ErrorHandler
    Debug.Print "***********************************************"
    Debug.Print "Not False", Not False
    Debug.Print "Not True  ", Not True
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
