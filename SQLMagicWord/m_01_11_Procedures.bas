Attribute VB_Name = "m_01_11_Procedures"
Option Explicit

Sub TestGlobalScopeProcedure()
    Normal.GlobalScopeProcedure
End Sub

Sub TestProjectScopeProcedure()
    'Normal.ProjectScopeProcedure
End Sub

Sub TestProjectScopeProcedure_01()
    ProjectScopeProcedure
End Sub

Sub TestModuleScopeProcedure()
    'ModuleScopeProcedure
End Sub

Public Sub PublicIsOptionalProcedure()
    Debug.Print "Ключевое слово Public опционально"
End Sub

Sub SameNameProcedure()
    Debug.Print "Я в модуле ProceduresAndFunctions"
End Sub

Sub TestSameNameProcedure()
    'Normal.GlobalScope.SameNameProcedure
End Sub

Sub Parameters(CourseName As String, _
                 Category As String, _
                  trainer As String, _
                   online As Boolean)
    Debug.Print "Название курса: " & CourseName
    Debug.Print "Категория     : " & Category
    Debug.Print "Тренер        : " & trainer
    Debug.Print "Онлайн        : " & online
End Sub

Sub TestParameters()
    Parameters "Программирование на VBA", _
        "Microsoft Office", "Михаил Гурьянов", True
End Sub

Sub CallDemonstration()
    Debug.Print "Я для демонстрации использования CALL"
End Sub

Sub TestCallDemonstration()
    CallDemonstration
    Call CallDemonstration
End Sub

Sub Parameters_01(par1 As String, par2 As String)
    Debug.Print par1 & " " & par2
End Sub

Sub TestParameters_01()
    Parameters_01 "Вызов", "без CALL"
    Call Parameters_01("Вызов", "c CALL")
End Sub

Sub Parameters_ByRef(par As String)
    Debug.Print "Значение при передаче в процедуру: " & par
    par = "Изменённое в процедуре значение"
End Sub

Sub TestParameters_ByRef()
    Dim strVar As String
    strVar = "Значение до вызова процедуры"
    Parameters_ByRef strVar
    Debug.Print "Значение после вызова процедуры: " & strVar
End Sub

Sub Parameters_ByRef_02(ByRef par As String)
    Debug.Print "Значение при передаче в процедуру: " & par
    par = "Изменённое в процедуре значение"
End Sub

Sub TestParameters_ByRef_02()
    Dim strVar As String
    strVar = "Значение до вызова процедуры"
    Parameters_ByRef_02 strVar
    Debug.Print "Значение после вызова процедуры: " & strVar
End Sub


Sub Parameters_ByVal(ByVal par As String)
    Debug.Print "Значение при передаче в процедуру: " & par
    par = "Изменённое в процедуре значение"
End Sub

Sub TestParameters_ByVal()
    Dim strVar As String
    strVar = "Значение до вызова процедуры"
    Parameters_ByVal strVar
    Debug.Print "Значение после вызова процедуры: " & strVar
End Sub


Sub OutputValues(Balance As Currency, outResult As String)
    On Error GoTo ErrorHandler
    If Balance > 0 Then
        outResult = "Положительный"
    Else
        outResult = "Нулевой или отрицательный"
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
Sub TestOutputValues()
    Dim curBalance As Currency
    Dim strResult As String
    curBalance = 100
    OutputValues curBalance, strResult
    Debug.Print strResult & " баланс"
End Sub


Sub PrintArray(arr() As Integer)
    On Error GoTo ErrorHandler
    Dim strResult As String
    Dim iCounter As Integer
    Dim iFirstIdx As Integer: iFirstIdx = LBound(arr())
    Dim iLastIdx As Integer: iLastIdx = UBound(arr())
    For iCounter = iFirstIdx To iLastIdx
        strResult = strResult & arr(iCounter)
        If iCounter < iLastIdx Then strResult = strResult & ", "
    Next iCounter
    Debug.Print strResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub TestPrintArray()
    Dim iArray(2 To 5) As Integer
    iArray(2) = 1
    iArray(3) = 11
    iArray(4) = 21
    iArray(5) = 31
    PrintArray iArray()
End Sub


Sub ParamArrayDemo(ParamArray arr())
    On Error GoTo ErrorHandler
    Dim strResult As String
    Dim iCounter As Integer
    Dim iLastIdx As Integer: iLastIdx = UBound(arr())
    For iCounter = 0 To iLastIdx
        strResult = strResult & arr(iCounter)
        If iCounter < iLastIdx Then strResult = strResult & ", "
    Next iCounter
    Debug.Print strResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub TestParamArray()
     ParamArrayDemo 1, "Two", 3, 4, 5
End Sub

Sub DefaultParameters(CourseName As String, _
                    Optional Category As String, _
                    Optional trainer As String = "Михаил Гурьянов", _
                    Optional online As Boolean = True)
    Debug.Print "Название курса: " & CourseName
    Debug.Print "Категория     : " & Category
    Debug.Print "Тренер        : " & trainer
    Debug.Print "Онлайн        : " & online
End Sub

Sub TestDefaultParameters()
    'DefaultParameters "Программирование на VBA"
    DefaultParameters "Программирование на VBA", "Microsoft Office", , False
End Sub

Sub ArgumentTypes(CourseName As String, _
                    Optional Category As String, _
                    Optional trainerFirstName As String, _
                    Optional trainerLastName As String, _
                    Optional online As Boolean = True)
    Debug.Print "Название курса : " & CourseName
    Debug.Print "Категория      : " & Category
    Debug.Print "Имя тренера    : " & trainerFirstName
    Debug.Print "Фамилия тренера: " & trainerLastName
    Debug.Print "Онлайн         : " & online
End Sub

Sub TestPositionalArguments()
    ArgumentTypes "Программирование на VBA", _
                   "Microsoft Office", _
                   "Гурьянов", _
                   "Михаил", _
                    online:=False

End Sub

Sub TestNamedArguments()
    ArgumentTypes trainerLastName:="Гурьянов", _
                 trainerFirstName:="Михаил", _
                       CourseName:="Программирование на VBA", _
                           online:=False
End Sub

Sub TestMixedArguments()
    ArgumentTypes "Программирование на VBA", online:=False
End Sub


Function GetBalanceState(Balance As Currency) As String
    On Error GoTo ErrorHandler
    If Balance > 0 Then
        GetBalanceState = "Положительный"
    Else
        GetBalanceState = "Нулевой или отрицательный"
    End If
    Exit Function
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Function
Sub TestOutputValues_02()
    Dim curBalance As Currency
    curBalance = 100
    Debug.Print GetBalanceState(curBalance) & " баланс"
End Sub

Sub TestOutputValues_03()
    Dim curBalance As Currency
    Dim strResult As String
    curBalance = 100
    OutputValues curBalance, strResult
    Debug.Print strResult & " баланс"
End Sub
' Тип возвращаемого значения
Function WithReturnedType() As String
    WithReturnedType = "Проверка возвращаемого типа"
    WithReturnedType = 100
End Function

Function WithoutReturnedType()
    WithoutReturnedType = "Проверка возвращаемого типа"
    WithoutReturnedType = 100
End Function
Sub TestReturnedType()
    Debug.Print "WithReturnedType", TypeName(WithReturnedType())
    Debug.Print "WithoutReturnedType", TypeName(WithoutReturnedType())
End Sub
' Функция без присваивания функции
Function WithoutReturnValue()
End Function

Sub TestWithoutReturnValue()
    Debug.Print "/" & WithoutReturnValue() & "/"
End Sub

