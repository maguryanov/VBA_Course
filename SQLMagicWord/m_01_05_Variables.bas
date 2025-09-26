Attribute VB_Name = "m_01_05_Variables"
'Option Explicit

'Переменная - это именованная область памяти для хранения данных.
Sub d_01_Variable()
    
    Dim lngProductQty As Long
    lngProductQty = 1000&
    Debug.Print lngProductQty
    
    'Какие преимущества дает использование переменной?
End Sub

'У переменной есть имя. Что нельзя?
Sub d_02_VariableNameRules()
'Длина имени не может превышать 255 знаков...

    Dim v As String
    
    Debug.Print "Успешное выполнение"
    
End Sub

'Имя переменной. Затенение
Sub d_03_Shadowing()
    
'    Dim ThisDocument As String
'    ThisDocument = "d:\Docs\report.docx"
    ThisDocument.Paragraphs(1).Alignment = wdAlignParagraphLeft
    
    Debug.Print "Успешное выполнение"

End Sub


'Имя переменной. Что рекомендуется?
Sub d_04_VariableNameBestPractices()
    
    Dim Count As Long
    Dim Price As Currency
    Dim Active As Boolean
    Dim BirthDate As Date
    Dim Qty As Long
    Dim Temp As Double
    Dim Temperature As Double

    Dim LastName As String  'PascalCase
    Dim FirstName As String 'camelCase
    Dim full_name As String 'snake_case
    
    Debug.Print "Успешное выполнение"
End Sub

'Имя переменной. Венгерская нотация
Sub d_05_VariableNameNotation()
    
    Dim strFirstName As String   ' str - строка
    Dim lngCount As Long         ' lng - long
    Dim dblTemperature As Double ' dbl - double
    Dim curPrice As Currency     ' cur - currency
    Dim boolActive As Boolean    ' bool - boolean
    Dim dtBirthDate As Date      ' dt - date
    
    Debug.Print "Успешное выполнение"
End Sub

'Имя переменной. Русские названия
Sub d_06_НазваниеПеременнойПоРусски()
    
    Dim strИмя As String
    Dim intКоличество As Integer
    Dim curЦена As Double
    Dim boolАктивен As Boolean
    Dim dtДатаРождения As Date
    
    Debug.Print "Успешное выполнение"
End Sub

'У переменной есть значение. Оператор Let
Sub d_07_VariableValue()
    
    Dim strLastName As String
    
    Let strLastName = "Иванова"
    Debug.Print strLastName
    
    strLastName = "Петрова"
    Debug.Print strLastName
   
End Sub



'Оператор Let. Подробно
Sub d_08_Let()
    
    Dim curPrice As Currency
    
    curPrice = 900@
    Debug.Print curPrice
    
    curPrice = curPrice + 100@
    Debug.Print curPrice
    
    curPrice = curPrice * 0.9@
    Debug.Print curPrice
   
End Sub


' Значения по умолчанию для разных типов переменных
Sub d_09_DefaultValues()
    
    Dim byteVar As Byte
    Dim intVar As Integer
    Dim lngVar As Long
    'Dim llVar As LongLong
    Dim sngVar As Single
    Dim dblVar As Double
    Dim curVar As Currency
    Dim strVarFix As String * 3
    Dim strVar As String
    Dim boolVar As Boolean
    Dim dtVar As Date
    Dim varVar As Variant
    Debug.Print TypeName(byteVar), byteVar
    Debug.Print TypeName(intVar), intVar
    Debug.Print TypeName(lngVar), lngVar
    'Debug.Print TypeName(llVar), llVar
    Debug.Print TypeName(sngVar), sngVar
    Debug.Print TypeName(dblVar), dblVar
    Debug.Print TypeName(curVar), curVar
    Debug.Print TypeName(strVarFix), "|" & strVarFix & "|"
    Debug.Print TypeName(strVar), "|" & strVar & "|"
    Debug.Print TypeName(boolVar), boolVar
    Debug.Print TypeName(dtVar), Format(dtVar, "dd.mm.yyyy hh:nn:ss")
    Debug.Print TypeName(varVar), varVar

End Sub


'У переменной есть тип. Если не указан - Variant
Sub d_10_VariableType()
    
    Debug.Print Qty, TypeName(Qty)
    Qty = 1000&
    Debug.Print Qty, TypeName(Qty)
    Qty = "40шт"
    Debug.Print Qty, TypeName(Qty)

End Sub


Sub d_11_SpeedOfVariant()
    Const lngIterations = 100000000
    Dim lngQty As Long
    Dim varQty As Variant
    
    Dim dblStartTime As Double
    Dim lngCounter As Long
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        lngQty = (lngQty + 1) * 2 / 2 - 1
    Next lngCounter
    
    Dim dblLong As Double: dblLong = Timer - dblStartTime
    Debug.Print "Long: "; dblLong
    
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        varQty = (varQty + 1) * 2 / 2 - 1
    Next lngCounter
    
    Dim dblVariant As Double: dblVariant = Timer - dblStartTime
    Debug.Print "Variant: "; dblVariant
    
    Debug.Print "Ratio: "; dblVariant / dblLong

End Sub



'Тип переменной можно определить суффиксом. Не все типы
Sub d_12_VariableType()
    '%   Integer
    '&   Long
    '^   LongLong
    '@   Currency
    '!   Single
    '#   Double
    '$   String
    
    Debug.Print TypeName(ProductQty&), TypeName(Price@), TypeName(FirstName$)

End Sub



'Почему важно декларировать переменные? IntelliSence
Sub d_14_VariableDeclaration()
    
    'Dim dtInvoicePaymentDeadline As Date
    dtInvoicePaymentDeadline = #9/25/2025#
    Debug.Print
    
End Sub


'Почему важно декларировать переменные? Option Explicit
Sub d_15_VariableDeclaration()
    Dim Price As Currency
    'Priсe = 1000@
    Debug.Print Price, TypeName(Price)

End Sub



'Декларирование переменной. Тип Variant по умолчанию
Sub d_16_VariantDefault()
    
    Dim Price
    Debug.Print Price, TypeName(Price)
    Price = 1000@
    Debug.Print Price, TypeName(Price)

End Sub


'Декларирование переменных
Sub d_17_VariableDeclaration()

    Dim strProductName As String
    Dim strFirstName, strLastName As String
    Dim dtBirthDate As Date, strGender As String
    
    Dim lngCounter As Long: lngCounter = 0

End Sub


'Константа – это именованное значение, которое не может быть изменено в ходе выполнения программы.
Sub d_18_Constants()

    Const curIncomeTax As Currency = 0.13@
    
    Debug.Print "Налог:" & (100000 * curIncomeTax)
    Debug.Print "На руки:" & (100000 * (1 - curIncomeTax))
    
End Sub

'Не получится изменить
Sub d_19_Constants()

    Const curIncomeTax As Currency = 0.13
    'curIncomeTax = 0.15
    Debug.Print "Налог: " & (100000 * curIncomeTax)
    Debug.Print "На руки: " & (100000 * (1 - curIncomeTax))
    
End Sub

' Можно не определять тип константы
Sub d_20_ConstantType()

    Const curIncomeTax = 0.13@
    
    Debug.Print curIncomeTax, TypeName(curIncomeTax)
    
End Sub


' Некоторые важные встроенные константы
Sub d_21_StandardConstants()
    
    MsgBox "Сообщение содержащее" & vbCrLf & "несколько строк"
    Debug.Print "Сообщение содержащее" & vbTab & "табуляцию"
    Debug.Print "vbRed = ", Hex(vbRed)
    Debug.Print "vbGreen = ", Hex(vbGreen)
    Debug.Print "vbBlue = ", Hex(vbBlue)
    
End Sub

' Использование выражений при определении констант
Sub d_22_Expressions()

    Const curIncomeTaxPercent As Currency = 13@
    Const curNetIncomeRatio As Currency = 1@ - curIncomeTaxPercent / 100@
    Debug.Print 100000@ * curNetIncomeRatio

End Sub


Sub d_23_SpeedOfConstants()
    Const lngIterations = 100000000
    Const curConst As Currency = 0.87@
    Dim curVar As Currency: curVar = 0.87@
    Dim curResult As Currency
    Dim dblStartTime As Double
    Dim lngCounter As Long
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        curResult = 100000@ * curConst
    Next lngCounter
    
    Dim dblConst As Double: dblConst = Timer - dblStartTime
    Debug.Print "Константа: ", dblConst
    
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        curResult = 100000@ * curVar
    Next lngCounter
    
    Dim dblVar As Double: dblVar = Timer - dblStartTime
    Debug.Print "Переменная: ", dblVar
    
    Debug.Print "Ratio: ", dblVar / dblConst

End Sub


' Грабли и их решение
Sub d_24_TypicalErrors()
    
    'Использование ключевых слов в идентификаторах
    'Dim Set As String
    'Использование пробелов в идентификаторах
    'Dim First Name As String
    'Использование вначале цифры или подчёркивания
    'Dim 01Module As String
    'Dim _Value As Currency
    'Затенение (Shadowing)
    'Dim ThisDocument As String
    'ThisDocument.Paragraphs(1).Alignment = wdAlignParagraphLeft
    
    Debug.Print "Успешное выполнение"
End Sub

' Вредные советы
Sub d_25_BadPractices()

    'Использовать не содержательные имена
    Dim a, b, var1, var2 As Long
    'Использовать неочевидные сокращения
    Dim vl As Long
    'Допускать ошибки и именовании
    Dim curPrice As Double
    
    'Не декларировать переменные
    
    'Не использовать опцию Option Explicit
    
    'Неоправданно использовать Variant
    Dim lngQty

End Sub
