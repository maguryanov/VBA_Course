Attribute VB_Name = "m_01_05_Literal"
Option Explicit

'Литерал - это прямое указание значения в коде.
Sub d_01_Literal()

    Debug.Print "Сумма: "; 2144.48; " руб."

End Sub


'У литерала есть тип данных
Sub d_02_DataType()

    Debug.Print TypeName("Сумма: ")
    Debug.Print TypeName(2144.48)
    
End Sub

'Целые типы
Sub d_03_IntegerLiteral()
    
    'Integer 2 байта     от –32 768 до 32 767
    Debug.Print 10, TypeName(10)
    
    'Long 4 байта     от –2 147 483 648 до 2 147 483 647
    Debug.Print 100500, TypeName(100500)

End Sub

'Числа с плавающей точкой
Sub d_04_FloatPointLiteral()
    
    'Double (число с плавающей точкой двойной точности)     8 байт
    Debug.Print 36.6; TypeName(36.6)
    
    'Single (число с плавающей точкой одинарной точности)   4 байта
    Debug.Print 36.6!; TypeName(36.6!)
    

End Sub

'Числа с фиксированной точкой
Sub d_05_FixedPointLiteral()
    
    'Currency (масштабируемое целое число)   8 байт
    Debug.Print TypeName(1200.57@); 1200.57@
    Debug.Print TypeName(123456789012345.1235@); 1200.57@

End Sub

'Строки
Sub d_06_StringLiteral()
    
    Debug.Print TypeName("Михаил Гурьянов"), "Михаил Гурьянов"
    Debug.Print TypeName("MA-144F5"), "MA-144F5"
    Debug.Print TypeName("4080281000000000000"), "4080281000000000000"
    Debug.Print TypeName("00123"), "00123"
    Debug.Print TypeName("123456,12345"), "123456,12345"
    Debug.Print TypeName(""), "|"; ""; "|"
    Debug.Print "Курс ""Основы VBA"""
    
End Sub

'Дата/время литерал
Sub d_07_DateLiteral()
    
    Debug.Print TypeName(#9/23/2025#), #9/23/2025#
    Debug.Print "Дата", #9/23/2025#
    Debug.Print "Время", #10:30:44 PM#
    Debug.Print "Дата и время", #9/23/2025 10:30:00 PM#
    Debug.Print #9/23/2025# + 1              'Дата
    
End Sub

'Логический тип
Sub d_08_BooleanLiteral()
    
    Debug.Print TypeName(True), True
    Debug.Print TypeName(False), False

End Sub

'Тип данных может определяться автоматически
Sub d_09_AutomaticType()

    Debug.Print TypeName(100); 100
    Debug.Print TypeName(100500); 100500
    
End Sub

'Тип может указываться явно
Sub d_10_ExplicitType()
    Debug.Print TypeName(100), 100        ' % - Integer"
    Debug.Print TypeName(100&), 100&     ' & - Long"
    Debug.Print TypeName(100.5!), 100.5! ' ! - Single"
    Debug.Print TypeName(100#), 100#     ' # - Double"
    Debug.Print TypeName(200@), 200@     ' @ - Currency
    
End Sub

'Почему полезно указывать тип явно?
Sub d_11_ExplicitType()
    
    'Использование типа по умолчанию может приводить к ошибкам
    Debug.Print 20000& + 20000
    
End Sub


'Почему полезно указывать тип явно?
Sub d_12_ExplicitType()
    
    'Проверка и автоматическое исправление на стадии печати значения
    Debug.Print 40000
    Debug.Print 12345678
    
End Sub



'Почему полезно указывать тип явно?
Sub d_14_ExplicitType()
            
    'Currency и Single автоматически не определяется
    Debug.Print 499.99
    Debug.Print 36.6!
        
End Sub


'Почему полезно указывать тип явно?
Sub d_15_ExplicitType()
    
    'Если нужно Double, а число не содержит дробной части
    Debug.Print 100#
    
End Sub


'Шестнадцатеричные и восьмеричные литералы
Sub d_16_HexAndOctalLiterals()

    Debug.Print "Значение &HF: "; TypeName(&HF); &HF
    Debug.Print "Значение &O10: "; &O10
    Debug.Print "Значение &HFFFF: "; TypeName(&HFFFF), &HFFFF
    
    Debug.Print "Hex(0) = "; Hex(0)
    Debug.Print "Hex(255) = "; Hex(255)
    Debug.Print "Oct(8) = "; Oct(8)
End Sub


'Экспоненциальная или научная запись
Sub d_17_ExponentialLiteral()

    Debug.Print 1.23E+20; "Это значит"; 1.23 * 10 ^ 20
    Debug.Print 1.23E-20; "Это значит"; 1.23 / 10 ^ 20
    'Преобразование в среде VBA может быть признаком округления даже в целой части!
    Debug.Print "12345678!" = 1.234568E+07!
    
End Sub

'Типичные ошибки при работе с литералами
Sub d_18_TypicalErrors()

    'Debug.Print Мороженое
    'Debug.Print "Лимонад "Колокольчик""
    'Debug.Print 32767 + 1
    'Debug.Print 1000 + "$"
    'Debug.Print "12345678!" = 12345678!

End Sub


