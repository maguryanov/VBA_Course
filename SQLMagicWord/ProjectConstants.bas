Attribute VB_Name = "ProjectConstants"
Option Explicit
' Байтовые значения
Public Const MAX_BYTE As Byte = 255
Public Const MIN_BYTE As Byte = 0

' Целочисленные
Public Const MAX_INTEGER As Integer = 32767
Public Const MIN_INTEGER As Integer = -32768

Public Const MAX_LONG As Long = 2147483647
Public Const MIN_LONG As Long = -2147483648#

'Public Const MAX_LONG_LONG As LongLong = 9223372036854775807
'Public Const MIN_LONG_LONG As LongLong = -9223372036854775808

'Single (одинарная точность)
Public Const MAX_SINGLE As Single = 3.402823E+38
Public Const MIN_SINGLE As Single = -3.402823E+38

' Double (двойная точность)
Public Const MAX_DOUBLE As Double = 1.79769313486231E+308
Public Const MIN_DOUBLE As Double = -1.79769313486231E+308

Public Const MAX_CURRENCY As Currency = 922337203685477.5807@
Public Const MIN_CURRENCY As Currency = -922337203685477.5807@

' Для работы с датами
Public Const MAX_DATE As Date = #12/31/9999#
Public Const MIN_DATE As Date = #1/1/100#

' Максимальная длина строки
Public Const MAX_STRING_LENGTH As Long = 2 ^ 31 - 1 ' ? около 2 миллиарда символов


Public Const COPYRIGHT_SIGN As String = "©"
