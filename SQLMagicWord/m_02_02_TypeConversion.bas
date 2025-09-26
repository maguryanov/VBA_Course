Attribute VB_Name = "m_02_02_TypeConversion"
Option Explicit

Sub Functions_Val_01()
On Error GoTo ErrorHandler
    Debug.Print "---------------------------------"
    Debug.Print Val("100.123456"), TypeName(Val("100.123456"))
    Debug.Print Val("100@"), TypeName(Val("100@"))
    Debug.Print Val("1000000$"), Val("1000000 руб.")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_Val_02()
On Error GoTo ErrorHandler
    Dim strVal As String
    Debug.Print "---------------------------------"
    Debug.Print Val("1 000 000$")
    strVal = "1" & vbTab & "000 руб"
    Debug.Print Val(strVal)
    strVal = "1" & vbCrLf & "000 руб"
    Debug.Print Val(strVal)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Если не может преобразовать возвращает 0
Sub Functions_Val_03()
On Error GoTo ErrorHandler
    Debug.Print "---------------------------------"
    Debug.Print Val("$1000000"), Val("Итого: 1000000 руб."), Val(True)
    Debug.Print Val(#2/13/1972#), Val("13.02.1972")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_Str()
On Error GoTo ErrorHandler
    Debug.Print "---------------------------------"
    Debug.Print str(100.123), TypeName(str(100.123)) 'Variant(String)
    Debug.Print str(100.123@), TypeName(str(100.123@)) 'Variant(String)
    Debug.Print str(Now), TypeName(str(Now)) 'Variant(String)
    Debug.Print str(True), TypeName(str(True)) 'Variant(String)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_Hex()
On Error GoTo ErrorHandler
    Dim varVar As Variant
    Debug.Print "---------------------------------"
    Debug.Print Hex(15), TypeName(Hex(15)) 'String
    Debug.Print Hex(31), Hex(31.123)
    Debug.Print Hex(varVar)
    Debug.Print Hex("One")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_Oct()
On Error GoTo ErrorHandler
    Dim varVar As Variant
    Debug.Print "---------------------------------"
    Debug.Print Oct(15), TypeName(Oct(15)) 'String
    Debug.Print Oct(31), Oct(31.123)
    Debug.Print Oct(varVar)
    Debug.Print Oct("One")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Functions_CBool()
On Error GoTo ErrorHandler
    Dim varVar As Variant
    Debug.Print "---------------------------------"
    Debug.Print CBool(15), TypeName(CBool(15))
    Debug.Print CBool(0), CBool("-134")
    Debug.Print CBool(Now), CBool(1E-19)
    Debug.Print CBool("True"), CBool("False"), CBool("100500")
    Debug.Print CBool("0$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_Cbyte()
On Error GoTo ErrorHandler
    Dim varVar As Variant
    Debug.Print "---------------------------------"
    Debug.Print CByte(15), TypeName(CByte(15))
    Debug.Print CByte(varVar)
    Debug.Print CByte(123.6789), CByte(123.1), CByte(123.5)
    Debug.Print CByte(True), CByte(False), CByte("200")
    Debug.Print CByte("")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Functions_CInt()
On Error GoTo ErrorHandler
    Dim varVar As Variant
    Debug.Print "---------------------------------"
    Debug.Print CInt(15), TypeName(CInt(15))
    Debug.Print CInt(0), CInt("-134")
    Debug.Print CInt(varVar), CInt(1E-19)
    Debug.Print CInt(123.6789), CInt(123.1), CInt(123.5)
    Debug.Print CInt(-123.6789), CInt(-123.1), CInt(-123.5)
    Debug.Print CInt(True), CInt(False), CInt("500")
    Debug.Print CInt("0$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_CLng()
On Error GoTo ErrorHandler
    Dim varVar As Variant
    Debug.Print "---------------------------------"
    Debug.Print CLng(15), TypeName(CLng(15))
    Debug.Print CLng(0), CLng("-134")
    Debug.Print CLng(varVar), CLng(1E-19)
    Debug.Print CLng(123.6789), CLng(123.1), CLng(123.5)
    Debug.Print CLng(-123.6789), CLng(-123.1), CLng(-123.5)
    Debug.Print CLng(True), CLng(False), CLng("500")
    Debug.Print CLng("0$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Functions_CSng()
On Error GoTo ErrorHandler
    Dim varVar As Variant
    Debug.Print "---------------------------------"
    Debug.Print CSng(15), TypeName(CSng(15))
    Debug.Print CSng(0), CSng("-134")
    Debug.Print CSng(varVar), CSng(1E-19)
    Debug.Print CSng(123.6789), CSng(123.1), CSng(1.23456789012346E+19)
    Debug.Print CSng(-123.6789), CSng(-123.1), CSng(-123.5)
    Debug.Print CSng(True), CSng(False), CSng("500,1234") '. - ошибка
    Debug.Print CSng("0$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_CDbl()
On Error GoTo ErrorHandler
    Dim varVar As Variant
    Debug.Print "---------------------------------"
    Debug.Print CDbl(15), TypeName(CDbl(15))
    Debug.Print CDbl(0), CDbl("-134")
    Debug.Print CDbl(varVar), CDbl(1E-19)
    Debug.Print CDbl(123.6789), CDbl(123.1), CDbl(1.23456789012346E+19)
    Debug.Print CDbl(-123.6789), CDbl(-123.1), CDbl(-123.5)
    Debug.Print CDbl(True), CDbl(False), CDbl("500,99999") '. - ошибка
    Debug.Print CDbl("0$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_CCur()
On Error GoTo ErrorHandler
    Dim varVar As Variant
    Debug.Print "---------------------------------"
    Debug.Print CCur(15), TypeName(CCur(15))
    Debug.Print CCur(0), CCur("-134")
    Debug.Print CCur(varVar), CCur(1E-19)
    Debug.Print CCur(123.6789), CCur(123.1), CCur(1.2345678)
    Debug.Print CCur(-123.6789), CCur(-123.1), CCur(-123.5)
    Debug.Print CCur(True), CCur(False), CCur("500,99999") '. - ошибка
    Debug.Print CCur("0$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_CDec()
On Error GoTo ErrorHandler
    Dim varVar As Variant
    Debug.Print "---------------------------------"
    Debug.Print CDec(15), TypeName(CDec(15))
    Debug.Print CDec(0), CDec("-134")
    Debug.Print CDec(varVar), CDec(1E-19)
    Debug.Print CDec(123.6789), CDec(123.1), CDec(1.2345678)
    Debug.Print CDec(-123.6789), CDec(-123.1), CDec(-123.5)
    Debug.Print CDec(True), CDec(False), CDec("500,99999") '. - ошибка
    Debug.Print CDec("0$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
'1 января 100 (-657 434), до 31 декабря 9999 года (2 958 465)
Sub Functions_CDate()
On Error GoTo ErrorHandler
    Dim varVar As Variant
    Debug.Print "---------------------------------"
    Debug.Print Format(CDate(0), "dd mmmm yyyy г. hh.nn.ss")
    Debug.Print CDate(15), TypeName(CDate(15))
    Debug.Print CDate(varVar), CDate(1E-19)
    Debug.Print CDate(123.6789), CDate(123.1), CDate(1.2345678)
    Debug.Print CDate(-123.6789), CDate(-123.1), CDate(-123.5)
    Debug.Print CDate(True), CDate(False), CDate("500,99999") '. - ошибка
    Debug.Print CDate(-657434), CDate(2958465)
    Debug.Print CDate("0$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_CVar_01()
On Error GoTo ErrorHandler
    Dim varVar As Variant
    Debug.Print "---------------------------------"
    Debug.Print CVar(15), TypeName(CVar(15))
    Debug.Print CVar(varVar), CVar(1E-19)
    Debug.Print CVar(123.6789), CVar(123.1), CVar(1.2345678)
    Debug.Print CVar(-123.6789), CVar(-123.1), CVar(-123.5)
    Debug.Print CVar(True), CVar(False), CVar("500,99999") '. - ошибка
    Debug.Print CVar("0$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_CVar_02()
On Error GoTo ErrorHandler
    Dim varVar As Variant
    Debug.Print "---------------------------------"
    Debug.Print TypeName(CVar(CByte(15))), TypeName(CVar(15)), TypeName(CVar(15&))
    Debug.Print TypeName(CVar(15!)), TypeName(CVar(15#)), TypeName(CVar(15@))
    Debug.Print TypeName(CVar(CDec(15))), TypeName(CVar(True))
    Debug.Print TypeName(CVar("Строка")), TypeName(CVar(Now)), TypeName(CVar(varVar))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'w   1–7 (день недели, начиная с воскресенья = 1)
'ww  1–53 (неделя года без нуля в начале; неделя 1 начинается с 1 января)
'y   1–366 (день года)
Sub Functions_FormatDates_01()
On Error GoTo ErrorHandler
    Dim varDate As Date: varDate = #2/1/2025 12:34:56 PM#
    Debug.Print "---------------------------------"
    Debug.Print Format(Now), TypeName(Format(Now)) 'Variant (String)
    Debug.Print Format(varDate, "d dd w ww m mm mmm mmmm y yy yyyy")
    Debug.Print Format(varDate, "dd.mm.yyyy d\/m\/yy d mmm yyyy г. dd mmmm yyyy г.")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_FormatDates_02()
On Error GoTo ErrorHandler
    Debug.Print "---------------------------------"
    Debug.Print Format(Now, "ddd dddd, w день недели", vbMonday)
    Debug.Print Format(Now, "ww неделя года начиная с полной", , vbFirstFullWeek)
    Debug.Print Format(Now, "ww неделя года с 1 января", , vbFirstJan1)
    Debug.Print Format(Now, "y день года")
    Debug.Print Format(Now, "q квартал")
    Debug.Print CInt(Format(Now, "y"))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Functions_FormatTime()
On Error GoTo ErrorHandler
    Dim varDate As Date: varDate = #2/1/2025 12:34:56 PM#
    Debug.Print "---------------------------------"
    Debug.Print Format(varDate), TypeName(Format(varDate)) 'Variant (String)
    Debug.Print Format(varDate, "hh:nn:ss hh:mm:ss")
    Debug.Print Format(varDate, "m мин.")
    varDate = #2/1/2025 2:04:06 AM#
    Debug.Print Format(varDate, "h час. n мин. s сек. hh:nn:ss")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_NamedFormat()
On Error GoTo ErrorHandler
    Dim varDate As Date: varDate = #2/1/2025 12:34:56 PM#
    Debug.Print "---------------------------------"
    Debug.Print Format(10234567, "Standard")
    Debug.Print Format(0.1, "Fixed")
    Debug.Print Format(0.43, "Percent")
    Debug.Print Format(varDate, "Long date")
    Debug.Print Format(0.1, "General date")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Functions_FormatNumbers()
On Error GoTo ErrorHandler
    Debug.Print "---------------------------------"
    Debug.Print Format(1.1, "00000.00")
    Debug.Print Format(1234.123456, "00000.00")
    Debug.Print Format(1.125456, "#.##")
    Debug.Print Format(100.1, "#.000")
    Debug.Print Format(123, "+#,##0;($#,##0);Ноль")
    Debug.Print Format(-100, "+#,##0;-#,##0;Ноль")
    Debug.Print Format(0, "+#,##0;#,##0;Ноль")
    Debug.Print Format(-100, "+;-;0")
    Debug.Print Format(0, "+;-;0")
    Debug.Print Format(100, "+;-;0")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_AscChr()
On Error GoTo ErrorHandler
    Debug.Print "---------------------------------"
    Dim iCounter As Integer
    For iCounter = Asc("а") To Asc("я")
        Debug.Print Chr(iCounter)
    Next iCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Functions_FormatNumber()
On Error GoTo ErrorHandler
    Debug.Print "---------------------------------"
    Debug.Print FormatNumber(1234567.56789), FormatNumber(-1234567.56789),
    Debug.Print FormatNumber(0.56789)
    Debug.Print FormatNumber(-1234567.56789, 4)
    Debug.Print FormatNumber(0.56789, , vbTrue), FormatNumber(0.56789, , vbFalse)
    Debug.Print FormatNumber(-0.56789, , , vbFalse), FormatNumber(-0.56789, , , vbTrue)
    Debug.Print FormatNumber(-1234567.56789, , , , vbTrue), _
                                    FormatNumber(-1234567.56789, , , , vbFalse)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_FormatCurrency()
On Error GoTo ErrorHandler
    Debug.Print "---------------------------------"
    Debug.Print FormatCurrency(1234567.56789), FormatCurrency(-1234567.56789),
    Debug.Print FormatCurrency(0.56789)
    Debug.Print FormatCurrency(-1234567.56789, 4)
    Debug.Print FormatCurrency(0.56789, , vbTrue), FormatCurrency(0.56789, , vbFalse)
    Debug.Print FormatCurrency(-0.56789, , , vbFalse), FormatCurrency(-0.56789, , , vbTrue)
    Debug.Print FormatCurrency(-1234567.56789, , , , vbTrue), _
                                    FormatCurrency(-1234567.56789, , , , vbFalse)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_FormatPercent()
On Error GoTo ErrorHandler
    Debug.Print "---------------------------------"
    Debug.Print FormatPercent(0.56789123), FormatPercent(-0.56789123),
    Debug.Print FormatPercent(0.56789123)
    Debug.Print FormatPercent(0.56789123, 4)
    Debug.Print FormatPercent(0.56789123, , vbTrue), FormatPercent(0.56789123, , vbFalse)
    Debug.Print FormatPercent(-0.56789123, , , vbFalse), FormatPercent(-0.56789123, , , vbTrue)
    Debug.Print FormatPercent(-0.56789123, , , , vbTrue), _
                                    FormatPercent(-0.56789123, , , , vbFalse)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Functions_FormatDateTime()
On Error GoTo ErrorHandler
    Debug.Print "---------------------------------"
    Debug.Print FormatDateTime(Now), FormatDateTime(Now, vbGeneralDate)
    Debug.Print FormatDateTime(Now, vbLongDate), FormatDateTime(Now, vbShortDate)
    Debug.Print FormatDateTime(Now, vbLongTime), FormatDateTime(Now, vbShortTime)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_IsDate_01()
On Error GoTo ErrorHandler
    Dim dtDate As Date
    dtDate = "0"
    Debug.Print FormatDateTime(dtDate, vbLongDate)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Аргумент - variant, содержащий выражение даты или строковое выражение,
'распознаваемое как дата или время

Sub Functions_IsDate_02()
On Error GoTo ErrorHandler
    Debug.Print "---------------------------------"
    Debug.Print IsDate(Now), IsDate("1 февраля 2025"), IsDate("14.02.2025")
    Debug.Print IsDate("12:15:11"), IsDate("23:15:11"), IsDate("24:15:11")
    Debug.Print IsDate("12:15:11 PM"), IsDate("23:15:11 OE"), IsDate("2/22/2025")
    Debug.Print IsDate("12:15:11 PM"), IsDate("23:15:11 OE"), IsDate("2/22/2025")
    Debug.Print IsDate(""), IsDate("23:15:11 OE"), IsDate("2/22/2025")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'аргумент выражения — это Variant, содержащий числовое или строковое выражение.
Sub Functions_IsNumeric_01()
On Error GoTo ErrorHandler
    Dim dVar As Double: dVar = Now
    Debug.Print "---------------------------------"
    Debug.Print IsNumeric(1), IsNumeric("100500")
    Debug.Print IsNumeric("100000$"), IsNumeric("100 000")
    Debug.Print IsNumeric("1000,87"), IsNumeric("1000.87")
    Debug.Print IsNumeric(False), IsNumeric(Now)
    Debug.Print IsNumeric(CDbl(Now)), dVar
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_IsNumeric_02()
    On Error GoTo ErrorHandler
    Dim dNumber As Double
    Dim strInputValue As String
    Do
        strInputValue = InputBox("Введите число", "Окно ввода")
        If IsNumeric(strInputValue) Then Exit Do
    Loop
    Debug.Print CLng(strInputValue)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_IsObject()
    On Error GoTo ErrorHandler
    Dim oPerson As Person
    Dim dtVar As Date
    Debug.Print "---------------------------------"
    Debug.Print "oPerson", IsObject(oPerson)
    Debug.Print "dtVar", IsObject(dtVar)
    Debug.Print "Nothing", IsObject(Nothing)
    Debug.Print "Empty", IsObject(Empty)
    Debug.Print "Null", IsObject(Null)
    Debug.Print "Before new", TypeName(oPerson)
    Set oPerson = New Person
    Debug.Print "After new", TypeName(oPerson)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_IsArray()
    On Error GoTo ErrorHandler
    Dim aArray(10) As Date
    Dim dtVar As Date
    Debug.Print "---------------------------------"
    Debug.Print "aArray", IsArray(aArray)
    Debug.Print "dtVar", IsArray(dtVar)
    Debug.Print "aArray", TypeName(aArray)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_IsNull()
    On Error GoTo ErrorHandler
    Dim oPerson As New Person
    Dim dtVar As Date
    Dim strVar As String
    Dim iVar As Integer
    Dim varVar As Variant
    Debug.Print "---------------------------------"
    Debug.Print "oPerson", IsNull(oPerson)
    Debug.Print "dtVar", IsNull(dtVar)
    Debug.Print "strVar", IsNull(strVar)
    Debug.Print "iVar", IsNull(iVar)
    Debug.Print "varVar", IsNull(varVar)
    Debug.Print "Null", IsNull(Null)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Functions_IsEmpty()
    On Error GoTo ErrorHandler
    Dim oPerson As New Person
    Dim dtVar As Date
    Dim strVar As String
    Dim iVar As Integer
    Dim varVar As Variant
    Debug.Print "---------------------------------"
    Debug.Print "oPerson", IsEmpty(oPerson)
    Debug.Print "dtVar", IsEmpty(dtVar)
    Debug.Print "strVar", IsEmpty(strVar)
    Debug.Print "iVar", IsEmpty(iVar)
    Debug.Print "varVar", IsEmpty(varVar)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Functions_IsError()
    On Error GoTo ErrorHandler
    Dim varResult As Variant
    Debug.Print "---------------------------------"
    varResult = AddItems(10)
    Debug.Print "AddItems(10)", IsError(varResult)
    Debug.Print "varResult", varResult, TypeName(varResult)
    varResult = AddItems(-10)
    Debug.Print "AddItems(-10)", IsError(varResult)
    Debug.Print "varResult", varResult, TypeName(varResult)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Function AddItems(Qty As Long) As Variant
    If Qty < 0 Then AddItems = CVErr(50001) Else AddItems = 100 + Qty
End Function
