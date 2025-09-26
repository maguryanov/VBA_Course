Attribute VB_Name = "m_02_03_Strings"
Option Explicit


Sub Strings_Chr()
On Error GoTo ErrorHandler
    Dim strResult As String
    strResult = "Александр Иванов" & Chr(13) & Chr(10) & "Андрей Петров"
    Debug.Print strResult
    strResult = "Александр Иванов" & vbCrLf & "Андрей Петров"
    Debug.Print strResult
    strResult = "Александр Иванов" & Chr(9) & "Андрей Петров"
    Debug.Print strResult
    strResult = "Александр Иванов" & vbTab & "Андрей Петров"
    Debug.Print strResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



Sub Strings_Asc_Chr()
On Error GoTo ErrorHandler
    Dim iCounter As Integer
    Dim strResult As String
    For iCounter = Asc("а") To Asc("я")
        strResult = strResult + Chr(iCounter)
    Next iCounter
    For iCounter = Asc("А") To Asc("Я")
        strResult = strResult + Chr(iCounter)
    Next iCounter
    Debug.Print strResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Strings_Case()
On Error GoTo ErrorHandler
    Dim strFullName As String
    strFullName = "оЛьгА пеТроВа"
    Debug.Print UCase(strFullName)
    Debug.Print LCase(strFullName)
    Debug.Print StrConv(strFullName, vbProperCase)
    Debug.Print StrConv(strFullName, vbUpperCase)
    Debug.Print StrConv(strFullName, vbLowerCase)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Strings_Trim()
On Error GoTo ErrorHandler
    Dim strFullName As String
    strFullName = "  Марина Иванова  "
    Debug.Print "|" & RTrim(strFullName) & "|"
    Debug.Print "|" & LTrim(strFullName) & "|"
    Debug.Print "|" & Trim(strFullName) & "|"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Strings_Space_String()
On Error GoTo ErrorHandler
    Dim strFullName As String
    Debug.Print "|" & Space(5) & "|"
    Debug.Print "|" & String(5, "*") & "|"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Strings_Space_InStr()
On Error GoTo ErrorHandler
    Dim strVar As String
    strVar = "12345678901234567890"
    Debug.Print InStr(1, strVar, "45")
    Debug.Print InStr(5, strVar, "45")
    strVar = "Microsoft SQL Server"
    Debug.Print InStr(1, strVar, "sql")
    Debug.Print InStr(1, strVar, "sql", vbTextCompare)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Strings_Space_InStrRev()
On Error GoTo ErrorHandler
    Dim strVar As String
    strVar = "12345678901234567890"
    Debug.Print InStrRev(strVar, "45")
    Debug.Print InStrRev(strVar, "45", 14)
    strVar = "Microsoft SQL Server"
    Debug.Print InStrRev(strVar, "sql")
    Debug.Print InStrRev(strVar, "sql", , vbTextCompare)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Strings_Space_Len()
On Error GoTo ErrorHandler
    Dim strVar As String
    strVar = "12345678901234567890"
    Debug.Print Len(strVar)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Strings_Left_Right()
On Error GoTo ErrorHandler
    Dim strVar As String
    strVar = "12345678901234567890"
    Debug.Print Left(strVar, 5)
    Debug.Print Right(strVar, 5)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Strings_Mid()
On Error GoTo ErrorHandler
    Dim strVar As String
    strVar = "12345678901234567890"
    Debug.Print Mid(strVar, 5)
    Debug.Print Mid(strVar, 5, 4)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Strings_ParseFullName()
On Error GoTo ErrorHandler
    Dim strFullName As String
    Dim strFirstName As String
    Dim strLastName As String
    strFullName = "Марина Иванова"
    strFirstName = Left(strFullName, InStr(strFullName, " "))
    strLastName = Mid(strFullName, InStr(strFullName, " ") + 1)
    Debug.Print strFirstName, strLastName
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Strings_Replace()
On Error GoTo ErrorHandler
    Dim strFullName As String
    strFullName = "Марина Иванова"
    strFullName = Replace(strFullName, "иванова", "Петрова", , , vbTextCompare)
    Debug.Print strFullName
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Strings_DeleteSpaces()
On Error GoTo ErrorHandler
    Dim strVar As String
    strVar = "В          синтаксисе    функции    InStr   используются следующие аргументы"
    Do While InStr(strVar, "  ") > 0
        strVar = Replace(strVar, "  ", " ")
    Loop
    Debug.Print strVar
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Strings_StrComp()
On Error GoTo ErrorHandler
    Debug.Print StrComp("A", "B")
    Debug.Print StrComp("B", "A")
    Debug.Print StrComp("A", "A")
    Debug.Print StrComp("A", "a")
    Debug.Print StrComp("A", "a", vbTextCompare)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Strings_PlasticCard()
On Error GoTo ErrorHandler
    Dim strCardNumber As String
    Dim strCleaned As String
    strCardNumber = "1234 5678 9012 3456"
    strCleaned = Replace(strCardNumber, " ", "")
    strCleaned = Left(strCleaned, 6) & "******" & Mid(strCleaned, 13)
    Debug.Print strCleaned
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
'Операторы
Sub Strings_LSet()
On Error GoTo ErrorHandler
    Dim strVar As String
    strVar = "abcde"
    LSet strVar = "1234567890"
    Debug.Print "|" & strVar & "|"
    strVar = "abcde"
    LSet strVar = "123"
    Debug.Print "|" & strVar & "|"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Strings_RSet()
On Error GoTo ErrorHandler
    Dim strVar As String
    strVar = "abcde"
    RSet strVar = "1234567890"
    Debug.Print "|" & strVar & "|"
    strVar = "abcde"
    RSet strVar = "123"
    Debug.Print "|" & strVar & "|"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

