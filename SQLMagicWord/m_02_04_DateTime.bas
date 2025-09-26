Attribute VB_Name = "m_02_04_DateTime"
Option Explicit

Public Const mcTestDate As Date = #1/31/2025 12:34:56 PM#
Public Const mcDateTimeFormat As String = "dd\.mm\.yyyy hh\:nn\:ss, Dddd"

Sub DateTime_Date_Time_Now()
    On Error GoTo ErrorHandler
    Debug.Print Date, Time, Now
    Debug.Print "------------------------------"
    Debug.Print Format(Date, mcDateTimeFormat)
    Debug.Print Format(Time, mcDateTimeFormat)
    Debug.Print Format(Now, mcDateTimeFormat)
    Debug.Print Format(CDate(0), mcDateTimeFormat)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub DateTime_DateAdd()
    On Error GoTo ErrorHandler
    Debug.Print "------------------------------"
    Debug.Print mcTestDate
    Debug.Print "yyyy", DateAdd("yyyy", 1, mcTestDate)
    Debug.Print "m", DateAdd("m", 1, mcTestDate)
    Debug.Print "ww", DateAdd("ww", 1, mcTestDate)
    Debug.Print "d", DateAdd("d", 1, mcTestDate)
    Debug.Print "h", DateAdd("h", 1, mcTestDate)
    Debug.Print "n", DateAdd("n", 1, mcTestDate)
    Debug.Print "s", DateAdd("s", 1, mcTestDate)
    Debug.Print "q", DateAdd("q", 1, mcTestDate)
    Debug.Print "-yyyy", DateAdd("yyyy", -1, mcTestDate)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub DateTime_DateDiff()
    On Error GoTo ErrorHandler
    Debug.Print "------------------------------"
    Debug.Print mcTestDate
    Debug.Print "yyyy", DateDiff("yyyy", mcTestDate, Now)
    Debug.Print "m", DateDiff("m", mcTestDate, Now)
    Debug.Print "ww", DateDiff("ww", mcTestDate, Now)
    Debug.Print "d", DateDiff("d", mcTestDate, Now)
    Debug.Print "h", DateDiff("h", mcTestDate, Now)
    Debug.Print "n", DateDiff("n", mcTestDate, Now)
    Debug.Print "s", DateDiff("s", mcTestDate, Now)
    Debug.Print "q", DateDiff("q", mcTestDate, Now)
    Debug.Print "ww", DateDiff("ww", mcTestDate, Now, vbMonday, vbFirstJan1)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub DateTime_DatePart()
    On Error GoTo ErrorHandler
    Debug.Print "------------------------------"
    Debug.Print "yyyy", DatePart("yyyy", mcTestDate)
    Debug.Print "m", DatePart("m", mcTestDate)
    Debug.Print "ww", DatePart("ww", mcTestDate)
    Debug.Print "d", DatePart("d", mcTestDate)
    Debug.Print "h", DatePart("h", mcTestDate)
    Debug.Print "n", DatePart("n", mcTestDate)
    Debug.Print "s", DatePart("s", mcTestDate)
    Debug.Print "q", DatePart("q", mcTestDate)
    Debug.Print "ww", DatePart("ww", mcTestDate, vbMonday, vbFirstFullWeek)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub DateTime_DateParts()
    On Error GoTo ErrorHandler
    Debug.Print "------------------------------"
    Debug.Print Format(mcTestDate, mcDateTimeFormat)
    Debug.Print "Day", Day(mcTestDate)
    Debug.Print "Month", Month(mcTestDate)
    Debug.Print "Year", Year(mcTestDate)
    Debug.Print "Hour", Hour(mcTestDate)
    Debug.Print "Minute", Minute(mcTestDate)
    Debug.Print "Second", Second(mcTestDate)
    Debug.Print "Weekday", Weekday(mcTestDate)
    Debug.Print "Weekday", Weekday(mcTestDate, vbMonday)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



Sub DateTime_Timer_01()
    On Error GoTo ErrorHandler
    Debug.Print "------------------------------"
    Debug.Print "Timer", Timer 'количество секунд, прошедших после полночи
    Randomize Timer
    Debug.Print "Rnd()", Rnd()
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub DateTime_Timer()
    On Error GoTo ErrorHandler
    Debug.Print "------------------------------"
    Debug.Print "Timer", Timer 'количество секунд, прошедших после полночи
    Randomize Timer
    Debug.Print "Rnd()", Rnd()
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub DateTime_Serial()
    On Error GoTo ErrorHandler
    Debug.Print "------------------------------"
    Debug.Print TimeSerial(12, 34, 56), TypeName(TimeSerial(12, 34, 56))
    Debug.Print DateSerial(2025, 1, 14), TypeName(DateSerial(2025, 1, 14))
    Debug.Print DateSerial(2025, 1, 14) + TimeSerial(12, 34, 56)
    Debug.Print DateSerial(99, 1, 14), DateSerial(25, 1, 14)
    Debug.Print DateSerial(49, 1, 14), DateSerial(50, 1, 14)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub DateTime_01()
    '#1/31/2025 12:34:56 PM#
    On Error GoTo ErrorHandler
    Debug.Print "------------------------------"
    If mcTestDate = #1/31/2025# Then Debug.Print "Да" Else Debug.Print "Нет"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub DateTime_02()
    On Error GoTo ErrorHandler
    Debug.Print "------------------------------"
    If DateValue(mcTestDate) = #1/31/2025# Then Debug.Print "Да" Else Debug.Print "Нет"
    If Fix(mcTestDate) = #1/31/2025# Then Debug.Print "Да" Else Debug.Print "Нет"
    If Int(mcTestDate) = #1/31/2025# Then Debug.Print "Да" Else Debug.Print "Нет"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub DateTime_03()
    On Error GoTo ErrorHandler
    Debug.Print "------------------------------"
    If Fix(#1/1/1861 12:34:56 PM#) = Fix(#1/1/1861 11:34:00 PM#) Then
        Debug.Print "Да"
    Else
        Debug.Print "Нет"
    End If
    If Int(#1/1/1861 12:34:56 PM#) = Int(#1/1/1861 11:34:00 PM#) Then
        Debug.Print "Да"
    Else
        Debug.Print "Нет"
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub DateTime_04()
    On Error GoTo ErrorHandler
    Debug.Print "------------------------------"
    Debug.Print Format(DateValue(mcTestDate), mcDateTimeFormat)
    Debug.Print Format(Int(mcTestDate), mcDateTimeFormat)
    Debug.Print Format(Fix(mcTestDate), mcDateTimeFormat)
    Debug.Print Format(DateValue(#1/1/1789 12:34:56 PM#), mcDateTimeFormat)
    Debug.Print Format(Int(#1/1/1789 12:34:56 PM#), mcDateTimeFormat)
    Debug.Print Format(Fix(#1/1/1789 12:34:56 PM#), mcDateTimeFormat)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub DateTime_DateValue()
    Dim sFirstTime As Single, dtResult As Date
    sFirstTime = Timer
    Dim iCounter As Long
    For iCounter = 0 To 10000000
         dtResult = DateValue(mcTestDate)
    Next iCounter
    Debug.Print Timer - sFirstTime
End Sub


Sub DateTime_Int()
    Dim sFirstTime As Single, dtResult As Date
    sFirstTime = Timer
    Dim iCounter As Long
    For iCounter = 0 To 10000000
         dtResult = Int(mcTestDate)
    Next iCounter
    Debug.Print Timer - sFirstTime
End Sub

Sub DateTime_Fix()
    Dim sFirstTime As Single, dtResult As Date
    sFirstTime = Timer
    Dim iCounter As Long
    For iCounter = 0 To 10000000
         dtResult = Fix(mcTestDate)
    Next iCounter
    Debug.Print Timer - sFirstTime
End Sub

Sub DateTime_DateValue_02()
    On Error GoTo ErrorHandler
    Debug.Print "------------------------------"
    Debug.Print Format(DateValue(Now), mcDateTimeFormat)
    Debug.Print Format(DateValue("31.01.2024"), mcDateTimeFormat)
    Debug.Print Format(DateValue("31.01.2024 12:34:56"), mcDateTimeFormat)
    Debug.Print Format(DateValue("31 января 2024 12:34:56"), mcDateTimeFormat)
    Debug.Print Format(DateValue("31 январ 2024 12:34:56"), mcDateTimeFormat)
    Debug.Print Format(DateValue("31 январ. 2024 12:34:56"), mcDateTimeFormat)
    Debug.Print Format(DateValue("31 января"), mcDateTimeFormat)
    Debug.Print Format(DateValue("31 января 12:34:61"), mcDateTimeFormat)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub DateTime_TimeValue()
    On Error GoTo ErrorHandler
    Debug.Print "------------------------------"
    Debug.Print Format(TimeValue(Now), mcDateTimeFormat)
    Debug.Print Format(TimeValue("31.01.2024"), mcDateTimeFormat)
    Debug.Print Format(TimeValue("31.01.2024 12:34:56"), mcDateTimeFormat)
    Debug.Print Format(TimeValue("31 января 2024 12:34:56"), mcDateTimeFormat)
    Debug.Print Format(TimeValue("31 январ 2024 12:34:56"), mcDateTimeFormat)
    Debug.Print Format(TimeValue("31 январ. 2024 12:34:56"), mcDateTimeFormat)
    Debug.Print Format(TimeValue("31 января"), mcDateTimeFormat)
    Debug.Print Format(TimeValue("31 января 12:34:61"), mcDateTimeFormat)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
