Attribute VB_Name = "m_01_10_Loops"
Option Explicit

Sub LoopFor_01()
    On Error GoTo ErrorHandler
    Dim iCounter As Integer
    For iCounter = 1 To 5
        Debug.Print iCounter
    Next iCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub

Sub LoopFor_02()
    On Error GoTo ErrorHandler
    Dim iCounter As Integer
    For iCounter = 1 To 10 Step 2
        Debug.Print iCounter
    Next iCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub

Sub LoopFor_03()
    On Error GoTo ErrorHandler
    Dim iCounter As Integer
    For iCounter = 5 To 1 Step -1
        Debug.Print iCounter
    Next iCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub

Sub LoopFor_04()
    On Error GoTo ErrorHandler
    Dim iCounter As Integer
    For iCounter = 5 To 1 Step -1
        Debug.Print iCounter
    Next iCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub

Sub LoopFor_05()
    On Error GoTo ErrorHandler
    Dim iCharQry As Integer: iCharQry = 10
    Dim strResult
    Dim iCounter As Integer
    For iCounter = 1 To iCharQry
        strResult = strResult & "*"
        'Debug.Print strResult
    Next iCounter
    Debug.Print "strResult = "; strResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub


Sub LoopFor_Single_Step_06()
    On Error GoTo ErrorHandler
    Dim sXValue As Single
    For sXValue = 0! To 10! Step 0.1!
        Debug.Print sXValue
    Next sXValue
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub

Sub LoopFor_Currency_Step_07()
    On Error GoTo ErrorHandler
    Dim curXValue As Currency
    For curXValue = 0 To 10 Step 0.1
        Debug.Print curXValue
    Next curXValue
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub


Sub LoopFor_08()
    On Error GoTo ErrorHandler
    Dim iArray(4, 4) As Integer
    Dim iCounterY As Integer
    Dim iCounterX As Integer
    
    For iCounterY = 0 To 4
        For iCounterX = 0 To 4
            iArray(iCounterX, iCounterY) = 5
        Next iCounterX
    Next iCounterY
    
    Debug.Print iArray(4, 4)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub


Sub LoopFor_09()
    On Error GoTo ErrorHandler
    Dim iArray(4, 4) As Integer
    Dim iCounterY As Integer
    Dim iCounterX As Integer
    For iCounterY = 0 To 4
        For iCounterX = 0 To 4
            iArray(iCounterX, iCounterY) = 5
            Debug.Print "iArray(" & iCounterX & "," & iCounterY & ") = " & _
                iArray(iCounterX, iCounterY)
        Next iCounterX
    Next iCounterY
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub



Sub LoopFor_10()
    On Error GoTo ErrorHandler
    Dim iCounterY As Integer
    Dim iCounterX As Integer
    Dim strResult As String
    For iCounterY = 1 To 10
        strResult = ""
        For iCounterX = 1 To 10
            strResult = strResult & iCounterX * iCounterY & vbTab
        Next iCounterX
        Debug.Print strResult
    Next iCounterY
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub


Sub LoopFor_ExitFor11()
    On Error GoTo ErrorHandler
    Dim strArray(99) As String
    Dim strNeededValue As String
    Dim iResult As Integer
    Dim iCounter
    strNeededValue = "SQL"
    strArray(2) = "SQL"
    For iCounter = 0 To 99
        If strArray(iCounter) = strNeededValue Then
            iResult = iCounter
            Exit For
        End If
    Next iCounter
    Debug.Print iResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub

Sub InfiniteLoop_01()
    On Error GoTo ErrorHandler
    Dim iNumber As Integer
    Do
        iNumber = iNumber + 1
    Loop
    Exit Sub
ErrorHandler:
    Debug.Print "Максимальное iNumber = "; iNumber
    Debug.Print Err.Number & " / "; Err.Description
End Sub


Sub InfiniteLoop_02()
    On Error GoTo ErrorHandler
    Dim dNumber As Double
    Dim strInputValue As String
    Do
        strInputValue = InputBox("Введите число", "Окно ввода")
        If IsNumeric(strInputValue) Then Exit Do
    Loop
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub




Sub WhileLoop_03()
    On Error GoTo ErrorHandler
    Dim iLinesTotal As Integer: iLinesTotal = 5
    Dim iLinesOnPage As Integer: iLinesOnPage = 20
    Dim iLinesQty As Integer: iLinesQty = iLinesTotal
    Do While iLinesQty > iLinesOnPage
        Debug.Print "Печатаем " & iLinesOnPage & " строк"
        iLinesQty = iLinesQty - iLinesOnPage
    Loop
    Debug.Print "Печатаем " & iLinesQty & " строк"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub


Sub DoWhileLoop_01()
    On Error GoTo ErrorHandler
    Dim dNumber As Double
    Dim strInputValue As String
    Do
        strInputValue = InputBox("Введите число", "Окно ввода")
        If IsNumeric(strInputValue) Then Exit Do
    Loop While Not IsNumeric(strInputValue)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub DoUntilLoop_02()
    On Error GoTo ErrorHandler
    Dim dNumber As Double
    Dim strInputValue As String
    Do
        strInputValue = InputBox("Введите число", "Окно ввода")
        If IsNumeric(strInputValue) Then Exit Do
    Loop Until IsNumeric(strInputValue)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Loops_01()
    On Error GoTo ErrorHandler
    Dim iCounter As Integer
    iCounter = 1
    Do While (iCounter <= 5)
        Debug.Print iCounter
        iCounter = iCounter + 1
    Loop
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub Loops_02()
    On Error GoTo ErrorHandler
    Dim iCounter As Integer
    iCounter = 1
    Do Until (iCounter > 5)
        Debug.Print iCounter
        iCounter = iCounter + 1
    Loop
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Loops_03()
    On Error GoTo ErrorHandler
    Dim iCounter As Integer
    iCounter = 1
    Do
        Debug.Print iCounter
        iCounter = iCounter + 1
    Loop While (iCounter <= 5)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub Loops_04()
    On Error GoTo ErrorHandler
    Dim iCounter As Integer
    iCounter = 1
    Do
        Debug.Print iCounter
        iCounter = iCounter + 1
    Loop Until (iCounter > 5)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub ForEachLoop_01()
    On Error GoTo ErrorHandler
    Dim oStyle As Style
    Dim oDoc As Document
    Set oDoc = ActiveDocument
    Dim iCounter As Integer
    For Each oStyle In oDoc.Styles
        If Not oStyle.BuiltIn Then
            Debug.Print oStyle.NameLocal
        End If
    Next oStyle
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
