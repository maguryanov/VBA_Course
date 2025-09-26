Attribute VB_Name = "m_01_09_FlowControl"
Option Explicit

Sub IfConstruction()
    On Error GoTo ErrorHandler
    Debug.Print "������� ������"
    'Debug.Print 1 / 0
    Debug.Print "����������� ������� ������"
Finalization:
    Debug.Print "���������� ������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

Sub SingleLineIfConstraction()
    On Error GoTo ErrorHandler
    Dim strUserName As String
    'strUserName = "������"
    Dim strGreeting As String
    strGreeting = "������ ����"
    If strUserName > "" Then strGreeting = strGreeting & ", " & strUserName
    strGreeting = strGreeting & "!"
    Debug.Print strGreeting
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub MultiLinesIfConstruction()
    On Error GoTo ErrorHandler
    Dim strUserName As String
    'strUserName = "������"
    Dim strGreeting As String
    strGreeting = "������ ����"
    If strUserName > "" Then
        strGreeting = strGreeting & ", " & strUserName
    End If
    strGreeting = strGreeting & "!"
    Debug.Print strGreeting
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub IfElseConstruction_01()
    On Error GoTo ErrorHandler
    If True Then
        Debug.Print "���������� �� True"
    Else
        Debug.Print "���������� �� False"
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



Sub MultiLinesIfElseConstruction_01()
    On Error GoTo ErrorHandler
    Dim strGender As String
    Dim strUserName As String
    Dim strGreeting As String
    strGender = "f"
    strUserName = "������ ����������"
    strUserName = "������ ���������"
    If strGender = "f" Then
        strGreeting = "��������� " & strUserName
    Else
        strGreeting = "��������� " & strUserName
    End If
    strGreeting = strGreeting & "!"
    Debug.Print strGreeting
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub MultiLinesIfElseConstruction_02()
    On Error GoTo ErrorHandler
    Dim curBalance As Currency
    Dim strResult As String
    curBalance = -100.2
    
    If curBalance > 0 Then
        strResult = "�������������"
    Else
        strResult = "������� ��� �������������"
    End If
    Debug.Print strResult
    
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub SingleLineIfElseConstruction_01()
    On Error GoTo ErrorHandler
    Dim bResult As Byte
    If False Then Debug.Print bResult = 100 Else bResult = 200
    Debug.Print bResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub SeveralIfElseConstruction()
    On Error GoTo ErrorHandler
    Dim curPrice As Currency
    Dim strPriceRange As String
    curPrice = 10000
    If curPrice < 1000 Then
        strPriceRange = "������"
    ElseIf curPrice < 2000 Then
        strPriceRange = "�������"
    Else
        strPriceRange = "�������"
    End If
    Debug.Print strPriceRange
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub SelectConstruction_01()
    On Error GoTo ErrorHandler
    Dim bPriceRangeNumber As String '1,2,3
    Dim strPriceRange As String
    bPriceRangeNumber = 5
    If bPriceRangeNumber = 1 Then
        strPriceRange = "������"
    ElseIf bPriceRangeNumber = 2 Then
        strPriceRange = "�������"
    ElseIf bPriceRangeNumber = 3 Then
        strPriceRange = "�������"
    Else
        strPriceRange = "�������� ��������"
    End If
    Debug.Print strPriceRange
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub SelectConstruction_02()
    On Error GoTo ErrorHandler
    Dim bPriceRangeNumber As String '1,2,3
    Dim strPriceRange As String
    bPriceRangeNumber = 3
    Select Case bPriceRangeNumber
    Case 1
        strPriceRange = "������"
    Case 2
        strPriceRange = "�������"
    Case 3
        strPriceRange = "�������"
    Case Else
        strPriceRange = "�������� ��������"
    End Select
    Debug.Print strPriceRange
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub SelectConstruction_03()
    On Error GoTo ErrorHandler
    Dim bDayNumber As Byte
    bDayNumber = 1
    Select Case bDayNumber
    Case 1 To 5
        Debug.Print "�������� �� �������� ���"
    Case 6 To 7
        Debug.Print "�������� �� ��������� ���"
    End Select
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub SelectConstruction_04()
    On Error GoTo ErrorHandler
    Dim curPrice As Currency
    Dim strPriceRange As String
    curPrice = 100@
    
    Select Case curPrice
        Case Is < 1000
            strPriceRange = "������"
        Case Is < 2000
            strPriceRange = "�������"
        Case Else
            strPriceRange = "�������"
    End Select
    
    Debug.Print strPriceRange
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
