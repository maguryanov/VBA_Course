Attribute VB_Name = "m_01_19_Debug"
Option Explicit
Private mintVar As Integer
Private mdtVar As Date
Public gstrVar As String

'1.  �������������� ������. ����� Auto syntax check
Private Sub d01_DisplayArithmeticProgression()
    ' ���������� ����������
    'Dim dblFirstElement Double  ' ������ ������� ����������
End Sub
'2.  ������ ����������.
Private Sub d02_DisplayArithmeticProgression()
    ' ���������� ����������
    Dim FirstElement As String  ' ���� ������
    'Dim FirstElement As Double  ' ��������������� ��������
    ' ��������� ������ ��� ����������� ���������
    ' �������������� ��������� ����� � �����
    FirstElement = CDbl("100.5")
    'Step = CDbl("2")
End Sub

'3.  ������ ������� ����������
Private Sub d03_DisplayArithmeticProgression()
    ' ���������� ����������
    Dim dblFirstElement As Double  ' ������ ������� ����������
    Dim dblStep As Double          ' ��� ����������
    Dim strResult As String        ' ������ ��� ������������ ����������
    Dim intCounter As Integer      ' ������� �����
    ' ��������� ������ ��� ����������� ���������
    ' �������������� ��������� ����� � �����
    dblFirstElement = CDbl("1,5")
    dblStep = CDbl("2")
    ' ������������ ����������
    For intCounter = 0 To 2
        strResult = strResult & "������� " & (intCounter + 1) & ": " & _
                   Format(dblFirstElement + intCounter, "0.#####") & vbCrLf
    Next intCounter
    ' ����� ����������
    Debug.Print strResult
    Exit Sub
End Sub

'4.  ���������� ������
Private Sub d04_DisplayArithmeticProgression()
    On Error GoTo ErrorHandler
    Dim dblFirstElement As Double, dblStep As Double
    Dim strFirstElement As String, strStep As String
    Dim strResult As String, intCounter As Integer
    ' ��������� ������ ��� ����������� ���������
    strFirstElement = "2": strStep = "0,5"
    ' �������� �� ���������� ����
    If IsNumeric(strFirstElement) And IsNumeric(strStep) Then
        dblFirstElement = strFirstElement: dblStep = strStep
        Else: MsgBox "���������� ������ �������� ��������!", vbCritical: Exit Sub
    End If
    For intCounter = 0 To 2
        strResult = strResult & "������� " & (intCounter + 1) & ": " & _
                   Format(dblFirstElement + intCounter, "0.#####") & vbCrLf
    Next intCounter
    Debug.Print strResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Immediate Window (Ctrl+G)
Private Sub d05_ImmediateWindow()
    Debug.Print "������������ ������ � Immediate Window"
'print 2 * 2
'? 2 * 2
'? Choose(10, 1, 2)
'? strVar
'dim intVar as Integer '������
'intVar  =  100
'? intVar
'MkDir "d:\Vba\FromImmediateWindow"
'Help �� F1
End Sub



' ����� ���������� � ����� ��������
Private Sub d06_Breakpoint()
' �������� ����� Immediate window, Watch, Tips, locals
' ��������� Auto Dat� Tips
' ��������� �������� ����������
    On Error GoTo ErrorHandler
    Dim intCounter As Integer, intResult As Integer
    Static intAttempt As Integer: intAttempt = intAttempt + 1
    For intCounter = 1 To 10
        intResult = intResult + intCounter
    Next intCounter
    Debug.Print "�������� ���������� #" & intAttempt
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


' Stop ����������� ���������
Private Sub d07_Stop()
    On Error GoTo ErrorHandler
    Dim intCounter As Integer, intResult As Integer
    For intCounter = 1 To 10
        intResult = intResult + intCounter
        Stop '����������� ���������
    Next intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


' Stop ��������� �� �������
Private Sub d08_Stop()
    On Error GoTo ErrorHandler
    Dim intCounter As Integer, intResult As Integer
    For intCounter = 1 To 10
        If intCounter = 5 Then Stop '��������� �� �������
        intResult = intResult + intCounter
    Next intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

' StepInto - F8
Private Sub d09_StepInto()
    On Error GoTo ErrorHandler
    Dim intCounter As Integer, intResult As Integer
    Dim intDivider As Integer: intDivider = 0
    For intCounter = 1 To 10
        intResult = intResult + intCounter / intDivider
    Next intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'���� ������� ������
Private Sub d10_DisplayArithmeticProgression_Demo()
    On Error GoTo ErrorHandler
    Dim dblFirstElement As Double, dblStep As Double
    Dim strFirstElement As String, strStep As String
    Dim strResult As String, intCounter As Integer
    ' ��������� ������ ��� ����������� ���������
    strFirstElement = "2": strStep = "0,5"
    ' �������� �� ���������� ����
    If IsNumeric(strFirstElement) And IsNumeric(strStep) Then
        dblFirstElement = strFirstElement: dblStep = strStep
        Else: MsgBox "���������� ������ �������� ��������!", vbCritical: Exit Sub
    End If
    For intCounter = 0 To 2
        strResult = strResult & "������� " & (intCounter + 1) & ": " & _
                   Format(dblFirstElement + intCounter, "0.#####") & vbCrLf
    Next intCounter
    Debug.Print strResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Step Over - ���������� ��������� (Shift+F8)
Private Sub d11_DisplayArithmeticProgression_Step_Over()
    On Error GoTo ErrorHandler
    Dim dblFirstElement As Double, dblStep As Double
    Dim strFirstElement As String, strStep As String
    Dim strResult As String, intCounter As Integer
    ' ��������� ������ ��� ����������� ���������
    strFirstElement = "2": strStep = "0,5"
    ' �������� �� ���������� ���� ��� ����������� ���������
    dblFirstElement = strFirstElement: dblStep = strStep
    For intCounter = 0 To 2
        strResult = strResult & "������� " & (intCounter + 1) & ": " & _
                   Format(d11_ElementN(dblFirstElement, intCounter, dblStep), "0.00") & vbCrLf
    Next intCounter
    Debug.Print strResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
'������� ���������� N �������� �������������� ������������������
Private Function d11_ElementN _
    (ByVal FirstElement As Double, ByVal Number As Long, ByVal Step As Double) As Double
    d11_ElementN = FirstElement + (Number - 1) * Step
End Function


'Step Out (Ctrl+Shift+F8)
Private Sub d12_DisplayArithmeticProgression_Step_Out()
    On Error GoTo ErrorHandler
    Dim dblFirstElement As Double, dblStep As Double
    Dim strFirstElement As String, strStep As String
    ' ��������� ������ ��� ����������� ���������
    strFirstElement = "2": strStep = "0,5"
    ' �������� �� ���������� ���� ��� ����������� ���������
    dblFirstElement = CDbl(strFirstElement)
    dblStep = CDbl(strStep)
    Debug.Print d12_ArithmeticProgressionText(dblFirstElement, dblStep)
    Debug.Print "�������� ����������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������� ��������� ������ ��� ������ �������������� ������������������
Private Function d12_ArithmeticProgressionText _
    (ByVal FirstElement As Double, ByVal Step As Double) As String
    Dim intCounter As Integer
    Dim strResult As String
    For intCounter = 1 To 3
        strResult = strResult & "������� " & intCounter & ": " & _
                   Format(d11_ElementN(FirstElement, intCounter, Step), "0.00") & vbCrLf
    Next intCounter
    d12_ArithmeticProgressionText = strResult
End Function

'Call Stack-���� �������. ����� ���� (Ctrl+L), Locals
Sub d13_CallStack()
    d12_DisplayArithmeticProgression_Step_Out
End Sub

'Run To Cursor (Ctrl+F8) ������� � ������� �������
Private Sub d14_RunToCursor()
    On Error GoTo ErrorHandler
    Dim dblFirstElement As Double, dblStep As Double
    Dim strFirstElement As String, strStep As String
    Dim strResult As String, intCounter As Integer
    ' ��������� ������ ��� ����������� ���������
    strFirstElement = "2": strStep = "0,5"
    ' �������� �� ���������� ����
    If IsNumeric(strFirstElement) And IsNumeric(strStep) Then
        dblFirstElement = strFirstElement: dblStep = strStep
        Else: MsgBox "���������� ������ �������� ��������!", vbCritical: Exit Sub
    End If
    For intCounter = 0 To 2
        strResult = strResult & "������� " & (intCounter + 1) & ": " & _
                   Format(dblFirstElement + intCounter, "0.#####") & vbCrLf
    Next intCounter
    Debug.Print strResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'��������� �� ���������. "�������������� ���������"
Private Sub d15_ExecutionBySteps()
    Debug.Print "��� ������"
    Debug.Print "��� ������"
    Debug.Print "��� ������"
    Debug.Print "��� ��������"
    Debug.Print "��� �����"
    Debug.Print "��� ������"
    Debug.Print "��� �������"
End Sub


'������� ����� ����. ����������� � Set Next Statement
Private Sub d17_SetNextStatement()
    Debug.Print "��� ������"
    Debug.Print "��� ������"
    Debug.Print "��� ������"
    Debug.Print "��� ��������"
    Debug.Print "��� �����"
    Debug.Print "��� ������"
    Debug.Print "��� �������"
End Sub


'������� ����� ����. Show Next Statement
Private Sub d17_ShowNextStatement()
    Debug.Print "������������ ������ Show Next Statement"
End Sub


'Watch Window - ���� ����������� ��������. Quick Watch (Shift+F9)
Private Sub d18_QuickWatch()
    On Error GoTo ErrorHandler
    Dim dblFirstElement As Double, dblStep As Double
    Dim strFirstElement As String, strStep As String
    Dim strResult As String, intCounter As Integer
    ' ��������� ������ � �������� �� ���������� ���� ��� ����������� ���������
    dblFirstElement = 2: dblStep = 0.5
    For intCounter = 0 To 2
        strResult = strResult & "������� " & (intCounter + 1) & ": " & _
                   Format(dblFirstElement + intCounter * dblStep, "0.00") & vbCrLf
    Next intCounter
    Debug.Print strResult
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Watch Window. ����������� ���������
Private Sub d19_WatchContext()
    Dim strVar As String
    strVar = "One"
    Debug.Print "������������ ���������"
'    ���������� ����������� � ������
'    Private mintVar As Integer
'    Private mdtVar As Date
'    Public gstrVar As String
End Sub

'Watch Window. ����� �������� ���������
Private Sub d20_ExpressionWatch()
    On Error GoTo ErrorHandler
    Dim intCounter As Integer, intResult As Integer
    For intCounter = 1 To 10
        intResult = intResult + intCounter
    Next intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Watch Window. ��������� �� �������
Private Sub d21_ConditionalWatch()
    On Error GoTo ErrorHandler
    Dim intCounter As Integer, intResult As Integer
    For intCounter = 1 To 10
        intResult = intResult + intCounter
    Next intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Watch Window. ��������� �� ��������� ��������
Private Sub d22_ChangeWatch()
    On Error GoTo ErrorHandler
    Dim intCounter As Integer, intResult As Integer
    For intCounter = 1 To 10
        intResult = intResult + intCounter
    Next intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Debug.Assert
Private Sub d24_Debug_Assert()
    On Error GoTo ErrorHandler
    
    Dim lngQty As Long '�������� ���������� �� ������ ���� �������������
    Dim intCounter As Integer
    '���������� �������� ��� ���������� ����� ����� �������������
    For intCounter = 100 To -100 Step -1
        lngQty = intCounter
        Debug.Assert lngQty >= 0
    Next intCounter

ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Debug.Print
Private Sub d24_Debug_Print()
    On Error GoTo ErrorHandler
    
    Debug.Print "����", "������ "; "Debug.Print",
    Debug.Print "�����������";
    Debug.Print " ������"
    Debug.Print "������������"; Spc(5); "������������� Spc"
    Debug.Print "1234567890"; Tab(12); "������������� Tab"

ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'����������� ������� ��� ������������� ��������� ������
Private Sub d25_ErrorHandling()
    On Error GoTo ErrorHandler
    Dim intResult As Integer
    Err.Raise 50000, , "�������� ������"
    intResult = intResult + 10000
    intResult = intResult + 10000
    intResult = intResult + 10000
    intResult = intResult + 10000
    intResult = intResult + 10000
        
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'����� Error Trapping. Break In �lass Module
Sub d26_BreakIn�lassModule()
    On Error GoTo ErrorHandler
    Dim varResult As Variant
    Dim oVideoLesson As New VideoLesson
    'varResult = oVideoLesson.ErrorHandlingDemo(0)
    varResult = oVideoLesson.ErrorHandlingRaiseDemo("100$")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub





