Attribute VB_Name = "m_01_15_Files"
Option Explicit

'������ ��������� ������. ����������
Sub File_Append_01()
    On Error GoTo ErrorHandler
    Static intAttempt As Integer
    intAttempt = intAttempt + 1
    Open "D:\VBA\������������ ������.txt" For Append As #1
    Print #1, "�������� �� ��������� �� VBA (Append) ������� " & intAttempt
    Debug.Print "�������� ������ � ���� ������� " & intAttempt
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'������ ��������� ������. ����������
Sub File_Output_01()
    On Error GoTo ErrorHandler
    Static intAttempt As Integer: intAttempt = intAttempt + 1
    Open "D:\VBA\������������ ������.txt" For Output As #1
    Print #1, "�������� �� ��������� �� VBA ������� " & intAttempt
    Debug.Print "�������� ������ � ���� ������� " & intAttempt
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'������� FreeFile. ��������� 0 � 1
Sub File_FreeFile_01()
    On Error GoTo ErrorHandler
    Debug.Print "����� ����� ��� FreeFile(0): " & FreeFile(0)
    Debug.Print "����� ����� ��� FreeFile(1): " & FreeFile(1)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������� FreeFile. ������������ ������
Sub File_FreeFile_02()
    On Error GoTo ErrorHandler
    Static intAttempt As Integer
    Dim intFile As Integer
    intAttempt = intAttempt + 1
    intFile = FreeFile
    Open "D:\VBA\������������ ������.txt" For Output As #intFile
    Print #intFile, "�������� �� ��������� �� VBA, ������� " & intAttempt
    Debug.Print "�������� ������ � ���� �������: " & intAttempt
Finalization:
    Debug.Print "��������� ��������� ���� �� ��������: " & FreeFile
    Close #intFile
    Debug.Print "��������� ��������� ���� ����� ��������: " & FreeFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'�������� Print. ������ � ������ ������
Sub File_DataTypes_Print_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Static intAttempt As Integer: intAttempt = intAttempt + 1
    Open "D:\VBA\������������ Print.txt" For Output As #intFile
    Print #intFile, "������ ������ Print, �������"; intAttempt
    Print #intFile, "Integer", "Double", "Boolean", "Null", "Date"
    Print #intFile, 100, 100.5, True, Null, Now
    Debug.Print "�������� ������ � ����, �������: " & intAttempt
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'�������� Write. ������ � ������ ������
Sub File_DataTypes_Write_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Static intAttempt As Integer: intAttempt = intAttempt + 1
    Open "D:\VBA\������������ Write.txt" For Output As #intFile
    Write #intFile, "������ ������ Write, �������"; intAttempt
    Write #intFile, "Integer", "Double", "Boolean", "Null", "Date"
    Write #intFile, 100, 100.5, True, Null, Now
    Debug.Print "�������� ������ � ����, �������: " & intAttempt
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'�������� Write. ������ ������
Sub File_DataTypes_Write_02()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Static intAttempt As Integer: intAttempt = intAttempt + 1
    Open "D:\VBA\���������� ������.txt" For Output As #intFile
    Write #intFile, 100, 100.5, True, Null, Now
    Write #intFile, 200, 100.6, False, Empty, Now
    Write #intFile, 300, 100.7, True, Null, Now
    Write #intFile, 400, 100.8, True, Null, Now
    Debug.Print "�������� ������ � ����, �������: " & intAttempt
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'������ ���������� Write ������
Sub File_DataTypes_Input_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim intVar As Integer
    Dim dblVar As Double
    Dim boolVar As Boolean
    Dim varVar As Variant
    Dim dtVar As Date
    Open "D:\VBA\���������� ������.txt" For Input As #intFile
    Input #intFile, intVar, dblVar, boolVar, varVar, dtVar
    Debug.Print intVar, dblVar, boolVar, varVar, dtVar
    Input #intFile, intVar, dblVar, boolVar, varVar, dtVar
    Debug.Print intVar, dblVar, boolVar, varVar, dtVar
    Debug.Print TypeName(varVar)
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'���������� ������ ������� Person
Sub WritePeople()
    Dim oPeople As Collection
    Dim oPerson As Person
    Set oPerson = New Person
    oPerson.LastName = "��������"
    oPerson.FirstName = "������"
    oPerson.Gender = "�"
    oPerson.BirthDate = #2/13/1972#
    oPerson.AddToFile ("D:\VBA\People.txt")
    Set oPeople = New Collection
    oPeople.Add oPerson
    Set oPerson = New Person
    oPerson.LastName = "������"
    oPerson.FirstName = "�������"
    oPerson.Gender = "f"
    oPerson.BirthDateIsUnknown = True
    oPerson.AddToFile ("D:\VBA\People.txt")

End Sub
'������ ������ ������� � Person ���������
Sub File_DataTypes_Input_02()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim oPeople As New Collection
    Dim oPerson As Person
    Dim strFirstName As String, strLastName As String, strGender As String, _
                dtBirthDate As Date, BirthDateIsUnknown As Boolean
    Open "D:\VBA\People.txt" For Input As #intFile
    Do Until EOF(intFile)
         Input #intFile, strFirstName, strLastName, strGender, _
                dtBirthDate, BirthDateIsUnknown
         Set oPerson = New Person
         oPerson.FirstName = strFirstName
         oPerson.LastName = strLastName
         oPerson.Gender = strGender
         oPerson.BirthDate = dtBirthDate
         oPerson.BirthDateIsUnknown = BirthDateIsUnknown
         oPeople.Add oPerson, strLastName & strFirstName
    Loop
    PrintPeopleCollection oPeople
Finalization:
    Set oPerson = Nothing
    Set oPeople = Nothing
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'�������� ������ ������ ����� Input
Sub File_DataTypes_Input_04()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim strText As String
    Open "D:\VBA\������ �����.txt" For Input As #intFile
    Input #intFile, strText
    MsgBox strText
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'������ ����� ������ Line Input � ������������� �����
Sub File_DataTypes_Input_06()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim strText As String
        Open "D:\VBA\������ �����.txt" For Input As #intFile
        Do Until EOF(intFile)
            Line Input #intFile, strText
            MsgBox strText
        Loop
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'���� ����� ���
Sub File_DataTypes_Input_07()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim strText As String
        Open "D:\VBA\��������� ��������.txt" For Input As #intFile
        Do Until EOF(intFile)
            Line Input #intFile, strText
            MsgBox strText
        Loop
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'������� ���������������� ������ Print
Sub File_Commands_Print()
    On Error GoTo ErrorHandler
    Static intAttempt As Integer
    intAttempt = intAttempt + 1
    Open "D:\VBA\������������ ������.txt" For Output As #1
    Print #1, "������� " & intAttempt
    Print #1, "������� ������� ", intAttempt
    Print #1, "����� � ������� ������� "; intAttempt
    Print #1, "Spc(2) ������� "; Spc(2); intAttempt
    Print #1, "Space(2) ������� " & Space(2) & intAttempt
    Print #1, "String(String(2, vbTab)) ������� " & String(2, vbTab) & intAttempt
    Print #1, "Tab ������� "; Tab; intAttempt
    Print #1, "Tab(1) ������� "; Tab(1); intAttempt
    Print #1, "Tab(2) ������� "; Tab(2); intAttempt
    Print #1, "Tab(10) ������� "; Tab(10); intAttempt
    Print #1, "Tab(11) ������� "; Tab(11); intAttempt
    Print #1, "������ ����������� ������ ";
    Print #1, "� �� �� ����� ������"
    
    Debug.Print "�������� ������ � ���� ������� " & intAttempt
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'������� ���������������� ������ Write
Sub File_Commands_Write()
    On Error GoTo ErrorHandler
    Static intAttempt As Integer
    intAttempt = intAttempt + 1
    Open "D:\VBA\������������ ������.txt" For Output As #1
    Write #1, "������� " & intAttempt
    Write #1, "������� ������� ", intAttempt
    Write #1, "����� � ������� ������� "; intAttempt
    Write #1, "Spc(2) ������� "; Spc(2); intAttempt
    Write #1, "Tab ������� "; Tab; intAttempt
    Write #1, "Tab(1) ������� "; Tab(1); intAttempt
    Write #1, "Tab(2) ������� "; Tab(2); intAttempt
    Write #1, "Tab(10) ������� "; Tab(10); intAttempt
    Write #1, "Tab(11) ������� "; Tab(11); intAttempt
    Write #1, "������ ����������� ������ ";
    Write #1, "� �� �� ����� ������"
    Debug.Print "�������� ������ � ���� ������� " & intAttempt
Finalization:
    Close #1
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'����������. ����������� Shared
Sub File_Lock_Shared_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Open "D:\VBA\������ ������.txt" For Output Shared As #intFile
    Write #intFile, "������ Word Shared", 200
    Debug.Print "�������� ������ ������ Word"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'����������. ����������� Lock Read
Sub File_Procedure_Lock_Read_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Open "D:\VBA\������ ������.txt" For Output Lock Read As #intFile
    Write #intFile, "������ Word Lock Read", 200
    Debug.Print "�������� ������ ������ Word"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'����������. ����������� Lock Write
Sub File_Procedure_Lock_Write_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Open "D:\VBA\������ ������.txt" For Output Lock Write As #intFile
    Write #intFile, "������ ���� ���������", 200
    Debug.Print "�������� ������ ������ Word"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'����������. ����������� Lock Read Write
Sub File_Procedure_Lock_Read_Write_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Open "D:\VBA\������ ������.txt" For Output Lock Read Write As #intFile
    Write #intFile, "������ ���� ���������", 200
    Debug.Print "�������� ������ ������ Word"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'����� ������������� �������. Put
Sub File_Random_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Open "D:\VBA\Random.txt" For Random As #intFile Len = 4
    Put #intFile, 1, 100
    Put #intFile, 2, 200
    Put #intFile, 3, 100500
    Debug.Print "�������� ������ ������"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'����� ������������� �������. Get
Sub File_Random_02()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim lngVar As Long
    Open "D:\VBA\Random.txt" For Random As #intFile Len = 4
    Get #intFile, 3, lngVar
    Debug.Print "�������� ������ ������:", lngVar
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'����� ������������� �������. ������ Write, Print
Sub File_Random_03()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Open "D:\VBA\������ ������.txt" For Random As #intFile
    Print #intFile, "������ ���� ���������", 200
    Debug.Print "�������� ������ ������"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'����� ������������� �������. ���������� ����� ������
Sub File_Random_04()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Open "D:\VBA\Random.txt" For Random As #intFile Len = 4
    Put #intFile, 10, 100
    Debug.Print "�������� ������ ������"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'����� ������������� �������. ���������������� ���������� ����� ������
Sub File_Random_05()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim byteRecordLenth As Byte
    byteRecordLenth = 4
    Open "D:\VBA\Random.txt" For Random As #intFile Len = byteRecordLenth
    Put #intFile, LOF(intFile) / byteRecordLenth + 1, 100
    Debug.Print "�������� ������ ������"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Random �������� �� ������ � ������. Read Write �� ���������
Sub File_Random_06()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim lngVar As Long
    Open "D:\VBA\Random.txt" For Random Access Read Write As #intFile Len = 4
    Put #intFile, 3, 200500
    Debug.Print "�������� ������ ������"
    Get #intFile, 3, lngVar
    Debug.Print "�������� ������ ������:"; lngVar
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'�������� ����� ������������� ������� ������ �� ������
Sub File_Random_07()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim lngVar As Long
    Open "D:\VBA\Random.txt" For Random Access Read As #intFile Len = 4
    'Put #intFile, 3, 200500
    'Debug.Print "�������� ������ ������"
    Get #intFile, 3, lngVar
    Debug.Print "�������� ������ ������:"; lngVar
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'�������� ����� ������������� ������� ������ �� ������
Sub File_Random_08()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim lngVar As Long
    Open "D:\VBA\Random.txt" For Random Access Write As #intFile Len = 4
    Put #intFile, 3, 300500
    Debug.Print "�������� ������ ������"
    'Get #intFile, 3, lngVar
    'Debug.Print "�������� ������ ������:"; lngVar
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'�������� ���������. �������� Seek, ������� Seek � Loc
Sub File_Random_09()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim lngVar As Long
    Open "D:\VBA\�������� ���������.txt" For Random As #intFile Len = 8
    Put #intFile, , "������"
    Debug.Print "Seek ����� ������:"; Seek(intFile)
    Put #intFile, , "������"
    Debug.Print "Seek ����� ������:"; Seek(intFile)
    Debug.Print "Loc ����� ������: "; Loc(intFile)
    Seek intFile, 1
    Debug.Print "Seek ����� Seek intFile, 1:"; Seek(intFile)
    Debug.Print "Loc ����� Seek intFile, 1:"; Loc(intFile)
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
'�������� �����. ������
Sub File_Binary_01()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim strFirstName As String * 20: strFirstName = "������"
    Dim strLastName As String * 20: strLastName = "��������"
    Dim strGender As String * 1: strGender = "M"
    Open "D:\VBA\Person.bin" For Binary As #intFile
    Put #intFile, , strFirstName
    Put #intFile, , strLastName
    Put #intFile, , strGender
    Put #intFile, , #2/13/1972#
    Put #intFile, , True
    Debug.Print "�������� ������ ������"
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'�������� �����. ������
Sub File_Binary_02()
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Dim strFirstName As String * 20
    Dim strLastName As String * 20
    Dim strGender As String * 1
    Dim dtBirthDate As Date
    Dim boolRFCitizen As Boolean
    Open "D:\VBA\Person.bin" For Binary As #intFile
    Get #intFile, , strFirstName
    Get #intFile, , strLastName
    Get #intFile, , strGender
    Get #intFile, , dtBirthDate
    Get #intFile, , boolRFCitizen
    Debug.Print strFirstName; strLastName; strGender, dtBirthDate; boolRFCitizen
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

