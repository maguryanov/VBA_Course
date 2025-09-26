Attribute VB_Name = "m_01_16_FilesOOP"
Option Explicit
'������ ����� �����
Sub FileReading_01()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Dim strContent As String
    Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\ANSI.txt", ForReading)
    strContent = oTextStream.ReadAll()
    MsgBox strContent
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'������ ��������� ������
Sub FileReading_02()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Dim strContent As String
    Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\ANSI.txt", ForReading)
    strContent = oTextStream.ReadLine()
    MsgBox strContent
    strContent = oTextStream.ReadLine()
    MsgBox strContent
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'��������� (Read) � ���������� (Skip) �������� ���������� ��������
Sub FileReading_03()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Dim strContent As String
    Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\ANSI.txt", ForReading)
    oTextStream.Skip 7
    strContent = oTextStream.Read(5)
    Debug.Print strContent
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'���������� ������ � �����
Sub FileReading_04()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Dim strContent As String
    Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\ANSI.txt", ForReading)
    oTextStream.SkipLine
    strContent = oTextStream.ReadLine()
    Debug.Print strContent
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'������� ����, ������������, ���� ����������
Sub FileWriting_01()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Static intAttemp As Integer: intAttemp = intAttemp + 1
    Set oTextStream = oFileSystem.CreateTextFile _
        ("D:\VBA\������������ FileSystemObject.txt", ForWriting)
    oTextStream.WriteLine ("������������ ForWriting, �������: " & intAttemp)
    Debug.Print "�������� ������ �����, �������: " & intAttemp; ""
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'������� ����, ��������
Sub FileWriting_02()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Static intAttemp As Integer: intAttemp = intAttemp + 1
    Set oTextStream = oFileSystem.OpenTextFile _
        ("D:\VBA\������������ FileSystemObject.txt", ForAppending)
    oTextStream.WriteLine ("������������ ForAppending, �������: " & intAttemp)
    Debug.Print "�������� ������ �����, �������: " & intAttemp; ""
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

' ������� ����� ����, ���� �� �� ����������
Sub FileWriting_03()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject, oTextStream As TextStream
    Dim strFilePath As String: strFilePath = "D:\VBA\����� ����.txt"
    If Not oFileSystem.FileExists(strFilePath) Then
' ������ �������� - True, ���� ���� ����� ������������, � False � ��������� ������
        Set oTextStream = oFileSystem.CreateTextFile(strFilePath, False)
        oTextStream.Close
        Debug.Print "����� ���� ������"
    Else
        Debug.Print "���� ����������"
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������ �������������� ����� ����� �������
Sub FileWriting_04()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Static intAttemp As Integer: intAttemp = intAttemp + 1
    Set oTextStream = oFileSystem.CreateTextFile _
        ("D:\VBA\������������� ����.txt", ForWriting)
    oTextStream.WriteLine _
        ("������ ������" & vbCrLf & "������ ������" & vbCrLf & "������ ������")
    Debug.Print "�������� ������ �����, �������: " & intAttemp; ""
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'������ ��� �������� ������ (���������� ������)
Sub FileWriting_05()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Static intAttemp As Integer: intAttemp = intAttemp + 1
    Set oTextStream = oFileSystem.CreateTextFile _
        ("D:\VBA\������������� ����.txt", ForWriting)
    oTextStream.Write ("������ ������ ")
    oTextStream.Write ("������ ������")
    Debug.Print "�������� ������ �����, �������: " & intAttemp; ""
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'������� ������ �����
Sub FileWriting_06()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Static intAttemp As Integer: intAttemp = intAttemp + 1
    Set oTextStream = oFileSystem.CreateTextFile _
        ("D:\VBA\������������� ����.txt", ForWriting)
    oTextStream.WriteLine ("������ ������")
    oTextStream.WriteBlankLines (5)
    oTextStream.WriteLine ("��������� ������")
    Debug.Print "�������� ������ �����, �������: " & intAttemp; ""
Finalization:
    If Not oTextStream Is Nothing Then oTextStream.Close
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'�������� ���������
Sub FileNavigation_01()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject, oTextStream As TextStream
    Dim strContent As String
    Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\ANSI.txt", ForReading)
    strContent = oTextStream.ReadLine()
    Debug.Print "------ ����� ������� ReadLine ------"
    Debug.Print "AtEndOfLine:", , oTextStream.AtEndOfLine
    Debug.Print "AtEndOfStream:", oTextStream.AtEndOfStream
    Debug.Print "Column:", , oTextStream.Column
    Debug.Print "Line:", , oTextStream.Line
    strContent = oTextStream.ReadLine()
    Debug.Print "------ ����� ������� ReadLine ------"
    Debug.Print "AtEndOfLine:", , oTextStream.AtEndOfLine
    Debug.Print "AtEndOfStream:", oTextStream.AtEndOfStream
    Debug.Print "Column:", , oTextStream.Column
    Debug.Print "Line:", , oTextStream.Line
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'�������� ��������� Skip, Read
Sub FileNavigation_02()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject, oTextStream As TextStream
    Dim strContent As String
    Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\ANSI.txt", ForReading)
    oTextStream.Skip (7)
    Debug.Print "------ ����� Skip(7) ------"
    Debug.Print "AtEndOfLine:", , oTextStream.AtEndOfLine
    Debug.Print "AtEndOfStream:", oTextStream.AtEndOfStream
    Debug.Print "Column:", , oTextStream.Column
    Debug.Print "Line:", , oTextStream.Line
    strContent = oTextStream.Read(6)
    Debug.Print "------ ����� Read(6) ------"
    Debug.Print "AtEndOfLine:", , oTextStream.AtEndOfLine
    Debug.Print "AtEndOfStream:", oTextStream.AtEndOfStream
    Debug.Print "Column:", , oTextStream.Column
    Debug.Print "Line:", , oTextStream.Line
    Debug.Print "strContent:", strContent
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'������ � UTF-16,  ������ � UTF-8 �� ��������������
Sub FileReading_10()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oFile As File
    Dim oTextStream As TextStream
    Dim strContent As String
    Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\ANSI.txt", ForReading, , False)
    Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\UTF-16.txt", ForReading, , True)
    'Set oTextStream = oFileSystem.OpenTextFile("D:\VBA\UTF-8.txt", ForReading, , True)
    strContent = oTextStream.ReadAll()
    MsgBox strContent
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'������ UTF-8. ������ ��
Sub ADODB_Reading_01()
    On Error GoTo ErrorHandler
    Dim oStream As New ADODB.stream
    With oStream
        .Type = 2 ' ��������� ���
        'HKEY_CLASSES_ROOT\MIME\Database\Charset
        .Charset = "utf-8"
        .Open
        .LoadFromFile "D:\VBA\UTF-8.txt"
        Debug.Print .ReadText
        .Close
    End With
Finalization:
    Set oStream = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'������ UTF-8. ������ ���������
Sub ADODB_Reading_02()
    On Error GoTo ErrorHandler
    Dim oStream As New ADODB.stream
    With oStream
        .Type = 2 ' ��������� ���
        .Charset = "utf-8"
        .Open
        .LoadFromFile "D:\VBA\UTF-8.txt"
        .SkipLine
        Debug.Print .ReadText(3)
        .Close
    End With
Finalization:
    Set oStream = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'������ UTF-8
Sub ADODB_Writing_01()
    On Error GoTo ErrorHandler
    Dim oStream As New ADODB.stream
    Dim strContent As String: strContent = "���, ���������� �� VBA, ������������� � ������������� Microsoft P-��� (����-���)"
    With oStream
        .Type = 2 ' ��������� ���
        .Charset = "utf-8"
        .Open
        .WriteText strContent
        .SaveToFile "D:\VBA\UTF-8 ������.txt", adSaveCreateOverWrite
        .Close
    End With
    Debug.Print "������ ������� ���������"
Finalization:
    Set oStream = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'������ cp866
Sub ADODB_Writing_02()
    On Error GoTo ErrorHandler
    Dim oStream As New ADODB.stream
    Dim strContent As String: strContent = "���, ���������� �� VBA, ������������� � ������������� Microsoft P-��� (����-���)"
    With oStream
        .Type = 2 ' ��������� ���
        'HKEY_CLASSES_ROOT\MIME\Database\Charset
        .Charset = "cp866"
        .Open
        .WriteText strContent
        .SaveToFile "D:\VBA\cp866 ������.txt", adSaveCreateOverWrite
        .Close
    End With
    Debug.Print "������ ������� ���������"
Finalization:
    Set oStream = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'������ cp866
Sub ADODB_Readingcp866_01()
    On Error GoTo ErrorHandler
    Dim oStream As New ADODB.stream
    With oStream
        .Type = adTypeText
        'HKEY_CLASSES_ROOT\MIME\Database\Charset
        .Charset = "cp866"
        .Open
        .LoadFromFile "D:\VBA\cp866 ������.txt"
        Debug.Print .ReadText
        .Close
    End With
Finalization:
    Set oStream = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub
