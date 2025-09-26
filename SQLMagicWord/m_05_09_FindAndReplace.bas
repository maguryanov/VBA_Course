Attribute VB_Name = "m_05_09_FindAndReplace"
Option Explicit

'������ ����� ����� "Visual Basic" ��� �������������
Private Sub d_01_CopyFileForExperiment()
    On Error GoTo ErrorHandler
    Dim strSourceFilename As String, strTargetFilename As String
    strSourceFilename = "D:\VBA\��� ������ Word\������� �����.docx"
    strTargetFilename = "D:\VBA\Word\Visual Basic.docx"
    Call CopyFile(SourceFilename:=strSourceFilename, TargetFilename:=strTargetFilename)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'����� �� ��� ������. ������� ������
Private Sub d_02_FindWholeDocument()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    With objDocument.Content.Find
        .ClearFormatting          ' ��� ����� ��������������
        .Text = "visual basic"    ' ��� ����
        .MatchCase = False        ' �� ��������� �������
        .MatchWholeWord = True    ' ������ ������ ����� �����
        .Execute
        If .Found Then
            Debug.Print "����� ������"
        Else
            Debug.Print "����� �� ������"
        End If
    End With
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



'����� �� ��� ������. ���������� ������
Private Sub d_03_FindAll()
    On Error GoTo ErrorHandler
    Dim objDocument As Document, rngItem As Range, intCounter As Integer
    Dim rngSentence
    Set objDocument = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    Set rngItem = objDocument.Content
    intCounter = 0
    With rngItem.Find
        .ClearFormatting          ' ��� ����� ��������������
        .Text = "visual basic"    ' ��� ����
        .MatchCase = False        ' �� ��������� �������
        .MatchWholeWord = True    ' ������ ������ ����� �����
        Do While .Execute 'True , ���� �������� ������ ��������� �������
            rngItem.Font.Bold = True ' ������� ���������
            Set rngSentence = rngItem.Duplicate
            rngSentence.Expand Unit:=wdSentence
            rngSentence.HighlightColorIndex = wdYellow ' ������� ����������� � ������� ������� �����
            rngSentence.Font.Color = vbBlue
            rngItem.Collapse wdCollapseEnd ' ������� ������ ������
            intCounter = intCounter + 1
        Loop
    End With
    Debug.Print "�������:"; intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'������ ������
Private Sub d_04_ReplaceAll()
    On Error GoTo ErrorHandler
    Dim objDocument As Document, rngItem As Range, intCounter As Integer
    Dim rngSentence
    Set objDocument = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    With objDocument.Content.Find
        .Text = "VBA"
        .Replacement.Text = "Visual Basic for Applications"
        .MatchCase = False
        .Execute Replace:=wdReplaceAll
    End With
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'����� �� ��������������. ����� ������ � ���������� ������� ������
Private Sub d_05_FindByFormat()
    On Error GoTo ErrorHandler
    Dim objDocument As Document, rngItem As Range, intCounter As Integer
    Dim rngSentence
    Set objDocument = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    Set rngItem = objDocument.Content
    intCounter = 0
    With rngItem.Find
        .ClearFormatting          ' ������� ��������������
        .Style = "��������� 1"
        .Text = "visual basic"    ' ��� ����
        .MatchCase = False        ' �� ��������� �������
        .MatchWholeWord = True    ' ������ ������ ����� �����
        Do While .Execute 'True , ���� �������� ������ ��������� �������
            rngItem.Font.Bold = True ' ������� ���������
            rngItem.Font.Italic = True ' ������� ���������
            rngItem.Collapse wdCollapseEnd ' ������� ������ ������
            intCounter = intCounter + 1
        Loop
    End With
    Debug.Print "�������:"; intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



'����� � ���������
Private Sub d_06_FindInSelection()
    On Error GoTo ErrorHandler
    Dim rngItem As Range, intCounter As Integer, lngEnd As Long
    Set rngItem = Selection.Range
    lngEnd = Selection.Range.End
    intCounter = 0
    With rngItem.Find
        .ClearFormatting          ' ��� ����� ��������������
        .Text = "visual basic"    ' ��� ����
        .MatchCase = False        ' �� ��������� �������
        .MatchWholeWord = True    ' ������ ������ ����� �����
        .Wrap = wdFindStop
        Do While .Execute And rngItem.End < lngEnd
            rngItem.Font.Bold = True ' ������� ���������
            rngItem.HighlightColorIndex = wdYellow
            rngItem.Collapse wdCollapseEnd ' ������� ������ ������
            intCounter = intCounter + 1
        Loop
    End With
    Debug.Print "�������:"; intCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
