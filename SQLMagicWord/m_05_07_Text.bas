Attribute VB_Name = "m_05_07_Text"
Option Explicit

'������ ����� ����� ��� �������������
Private Sub d_01_CopyFile()
    On Error GoTo ErrorHandler
    Dim strSourceFilename As String, strTargetFilename As String
    strSourceFilename = "D:\VBA\��� ������ Word\������������.docx"
    strTargetFilename = "D:\VBA\Word\������ � �������.docx"
    Application.DisplayAlerts = wdAlertsNone   ' ��������� ��������������
    If IsDocumentOpen(strTargetFilename) Then Documents(strTargetFilename).Close
    Documents.Add(strSourceFilename).SaveAs2 strTargetFilename
Finalization:
    Application.DisplayAlerts = wdAlertsAll   ' ���������� �������
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub



'��� �������� Selection �� Range?
Private Sub d_02_SelectionFromRange()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\������ � �������.docx")
    objDocument.Paragraphs(3).Range.Select
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'��� �������� Range �� Selection?
Private Sub d_03_RangeFromSelection()
    On Error GoTo ErrorHandler
    Dim rngFromSelection As Range
    Debug.Print Selection.Text
    Debug.Print Selection.Range.Text
    Set rngFromSelection = Selection.Range
    Debug.Print rngFromSelection.Text
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'���������� Selection
Private Sub d_04_ExpandSelection()
    On Error GoTo ErrorHandler
    Dim byteSwitch As Byte
    byteSwitch = 3
    Select Case byteSwitch
        Case 1: Selection.Expand Unit:=wdWord
        Case 2: Selection.Expand Unit:=wdSentence
        Case 3: Selection.Expand Unit:=wdParagraph
    End Select
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'���������� Range
Private Sub d_05_ExpandRange()
    On Error GoTo ErrorHandler
    Dim byteSwitch As Byte
    Dim objDocument As Document
    Dim rngDemo As Range
    Set objDocument = Documents.Open("D:\VBA\Word\������ � �������.docx")
    Set rngDemo = objDocument.Range(200, 200)
    byteSwitch = 3
    Select Case byteSwitch
        Case 1: rngDemo.Expand Unit:=wdWord
        Case 2: rngDemo.Expand Unit:=wdSentence
        Case 3: rngDemo.Expand Unit:=wdParagraph
    End Select
    Debug.Print rngDemo.Text
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'����������� ������ � �����
Private Sub d_06_MoveStartMoveEnd()
    On Error GoTo ErrorHandler
    Dim byteSwitch As Byte
    Dim objDocument As Document
    Dim rngDemo As Range
    Set objDocument = Documents.Open("D:\VBA\Word\������ � �������.docx")
    Set rngDemo = objDocument.Paragraphs(3).Range
    rngDemo.MoveStart wdWord, 2
    rngDemo.MoveEnd wdWord, -3
    rngDemo.Select
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������ ���������� ���������
Private Sub d_07_StartOfEndOf()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Dim rngDemo As Range
    Set objDocument = Documents.Open("D:\VBA\Word\������ � �������.docx")
    Set rngDemo = objDocument.Range(200, 200)
    rngDemo.StartOf Unit:=wdWord, Extend:=wdMove
    rngDemo.EndOf Unit:=wdSentence, Extend:=wdExtend
    rngDemo.Select
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'� ����� ������� ���������� �����������, ������� ������ �� 250 �������
Private Sub d_08_SentenceStart()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Dim rngDemo As Range
    Set objDocument = Documents.Open("D:\VBA\Word\������ � �������.docx")
    Selection.Collapse Direction:=wdCollapseStart
    Set rngDemo = objDocument.Range(250, 250)
    rngDemo.StartOf Unit:=wdSentence, Extend:=wdMove
    Debug.Print "������ �����������: " & rngDemo.Start
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'���������� ��������� wdCollapseEnd
Private Sub d_09_CollapseEndSelection()
    On Error GoTo ErrorHandler
    Dim rngDemo As Range
    Selection.Collapse Direction:=wdCollapseEnd
    Set rngDemo = Selection.Range
    Debug.Print "������ ���������: " & rngDemo.Start
    Debug.Print "����� ���������: " & rngDemo.End
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'���������� ��������� wdCollapseStart
Private Sub d_10_CollapseStartSelection()
    On Error GoTo ErrorHandler
    Dim rngDemo As Range
    Selection.Collapse Direction:=wdCollapseStart
    Set rngDemo = Selection.Range
    Debug.Print "������ ���������: " & rngDemo.Start
    Debug.Print "����� ���������: " & rngDemo.End
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



'�������� ������. ����� ��������
Private Sub d_11_Delete()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Dim rngDemo As Range
    Set objDocument = Documents.Open("D:\VBA\Word\������ � �������.docx")
    objDocument.Paragraphs(3).Range.Delete
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'�������� ������. ����� ��������
Private Sub d_12_DeleteAllText()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Dim rngDemo As Range
    Set objDocument = Documents.Open("D:\VBA\Word\������ � �������.docx")
    objDocument.Range.Delete
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'�������� � ����� ������
Private Sub d_13_CopyCut()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Dim rngDemo As Range
    Set objDocument = Documents.Open("D:\VBA\Word\������ � �������.docx")
    objDocument.Paragraphs(3).Range.Copy
    'objDocument.Paragraphs(3).Range.Cut
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������� �� ������ ������ � ����������
Private Sub d_14_PasteWithReplacement()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Dim rngDemo As Range
    Set objDocument = Documents.Open("D:\VBA\Word\������ � �������.docx")
    objDocument.Paragraphs(2).Range.Paste
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������� �� ������ ������ �����������
Private Sub d_15_PasteWithAddition()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Dim rngDemo As Range
    Set objDocument = Documents.Open("D:\VBA\Word\������ � �������.docx")
    Set rngDemo = objDocument.Paragraphs(3).Range
    rngDemo.Collapse Direction:=wdCollapseEnd
    rngDemo.Paste
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
'����� InsertAfter
Private Sub d_16_InsertAfter()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Dim rngDemo As Range
    Set objDocument = Documents.Open("D:\VBA\Word\������ � �������.docx")
    Set rngDemo = objDocument.Paragraphs(4).Range
    rngDemo.InsertAfter rngDemo.Text
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'����� InsertBefore
Private Sub d_17_InsertBefore()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Dim rngDemo As Range
    Set objDocument = Documents.Open("D:\VBA\Word\������ � �������.docx")
    Set rngDemo = objDocument.Paragraphs(4).Range
    rngDemo.InsertBefore rngDemo.Text
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'�������� FormatedText, ��������� ����� � ���������������
Private Sub d_18_InsertAfter()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Dim rngDemo As Range
    ' ��������� ��������
    Set objDocument = Documents.Open("D:\VBA\Word\������ � �������.docx")
    ' ���� 4-� �����
    Set rngDemo = objDocument.Paragraphs(4).Range
    ' ������ ��� � ����� (����� ��������� ����� ������)
    rngDemo.Collapse Direction:=wdCollapseEnd
    ' ��������� ���� ��������������� ����� �� 4-�� ������
    rngDemo.FormattedText = objDocument.Paragraphs(4).Range.FormattedText
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub




