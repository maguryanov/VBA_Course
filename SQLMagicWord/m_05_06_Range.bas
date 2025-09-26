Attribute VB_Name = "m_05_06_Range"
Option Explicit

'������ ����� ����� ��� �������������
Private Sub d_01_CopyFile()
    On Error GoTo ErrorHandler
    Dim strSourceFilename As String, strTargetFilename As String
    strSourceFilename = "D:\VBA\��� ������ Word\������������.docx"
    strTargetFilename = "D:\VBA\Word\������ � Range.docx"
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

'��� �������� ������ � Range? Paragraph
Private Sub d_02_Paragraph()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\������������.docx")
    Debug.Print objDocument.Paragraphs(2).Range.Text
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'��� �������� ������ � Range? Sentence
Private Sub d_03_Sentence()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\������������.docx")
    Debug.Print objDocument.Sentences(4).Text
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'��� �������� ������ � Range? Words
Private Sub d_04_Rows()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\������������.docx")
    Debug.Print objDocument.Words(2).Text
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'��� �������� ������ � Range? Characters
Private Sub d_05_Characters()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\������������.docx")
    Debug.Print objDocument.Characters(1).Text
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'��� �������� ������ � �������? Range Tables
Private Sub d_07_Tables()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\������������.docx")
    Debug.Print objDocument.Tables(1).Range.Text
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'��� �������� ������ � ����������? Comments
Private Sub d_08_Comments()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\������������.docx")
    Debug.Print objDocument.Comments(1).Range.Text
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'��������� ������� � ����� Range
Private Sub d_09_CustomRange()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\������������.docx")
    Debug.Print objDocument.Paragraphs(1).Range.Text
    Debug.Print objDocument.Range(Start:=5, End:=9).Text
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'����� � ���������
Private Sub d_10_PointRange()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\������������.docx")
    Debug.Print objDocument.Paragraphs(1).Range.Text
    Debug.Print objDocument.Range(Start:=8, End:=8)
    objDocument.Range(Start:=9, End:=9).InsertAfter COPYRIGHT_SIGN
    Debug.Print objDocument.Paragraphs(1).Range.Text
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'��������� �� ���������
Private Sub d_11_CustomRange()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\������������.docx")
    Debug.Print objDocument.Paragraphs(2).Range.Sentences(2).Words(1).Text
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'������������ ������ �� Range ����������
Private Sub d_12_Variable()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Dim rngParagraph02 As Range
    Set objDocument = Documents.Open("D:\VBA\Word\������������.docx")
    Set rngParagraph02 = objDocument.Paragraphs(2).Range
    Debug.Print rngParagraph02.Text
    rngParagraph02.Bold = True
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'��� �������� ������ ������?
Private Sub d_13_SeveralObjects()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\������������.docx")
    objDocument.Range( _
        Start:=objDocument.Paragraphs(1).Range.Start, _
        End:=objDocument.Paragraphs(4).Range.End _
        ).Select
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
