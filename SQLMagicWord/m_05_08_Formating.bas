Attribute VB_Name = "m_05_08_Formating"
Option Explicit

'������ ����� ����� "���� VBA" ��� �������������
Private Sub d_01_CopyFile()
    On Error GoTo ErrorHandler
    Dim strSourceFilename As String, strTargetFilename As String
    strSourceFilename = "D:\VBA\��� ������ Word\��������� �����.docx"
    strTargetFilename = "D:\VBA\Word\���� VBA.docx"
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


'���������� ������
Private Sub d_02_Style()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\���� VBA.docx")
    objDocument.Paragraphs(1).Style = "��������� 1"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'�������������� �������������� � �������������� ���������
Private Sub d_03_AutomaticFormating()
    On Error GoTo ErrorHandler
    Dim objDocument As Document, rngItem As Range, objParagraph As Paragraph
    Dim strWord As String
    Set objDocument = Documents.Open("D:\VBA\Word\���� VBA.docx")
    For Each objParagraph In objDocument.Paragraphs
        strWord = LCase(objParagraph.Range.Words(2).Text)
        Select Case strWord
            Case "�����": objParagraph.Range.Style = "��������� 1"
                objParagraph.Range.Sentences(1).Delete
            Case "����": objParagraph.Range.Style = "��������� 2"
                objParagraph.Range.Sentences(1).Delete
            Case "����": objParagraph.Range.Style = "��������� 3"
                objParagraph.Range.Sentences(1).Delete
        End Select
    Next objParagraph
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'��������� ������
Private Sub d_05_Borders()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\���� VBA.docx")
    objDocument.Paragraphs(1).Borders.Enable = _
        Not objDocument.Paragraphs(1).Borders.Enable
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'��������� ��������� ������
Private Sub d_06_SetInnerBorders()
    On Error GoTo ErrorHandler
    
    With Selection.Borders
        .InsideLineStyle = wdLineStyleDot
        .InsideLineWidth = wdLineWidth150pt
        .InsideColor = wdColorAqua
    End With
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'������ ���������� ������
Private Sub d_07_UnsetInnerBorders()
    On Error GoTo ErrorHandler
    
    Selection.Borders.InsideLineStyle = wdLineStyleNone
    
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'��������� ������� ������
Private Sub d_08_SetOuterBorders()
    On Error GoTo ErrorHandler
    
    With Selection.Borders
        .OutsideLineStyle = wdLineStyleDouble
        .OutsideLineWidth = wdLineWidth050pt
        .OutsideColor = wdColorLightBlue
    End With
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������ ������� ������
Private Sub d_09_UnsetOuterBorders()
    On Error GoTo ErrorHandler
    
    Selection.Borders.OutsideLineStyle = wdLineStyleNone
    
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'������ ����� ����� "������ � �������" ��� �������������
Private Sub d_10_CopyFile()
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


'������ �� �������
Private Sub d_11_Font()
    On Error GoTo ErrorHandler
    With Selection.Font
        .Name = "Times New Roman"
        .Bold = True
        .Italic = True
        .TextColor = RGB(0, 255, 0)
        .Parent.Shading.BackgroundPatternColor = wdColorBlack
    End With
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'����������� � ������� ����������� Visual Basic ������� ��������
Private Sub d_12_AutomaticFormating()
    On Error GoTo ErrorHandler
    Dim objDocument As Document, rngWord As Range
    Set objDocument = Documents.Open("D:\VBA\Word\������ � �������.docx")
    For Each rngWord In objDocument.Words
        If LCase(rngWord.Text) = "basic" Then
            rngWord.Font.Bold = True
            rngWord.Expand Unit:=wdSentence
            rngWord.Font.Italic = True
        End If
    Next rngWord
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



'�o������� ���������
Private Sub d_13_StripedParagraphs()
    On Error GoTo ErrorHandler
    Dim objDocument As Document, objItem As Paragraph
    Dim intParagraph As Integer: intParagraph = 0
    Set objDocument = Documents.Open("D:\VBA\Word\������ � �������.docx")
    For Each objItem In objDocument.Paragraphs
        intParagraph = intParagraph + 1
        If intParagraph Mod 2 = 0 Then
            objItem.Range.Font.ColorIndex = wdGreen
        Else
            objItem.Range.Font.ColorIndex = wdDarkBlue
        End If
    Next objItem
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������ � ����������
Private Sub d_14_Paragraph()
    On Error GoTo ErrorHandler
    With Selection.ParagraphFormat
        .Alignment = wdAlignParagraphJustify
        .FirstLineIndent = CentimetersToPoints(2)
        .SpaceAfter = 0
        .SpaceBefore = 0
        .LineSpacingRule = wdLineSpaceDouble
    End With
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



'������ ����� ����� "Visual Basic" ��� �������������
Private Sub d_15_CopyFile()
    On Error GoTo ErrorHandler
    Dim strSourceFilename As String, strTargetFilename As String
    strSourceFilename = "D:\VBA\��� ������ Word\������� �����.docx"
    strTargetFilename = "D:\VBA\Word\Visual Basic.docx"
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



'������ ��������� �������������� ��� ������ ����� ��������� � �����������
Private Sub d_16_AutomationParagraphFormating()
    On Error GoTo ErrorHandler
    Dim objDocument As Document, objItem As Paragraph
    Dim boolAfterHeader As Boolean
    Set objDocument = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    For Each objItem In objDocument.Paragraphs
        If objItem.Style Like "��������� *" Then
            boolAfterHeader = True
        ElseIf boolAfterHeader = True Then
            With objItem.Range
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .ParagraphFormat.FirstLineIndent = 0
                .ParagraphFormat.SpaceAfter = 18
                .Font.ColorIndex = wdBrightGreen
                .Shading.BackgroundPatternColor = wdColorBlack
            End With
            boolAfterHeader = False
        Else
            With objItem.Range.ParagraphFormat
                .Alignment = wdAlignParagraphJustify
                .FirstLineIndent = CentimetersToPoints(1.4)
                .SpaceBefore = 0
                .SpaceAfter = 0
            End With
        End If
    Next objItem
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������ � ������� ���������
Private Sub d_17_AddTabStops()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    objDocument.Range.InsertAfter vbCrLf & "������������" & vbTab & " ������ " & _
            vbTab & "� ������� " & vbTab & "���������"
    objDocument.Paragraphs(objDocument.Paragraphs.Count).TabStops.Add 120
    objDocument.Paragraphs(objDocument.Paragraphs.Count).TabStops.Add 240
    objDocument.Paragraphs(objDocument.Paragraphs.Count).TabStops.Add 360
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'��������� ����� ���������
Private Sub d_18_ChangeTabStops()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    objDocument.Paragraphs(objDocument.Paragraphs.Count).TabStops(2).Position = 280
    objDocument.Paragraphs(objDocument.Paragraphs.Count).TabStops(3).Clear
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



'�������� ���� ����� ���������
Private Sub d_19_ClearTabStops()
    On Error GoTo ErrorHandler
    Dim objDocument As Document
    Set objDocument = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    objDocument.Paragraphs(objDocument.Paragraphs.Count).TabStops.ClearAll
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

