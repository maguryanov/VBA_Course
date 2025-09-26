Attribute VB_Name = "m_05_11_Bookmarks"
Option Explicit

'������ ����� ����� "��������� �� ������" ��� �������������
Private Sub d_01_CopyFileForExperiment()
    On Error GoTo ErrorHandler
    Dim strSourceFilename As String, strTargetFilename As String
    strSourceFilename = "D:\VBA\��� ������ Word\��������� �� ������.docx"
    strTargetFilename = "D:\VBA\Word\���������.docx"
    Call CopyFile(SourceFilename:=strSourceFilename, TargetFilename:=strTargetFilename)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'�������� ��������
Private Sub d_02_ShowBookmarks()
    On Error GoTo ErrorHandler
    Dim docStatement As Document, objBookmarkItem As Bookmark
    Set docStatement = Documents.Open("D:\VBA\Word\���������.docx")
    For Each objBookmarkItem In docStatement.Bookmarks
        Debug.Print objBookmarkItem.Name
    Next objBookmarkItem
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'�������� ������� ��������
Private Sub d_03_Exists()
    On Error GoTo ErrorHandler
    Dim docStatement As Document, objBookmarkItem As Bookmark
    Set docStatement = Documents.Open("D:\VBA\Word\���������.docx")
    If docStatement.Bookmarks.Exists("���������") Then
        Debug.Print "�������� ����������"
    Else
        Debug.Print "�������� �� ����������"
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'�������������� ��������
Private Sub d_04_FormatBookmarks()
    On Error GoTo ErrorHandler
    Dim docStatement As Document, objBookmarkItem As Bookmark
    Set docStatement = Documents.Open("D:\VBA\Word\���������.docx")
    For Each objBookmarkItem In docStatement.Bookmarks
        objBookmarkItem.Range.Font.Bold = False
    Next objBookmarkItem
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



'������������ �������� ������ ��������
Private Sub d_05_SetBookmarkText()
    On Error GoTo ErrorHandler
    Dim docStatement As Document
    Dim strEmployee As String, dtStart As Date, dtEnd As Date
    strEmployee = "��������� �. �. �������"
    dtStart = #9/15/2025#
    dtEnd = #9/26/2025#
    Set docStatement = Documents.Add("D:\VBA\��� ������ Word\��������� �� ������.docx")
    'Set docStatement = Documents.Add("D:\VBA\��� ������ Word\������� �����.docx")
    docStatement.Bookmarks("���������").Range.Text = strEmployee
    docStatement.Bookmarks("����������").Range.Text = Format(dtStart, "�dd� mmmm yyyy")
    docStatement.Bookmarks("���������").Range.Text = Format(dtEnd, "�dd� mmmm yyyy")
    Exit Sub
ErrorHandler:
    MsgBox "O����� ��� ������������ ���������: " & Err.Number & vbCrLf _
        & Err.Description & vbCrLf & "��������� �������� � ���������� � ������������"
End Sub

'���������� �������� �� ���������
Private Sub d_06_AddBookmark()
    On Error GoTo ErrorHandler
    Dim docStatement As Document
    Dim strEmployee As String, dtStart As Date, dtEnd As Date
    Set docStatement = Documents.Open("D:\VBA\Word\���������.docx")
    docStatement.Bookmarks.Add "�����������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'�������� ��������
Private Sub d_07_DeleteBookmark()
    On Error GoTo ErrorHandler
    Dim docStatement As Document
    Dim strEmployee As String, dtStart As Date, dtEnd As Date
    Set docStatement = Documents.Open("D:\VBA\Word\���������.docx")
    docStatement.Bookmarks("�����������").Delete
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'���������� �������� �� ���������
Private Sub d_08_AddBookmarkFromRange()
    On Error GoTo ErrorHandler
    Dim docStatement As Document
    Dim rngPositionOfBoss As Range
    Set docStatement = Documents.Open("D:\VBA\Word\���������.docx")
    Set rngPositionOfBoss = docStatement.Range(0, docStatement.Range.Words(2).End - 1)
    docStatement.Bookmarks.Add "��������������������", rngPositionOfBoss
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'������������ �������� ������ �������� c ����������� ��������
Private Sub d_09_BookmarksSaving()
    On Error GoTo ErrorHandler
    Dim docStatement As Document, rngItem As Range
    Dim strEmployee As String, dtStart As Date, dtEnd As Date
    strEmployee = "��������� �. �. �������"
    dtStart = #9/15/2025#
    dtEnd = #9/26/2025#
    Set docStatement = Documents.Add("D:\VBA\��� ������ Word\��������� �� ������.docx")
    
    If docStatement.Bookmarks.Exists("���������") Then
        Set rngItem = docStatement.Bookmarks("���������").Range
        docStatement.Bookmarks("���������").Range.Text = strEmployee
        docStatement.Bookmarks.Add "���������", rngItem
    End If
    
    If docStatement.Bookmarks.Exists("����������") Then
        Set rngItem = docStatement.Bookmarks("����������").Range
        docStatement.Bookmarks("����������").Range.Text = Format(dtStart, "�dd� mmmm yyyy")
        docStatement.Bookmarks.Add "����������", rngItem
    End If
    
    If docStatement.Bookmarks.Exists("���������") Then
        Set rngItem = docStatement.Bookmarks("���������").Range
        docStatement.Bookmarks("���������").Range.Text = Format(dtEnd, "�dd� mmmm yyyy")
        docStatement.Bookmarks.Add "���������", rngItem
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



