Attribute VB_Name = "Macroses"
Sub Balabolka2()
'
' Balabolka2 ������
'
'
Dim Result As String
Dim Txt As String
Dim i As Integer

For i = 1 To ActiveDocument.Tables(1).Rows.Count
    Txt = ActiveDocument.Tables(1).Cell(i, 1).Range.Text
    Txt = Left(Txt, Len(Txt) - 2)
    Result = Result & "<silence msec=""1000""/><voice required=""Language=409"">" & Txt & "</voice>"
    Txt = ActiveDocument.Tables(1).Cell(i, 2).Range.Text
    Txt = Left(Txt, Len(Txt) - 2)
    Result = Result & " - <silence msec=""3000""/><voice required=""Language=419"">" & Txt & "</voice>" & Chr(13) & Chr(10)
    If (i Mod 80) = 0 Then
        Result = Result & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    End If
Next i

Dim MyDataObj As New DataObject
MyDataObj.SetText Result
MyDataObj.PutInClipboard

End Sub


Sub Balabolka3()
'
' Balabolka3 ������
'
'
Dim Result As String
Dim Txt As String
Dim i As Integer

For i = 1 To ActiveDocument.Tables(1).Rows.Count
    Txt = ActiveDocument.Tables(1).Cell(i, 2).Range.Text
    Txt = Left(Txt, Len(Txt) - 2)
    Result = Result & "<silence msec=""1000""/><voice required=""Language=409"">" & Txt & "</voice>"
    Txt = ActiveDocument.Tables(1).Cell(i, 3).Range.Text
    Txt = Left(Txt, Len(Txt) - 2)
    Result = Result & " - <silence msec=""3000""/><voice required=""Language=419"">" & Txt & "</voice>" & Chr(13) & Chr(10)
    If (i Mod 80) = 0 Then
        Result = Result & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    End If
Next i

Dim MyDataObj As New DataObject
MyDataObj.SetText Result
MyDataObj.PutInClipboard

End Sub

Sub ����������_����_������_����_���_����_���������()
    Dim selRange As Range
    Set selRange = Selection.Range
    Dim nextWord As Range
    Dim nextCharacter As Range
    Const charactersLAT As String = " abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim isLAT As Boolean
    For Each nextWord In selRange.Words
        isLAT = True
        For Each nextCharacter In nextWord.Characters
            Dim pos As Integer
            pos = InStr(charactersLAT, nextCharacter.Text)
            If pos = 0 Then
                isLAT = False
                Exit For
            End If
        Next
        If isLAT = True Then
            nextWord.LanguageID = wdEnglishUS
            nextWord.NoProofing = False
        End If
    Next
End Sub
Sub �������()
Attribute �������.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.�������"
'
' ������� ������
'
'
    MsgBox ("��������")
    Selection.EndKey Unit:=wdStory
    Selection.TypeText Text:="������������� ����������� �������� �. �."
    Selection.TypeParagraph
    Selection.TypeText Text:="�. 55-55"
End Sub

Sub �������2()
'
' ������� ������
'
'
    Selection.EndKey Unit:=wdStory
    Author = InputBox("������� ���� ���", "������ ����������")
    Selection.TypeText Text:="������������� ����������� " & Author
    Selection.TypeParagraph
    Selection.TypeText Text:="�. 55-55"
End Sub


Sub ExitWOSaving()
Attribute ExitWOSaving.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.ExitWOSaving"
    ActiveWindow.Close SaveChanges:=False
End Sub


Sub SortSelectionAsc()
    Selection.Sort SortOrder:=wdSortOrderAscending
End Sub

Sub SortSelectionDesc()
    Selection.Sort SortOrder:=wdSortOrderDescending
End Sub

Sub ShowContractForm()
    ContractForm.Show
End Sub

Sub KeepOnlyVBAComments()
    Dim para As Paragraph
    Dim strLine As String
    Dim lngPos As Long
    Dim rngPara As Range
    Dim strComment As String
    
    ' �������� �� ���� ������� � ����� (����� �������� �� ������ ����)
    For i = ActiveDocument.Paragraphs.Count To 1 Step -1
        Set rngPara = ActiveDocument.Paragraphs(i).Range
        strLine = rngPara.Text
        strLine = Replace(strLine, vbCr, "") ' ������� �������
        
        ' ���� �������� (������ �����������)
        lngPos = InStr(strLine, "'")
        
        If lngPos > 0 Then
            ' ����� �����������
            strComment = Mid(strLine, lngPos)
            rngPara.Text = strComment & vbCr
        Else
            ' ������� ������ ��� �����������
            rngPara.Delete
        End If
    Next i
    With ActiveDocument.Content.Find
        .Text = "^p'"
        .Replacement.Text = "^p"
        .MatchCase = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub

