Attribute VB_Name = "Macroses"
Sub Balabolka2()
'
' Balabolka2 Макрос
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
' Balabolka3 Макрос
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

Sub Установить_язык_правоп_англ_для_слов_латиницей()
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
Sub Подпись()
Attribute Подпись.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Подпись"
'
' Подпись Макрос
'
'
    MsgBox ("Проверка")
    Selection.EndKey Unit:=wdStory
    Selection.TypeText Text:="Ответственный исполнитель Гурьянов М. А."
    Selection.TypeParagraph
    Selection.TypeText Text:="т. 55-55"
End Sub

Sub Подпись2()
'
' Подпись Макрос
'
'
    Selection.EndKey Unit:=wdStory
    Author = InputBox("Введите Ваше имя", "Запрос информации")
    Selection.TypeText Text:="Ответственный исполнитель " & Author
    Selection.TypeParagraph
    Selection.TypeText Text:="т. 55-55"
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
    
    ' Проходим по всем абзацам с конца (чтобы удаление не ломало цикл)
    For i = ActiveDocument.Paragraphs.Count To 1 Step -1
        Set rngPara = ActiveDocument.Paragraphs(i).Range
        strLine = rngPara.Text
        strLine = Replace(strLine, vbCr, "") ' убираем перенос
        
        ' Ищем апостроф (начало комментария)
        lngPos = InStr(strLine, "'")
        
        If lngPos > 0 Then
            ' Берем комментарий
            strComment = Mid(strLine, lngPos)
            rngPara.Text = strComment & vbCr
        Else
            ' Удаляем строку без комментария
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

