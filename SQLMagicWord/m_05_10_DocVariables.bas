Attribute VB_Name = "m_05_10_DocVariables"
Option Explicit

'—оздаЄм копию файла "Visual Basic" дл€ экспериментов
Private Sub d_01_CopyFileForExperiment()
    On Error GoTo ErrorHandler
    Dim strSourceFilename As String, strTargetFilename As String
    strSourceFilename = "D:\VBA\ƒл€ чтени€ Word\ѕрограмма курса.docx"
    strTargetFilename = "D:\VBA\Word\ урс VBA.docx"
    Call CopyFile(SourceFilename:=strSourceFilename, TargetFilename:=strTargetFilename)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub d_02_WhatToStore()
'„то можно хранить в переменной документа?
'ћетаданные документа
'DocAuthor Ц автор документа
'DocEditor Ц последний редактор
'DocVersion Ц верси€ документа
'DocStatus Ц статус документа (Draft, Approved, Final)
'DocCreatedDate Ц дата создани€
'DocApprovedBy Ц кто утвердил

'ƒанные проекта / договора
'ProjectCode Ц код проекта
'ProjectName Ц название проекта
'ClientName Ц им€ заказчика
'ContractNumber Ц номер договора
'ContractDate Ц дата договора

'Ќастройки макросов / шаблона
'Style_FirstLineIndent Ц отступ первой строки
'Style_SpaceAfter Ц интервал после абзаца
'PrintComments Ц печатать ли комментарии (True/False)
'Lang_Main Ц основной €зык текста (например, en-US, ru-RU)
'ThemeMode Ц тема оформлени€ (Light/Dark)

'—лужебные и технические
'DocGUID Ц уникальный идентификатор документа
'LastCursorPos Ц позици€ курсора при последнем закрытии
'UpdateTOC Ц обновл€ть ли оглавление (True/False)
'ProcessedByMacro Ц обработан ли документ макросом
'ExportPath Ц путь дл€ выгрузки документа
End Sub

'ƒобавить переменные документа
Private Sub d_03_AddVariables()
    On Error GoTo ErrorHandler
    Dim docVBA As Document
    Set docVBA = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    With docVBA.Variables
        .Item("DocAuthorLastName").Value = "√урь€нов"
        .Item("DocAuthorPatronymic").Value = "ћихаил"
        .Item("DocAuthorFirstName").Value = "јлексеевич"
        .Item("DocVersion").Value = "0.1"
        .Item("DocStatus").Value = "„ерновик"
        .Item("ProjectCode").Value = "VBACourse"
        .Item("PrintComments").Value = "False"
    End With
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'ѕросмотреть все переменные документа
Private Sub d_04_ShowVariables()
    On Error GoTo ErrorHandler
    Dim varItem As Variant 'ƒл€ использовани€ for each
    'ѕеременна€ документа всегда содержит строку
    Dim docVBA As Document
    Set docVBA = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    For Each varItem In docVBA.Variables
        Debug.Print varItem.Name & " = " & varItem.Value
    Next varItem
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'ќбращение к переменной документа
Private Sub d_05_GetVariableValue()
    On Error GoTo ErrorHandler
    Dim strItem As String
    Dim docVBA As Document
    Set docVBA = Documents.Open("D:\VBA\Word\Visual Basic.docx")
    Debug.Print "DocStatus="; docVBA.Variables("DocStatus")
    Debug.Print "ProjectCode="; docVBA.Variables("ProjectCode")
    Debug.Print "DocVersion="; docVBA.Variables("DocVersion")
    docVBA.Variables("DocVersion") = "0.2"
    Debug.Print "DocVersion="; docVBA.Variables("DocVersion")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



