VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event evtIcorrectBirthDate(IcorrectBirthDate As Date)
Dim strFirstName As String
Dim strLastName As String
Dim strGender As String
Dim dtBirthDate As Date
Dim boolBirthDateIsUnknown As Boolean


Public Property Let FirstName(NewFirstName As String)
    strFirstName = StrConv(NewFirstName, vbProperCase)
End Property

Public Property Let LastName(NewLastName As String)
    strLastName = StrConv(NewLastName, vbProperCase)
End Property

Public Property Let Gender(NewGender As String)
    strGender = LCase(NewGender)
End Property

Public Property Let BirthDate(NewBirthDate As Date)
    Dim dtOnlyDate As Date
    dtOnlyDate = DateValue(NewBirthDate)
    If dtOnlyDate > Date Then
        RaiseEvent evtIcorrectBirthDate(dtOnlyDate)
    Else
        dtBirthDate = dtOnlyDate
        boolBirthDateIsUnknown = False
    End If
    
End Property

Public Property Let BirthDateIsUnknown(NewBirthDateIsUnknown As Boolean)
    boolBirthDateIsUnknown = NewBirthDateIsUnknown
End Property

Public Property Get FirstName() As String
    FirstName = strFirstName
End Property
Public Property Get LastName() As String
    LastName = strLastName
End Property


Public Property Get Gender() As String
    Gender = strGender
End Property

Public Property Get BirthDate() As Date
    BirthDate = dtBirthDate
End Property

Public Property Get BirthDateIsUnknown() As Boolean
    BirthDateIsUnknown = boolBirthDateIsUnknown
End Property

Public Property Get fullName() As String
    fullName = FirstName & " " & LastName
    fullName = Trim(fullName)
End Property

Public Property Get FullYears() As Integer
    FullYears = GetFullYears()
End Property

Sub PrintForm()
    Debug.Print GetForm()
End Sub

Function GetForm() As String
    Dim strResult As String
    Dim strBirthDate As String
    If boolBirthDateIsUnknown Then
        strBirthDate = ""
    Else
        strBirthDate = Format(dtBirthDate, "dd mmmm yyyy г.")
    End If
    strResult = "-----------------------------------" & vbCrLf
    strResult = strResult + "Имя          : " & Me.FirstName & vbCrLf
    strResult = strResult + "Фамилия      : " & Me.LastName & vbCrLf
    strResult = strResult + "Пол          : " & Me.Gender & vbCrLf
    strResult = strResult + "Дата рождения: " & strBirthDate
    GetForm = strResult
End Function

Function GetFullYears(Optional TargetDate As Date = 0) As Integer
    If boolBirthDateIsUnknown Then
        GetFullYears = -1
        Exit Function
    End If
    If TargetDate = 0 Then TargetDate = Date Else TargetDate = DateValue(TargetDate)
    If TargetDate < BirthDate Then
        GetFullYears = -2
        Exit Function
    End If
    Dim Result As Integer
    ' Вычисляем разницу в годах
    Result = Year(TargetDate) - Year(BirthDate)
    ' Корректируем, если день рождения в этом году еще не наступил
    If DateSerial(Year(TargetDate), Month(BirthDate), Day(BirthDate)) > TargetDate Then
        Result = Result - 1
    End If
    GetFullYears = Result
End Function

Private Sub Class_Initialize()
    boolBirthDateIsUnknown = True
End Sub

'Private Sub Class_Terminate()
'    Debug.Print "Объект будет уничтожен"
'End Sub

Sub AddToFile(ByVal FileName As String)
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Open FileName For Append As #intFile
    Write #intFile, strFirstName, strLastName, strGender, dtBirthDate, boolBirthDateIsUnknown
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

Sub ReadFirstFromFile(ByVal FileName As String)
    On Error GoTo ErrorHandler
    Dim intFile As Integer: intFile = FreeFile
    Open FileName For Input As #intFile
    Input #intFile, strFirstName, strLastName, strGender, dtBirthDate, boolBirthDateIsUnknown
Finalization:
    Close #intFile
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

