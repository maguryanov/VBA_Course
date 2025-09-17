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

' Класс: Person
' Приватные поля класса
Private lngID As Long
Private strLastName As String
Private strFirstName As String
Private strPatronymic As String
Private strPhone As String
Private strEmail As String
Private strCity As String
Private strAddress As String
Private dtBirthDate As Date
Private strGender As String
Private strPassportNumber As String
Private strPassportSeries As String
Private strPassportIssuedBy As String

' Свойство для ID
Public Property Get ID() As Long
    ID = lngID
End Property

Public Property Let ID(ByVal value As Long)
    lngID = value
End Property

' Свойство для LastName
Public Property Get LastName() As String
    LastName = strLastName
End Property

Public Property Let LastName(ByVal value As String)
    strLastName = value
End Property

' Свойство для FirstName
Public Property Get FirstName() As String
    FirstName = strFirstName
End Property

Public Property Let FirstName(ByVal value As String)
    strFirstName = value
End Property

' Свойство для Patronymic
Public Property Get Patronymic() As String
    Patronymic = strPatronymic
End Property

Public Property Let Patronymic(ByVal value As String)
    strPatronymic = value
End Property

' Свойство для Phone
Public Property Get Phone() As String
    Phone = strPhone
End Property

Public Property Let Phone(ByVal value As String)
    strPhone = value
End Property

' Свойство для Email
Public Property Get Email() As String
    Email = strEmail
End Property

Public Property Let Email(ByVal value As String)
    strEmail = value
End Property

' Свойство для City
Public Property Get City() As String
    City = strCity
End Property

Public Property Let City(ByVal value As String)
    strCity = value
End Property

' Свойство для Address
Public Property Get Address() As String
    Address = strAddress
End Property

Public Property Let Address(ByVal value As String)
    strAddress = value
End Property

' Свойство для BirthDate
Public Property Get BirthDate() As Date
    BirthDate = dtBirthDate
End Property

Public Property Let BirthDate(ByVal value As Date)
    dtBirthDate = value
End Property

' Свойство для Gender
Public Property Get Gender() As String
    Gender = strGender
End Property

Public Property Let Gender(ByVal value As String)
    strGender = value
End Property

' Свойство для PassportNumber
Public Property Get PassportNumber() As String
    PassportNumber = strPassportNumber
End Property

Public Property Let PassportNumber(ByVal value As String)
    strPassportNumber = value
End Property

' Свойство для PassportSeries
Public Property Get PassportSeries() As String
    PassportSeries = strPassportSeries
End Property

Public Property Let PassportSeries(ByVal value As String)
    strPassportSeries = value
End Property

' Свойство для PassportIssuedBy
Public Property Get PassportIssuedBy() As String
    PassportIssuedBy = strPassportIssuedBy
End Property

Public Property Let PassportIssuedBy(ByVal value As String)
    strPassportIssuedBy = value
End Property
