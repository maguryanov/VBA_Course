VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PeopleWithAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Класс: People
' Описание: Коллекция объектов Person с использованием Dictionary и поддержкой For Each

Private dictPeople As Object ' Scripting.Dictionary
Private colOrder As Collection ' Для сохранения порядка элементов

Private Sub Class_Initialize()
    Set dictPeople = CreateObject("Scripting.Dictionary")
    Set colOrder = New Collection
End Sub

Private Sub Class_Terminate()
    Set dictPeople = Nothing
    Set colOrder = Nothing
End Sub

' Добавление объекта Person в коллекцию
Public Sub Add(objPerson As Person)
    Dim strKey As String
    strKey = CStr(objPerson.ID)
    
    If Not dictPeople.Exists(strKey) Then
        dictPeople.Add strKey, objPerson
        colOrder.Add objPerson, strKey
    End If
End Sub

' Удаление объекта из коллекции по индексу
Public Sub RemoveByIndex(intIndex As Integer)
    If intIndex > 0 And intIndex <= colOrder.Count Then
        Dim objPerson As Person
        Set objPerson = colOrder(intIndex)
        dictPeople.Remove CStr(objPerson.ID)
        colOrder.Remove intIndex
    End If
End Sub

' Удаление объекта из коллекции по ID
Public Sub RemoveByID(lngID As Long)
    Dim strKey As String
    strKey = CStr(lngID)
    
    If dictPeople.Exists(strKey) Then
        dictPeople.Remove strKey
        
        ' Удаляем объект из коллекции порядка
        Dim i As Integer
        For i = 1 To colOrder.Count
            If colOrder(i).ID = lngID Then
                colOrder.Remove i
                Exit For
            End If
        Next i
    End If
End Sub

' Получение количества элементов в коллекции
Public Property Get Count() As Long
    Count = dictPeople.Count
End Property

' Обращение к элементу по индексу (People(1))
Public Property Get Item(intIndex As Integer) As Person
    Attribute Item.VB_UserMemId = 0
    If intIndex > 0 And intIndex <= colOrder.Count Then
        Set Item = colOrder(intIndex)
    Else
        Set Item = Nothing
    End If
End Property

' Обращение к элементу по ID (People("1"))
Public Property Get ItemByID(lngID As Long) As Person
    Dim strKey As String
    strKey = CStr(lngID)
    
    If dictPeople.Exists(strKey) Then
        Set ItemByID = dictPeople(strKey)
    Else
        Set ItemByID = Nothing
    End If
End Property

' Проверка существования элемента по ID
Public Function Exists(lngID As Long) As Boolean
    Exists = dictPeople.Exists(CStr(lngID))
End Function

' Реализация поддержки For Each (правильная версия)
Public Function NewEnum() As IUnknown
    Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = colOrder.[_NewEnum]
End Function
