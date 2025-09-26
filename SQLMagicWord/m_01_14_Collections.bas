Attribute VB_Name = "m_01_14_Collections"
Option Explicit
Private moPeople As Collection

Sub Collection_01()
    Dim oPerson As Person
    Set oPerson = New Person
    oPerson.LastName = "��������"
    oPerson.FirstName = "������"
    oPerson.Gender = "�"
    oPerson.BirthDate = #2/13/1972#
    Dim oPeople As Collection
    Set oPeople = New Collection
    oPeople.Add oPerson
    Set oPerson = New Person
    oPerson.LastName = "������"
    oPerson.FirstName = "�������"
    oPerson.Gender = "f"
    oPerson.BirthDate = #3/3/2000#
    oPeople.Add oPerson
    PrintPeopleCollection oPeople
End Sub

Sub PrintPeopleCollection(People As Collection)
    Dim oPerson As Person
    For Each oPerson In People
        oPerson.PrintForm
    Next oPerson
End Sub

Sub PrintPeopleFullName(People As Collection)
    Dim oPerson As Person
    For Each oPerson In People
        Debug.Print oPerson.fullName
    Next oPerson
End Sub

Sub Collection_02()
    Dim oPerson As Person
    Set oPerson = New Person
    Dim oPeople As Collection
    Set oPeople = New Collection
    With oPerson
        .LastName = "��������"
        .FirstName = "������"
        .Gender = "�"
        .BirthDate = #2/13/1972#
    End With
    oPeople.Add oPerson
    Set oPerson = New Person
    With oPerson
        .LastName = "������"
        .FirstName = "�������"
        .Gender = "f"
        .BirthDate = #3/3/2000#
    End With
    oPeople.Add oPerson
    PrintPeopleCollection oPeople
End Sub


Sub Collection_WithKey()
    Dim oPerson As Person
    Set oPerson = New Person
    Dim oPeople As Collection
    Set oPeople = New Collection
    With oPerson
        .LastName = "��������"
        .FirstName = "������"
        .Gender = "�"
        .BirthDate = #2/13/1972#
    End With
    oPeople.Add oPerson, Key:="��������"
    Set oPerson = New Person
    With oPerson
        .LastName = "������"
        .FirstName = "�������"
        .Gender = "f"
        .BirthDate = #3/3/2000#
    End With
'    oPeople.Add oPerson, Key:="������"
'    oPeople.Item(1).PrintForm
'    oPeople.Item("������").PrintForm
'    oPeople(1).PrintForm
    oPeople("������").PrintForm
    Collection_GetByIndex_1 oPeople, 2
    Collection_GetByIndex_2 oPeople, 2
'    Collection_GetByKey_1 oPeople, "������"
'    Collection_GetByKey_2 oPeople, "������"
End Sub

' ������ ������
Sub Collection_GetByIndex_1(People As Collection, Idx As Long)
    People.Item(Idx).PrintForm
End Sub

Sub Collection_GetByKey_1(People As Collection, Key As String)
    People.Item(Key).PrintForm
End Sub


' ������ ������
Sub Collection_GetByIndex_2(People As Collection, Idx As Long)
    People(Idx).PrintForm
End Sub

Sub Collection_GetByKey_2(People As Collection, Key As String)
    People(Key).PrintForm
End Sub


Sub Collection_Deletion()
    Dim oPerson As Person
    Set oPerson = New Person
    Dim oPeople As Collection
    Set oPeople = New Collection
    With oPerson
        .LastName = "��������"
        .FirstName = "������"
        .Gender = "�"
        .BirthDate = #2/13/1972#
    End With
    oPeople.Add oPerson, Key:="��������"
    Set oPerson = New Person
    With oPerson
        .LastName = "������"
        .FirstName = "�������"
        .Gender = "f"
        .BirthDate = #3/3/2000#
    End With
    oPeople.Add oPerson, Key:="������"
    Set oPerson = New Person
    With oPerson
        .LastName = "������"
        .FirstName = "���������"
        .Gender = "m"
        .BirthDate = Date
    End With
    oPeople.Add oPerson, Key:="�������"
    PrintPeopleFullName oPeople
    oPeople("�������").PrintForm
    oPeople.Remove ("�������")
    PrintPeopleFullName oPeople
End Sub


Sub ModuleCollection_Creation()
On Error GoTo ErrorHandler
    Set moPeople = New Collection
    Dim oPerson As Person
    Set oPerson = New Person
    With oPerson
        .LastName = "��������"
        .FirstName = "������"
        .Gender = "�"
        .BirthDate = #2/13/1972#
    End With
    moPeople.Add oPerson, Key:=oPerson.fullName
    Set oPerson = New Person
    With oPerson
        .LastName = "������"
        .FirstName = "�������"
        .Gender = "f"
        .BirthDate = #3/3/2000#
    End With
    moPeople.Add oPerson, Key:=oPerson.fullName
    Set oPerson = New Person
    With oPerson
        .FirstName = "���������"
        .LastName = "������"
        .Gender = "m"
        .BirthDate = Date
    End With
    moPeople.Add oPerson, Key:=oPerson.fullName
    Debug.Print "��������� �������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub ModuleCollection_DuplicateKey()
On Error GoTo ErrorHandler
    Dim oPerson As Person
    Set oPerson = New Person
    With oPerson
        .LastName = "�������������"
        .FirstName = "����"
        .Gender = "�"
        .BirthDate = Date
    End With
    moPeople.Add oPerson, Key:="��������� ������"
    Debug.Print "������� ��������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub ModuleCollection_CorrectDateAdding()
On Error GoTo ErrorHandler
    Dim oPerson As Person
    Set oPerson = New Person
    With oPerson
        .FirstName = "�������"
        .LastName = "�������"
        .Gender = "�"
        .BirthDate = Date - 15000
    End With
    moPeople.Add oPerson, Key:=oPerson.fullName
    Debug.Print "������� ��������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub ModuleCollection_Viewing()
On Error GoTo ErrorHandler
    PrintPeopleFullName moPeople
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Sub ModuleCollection_Getting_01()
On Error GoTo ErrorHandler
    Dim oPerson As Person
    Set oPerson = moPeople("������� ������")
    oPerson.PrintForm
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub ModuleCollection_ItemRemoving()
On Error GoTo ErrorHandler
    moPeople.Remove ("��������� ������")
    Debug.Print "������ �������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Sub ModuleCollection_CollectionRemoving()
On Error GoTo ErrorHandler
    Set moPeople = Nothing
    Debug.Print "��������� �������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

