Attribute VB_Name = "m_01_07_Arrays"
Option Explicit
Dim marrStockBalances(5 To 8) As Integer
Dim marrCategories(1 To 3) As TCategory

'����������� ���������� ������
Sub d_01_StaticOneDimensionalArray()
    Dim arrStockBalances(5) As Integer
    arrStockBalances(0) = 1
    arrStockBalances(1) = 101
    arrStockBalances(2) = 201
    arrStockBalances(3) = 301
    arrStockBalances(4) = 401
    arrStockBalances(5) = 501
    'arrStockBalances(6) = 601 '������
    Debug.Print arrStockBalances(0), arrStockBalances(4), arrStockBalances(5)
    Debug.Print TypeName(arrStockBalances)
    
End Sub

'����������� ���������� ������
Sub d_02_StaticOneDimensionalArray()

'    arrStockBalances(4) = 1   '������
    marrStockBalances(5) = 101
    marrStockBalances(6) = 201
    marrStockBalances(7) = 301
    marrStockBalances(8) = 401
'   arrStockBalances(9) = 501 '������

End Sub

' ������� ��������� �������
Sub d_03_ForEach()
    
    Dim varItem As Variant
    For Each varItem In marrStockBalances
        Debug.Print varItem
    Next varItem
    
End Sub


' ��������� ��������� �������
Sub d_04_For()
    Dim lngCounter As Long
    
    ' ��������� ��������� �������
    For lngCounter = LBound(marrStockBalances) To UBound(marrStockBalances)
        marrStockBalances(lngCounter) = 500
    Next lngCounter
    
    Debug.Print "�������� ����������. ��������� ����������: " & _
                UBound(marrStockBalances) - LBound(marrStockBalances) + 1
End Sub

'����������� ���������� ������ UDT
Sub d_04_ArrayUDT()

    With marrCategories(1)
        .CategoryCode = "MSO"
        .Name = "Microsoft Office"
    End With
    With marrCategories(2)
        .CategoryCode = "Prog"
        .Name = "����������������"
    End With
    With marrCategories(3)
        .CategoryCode = "DB"
        .Name = "���� ������"
    End With
End Sub


' ������� ��������� ������� UDT
Sub d_05_For()
    Dim lngCounter As Long
    
    For lngCounter = LBound(marrCategories) To UBound(marrCategories)
        With marrCategories(lngCounter)
            Debug.Print .CategoryCode, .Name
        End With
    Next lngCounter
    
End Sub


' ����������� ��������� ������: [���������, �������]
Sub d_04_StaticTwoDimensionalArray()

    Dim arrSalesData(1 To 2, 1 To 2) As Double
    Dim arrRegionNames(1 To 2) As String
    Dim arrCategoryNames(1 To 2) As String
    ' ������������� �������� ��������
    arrRegionNames(1) = "�����������"
    arrRegionNames(2) = "������-��������"
    ' ������������� �������� ���������
    arrCategoryNames(1) = "���������"
    arrCategoryNames(2) = "��������"
    ' ���������� ������� ��������� ������� (������� � ��� ���.)
    arrSalesData(1, 1) = 120.5  ' ��������� - �����������
    arrSalesData(1, 2) = 85.2   ' ��������� - ������-��������
    arrSalesData(2, 1) = 89.7   ' �������� - �����������
    arrSalesData(2, 2) = 67.4   ' �������� - ������-��������

End Sub


Sub ������������������()
    On Error GoTo ErrorHandler
    Dim aTemp() As Integer
    'aTemp(0) = 01
    ReDim aTemp(1)
    aTemp(0) = 1
    aTemp(1) = 11
    Debug.Print aTemp(0), aTemp(1)
    ReDim aTemp(2)
    Debug.Print aTemp(0), aTemp(1)
    Debug.Print TypeName(aTemp), VarType(aTemp)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub

Sub ������������������������������()
    On Error GoTo ErrorHandler
    Dim aTemp() As Integer
    'aTemp(0) = 01
    ReDim aTemp(1)
    aTemp(0) = 1
    aTemp(1) = 11
    Debug.Print aTemp(0), aTemp(1)
    ReDim Preserve aTemp(2)
    Debug.Print aTemp(0), aTemp(1)
    Debug.Print TypeName(aTemp), VarType(aTemp)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub


Sub ������������������������������()
    On Error GoTo ErrorHandler
    Dim aTemp(1) As Integer
    Dim aNewTemp() As Integer
    ReDim aNewTemp(1)
    aTemp(0) = 1
    aTemp(1) = 11
    Debug.Print aTemp(0), aTemp(1)
    aNewTemp = aTemp
    Debug.Print aNewTemp(0), aNewTemp(1)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub


Sub �������������������������������()
    On Error GoTo ErrorHandler
    Dim aTemp() As Integer
    Dim aNewTemp() As Integer
    ReDim aTemp(1)
    ReDim aNewTemp(1)
    aTemp(0) = 1
    aTemp(1) = 11
    Debug.Print aTemp(0), aTemp(1)
    aNewTemp = aTemp
    Debug.Print aNewTemp(0), aNewTemp(1)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub

Sub ���������������������������������()
    On Error GoTo ErrorHandler
    Dim aTemp(10) As Integer
    Dim aNewTemp() As Integer
    ReDim aNewTemp(1)
    aTemp(0) = 1
    aTemp(1) = 11
    aTemp(2) = 21
    aTemp(3) = 21
    Debug.Print aTemp(0), aTemp(1)
    aNewTemp = aTemp
    Debug.Print aNewTemp(0), aNewTemp(1)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub

Sub ��������������������������������������()
    On Error GoTo ErrorHandler
    Dim aTemp(10) As Integer
    Dim aNewTemp() As String
    ReDim aNewTemp(1)
    aTemp(0) = 1
    aTemp(1) = 11
    aTemp(2) = 21
    aTemp(3) = 31
    aNewTemp(0) = aTemp(0)
    Debug.Print aTemp(0), aTemp(1)
    'aNewTemp = aTemp
    Debug.Print aNewTemp(0), aNewTemp(1)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub


Sub StaticArrayErasion()
    On Error GoTo ErrorHandler
    Dim aTemp(10) As Integer
    aTemp(0) = 1
    aTemp(1) = 11
    aTemp(2) = 21
    aTemp(3) = 31
    Erase aTemp
    Debug.Print aTemp(0), aTemp(1)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub


Sub DynamicArrayErasion()
    On Error GoTo ErrorHandler
    Dim aTemp() As Integer
    ReDim aTemp(10)
    aTemp(0) = 1
    aTemp(1) = 11
    aTemp(2) = 21
    aTemp(3) = 31
    Erase aTemp
    Debug.Print aTemp(0), aTemp(1)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / "; Err.Description
End Sub
