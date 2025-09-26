Attribute VB_Name = "m_01_03_DataTypes"
Option Explicit
' % - Integer
' & - Long
' ! - Single
' # - Double
' @ - Currency'

'����� ���� ������ ��� ���������� ��� ���������
'����������� �������������� ��������
Sub d_01_ChooseNumeric()
    
    Dim lngCounter As Long
    lngCounter = lngCounter + 1
    Debug.Print lngCounter
    
End Sub

'����������� �������� � ������
Sub d_03_ChooseDate()
    
    Dim dtShipDate As Date
    dtShipDate = Date + 1
    Debug.Print dtShipDate
    
End Sub

'����� ���������� ���� � ���������
Sub d_04_ChooseString()
    
    Dim strAccount As String
    strAccount = "4080281000000000000"
    Debug.Print Left(strAccount, 3)
    
End Sub

' ����������� ������������ � ������ ������ ���������
Sub d_05_ChooseBoolean()
    Dim boolGoldStatus As Boolean
    boolGoldStatus = True
    If boolGoldStatus Then
        Debug.Print "��������� ������"
    End If
End Sub


'����� �������� ��� �������? ������������� ����
'Byte ������������ ��� �������� �������� ����������, �� ����� ������� �����
Sub d_07_Byte()
'������������ � ����������� �������� ���� Byte: 0-255

    Dim byteVar As Byte
    byteVar = 255
    Debug.Print TypeName(byteVar), byteVar

End Sub


Sub d_08_Integer()
'������������ � ����������� �������� ����:-32,768 - 32,767, 2 �����
    
    Dim intVar As Integer
    intVar = 35000
    Debug.Print TypeName(intVar), intVar

End Sub


Sub d_08_long()
'4 bytes -2,147,483,648 to 2,147,483,647
    
    Dim lngVar As Long
    lngVar = 123456
    Debug.Print TypeName(lngVar), lngVar

End Sub


Sub d_09_LongLong()
'8 bytes     -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807
    
'    Dim llVar As LongLong
'    llVar = 123456
'    Debug.Print TypeName(llVar), llVar

End Sub


Sub d_10_SpeedOfIntegers()
    Const lngIterations = 100000000
    Dim byteVar As Byte
    Dim intVar As Integer
    Dim lngVar As Long
    Dim dblStartTime As Double
    Dim lngCounter As Long
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        byteVar = (byteVar + 1) * 2 / 2 - 1
    Next lngCounter
    Debug.Print "Byte"; Timer - dblStartTime
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        intVar = (intVar + 1) * 2 / 2 - 1
    Next lngCounter
    Debug.Print "Integer"; Timer - dblStartTime
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        lngVar = (lngVar + 1&) * 2& / 2& - 1&
    Next lngCounter
    Debug.Print "Long"; Timer - dblStartTime
End Sub


Sub d_11_SpeedOfIntegers_2()
    Const lngIterations = 100000000
    Dim intVar As Integer
    Dim lngVar As Long
    Dim dblStartTime As Double
    Dim lngCounter As Long
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        intVar = (CLng(intVar) + 1) * 2 / 2 - 1
    Next lngCounter
    Debug.Print "Integer c ���������������"; Timer - dblStartTime
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        intVar = (intVar + 1) * 2 / 2 - 1
    Next lngCounter
    Debug.Print "Integer ��� ��������������"; Timer - dblStartTime
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        lngVar = (lngVar + 1&) * 2& / 2& - 1&
    Next lngCounter
    Debug.Print "Long"; Timer - dblStartTime

End Sub

'���� ������ � ��������� ������. Single! 4 bytes
Sub d_12_SingleDataType()
    Dim sngXValue As Single
    sngXValue = 1.23456789
    Debug.Print TypeName(sngXValue), sngXValue
End Sub

'���������� ���� ������ Single
Sub d_13_SingleDataType()
        
    Dim sngXValue As Single
    For sngXValue = 0! To 1! Step 0.1!
        Debug.Print sngXValue
    Next sngXValue

End Sub


'������ �������� ���� ������ Single
Sub d_14_SingleDataType()
        
    Dim sngBadForMoney As Single
    sngBadForMoney = 1234567890.12
                  '1234568000      1,234568E+09 7 �������� ��������
    Debug.Print sngBadForMoney

End Sub



'��� ������ Double' #  8 bytes
Sub d_15_DoubleDataType()
    Dim dblValue As Double
    dblValue = 1.23456789
    Debug.Print TypeName(dblValue), dblValue
End Sub


'��� ������ Double. ������ ����������
Sub d_16_DoubleDataType()

    Dim dblXValue As Double
    For dblXValue = 0 To 1 Step 0.01
        Debug.Print dblXValue
    Next dblXValue

End Sub


'������ �������� ���� ������ Double
Sub d_17_DoubleDataType()
        
    Dim dblBadForMoney As Double
    dblBadForMoney = "12345678901234567890"
                         '12345678901234600000 1,23456789012346E+19
                         '15 �������� ��������
    Debug.Print dblBadForMoney

End Sub

'��� ������ Currency - @ 8 bytes, �������������� ����� �����
Sub d_18_Currency()
'-922,337,203,685,477.5808 to 922,337,203,685,477.5807
'���������� ���������� ���������� (Round-half-to-even) ���������� � ���������� �������

    Dim curForMoney As Currency
    curForMoney = "123456789012345,1234"  '������� 15.4
    Debug.Print curForMoney

End Sub

'��� ������ Currency. ��� ������ ����������
Sub d_19_Currency()
    
    Dim curXValue As Currency
    For curXValue = 0 To 1 Step 0.01
        Debug.Print curXValue
    Next curXValue
End Sub

'��� ������ Decimal 14 bytes, 28 ��������
Sub d_20_DecimalDataType()
    
    Dim decValue As Variant
    decValue = CDec("12345678901234567890123456789")
    Debug.Print TypeName(decValue), decValue
    decValue = CDec("1234567890123,4567890123456789")
    Debug.Print TypeName(decValue), decValue
    Debug.Print TypeName(decValue), decValue + 1
    decValue = CDec("0,0000000000000000000000000001")
    Debug.Print TypeName(decValue), decValue, decValue + decValue

End Sub


'������������������ ������� ����� ������
Sub d_21_PerformanceOfDecimals()
    Const lngIterations = 100000000
    Dim sngVar As Single
    Dim dblVar As Double
    Dim curVar As Currency
    Dim decVar As Variant
    Dim dblStartTime As Double
    Dim lngCounter As Long
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        sngVar = (sngVar + 1) * 2 / 2 - 1
    Next lngCounter
    Debug.Print "Single"; Timer - dblStartTime
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        dblVar = (dblVar + 1) * 2 / 2 - 1
    Next lngCounter
    Debug.Print "Double"; Timer - dblStartTime
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        curVar = (curVar + 1) * 2 / 2 - 1
    Next lngCounter
    Debug.Print "Currency"; Timer - dblStartTime
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        decVar = (decVar + 1) * 2 / 2 - 1
    Next lngCounter
    Debug.Print "Decimal"; Timer - dblStartTime
End Sub

'���������� ��� ������. 2 bytes
Sub d_21_BooleanDataType()
    Dim BoolValue As Boolean
    BoolValue = True
    Debug.Print BoolValue
    BoolValue = 1 = 1
    Debug.Print BoolValue
    BoolValue = 1
    Debug.Print BoolValue
    BoolValue = 0
    Debug.Print BoolValue
End Sub

'��� ������ ���� � �������. 8 ����
Sub d_22_DateDataType()
    Dim dtValue As Date
    dtValue = #2/13/1972#
    Debug.Print dtValue
    dtValue = dtValue + 1
    Debug.Print dtValue
    dtValue = 1
    Debug.Print dtValue
    dtValue = -1
    Debug.Print dtValue
    dtValue = #2/13/1972 2:45:55 PM#
    Debug.Print dtValue
    Debug.Print Format(dtValue, "dd mmmm yyyy HH:nn:ss")
    Debug.Print Format(Date, "dd mmmm yyyy hh:nn:ss")
End Sub

'������ ���������� �����, 10 ������ + ����� ������, �� 0 �� �������������� 2 ����������
Sub d_23_StringDataTypes_VariableLength()
    
    Dim strVariableLength As String
    strVariableLength = "1234567890"
        Debug.Print "|" & strVariableLength & "|"
    strVariableLength = ""
        Debug.Print "|" & strVariableLength & "|"

End Sub
'������ ���������������� �����, ����� ������, �� 1 �� �������������� 65 400
Sub d_24_StringDataTypes_FixedLength()
    
    Dim strFixedLength As String * 8
    strFixedLength = "1234567890"
        Debug.Print "|" & strFixedLength & "|"
    strFixedLength = ""
        Debug.Print "|" & strFixedLength & "|"

End Sub

Sub d_25_VariantDataType()

    Dim varValue As Variant
    varValue = "��������� �������� "
    Debug.Print varValue & 1
    varValue = #2/13/1972#
    Debug.Print varValue + 1
    varValue = 100
    Debug.Print varValue + 1

End Sub


Sub d_27_VariantDataType()

    Dim varValue As Variant
    varValue = "��������� �������� "
    Debug.Print varValue & 1, TypeName(varValue)
    varValue = #2/13/1972#
    Debug.Print varValue + 1, TypeName(varValue)
    varValue = 100
    Debug.Print varValue + 1, TypeName(varValue)

End Sub

