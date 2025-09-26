Attribute VB_Name = "m_01_05_Variables"
'Option Explicit

'���������� - ��� ����������� ������� ������ ��� �������� ������.
Sub d_01_Variable()
    
    Dim lngProductQty As Long
    lngProductQty = 1000&
    Debug.Print lngProductQty
    
    '����� ������������ ���� ������������� ����������?
End Sub

'� ���������� ���� ���. ��� ������?
Sub d_02_VariableNameRules()
'����� ����� �� ����� ��������� 255 ������...

    Dim v As String
    
    Debug.Print "�������� ����������"
    
End Sub

'��� ����������. ���������
Sub d_03_Shadowing()
    
'    Dim ThisDocument As String
'    ThisDocument = "d:\Docs\report.docx"
    ThisDocument.Paragraphs(1).Alignment = wdAlignParagraphLeft
    
    Debug.Print "�������� ����������"

End Sub


'��� ����������. ��� �������������?
Sub d_04_VariableNameBestPractices()
    
    Dim Count As Long
    Dim Price As Currency
    Dim Active As Boolean
    Dim BirthDate As Date
    Dim Qty As Long
    Dim Temp As Double
    Dim Temperature As Double

    Dim LastName As String  'PascalCase
    Dim FirstName As String 'camelCase
    Dim full_name As String 'snake_case
    
    Debug.Print "�������� ����������"
End Sub

'��� ����������. ���������� �������
Sub d_05_VariableNameNotation()
    
    Dim strFirstName As String   ' str - ������
    Dim lngCount As Long         ' lng - long
    Dim dblTemperature As Double ' dbl - double
    Dim curPrice As Currency     ' cur - currency
    Dim boolActive As Boolean    ' bool - boolean
    Dim dtBirthDate As Date      ' dt - date
    
    Debug.Print "�������� ����������"
End Sub

'��� ����������. ������� ��������
Sub d_06_��������������������������()
    
    Dim str��� As String
    Dim int���������� As Integer
    Dim cur���� As Double
    Dim bool������� As Boolean
    Dim dt������������ As Date
    
    Debug.Print "�������� ����������"
End Sub

'� ���������� ���� ��������. �������� Let
Sub d_07_VariableValue()
    
    Dim strLastName As String
    
    Let strLastName = "�������"
    Debug.Print strLastName
    
    strLastName = "�������"
    Debug.Print strLastName
   
End Sub



'�������� Let. ��������
Sub d_08_Let()
    
    Dim curPrice As Currency
    
    curPrice = 900@
    Debug.Print curPrice
    
    curPrice = curPrice + 100@
    Debug.Print curPrice
    
    curPrice = curPrice * 0.9@
    Debug.Print curPrice
   
End Sub


' �������� �� ��������� ��� ������ ����� ����������
Sub d_09_DefaultValues()
    
    Dim byteVar As Byte
    Dim intVar As Integer
    Dim lngVar As Long
    'Dim llVar As LongLong
    Dim sngVar As Single
    Dim dblVar As Double
    Dim curVar As Currency
    Dim strVarFix As String * 3
    Dim strVar As String
    Dim boolVar As Boolean
    Dim dtVar As Date
    Dim varVar As Variant
    Debug.Print TypeName(byteVar), byteVar
    Debug.Print TypeName(intVar), intVar
    Debug.Print TypeName(lngVar), lngVar
    'Debug.Print TypeName(llVar), llVar
    Debug.Print TypeName(sngVar), sngVar
    Debug.Print TypeName(dblVar), dblVar
    Debug.Print TypeName(curVar), curVar
    Debug.Print TypeName(strVarFix), "|" & strVarFix & "|"
    Debug.Print TypeName(strVar), "|" & strVar & "|"
    Debug.Print TypeName(boolVar), boolVar
    Debug.Print TypeName(dtVar), Format(dtVar, "dd.mm.yyyy hh:nn:ss")
    Debug.Print TypeName(varVar), varVar

End Sub


'� ���������� ���� ���. ���� �� ������ - Variant
Sub d_10_VariableType()
    
    Debug.Print Qty, TypeName(Qty)
    Qty = 1000&
    Debug.Print Qty, TypeName(Qty)
    Qty = "40��"
    Debug.Print Qty, TypeName(Qty)

End Sub


Sub d_11_SpeedOfVariant()
    Const lngIterations = 100000000
    Dim lngQty As Long
    Dim varQty As Variant
    
    Dim dblStartTime As Double
    Dim lngCounter As Long
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        lngQty = (lngQty + 1) * 2 / 2 - 1
    Next lngCounter
    
    Dim dblLong As Double: dblLong = Timer - dblStartTime
    Debug.Print "Long: "; dblLong
    
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        varQty = (varQty + 1) * 2 / 2 - 1
    Next lngCounter
    
    Dim dblVariant As Double: dblVariant = Timer - dblStartTime
    Debug.Print "Variant: "; dblVariant
    
    Debug.Print "Ratio: "; dblVariant / dblLong

End Sub



'��� ���������� ����� ���������� ���������. �� ��� ����
Sub d_12_VariableType()
    '%   Integer
    '&   Long
    '^   LongLong
    '@   Currency
    '!   Single
    '#   Double
    '$   String
    
    Debug.Print TypeName(ProductQty&), TypeName(Price@), TypeName(FirstName$)

End Sub



'������ ����� ������������� ����������? IntelliSence
Sub d_14_VariableDeclaration()
    
    'Dim dtInvoicePaymentDeadline As Date
    dtInvoicePaymentDeadline = #9/25/2025#
    Debug.Print
    
End Sub


'������ ����� ������������� ����������? Option Explicit
Sub d_15_VariableDeclaration()
    Dim Price As Currency
    'Pri�e = 1000@
    Debug.Print Price, TypeName(Price)

End Sub



'�������������� ����������. ��� Variant �� ���������
Sub d_16_VariantDefault()
    
    Dim Price
    Debug.Print Price, TypeName(Price)
    Price = 1000@
    Debug.Print Price, TypeName(Price)

End Sub


'�������������� ����������
Sub d_17_VariableDeclaration()

    Dim strProductName As String
    Dim strFirstName, strLastName As String
    Dim dtBirthDate As Date, strGender As String
    
    Dim lngCounter As Long: lngCounter = 0

End Sub


'��������� � ��� ����������� ��������, ������� �� ����� ���� �������� � ���� ���������� ���������.
Sub d_18_Constants()

    Const curIncomeTax As Currency = 0.13@
    
    Debug.Print "�����:" & (100000 * curIncomeTax)
    Debug.Print "�� ����:" & (100000 * (1 - curIncomeTax))
    
End Sub

'�� ��������� ��������
Sub d_19_Constants()

    Const curIncomeTax As Currency = 0.13
    'curIncomeTax = 0.15
    Debug.Print "�����: " & (100000 * curIncomeTax)
    Debug.Print "�� ����: " & (100000 * (1 - curIncomeTax))
    
End Sub

' ����� �� ���������� ��� ���������
Sub d_20_ConstantType()

    Const curIncomeTax = 0.13@
    
    Debug.Print curIncomeTax, TypeName(curIncomeTax)
    
End Sub


' ��������� ������ ���������� ���������
Sub d_21_StandardConstants()
    
    MsgBox "��������� ����������" & vbCrLf & "��������� �����"
    Debug.Print "��������� ����������" & vbTab & "���������"
    Debug.Print "vbRed = ", Hex(vbRed)
    Debug.Print "vbGreen = ", Hex(vbGreen)
    Debug.Print "vbBlue = ", Hex(vbBlue)
    
End Sub

' ������������� ��������� ��� ����������� ��������
Sub d_22_Expressions()

    Const curIncomeTaxPercent As Currency = 13@
    Const curNetIncomeRatio As Currency = 1@ - curIncomeTaxPercent / 100@
    Debug.Print 100000@ * curNetIncomeRatio

End Sub


Sub d_23_SpeedOfConstants()
    Const lngIterations = 100000000
    Const curConst As Currency = 0.87@
    Dim curVar As Currency: curVar = 0.87@
    Dim curResult As Currency
    Dim dblStartTime As Double
    Dim lngCounter As Long
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        curResult = 100000@ * curConst
    Next lngCounter
    
    Dim dblConst As Double: dblConst = Timer - dblStartTime
    Debug.Print "���������: ", dblConst
    
    dblStartTime = Timer
    For lngCounter = 1 To lngIterations
        curResult = 100000@ * curVar
    Next lngCounter
    
    Dim dblVar As Double: dblVar = Timer - dblStartTime
    Debug.Print "����������: ", dblVar
    
    Debug.Print "Ratio: ", dblVar / dblConst

End Sub


' ������ � �� �������
Sub d_24_TypicalErrors()
    
    '������������� �������� ���� � ���������������
    'Dim Set As String
    '������������� �������� � ���������������
    'Dim First Name As String
    '������������� ������� ����� ��� �������������
    'Dim 01Module As String
    'Dim _Value As Currency
    '��������� (Shadowing)
    'Dim ThisDocument As String
    'ThisDocument.Paragraphs(1).Alignment = wdAlignParagraphLeft
    
    Debug.Print "�������� ����������"
End Sub

' ������� ������
Sub d_25_BadPractices()

    '������������ �� �������������� �����
    Dim a, b, var1, var2 As Long
    '������������ ����������� ����������
    Dim vl As Long
    '��������� ������ � ����������
    Dim curPrice As Double
    
    '�� ������������� ����������
    
    '�� ������������ ����� Option Explicit
    
    '������������ ������������ Variant
    Dim lngQty

End Sub
