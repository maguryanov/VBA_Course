Attribute VB_Name = "m_01_05_Literal"
Option Explicit

'������� - ��� ������ �������� �������� � ����.
Sub d_01_Literal()

    Debug.Print "�����: "; 2144.48; " ���."

End Sub


'� �������� ���� ��� ������
Sub d_02_DataType()

    Debug.Print TypeName("�����: ")
    Debug.Print TypeName(2144.48)
    
End Sub

'����� ����
Sub d_03_IntegerLiteral()
    
    'Integer 2 �����     �� �32 768 �� 32 767
    Debug.Print 10, TypeName(10)
    
    'Long 4 �����     �� �2 147 483 648 �� 2 147 483 647
    Debug.Print 100500, TypeName(100500)

End Sub

'����� � ��������� ������
Sub d_04_FloatPointLiteral()
    
    'Double (����� � ��������� ������ ������� ��������)     8 ����
    Debug.Print 36.6; TypeName(36.6)
    
    'Single (����� � ��������� ������ ��������� ��������)   4 �����
    Debug.Print 36.6!; TypeName(36.6!)
    

End Sub

'����� � ������������� ������
Sub d_05_FixedPointLiteral()
    
    'Currency (�������������� ����� �����)   8 ����
    Debug.Print TypeName(1200.57@); 1200.57@
    Debug.Print TypeName(123456789012345.1235@); 1200.57@

End Sub

'������
Sub d_06_StringLiteral()
    
    Debug.Print TypeName("������ ��������"), "������ ��������"
    Debug.Print TypeName("MA-144F5"), "MA-144F5"
    Debug.Print TypeName("4080281000000000000"), "4080281000000000000"
    Debug.Print TypeName("00123"), "00123"
    Debug.Print TypeName("123456,12345"), "123456,12345"
    Debug.Print TypeName(""), "|"; ""; "|"
    Debug.Print "���� ""������ VBA"""
    
End Sub

'����/����� �������
Sub d_07_DateLiteral()
    
    Debug.Print TypeName(#9/23/2025#), #9/23/2025#
    Debug.Print "����", #9/23/2025#
    Debug.Print "�����", #10:30:44 PM#
    Debug.Print "���� � �����", #9/23/2025 10:30:00 PM#
    Debug.Print #9/23/2025# + 1              '����
    
End Sub

'���������� ���
Sub d_08_BooleanLiteral()
    
    Debug.Print TypeName(True), True
    Debug.Print TypeName(False), False

End Sub

'��� ������ ����� ������������ �������������
Sub d_09_AutomaticType()

    Debug.Print TypeName(100); 100
    Debug.Print TypeName(100500); 100500
    
End Sub

'��� ����� ����������� ����
Sub d_10_ExplicitType()
    Debug.Print TypeName(100), 100        ' % - Integer"
    Debug.Print TypeName(100&), 100&     ' & - Long"
    Debug.Print TypeName(100.5!), 100.5! ' ! - Single"
    Debug.Print TypeName(100#), 100#     ' # - Double"
    Debug.Print TypeName(200@), 200@     ' @ - Currency
    
End Sub

'������ ������� ��������� ��� ����?
Sub d_11_ExplicitType()
    
    '������������� ���� �� ��������� ����� ��������� � �������
    Debug.Print 20000& + 20000
    
End Sub


'������ ������� ��������� ��� ����?
Sub d_12_ExplicitType()
    
    '�������� � �������������� ����������� �� ������ ������ ��������
    Debug.Print 40000
    Debug.Print 12345678
    
End Sub



'������ ������� ��������� ��� ����?
Sub d_14_ExplicitType()
            
    'Currency � Single ������������� �� ������������
    Debug.Print 499.99
    Debug.Print 36.6!
        
End Sub


'������ ������� ��������� ��� ����?
Sub d_15_ExplicitType()
    
    '���� ����� Double, � ����� �� �������� ������� �����
    Debug.Print 100#
    
End Sub


'����������������� � ������������ ��������
Sub d_16_HexAndOctalLiterals()

    Debug.Print "�������� &HF: "; TypeName(&HF); &HF
    Debug.Print "�������� &O10: "; &O10
    Debug.Print "�������� &HFFFF: "; TypeName(&HFFFF), &HFFFF
    
    Debug.Print "Hex(0) = "; Hex(0)
    Debug.Print "Hex(255) = "; Hex(255)
    Debug.Print "Oct(8) = "; Oct(8)
End Sub


'���������������� ��� ������� ������
Sub d_17_ExponentialLiteral()

    Debug.Print 1.23E+20; "��� ������"; 1.23 * 10 ^ 20
    Debug.Print 1.23E-20; "��� ������"; 1.23 / 10 ^ 20
    '�������������� � ����� VBA ����� ���� ��������� ���������� ���� � ����� �����!
    Debug.Print "12345678!" = 1.234568E+07!
    
End Sub

'�������� ������ ��� ������ � ����������
Sub d_18_TypicalErrors()

    'Debug.Print ���������
    'Debug.Print "������� "�����������""
    'Debug.Print 32767 + 1
    'Debug.Print 1000 + "$"
    'Debug.Print "12345678!" = 12345678!

End Sub


