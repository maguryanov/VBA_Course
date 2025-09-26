Attribute VB_Name = "m_01_06_Scopes"
Option Explicit

Private Const MODULE_SCOPE_PRIVATE As String = _
    "��������� ��������� �� ������ ������ � ������������� Private"
Const MODULE_SCOPE As String = _
    "��������� ��������� �� ������ ������ ��� ������������"
Public Const PROJECT_SCOPE_PUBLIC As String = _
    "��������� ��������� �� ������ ������ � ������������� Public"

Public Const PROJECT_SCOPE_SAME_NAME As String = _
    "��������� � �������� ��������� ������� � ������ Scopes"

Private mstrModule_Private As String
Dim mstrModule_Dim As String
Public pstrProject_Public As String

Public strName As String
'Public strName As String


'� �������������� ���� ���������. ��������� ���������
Sub d_01_ConstantScope()

    Const PROCEDURE_SCOPE As String = "��������� ��������� �� ������ ���������"
    Debug.Print PROCEDURE_SCOPE

End Sub

'�������� ��������� �������������� �� ������� ���������
Sub d_02_ConstantVisibilityTest()

    'Debug.Print PROCEDURE_SCOPE
    Debug.Print MODULE_SCOPE
    Debug.Print MODULE_SCOPE_PRIVATE
    Debug.Print PROJECT_SCOPE_PUBLIC
    Debug.Print Normal.GLOBAL_SCOPE_PUBLIC
    'Debug.Print Normal.PROJECT_NORMAL_SCOPE_PUBLIC

End Sub


'� ���������� ��� �� ���� ���������. � �� �� �������
Sub d_03_Variables()
    
    Dim strLocal As String
    strLocal = "���������� ��������� �� ������ ���������"
    mstrModule_Private = "���������� ��������� �� ������ ������ � ������������� Private"
    mstrModule_Dim = "���������� ��������� �� ������ ������ � Dim"
    pstrProject_Public = "���������� ��������� �� ������ ������ � ������������� Public"
    
End Sub

' ��������� ����������
Sub d_04_VariablesVisibility()
    
'    Debug.Print strLocal
'    Debug.Print mstrModule_Private
'    Debug.Print mstrModule_Dim
'    Debug.Print pstrProject_Public
'    Debug.Print Normal.gstrGlobal
'    Debug.Print Normal.pstrNormalProject
        
    Debug.Print "��������� �������"
End Sub


'������� ��������� ���������� Public ������
Sub d_05_PublicLocal()
    
    'Public strLocal As String

End Sub

'��� ����� ���� ����� ����������� � ����� ��������� ��� ������
Sub d_06_SameNamesInProcedure()
        
    Dim strName As String
    
    'Dim strName As String

End Sub


'��� �����, ���� ����� ����������� � ����� �������. ��������� (Shadowing)
Sub d_07_SameNameInProject()
        
    Debug.Print PROJECT_SCOPE_SAME_NAME
    Debug.Print m_IdentifiresVisibility.PROJECT_SCOPE_SAME_NAME
    Debug.Print

End Sub

'��� �����, ���� ����� �� ������ ������� Scope. ��������� (Shadowing)
Sub d_08_SameNameInDifferentScopes()
        
    Const PROJECT_SCOPE_SAME_NAME = "������� ���������"
    Debug.Print PROJECT_SCOPE_SAME_NAME
    Debug.Print m_01_06_Scopes.PROJECT_SCOPE_SAME_NAME
    Debug.Print m_IdentifiresVisibility.PROJECT_SCOPE_SAME_NAME
    Debug.Print
    
End Sub

'��� ����������. ���������
Sub d_09_Shadowing()
    
    Dim ThisDocument As String
    ThisDocument = "d:\Docs\report.docx"
'    ThisDocument.Paragraphs(1).Alignment = wdAlignParagraphLeft
    Word.ActiveDocument.Paragraphs(1).Alignment = wdAlignParagraphLeft
    Debug.Print "�������� ����������"

End Sub

'������� ��������� ���������� �� ��������� �������� ����� ��������
Sub d_10_OrdinaryLocalVariable()

    Dim lngOrdinary As Long
    lngOrdinary = lngOrdinary + 1
    Debug.Print "lngOrdinary = "; lngOrdinary

End Sub


' ����������� ��������� ���������� ��������� �������� ����� ��������
Sub d_11_StaticLocalVariable()

    Static lngStatic As Long
    lngStatic = lngStatic + 1
    Debug.Print "lngStatic = "; lngStatic
    
End Sub


' ����� ����� ����������
Sub d_12_LifeTime()

    '��������� ���������� �� ���������� ���������
    '���������� � ���������� ������ � �������, ����������� - �� �������� ������� ���
    '   ���������� ����������
    '���������� �� ���������� ����������
       
End Sub

' ��������� ���������� ���������
Sub d_14_UnhandledErrors()

    Dim intValue As Integer
    intValue = 35000
    
End Sub

' ������ � �� �������
Sub d_15_TypicalErrors()
    
    '������������ ���
        Dim strName As String
        'Dim strName As String
        
    '������ ��������� � ����������
        Dim ThisDocument As String
        'ThisDocument.Paragraphs(1).Alignment = wdAlignParagraphLeft
        
    Debug.Print "�������� ����������"
End Sub

' ������ ��������
Sub d_16_BestPractices()

    '������������� �������� ����� ������� ��������� ����������
    
    '� �������� ��������� ������� ��������� ����������
    
    '�������� ���������� ���������� ���������
    
    Debug.Print "�������� ����������"
End Sub

