Attribute VB_Name = "m_IdentifiresVisibility"
Option Explicit

Public Const PROJECT_SCOPE_SAME_NAME As String = _
    "��������� � �������� ��������� ������� � ������ IdentifiresVisibility"

' ��������� ��������
Sub ConstantVisibilityTest()

'    Debug.Print PROCEDURE_SCOPE
'    Debug.Print MODULE_SCOPE_PRIVATE
'    Debug.Print MODULE_SCOPE
    Debug.Print PROJECT_SCOPE_PUBLIC
    
    Debug.Print "��������� �������"
End Sub

Sub VariablesVisibility()
    
'    Debug.Print strLocal
'    Debug.Print mstrModule_Private
'    Debug.Print mstrModule_Dim
'    Debug.Print pstrProject_Public
'    Debug.Print Normal.gstrGlobal
'    Debug.Print Normal.pstrNormalProject
        
    Debug.Print "��������� �������"
End Sub


