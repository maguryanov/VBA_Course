Attribute VB_Name = "m_07_03_Dialogs"
Option Explicit

'����� �����
Private Sub d_01_FileDialogFolderPicker()
    On Error Resume Next
    Dim objFileDialog As FileDialog
    Set objFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With objFileDialog
        .Title = "�������� �����"
        .InitialFileName = "d:\vba" ' ��������� �����
        If .Show = -1 Then ' -1 - ������� ������ ��������, 0 - ������
            Debug.Print .SelectedItems(1)
        End If
    End With
    Exit Sub
ErrorHandler:
    MsgBox "������: " & Err.Description, vbCritical
End Sub


'����� ������
Private Sub d_02_FileDialogFilePicker()
    On Error Resume Next
    Dim objFileDialog As FileDialog, varFileName As Variant
    Set objFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    With objFileDialog
        .Title = "�������� �����. ����� ���������, ������� Shift � Ctrl"
        .InitialFileName = "d:\vba\*.*" ' ��������� �����
        .AllowMultiSelect = True
        If .Show = -1 Then '-1 - ������� ������ ��������, 0 - ������
            For Each varFileName In .SelectedItems
                Debug.Print varFileName
            Next varFileName
        End If
    End With
    Exit Sub
ErrorHandler:
    MsgBox "������: " & Err.Description, vbCritical
End Sub


'�������� ������
Private Sub d_03_FileDialogOpen()
    On Error Resume Next
    Dim objFileDialog As FileDialog, varFileName As Variant
    Set objFileDialog = Application.FileDialog(msoFileDialogOpen)
    With objFileDialog
        .Title = "�������� �����. ����� ���������, ������� Shift � Ctrl"
        .InitialFileName = "D:\VBA\��� ������ Word\*.*" ' ��������� �����
        .Filters.Add "����� Word", "*.docx; *.docm", 1 '��������� �������
        .FilterIndex = 1 '������ �� ���������
        .AllowMultiSelect = True
        If .Show = -1 Then '-1 - ������� ������ ��������, 0 - ������
            For Each varFileName In .SelectedItems
                Debug.Print varFileName
            Next varFileName
        End If
        If .SelectedItems.Count > 3 Then
            MsgBox "���������� ������ ��������� ������������ ���������� - ���"
        Else
            .Execute
        End If
    End With
    Exit Sub
ErrorHandler:
    MsgBox "������: " & Err.Description, vbCritical
End Sub


'���������� ������
Private Sub d_04_FileDialogSaveAs()
    On Error Resume Next
    Dim objFileDialog As FileDialog, varFileName As Variant
    Set objFileDialog = Application.FileDialog(msoFileDialogSaveAs)
    With objFileDialog
        .Title = "��������� ���� ���"
        .InitialFileName = "d:\vba\Word\SaveAs.docx" ' ��������� ����
        .ButtonName = "��������� ���"
        If .Show = -1 Then '-1 - ������� ������ ��������, 0 - ������
            Debug.Print .SelectedItems(1)
        End If
        .Execute
    End With
    Exit Sub
ErrorHandler:
    MsgBox "������: " & Err.Description, vbCritical
End Sub
