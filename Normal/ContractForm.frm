VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ContractForm 
   Caption         =   "������ ��������"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14370
   OleObjectBlob   =   "ContractForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ContractForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    ContractForm.Hide
End Sub

Private Sub SformCommand_Click()
    Documents.Add "D:\Projects\VBA\������\�������.dotx"

    Dim bkm As Bookmark
    Dim rng As Range
    '�����
    If ActiveDocument.Bookmarks.Exists("�����") Then
        Set bkm = ActiveDocument.Bookmarks("�����")
        Set rng = bkm.Range
        rng.Text = ContractForm.Nomer.Value
        ActiveDocument.Bookmarks.Add "�����", rng
    End If
    '���������
    If ActiveDocument.Bookmarks.Exists("���������") Then
        Set bkm = ActiveDocument.Bookmarks("���������")
        Set rng = bkm.Range
        rng.Text = ContractForm.Dolgnost.Value
        ActiveDocument.Bookmarks.Add "���������", rng
    End If
    '��������
    If ActiveDocument.Bookmarks.Exists("��������") Then
        Set bkm = ActiveDocument.Bookmarks("��������")
        Set rng = bkm.Range
        rng.Text = ContractForm.Naimenovanie.Value
        ActiveDocument.Bookmarks.Add "��������", rng
    End If
    '����
    If ActiveDocument.Bookmarks.Exists("����") Then
        Set bkm = ActiveDocument.Bookmarks("����")
        Set rng = bkm.Range
        rng.Text = ContractForm.Data.Value
        ActiveDocument.Bookmarks.Add "����", rng
    End If
    '���������
    If ActiveDocument.Bookmarks.Exists("���������") Then
        Set bkm = ActiveDocument.Bookmarks("���������")
        Set rng = bkm.Range
        rng.Text = ContractForm.Deystvuuchego.Value
        ActiveDocument.Bookmarks.Add "���������", rng
    End If
    '���
    If ActiveDocument.Bookmarks.Exists("���") Then
        Set bkm = ActiveDocument.Bookmarks("���")
        Set rng = bkm.Range
        rng.Text = ContractForm.V_lice_kogo.Value
        ActiveDocument.Bookmarks.Add "���", rng
    End If
    '�����
    If ActiveDocument.Bookmarks.Exists("�����") Then
        Set bkm = ActiveDocument.Bookmarks("�����")
        Set rng = bkm.Range
        rng.Text = ContractForm.Gorod.Value
        ActiveDocument.Bookmarks.Add "�����", rng
    End If
    ContractForm.Hide
    
End Sub

