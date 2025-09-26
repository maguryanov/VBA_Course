VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ContractForm 
   Caption         =   "Данные договора"
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
    Documents.Add "D:\Projects\VBA\Михеев\ДОГОВОР.dotx"

    Dim bkm As Bookmark
    Dim rng As Range
    'Номер
    If ActiveDocument.Bookmarks.Exists("Номер") Then
        Set bkm = ActiveDocument.Bookmarks("Номер")
        Set rng = bkm.Range
        rng.Text = ContractForm.Nomer.Value
        ActiveDocument.Bookmarks.Add "Номер", rng
    End If
    'Должность
    If ActiveDocument.Bookmarks.Exists("Должность") Then
        Set bkm = ActiveDocument.Bookmarks("Должность")
        Set rng = bkm.Range
        rng.Text = ContractForm.Dolgnost.Value
        ActiveDocument.Bookmarks.Add "Должность", rng
    End If
    'Заказчик
    If ActiveDocument.Bookmarks.Exists("Заказчик") Then
        Set bkm = ActiveDocument.Bookmarks("Заказчик")
        Set rng = bkm.Range
        rng.Text = ContractForm.Naimenovanie.Value
        ActiveDocument.Bookmarks.Add "Заказчик", rng
    End If
    'Дата
    If ActiveDocument.Bookmarks.Exists("Дата") Then
        Set bkm = ActiveDocument.Bookmarks("Дата")
        Set rng = bkm.Range
        rng.Text = ContractForm.Data.Value
        ActiveDocument.Bookmarks.Add "Дата", rng
    End If
    'Основание
    If ActiveDocument.Bookmarks.Exists("Основание") Then
        Set bkm = ActiveDocument.Bookmarks("Основание")
        Set rng = bkm.Range
        rng.Text = ContractForm.Deystvuuchego.Value
        ActiveDocument.Bookmarks.Add "Основание", rng
    End If
    'ФИО
    If ActiveDocument.Bookmarks.Exists("ФИО") Then
        Set bkm = ActiveDocument.Bookmarks("ФИО")
        Set rng = bkm.Range
        rng.Text = ContractForm.V_lice_kogo.Value
        ActiveDocument.Bookmarks.Add "ФИО", rng
    End If
    'Город
    If ActiveDocument.Bookmarks.Exists("Город") Then
        Set bkm = ActiveDocument.Bookmarks("Город")
        Set rng = bkm.Range
        rng.Text = ContractForm.Gorod.Value
        ActiveDocument.Bookmarks.Add "Город", rng
    End If
    ContractForm.Hide
    
End Sub

