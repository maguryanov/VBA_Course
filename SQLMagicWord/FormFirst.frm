VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormFirst 
   Caption         =   "UserForm1"
   ClientHeight    =   4320
   ClientLeft      =   1125
   ClientTop       =   1470
   ClientWidth     =   10665
   OleObjectBlob   =   "FormFirst.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "FormFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
On Error GoTo ErrorHandler
    FormFirst.Top = CSng(GetSetting(AppName:="SQLMagic", Section:="PersonForm", Key:="Top"))
    FormFirst.Left = CSng(GetSetting(AppName:="SQLMagic", Section:="PersonForm", Key:="Left"))
    FormFirst.Height = CSng(GetSetting(AppName:="SQLMagic", Section:="PersonForm", Key:="Hight"))
    FormFirst.Width = CSng(GetSetting(AppName:="SQLMagic", Section:="PersonForm", Key:="Width"))
    FormFirst.Caption = GetSetting(AppName:="SQLMagic", Section:="PersonForm", Key:="Caption")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
On Error GoTo ErrorHandler
    SaveSetting AppName:="SQLMagic", Section:="PersonForm", Key:="Top", Setting:=FormFirst.Top
    SaveSetting AppName:="SQLMagic", Section:="PersonForm", Key:="Left", Setting:=FormFirst.Left
    SaveSetting AppName:="SQLMagic", Section:="PersonForm", Key:="Hight", Setting:=FormFirst.Height
    SaveSetting AppName:="SQLMagic", Section:="PersonForm", Key:="Width", Setting:=FormFirst.Width
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

