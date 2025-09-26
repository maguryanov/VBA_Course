VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Appearance_06 
   Caption         =   "Демонстрация свойств группы Appearance для TextBox"
   ClientHeight    =   12180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16560
   OleObjectBlob   =   "Appearance_06.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Appearance_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CheckBoxVisible_Change()
    TextBoxVisible.Visible = CheckBoxVisible.Value
End Sub

'Raised- рельефный, выпуклый
'Sunken - затонувший , впалый
'Etched -выгравированный
'Bump -шишка, выпуклость
Private Sub UserForm_Click()
    TextBoxFont.Text = TextBoxFont.Font.Name & ", " & TextBoxFont.Font.Charset _
        & ", Bold = " & TextBoxFont.Font.Bold
End Sub
