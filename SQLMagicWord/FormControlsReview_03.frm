VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormControlsReview_03 
   Caption         =   "UserForm1"
   ClientHeight    =   9945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16020
   OleObjectBlob   =   "FormControlsReview_03.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormControlsReview_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngTextBoxBackColor As Long

Private Sub TextBoxMove_Enter()
    lngTextBoxBackColor = TextBoxMove.BackColor
    TextBoxMove.BackColor = RGB(220, 255, 220)
End Sub

Private Sub TextBoxMove_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    TextBoxMove.BackColor = lngTextBoxBackColor
End Sub

