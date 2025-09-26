VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tabs_08 
   Caption         =   "ƒемонстраци€ свойств дл€ перемещени€ фокуса"
   ClientHeight    =   8910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17625
   OleObjectBlob   =   "Tabs_08.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Tabs_08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CheckBoxTabStop_Change()
    TextBoxComment.TabStop = CheckBoxTabStop.Value
End Sub

Private Sub UserForm_Initialize()
     CheckBoxTabStop.Value = TextBoxComment.TabStop
End Sub
