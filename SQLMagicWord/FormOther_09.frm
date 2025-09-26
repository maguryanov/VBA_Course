VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormOther_09 
   Caption         =   "Демонстрация остальных свойств для TextBox"
   ClientHeight    =   11925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16830
   OleObjectBlob   =   "FormOther_09.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormOther_09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    TextBoxValue.Value = Null
    'TextBoxValue.Text = Null
End Sub

Private Sub CommandButton2_Click()
    MsgBox TypeName(TextBoxValue.Value) & "|" & TextBoxValue.Value & "|"
End Sub

Private Sub CommandButton3_Click()
    Dim sFirstTime As Single, dtResult As Date
    sFirstTime = Timer
    Dim iCounter As Long
    For iCounter = 0 To 5000000
         TextBoxValue.Text = "Text"
    Next iCounter
    LabelResultText.Caption = Timer - sFirstTime
End Sub

Private Sub CommandButton4_Click()
    Dim sFirstTime As Single, dtResult As Date
    sFirstTime = Timer
    Dim iCounter As Long
    For iCounter = 0 To 5000000
         TextBoxValue.Value = "Value"
    Next iCounter
    LabelResulValue.Caption = Timer - sFirstTime
End Sub

Private Sub CommandButton6_Click()
    TextBoxTag.Text = TextBoxTag.Tag
End Sub



Private Sub ListBoxScrollBars_Change()
    TextBoxScrollBars.ScrollBars = ListBoxScrollBars.ListIndex
    TextBoxScrollBars.SetFocus
End Sub

Private Sub UserForm_Initialize()
    ListBoxScrollBars.AddItem "fmScrollBarsNone", 0
    ListBoxScrollBars.AddItem "fmScrollBarsHorizontal", 1
    ListBoxScrollBars.AddItem "fmScrollBarsVertical", 2
    ListBoxScrollBars.AddItem "fmScrollBarsBoth", 3
End Sub
