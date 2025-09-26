VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Behavior_07 
   Caption         =   "Демонстрация свойств группы Behavior для TextBox"
   ClientHeight    =   10755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17850
   OleObjectBlob   =   "Behavior_07.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Behavior_07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBox10_Change()
    TextBoxMultiLine.WordWrap = CheckBox10.Value
End Sub

Private Sub CheckBox11_Change()
    TextBoxMultiLine.MultiLine = CheckBox11.Value
End Sub

Private Sub CheckBox12_Change()
    TextBoxMultiLine.EnterKeyBehavior = CheckBox12.Value
End Sub

Private Sub CheckBox13_Click()
    TextBoxMultiLine.HideSelection = CheckBox13.Value
End Sub


Private Sub CheckBox15_Change()
    TextBoxMultiLine.Locked = CheckBox15.Value
End Sub

Private Sub CheckBox16_Change()
    TextBoxMultiLine.SelectionMargin = CheckBox16.Value
End Sub

Private Sub CheckBox17_Change()
    TextBoxMaxLength.AutoTab = CheckBox17.Value
End Sub

Private Sub CheckBox18_Change()
    TextBoxMultiLine.TabKeyBehavior = CheckBox18.Value
End Sub

Private Sub CheckBox8_Change()
    TextBoxMultiLine.AutoWordSelect = CheckBox8.Value
End Sub

Private Sub CheckBox9_Change()
    TextBoxMultiLine.Enabled = CheckBox9.Value
End Sub


Private Sub UserForm_Initialize()
    CheckBox10.Value = TextBoxMultiLine.WordWrap
    CheckBox11.Value = TextBoxMultiLine.MultiLine
    CheckBox12.Value = TextBoxMultiLine.EnterKeyBehavior
    CheckBox13.Value = TextBoxMultiLine.HideSelection
    CheckBox15.Value = TextBoxMultiLine.Locked
    CheckBox16.Value = TextBoxMultiLine.SelectionMargin
    CheckBox17.Value = TextBoxMaxLength.AutoTab
    CheckBox8.Value = TextBoxMultiLine.AutoWordSelect
    CheckBox9.Value = TextBoxMultiLine.Enabled
    CheckBox18.Value = TextBoxMultiLine.TabKeyBehavior
End Sub
