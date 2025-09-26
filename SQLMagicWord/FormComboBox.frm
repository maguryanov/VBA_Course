VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormComboBox 
   Caption         =   "UserForm1"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18165
   OleObjectBlob   =   "FormComboBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBoxEnabled_Change()
    ComboBoxRegion.Enabled = CheckBoxEnabled.Value
End Sub

Private Sub CheckBoxLocked_Change()
    ComboBoxRegion.Locked = CheckBoxLocked.Value
End Sub

Private Sub CheckBoxMatchRequired_Change()
    ComboBoxRegion.MatchRequired = CheckBoxMatchRequired.Value
End Sub

Private Sub ComboBoxBorderStyle_Change()
    ComboBoxRegion.BorderStyle = ComboBoxBorderStyle.ListIndex
End Sub

Private Sub ComboBoxDropButtonStyle_Change()
    ComboBoxRegion.DropButtonStyle = ComboBoxDropButtonStyle.ListIndex
End Sub

Private Sub ComboBoxMatchEntry_Change()
    ComboBoxRegion.MatchEntry = ComboBoxMatchEntry.ListIndex
End Sub

Private Sub ComboBoxShowDropButtonWhen_Change()
    ComboBoxRegion.ShowDropButtonWhen = ComboBoxShowDropButtonWhen.ListIndex
End Sub

Private Sub ComboBoxSpecialEffect_Change()
    Dim intInx As Integer
    intInx = IIf(ComboBoxSpecialEffect.ListIndex = 4, 6, ComboBoxSpecialEffect.ListIndex)
    ComboBoxRegion.SpecialEffect = intInx
    
End Sub

Private Sub ComboBoxStyle_Change()
    ComboBoxRegion.Style = Switch _
        (ComboBoxStyle.ListIndex = 0, 0, ComboBoxStyle.ListIndex = 1, 2)
End Sub





Private Sub UserForm_Initialize()
On Error GoTo ErrorHandler
    ComboBoxRegion.AddItem "Москва", "0"
    ComboBoxRegion.AddItem "Санкт-Петербург", "1"
    ComboBoxRegion.AddItem "Нижний Новгород", 2
    ComboBoxRegion.AddItem "Новосибирск", 3
    
    ComboBoxStyle.AddItem "fmStyleDropDownCombo", 0
    ComboBoxStyle.AddItem "fmStyleDropDownList", 1
    
    ComboBoxBorderStyle.AddItem "fmBorderStyleNone", 0
    ComboBoxBorderStyle.AddItem "fmBorderStyleSingle", 1

    ComboBoxSpecialEffect.AddItem "fmSpecialEffectFlat", 0
    ComboBoxSpecialEffect.AddItem "fmSpecialEffectRaised", 1
    ComboBoxSpecialEffect.AddItem "fmSpecialEffectSunken", 2
    ComboBoxSpecialEffect.AddItem "fmSpecialEffectEtched", 3
    ComboBoxSpecialEffect.AddItem "fmSpecialEffectBump", 4
    
    ComboBoxDropButtonStyle.AddItem "fmDropButtonStylePlain", 0
    ComboBoxDropButtonStyle.AddItem "fmDropButtonStyleArrow", 1
    ComboBoxDropButtonStyle.AddItem "fmDropButtonStyleEllipsis", 2
    ComboBoxDropButtonStyle.AddItem "fmDropButtonStyleReduce", 3

    ComboBoxShowDropButtonWhen.AddItem "fmShowDropButtonWhenNever", 0
    ComboBoxShowDropButtonWhen.AddItem "fmShowDropButtonWhenFocus", 1
    ComboBoxShowDropButtonWhen.AddItem "fmShowDropButtonWhenAlways", 2
    
    ComboBoxMatchEntry.AddItem "fmMatchEntryFirstLetter", 0
    ComboBoxMatchEntry.AddItem "FmMatchEntryComplete", 1
    ComboBoxMatchEntry.AddItem "FmMatchEntryNone", 2
    
    
    CheckBoxEnabled.Value = ComboBoxRegion.Enabled
    CheckBoxLocked.Value = ComboBoxRegion.Locked
    CheckBoxMatchRequired.Value = ComboBoxRegion.MatchRequired
    

    Exit Sub
ErrorHandler:
Debug.Print Err.Number & " / " & Err.Description
End Sub
