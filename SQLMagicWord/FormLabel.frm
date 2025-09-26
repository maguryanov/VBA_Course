VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormLabel 
   Caption         =   "Демонстрация возможностей элемента Label"
   ClientHeight    =   11265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17445
   OleObjectBlob   =   "FormLabel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ComboBoxBorderStyle_Change()
    LabelSingleLine.BorderStyle = ComboBoxBorderStyle.ListIndex
    LabelMultiLine.BorderStyle = ComboBoxBorderStyle.ListIndex
    LabelPicture.BorderStyle = ComboBoxBorderStyle.ListIndex
    LabelDemo.BorderStyle = ComboBoxBorderStyle.ListIndex
    LabelEmpty.BorderStyle = ComboBoxBorderStyle.ListIndex
End Sub

Private Sub ComboBoxSpecialEffect_Change()
    Dim intInx As Integer
    intInx = IIf(ComboBoxSpecialEffect.ListIndex = 4, 6, ComboBoxSpecialEffect.ListIndex)
    LabelSingleLine.SpecialEffect = intInx
    LabelMultiLine.SpecialEffect = intInx
    LabelPicture.SpecialEffect = intInx
    LabelDemo.SpecialEffect = intInx
    LabelEmpty.SpecialEffect = intInx
    
End Sub

Private Sub ListBoxDemo_Change()
    LabelDemo.PicturePosition = ListBoxDemo.ListIndex
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    ListBoxDemo.AddItem "fmPicturePositionLeftTop", 0
    ListBoxDemo.AddItem "fmPicturePositionLeftCenter", 1
    ListBoxDemo.AddItem "fmPicturePositionLeftBottom", 2
    ListBoxDemo.AddItem "fmPicturePositionRightTop", 3
    ListBoxDemo.AddItem "fmPicturePositionRightCenter", 4
    ListBoxDemo.AddItem "fmPicturePositionRightBottom", 5
    ListBoxDemo.AddItem "fmPicturePositionAboveLeft", 6
    ListBoxDemo.AddItem "fmPicturePositionAboveCenter", 7
    ListBoxDemo.AddItem "fmPicturePositionAboveRight", 8
    ListBoxDemo.AddItem "fmPicturePositionBelowLeft", 9
    ListBoxDemo.AddItem "fmPicturePositionBelowCenter", 10
    ListBoxDemo.AddItem "fmPicturePositionBelowRight", 11
    ListBoxDemo.AddItem "fmPicturePositionCenter", 12
    ComboBoxBorderStyle.AddItem "fmBorderStyleNone", 0
    ComboBoxBorderStyle.AddItem "fmBorderStyleSingle", 1
    
    ComboBoxSpecialEffect.AddItem "fmSpecialEffectFlat", 0
    ComboBoxSpecialEffect.AddItem "fmSpecialEffectRaised", 1
    ComboBoxSpecialEffect.AddItem "fmSpecialEffectSunken", 2
    ComboBoxSpecialEffect.AddItem "fmSpecialEffectEtched", 3
    ComboBoxSpecialEffect.AddItem "fmSpecialEffectBump", 4
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
