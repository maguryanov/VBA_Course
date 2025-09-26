VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents oPerson As Person
Attribute oPerson.VB_VarHelpID = -1


Private Sub oPerson_evtIcorrectBirthDate(IcorrectBirthDate As Date)
    Debug.Print "Обработка события в классе EventProcessor"
    Debug.Print "Некорректное значение даты рождения:" & IcorrectBirthDate
End Sub
