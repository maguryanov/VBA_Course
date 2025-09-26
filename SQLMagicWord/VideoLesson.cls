VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VideoLesson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public strLessonName As String
Private udtCourse As TCourse
Public intDuration As Integer

Function ErrorHandlingDemo(ByVal Divider As Variant) As Variant
    On Error GoTo ErrorHandler
    ErrorHandlingDemo = 1 / Divider
    Exit Function
ErrorHandler:
    ErrorHandlingDemo = Null
End Function


Function ErrorHandlingRaiseDemo(ByVal Divider As Variant) As Variant
    On Error GoTo ErrorHandler
    ErrorHandlingRaiseDemo = 1 / Divider
    Exit Function
ErrorHandler:
    ErrorHandlingRaiseDemo = Null
    Err.Raise Err.Number
End Function

