Attribute VB_Name = "Template"
Private Sub d_0_()
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Private Sub d_01_()
    On Error GoTo ErrorHandler

Finalization:
    
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub



Sub test()
    Dim Price As Currency
    Price = 100
    For i = 100 To -100 Step -1
        Price = i
        Debug.Assert Price > 0
    Next i
End Sub

'Sub TestCourse_Events()
'    Dim oCourse As New Course
'    oCourse.strCourseName = "VBA"
'    oCourse.oTrainer.FirstName = "Михаил"
'    oCourse.oTrainer.LastName = "Гурьянов"
'    oCourse.oTrainer.FirstName = "Михаил"
'    oCourse.oTrainer.Gender = "М"
'    oCourse.oTrainer.BirthDate = Now - 10000
'    oCourse.oTrainer.PrintForm
'End Sub


Sub TestRubberDuck()
    Dim a As String
    Debug.Print
End Sub







