Attribute VB_Name = "m_04_01_ExcelObjects"
Option Explicit
'Создаём объект Excel из Word
Private Sub d_01_CreateExcelObject()
    On Error GoTo ErrorHandler
    Dim oExcel As New Excel.Application
    Dim oWorkbook As Workbook
    Dim oWorksheet As Worksheet
    Set oWorkbook = oExcel.Workbooks.Add()
    Set oWorksheet = oWorkbook.ActiveSheet
    oWorksheet.Cells(1, 1).Value = "Данные из Word"
    oExcel.Visible = True
Finalization:
    Set oWorksheet = Nothing
    Set oWorkbook = Nothing
    Set oExcel = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Создаём объект Excel из Word. Coкращаем код
Private Sub d_02_CreateExcelObject()
    On Error GoTo ErrorHandler
    Dim oExcel As New Excel.Application
    oExcel.Workbooks.Add().ActiveSheet.Cells(1, 1).Value = "Данные из Word"
    oExcel.Visible = True
Finalization:
    Set oExcel = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Объект Application. Пример метода. Pi
Private Sub d_03_Pi()
    On Error GoTo ErrorHandler
    Dim oExcel As New Excel.Application
    Debug.Print oExcel.PI
Finalization:
    Set oExcel = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'Объект Application. Пример метода. Wait
Private Sub d_03_Wait()
    On Error GoTo ErrorHandler
    Dim oExcel As New Excel.Application
    Debug.Print Now
    oExcel.Wait (Now + TimeValue("00:00:02"))
    Debug.Print Now
Finalization:
    Set oExcel = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'ActiveCell
Private Sub d_04_ActiveCell()
    On Error GoTo ErrorHandler
    Dim oExcel As New Excel.Application
    oExcel.Workbooks.Add
    oExcel.ActiveCell.Value = "Данные из Word"
    oExcel.Visible = True
Finalization:
    Set oExcel = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'ActiveSheet
Private Sub d_05_ActiveSheet()
    On Error GoTo ErrorHandler
    Dim oExcel As New Excel.Application
    oExcel.Workbooks.Add
    oExcel.ActiveSheet.Cells(1, 1).Value = "Данные из Word"
    oExcel.Visible = True
Finalization:
    Set oExcel = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'ActiveWorkbook
Private Sub d_06_ActiveWorkbook()
    On Error GoTo ErrorHandler
    Dim oExcel As New Excel.Application
    oExcel.Workbooks.Add
    oExcel.ActiveWorkbook.Worksheets(1).Cells(1, 1).Value = "Данные из Word"
    oExcel.Visible = True
Finalization:
    Set oExcel = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'Видимый вывод данных
Private Sub d_06_Visible()
    On Error GoTo ErrorHandler
    Dim oExcel As New Excel.Application
    Dim lngCounter As Long
    Dim dblTimerStart As Double
    oExcel.Visible = True
    oExcel.Workbooks.Add
    dblTimerStart = Timer
    For lngCounter = 1 To 500
        oExcel.ActiveSheet.Cells(lngCounter, 1).Value = "Данные из Word"
    Next lngCounter
    Debug.Print "Время: " & Format(Timer - dblTimerStart, "0.000") & " сек"
Finalization:
    Set oExcel = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'Невидимый вывод данных
Private Sub d_06_NotVisible()
    On Error GoTo ErrorHandler
    Dim oExcel As New Excel.Application
    Dim lngCounter As Long
    Dim dblTimerStart As Double
    oExcel.Workbooks.Add
    dblTimerStart = Timer
    For lngCounter = 1 To 500
        oExcel.ActiveSheet.Cells(lngCounter, 1).Value = "Данные из Word"
    Next lngCounter
    Debug.Print "Время: " & Format(Timer - dblTimerStart, "0.000") & " сек"
Finalization:
    oExcel.Visible = True
    Set oExcel = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'ScreenUpdating
Private Sub d_07_ScreenUpdating()
    On Error GoTo ErrorHandler
    Dim oExcel As New Excel.Application
    Dim lngCounter As Long
    Dim dblTimerStart As Double
    oExcel.Visible = True
    oExcel.ScreenUpdating = False
    oExcel.Workbooks.Add
    dblTimerStart = Timer
    For lngCounter = 1 To 500
        oExcel.ActiveSheet.Cells(lngCounter, 1).Value = "Данные из Word"
    Next lngCounter
    Debug.Print "Время: " & Format(Timer - dblTimerStart, "0.000") & " сек"
Finalization:
    oExcel.ScreenUpdating = True
    Set oExcel = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub

'Переходим в Excel



