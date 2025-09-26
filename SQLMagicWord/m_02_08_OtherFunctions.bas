Attribute VB_Name = "m_02_08_OtherFunctions"
Option Explicit

Private Sub DoEvents_01()
On Error GoTo ErrorHandler
    Dim lCounter As Long
    Dim strVar As String
'    For lCounter = 1 To 10000000
'        strVar = CStr(Now)
'    Next lCounter
'    Exit Sub
    
    For lCounter = 1 To 2000000000
         strVar = CStr(Now)
        If lCounter Mod 100000 = 0 Then DoEvents
    Next lCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub Environ_01()
On Error GoTo ErrorHandler
    Debug.Print Environ("ProgramData"), TypeName(Environ("ProgramData"))
    Debug.Print Environ("ProgramFiles(x86)")
    Debug.Print Environ("WinDir")
    Debug.Print Environ("SQL"), TypeName(Environ("SQL"))
    Debug.Print Environ(1)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Private Sub RGB_01()
On Error GoTo ErrorHandler
    Selection.Font.Color = RGB(0, 0, 0)
    Selection.Font.Color = RGB(255, 255, 555)
    Selection.Font.Color = RGB(50, 205, 50) ' LimeGreen
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub Shell_01()
On Error GoTo ErrorHandler
    Debug.Print Shell("notepad.exe", vbMaximizedFocus)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
'Operator Computer\HKEY_CURRENT_USER\Software\VB and VBA Program Settings
Private Sub SaveSetting_01()
On Error GoTo ErrorHandler
    SaveSetting AppName:="SQLMagic", Section:="PersonForm", Key:="Top", Setting:="50"
    SaveSetting AppName:="SQLMagic", Section:="PersonForm", Key:="Left", Setting:="50"
    SaveSetting AppName:="SQLMagic", Section:="PersonForm", Key:="Hight", Setting:="300"
    SaveSetting AppName:="SQLMagic", Section:="PersonForm", Key:="Width", Setting:="400"
    SaveSetting "SQLMagic", "PersonForm", "Caption", "Данные о персоналиях. Версия 1.0"
    Debug.Print "Успешно записано в Реестр"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
Private Sub SaveSetting_02()
On Error GoTo ErrorHandler
    SaveSetting AppName:="SQLMagic", Section:="RecectDocuments", Key:="1", Setting:="Отчет.doc"
    SaveSetting AppName:="SQLMagic", Section:="RecectDocuments", Key:="2", Setting:="Расчет.xlsm"
    SaveSetting AppName:="SQLMagic", Section:="RecectDocuments", Key:="3", Setting:="Продажи.xlsm"
    Debug.Print "Успешно записано в Реестр"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub GetSetting_01()
On Error GoTo ErrorHandler
    Debug.Print GetSetting(AppName:="SQLMagic", Section:="PersonForm", Key:="Top")
    Debug.Print GetSetting(AppName:="SQLMagic", Section:="PersonForm", Key:="Left")
    Debug.Print GetSetting(AppName:="SQLMagic", Section:="PersonForm", Key:="Hight")
    Debug.Print GetSetting(AppName:="SQLMagic", Section:="PersonForm", Key:="Width")
    Debug.Print GetSetting(AppName:="SQLMagic", Section:="PersonForm", Key:="Caption")
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub GetAllSettings_01()
On Error GoTo ErrorHandler
    Dim aSettings As Variant
    aSettings = GetAllSettings(AppName:="SQLMagic", Section:="RecectDocuments")
    Debug.Print aSettings(0, 0), aSettings(0, 1)
    Debug.Print aSettings(1, 0), aSettings(1, 1)
    Debug.Print aSettings(2, 0), aSettings(2, 1)
    Debug.Print "LBound", LBound(aSettings, 1)
    Debug.Print "UBound", UBound(aSettings, 1)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Private Sub GetAllSettings_02()
On Error GoTo ErrorHandler
    Dim aSettings As Variant
    aSettings = GetAllSettings(AppName:="SQLMagic", Section:="RecectDocuments")
    Dim iCounter As Integer
    For iCounter = LBound(aSettings, 1) To UBound(aSettings, 1)
        Debug.Print aSettings(iCounter, 0), aSettings(iCounter, 1)
    Next iCounter
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
'Оператор
Private Sub DeleteSetting_01()
On Error GoTo ErrorHandler
    DeleteSetting AppName:="SQLMagic", Section:="RecectDocuments"
    DeleteSetting AppName:="SQLMagic", Section:="PersonForm"
    Debug.Print "Успешно удалены из Реестра"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub Spc_01()
On Error GoTo ErrorHandler
    Dim strVar As String
    Debug.Print "|"; Spc(5); "|"
    'strVar = "|" & Spc(5) & "|"
    'Debug.Print "|" & Spc(5) & "|"
    strVar = "|" & Space(5) & "|"
    Debug.Print Spc(5); strVar
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Private Sub Tab_01()
On Error GoTo ErrorHandler
    Dim strVar As String
    Debug.Print "|"; Tab; "|"
    'strVar = "|" & Tab & "|"
    'Debug.Print "|" & Tab & "|"
    strVar = "|" & vbTab & "|"
    Debug.Print Tab; strVar
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub Len_01()
On Error GoTo ErrorHandler
    Dim strVar As String
    Dim iVar As Integer
    Dim lVar As Long
    Dim aArr(100) As Integer
    Debug.Print "Len(strVar)", Len(strVar)
    Debug.Print "Len(iVar)", Len(iVar)
    Debug.Print "Len(lVar)", Len(lVar)
    'Debug.Print Len(aArr)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub Len_02()
On Error GoTo ErrorHandler
    Debug.Print "Len(строка)", Len("Михаил" & "Гурьянов")
    'Debug.Print "Len(другие типы)", Len(1 + 2)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub Len_03()
On Error GoTo ErrorHandler
    Dim varVar As Variant
    Debug.Print "--------------------------------"
    Debug.Print TypeName(varVar), Len(varVar)
    varVar = "1234567890"
    Debug.Print TypeName(varVar), Len(varVar)
    varVar = 10000
    Debug.Print TypeName(varVar), Len(varVar)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub Len_04()
On Error GoTo ErrorHandler
    Dim strFix As String * 10
    Debug.Print "--------------------------------"
    Debug.Print "|"; strFix; "|", Len(strFix)
    strFix = 12345
    Debug.Print "|"; strFix; "|", Len(strFix)
    Debug.Print "|"; strFix; "|", Len(RTrim(strFix))
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
'Что нельзя?
Private Sub TypeName_01()
On Error GoTo ErrorHandler
    Dim oPerson As New Person 'Class
    Dim uCategory As Category 'User Type
    Debug.Print "--------------------------------"
    Debug.Print TypeName(1 + 1)
    Debug.Print TypeName(1 + 1&)
    Debug.Print TypeName(oPerson)
    'Debug.Print TypeName(uCourseCategory)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


Private Sub TypeName_02()
On Error GoTo ErrorHandler
    Dim oObject As Object
    Debug.Print "--------------------------------"
    Debug.Print TypeName(oObject)
    Set oObject = New Person
    Debug.Print TypeName(oObject)
    Set oObject = New VideoLesson
    Debug.Print TypeName(oObject)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub TypeName_03()
On Error GoTo ErrorHandler
    Dim varVar As Variant
    Debug.Print "--------------------------------"
    Debug.Print TypeName(varVar)
    varVar = "Строка"
    Debug.Print TypeName(varVar)
    varVar = 100
    Debug.Print TypeName(varVar)
    varVar = Null
    Debug.Print TypeName(varVar)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Что нельзя?
Private Sub VarType_01()
On Error GoTo ErrorHandler
    Dim oPerson As New Person 'Class
    Dim uCategory As Category 'User Type
    Debug.Print "--------------------------------"
    Debug.Print VarType(1 + 1) 'vbInteger   2   Integer
    Debug.Print VarType(1 + 1&) 'vbLong  3   Длинное целое
    Debug.Print VarType(oPerson) 'vbObject    9   Объект
    'Debug.Print VarType(uCourseCategory)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub VarType_02()
On Error GoTo ErrorHandler
    Dim oObject As Object
    Debug.Print "--------------------------------"
    Debug.Print VarType(oObject) 'vbObject    9   Объект
    Set oObject = New Person
    Debug.Print VarType(oObject)
    Set oObject = New VideoLesson
    Debug.Print VarType(oObject)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

Private Sub VarType_03()
On Error GoTo ErrorHandler
    Dim varVar As Variant
    Debug.Print "--------------------------------"
    Debug.Print VarType(varVar) 'vbEmpty     0   Пустое (не инициализированный)
    varVar = "Строка"
    Debug.Print VarType(varVar) 'vbString    8   String
    varVar = 100
    Debug.Print VarType(varVar) 'vbInteger   2   Integer
    varVar = Null
    Debug.Print VarType(varVar) 'vbNull  1   Null (данные отсутствуют)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub
