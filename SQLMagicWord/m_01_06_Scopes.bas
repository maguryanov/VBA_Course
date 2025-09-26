Attribute VB_Name = "m_01_06_Scopes"
Option Explicit

Private Const MODULE_SCOPE_PRIVATE As String = _
    "Константа объявлена на уровне модуля с модификатором Private"
Const MODULE_SCOPE As String = _
    "Константа объявлена на уровне модуля без модификатора"
Public Const PROJECT_SCOPE_PUBLIC As String = _
    "Константа объявлена на уровне модуля с модификатором Public"

Public Const PROJECT_SCOPE_SAME_NAME As String = _
    "Константа с областью видимости проекта в модуле Scopes"

Private mstrModule_Private As String
Dim mstrModule_Dim As String
Public pstrProject_Public As String

Public strName As String
'Public strName As String


'У идентификатора есть видимость. Локальная константа
Sub d_01_ConstantScope()

    Const PROCEDURE_SCOPE As String = "Константа объявлена на уровне процедуры"
    Debug.Print PROCEDURE_SCOPE

End Sub

'Проверка видимости идентификатора на примере константы
Sub d_02_ConstantVisibilityTest()

    'Debug.Print PROCEDURE_SCOPE
    Debug.Print MODULE_SCOPE
    Debug.Print MODULE_SCOPE_PRIVATE
    Debug.Print PROJECT_SCOPE_PUBLIC
    Debug.Print Normal.GLOBAL_SCOPE_PUBLIC
    'Debug.Print Normal.PROJECT_NORMAL_SCOPE_PUBLIC

End Sub


'У переменной так же есть видимость. И те же правила
Sub d_03_Variables()
    
    Dim strLocal As String
    strLocal = "Переменная объявлена на уровне процедуры"
    mstrModule_Private = "Переменная объявлена на уровне модуля с модификатором Private"
    mstrModule_Dim = "Переменная объявлена на уровне модуля с Dim"
    pstrProject_Public = "Переменная объявлена на уровне модуля с модификатором Public"
    
End Sub

' Видимость переменных
Sub d_04_VariablesVisibility()
    
'    Debug.Print strLocal
'    Debug.Print mstrModule_Private
'    Debug.Print mstrModule_Dim
'    Debug.Print pstrProject_Public
'    Debug.Print Normal.gstrGlobal
'    Debug.Print Normal.pstrNormalProject
        
    Debug.Print "Выполнено успешно"
End Sub


'Сделать локальную переменную Public нельзя
Sub d_05_PublicLocal()
    
    'Public strLocal As String

End Sub

'Что будет если имена дублируются в одной процедуре или модуле
Sub d_06_SameNamesInProcedure()
        
    Dim strName As String
    
    'Dim strName As String

End Sub


'Что будет, если имена дублируются в одном проекте. Затенение (Shadowing)
Sub d_07_SameNameInProject()
        
    Debug.Print PROJECT_SCOPE_SAME_NAME
    Debug.Print m_IdentifiresVisibility.PROJECT_SCOPE_SAME_NAME
    Debug.Print

End Sub

'Что будет, если имена на разных уровнях Scope. Затенение (Shadowing)
Sub d_08_SameNameInDifferentScopes()
        
    Const PROJECT_SCOPE_SAME_NAME = "Уровень процедуры"
    Debug.Print PROJECT_SCOPE_SAME_NAME
    Debug.Print m_01_06_Scopes.PROJECT_SCOPE_SAME_NAME
    Debug.Print m_IdentifiresVisibility.PROJECT_SCOPE_SAME_NAME
    Debug.Print
    
End Sub

'Имя переменной. Затенение
Sub d_09_Shadowing()
    
    Dim ThisDocument As String
    ThisDocument = "d:\Docs\report.docx"
'    ThisDocument.Paragraphs(1).Alignment = wdAlignParagraphLeft
    Word.ActiveDocument.Paragraphs(1).Alignment = wdAlignParagraphLeft
    Debug.Print "Успешное выполнение"

End Sub

'Обычная локальная переменная не сохраняет значение между вызовами
Sub d_10_OrdinaryLocalVariable()

    Dim lngOrdinary As Long
    lngOrdinary = lngOrdinary + 1
    Debug.Print "lngOrdinary = "; lngOrdinary

End Sub


' Статическая локальная переменная сохраняет значение между вызовами
Sub d_11_StaticLocalVariable()

    Static lngStatic As Long
    lngStatic = lngStatic + 1
    Debug.Print "lngStatic = "; lngStatic
    
End Sub


' Время жизни переменных
Sub d_12_LifeTime()

    'Локальные переменные до завершения процедуры
    'Переменные с видимостью модуля и проекта, статические - до выгрузки проекта или
    '   завершения приложения
    'Глобальные до завершения приложения
       
End Sub

' Аварийное завершение программы
Sub d_14_UnhandledErrors()

    Dim intValue As Integer
    intValue = 35000
    
End Sub

' Грабли и их решение
Sub d_15_TypicalErrors()
    
    'Дублирование имён
        Dim strName As String
        'Dim strName As String
        
    'Ошибки связанные с затенением
        Dim ThisDocument As String
        'ThisDocument.Paragraphs(1).Alignment = wdAlignParagraphLeft
        
    Debug.Print "Успешное выполнение"
End Sub

' Лучшие практики
Sub d_16_BestPractices()

    'Устанавливать наиболее узкую область видимости переменных
    
    'В префиксе указывать область видимости переменных
    
    'Избегать аварийного завершения программы
    
    Debug.Print "Успешное выполнение"
End Sub

