Attribute VB_Name = "m_01_04_UserDefinedType"
Option Explicit
'Определённый пользователем тип - User-Defined Type (UDT)
' Продажи по месяцам
Type TMonthlySales
    Year As Integer
    MonthNumber As Byte
    ProductID As Long
    Value As Currency
    Volume As Long
End Type

'UDT Категория курсов
Type TCategory
    CategoryCode As String
    Name As String
End Type

'UDT Подкатегория курсов
Type TSubCategory
    SubCategoryCode As String
    Name As String
    Category As TCategory
End Type

'UDT Курс
Type TCourse
    CourseCode As String
    Name As String
    SubCategory As TSubCategory
End Type

'UDT Видеоурок


'Работа с пользовательским типом UDT
Private Sub d_01_TestSales()
    
    Dim udtSales As TMonthlySales
    
    udtSales.MonthNumber = 9
    Debug.Print udtSales.MonthNumber

End Sub

'Оператор With
Private Sub d_02_TestSales()
    Dim udtSales As TMonthlySales
    With udtSales
        .MonthNumber = 9
        .Year = 2025
        .ProductID = 777
        .Value = 100500.5@
        .Volume = 100000
        Debug.Print MonthName(.MonthNumber); .Year; .ProductID; .Volume; .Value
    End With
End Sub

'Работа с типом TCourse.
Private Sub d_03_TestCourse()
    
    Dim udtCourse As TCourse
    With udtCourse
        .CourseCode = "VBA"
        .Name = "Основы VBA"
        .SubCategory.Name = "VBA"
        .SubCategory.Category.Name = "Программирование"
        Debug.Print .SubCategory.Category.Name
    End With
    
End Sub


Private Sub d_04_TestCourse()
    Dim udtDataBases As TCategory
    Dim udtPostgreSQL As TSubCategory
    Dim udtSQL As TCourse
    Dim udtAdvancedSQL As TCourse
    
    With udtDataBases
        .CategoryCode = "DB"
        .Name = "Базы данных"
    End With
    With udtPostgreSQL
        .SubCategoryCode = "PostgreSQL"
        .Name = "PostgreSQL"
        .Category = udtDataBases
    End With
    With udtSQL
        .CourseCode = "SQL"
        .Name = "Запросы на SQL"
        .SubCategory = udtPostgreSQL
    End With
    With udtAdvancedSQL
        .CourseCode = "AdvSQL"
        .Name = "Расширенный SQL"
        .SubCategory = udtPostgreSQL
    End With
    With udtSQL.SubCategory.Category
        Debug.Print .Name, .CategoryCode
    End With
    With udtAdvancedSQL.SubCategory.Category
        Debug.Print .Name, .CategoryCode
    End With
    
End Sub

