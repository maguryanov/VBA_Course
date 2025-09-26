Attribute VB_Name = "m_01_17_FilesManipulation"
Option Explicit
' Константы атрибутов для файлов и папок
Private Const ATTR_READONLY = 1
Private Const ATTR_HIDDEN = 2
Private Const ATTR_SYSTEM = 4
Private Const ATTR_DIRECTORY = 16
Private Const ATTR_ARCHIVE = 32
Private Const ATTR_COMPRESSED = 2048
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Получение информации o диске (Drive)
Sub DrivesInfo_01()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oDrive As Drive
    Debug.Print "-----------------------------------------------------------------------"
    For Each oDrive In oFileSystem.Drives
        If oDrive.IsReady Then
            Debug.Print oDrive.DriveLetter; oDrive.AvailableSpace, Tab(30), _
                oDrive.DriveType, oDrive.FileSystem, , oDrive.IsReady
        Else
            Debug.Print oDrive.Path, "Диск не готов"
        End If
    Next
    If oFileSystem.DriveExists("D") Then
        Debug.Print "D существует"
        Debug.Print oFileSystem.Drives("D").FreeSpace
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Получение Folder
Sub FolderInfo_01()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oFolder As Folder
    Debug.Print "-----------------------------------------------------------------------"
    Debug.Print oFileSystem.Drives("D").RootFolder.SubFolders.Count
    Debug.Print oFileSystem.GetFolder("D:\").SubFolders.Count
    For Each oFolder In oFileSystem.GetFolder("C:\Windows\Help").SubFolders
        Debug.Print oFolder.Path, oFolder.Files.Count
    Next
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Информация о папке (Folder)
Sub FolderInfo_02()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oFolder As Folder
    Debug.Print "-----------------------------------------------------------------------"
    Set oFolder = oFileSystem.GetFolder("C:\Windows\Help")
        Debug.Print "Drive", , oFolder.Drive
        Debug.Print "IsRootFolder", , oFolder.IsRootFolder
        Debug.Print "Attributes", , oFolder.Attributes
        Debug.Print "DateCreated", , oFolder.DateCreated
        Debug.Print "DateLastAccessed", oFolder.DateLastAccessed
        Debug.Print "DateLastModified", oFolder.DateLastModified
        Debug.Print "ParentFolder", , oFolder.ParentFolder
        Debug.Print "ShortName", , oFolder.ShortName
        Debug.Print "Type", , oFolder.Type
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub



'Создание файла в папке, новой папки и файла в ней
Sub FolderActions_01()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oFolder As Folder
    Debug.Print "-----------------------------------------------------------------------"
    Set oFolder = oFileSystem.GetFolder("D:\VBA")
    Debug.Print oFolder.Path
    oFolder.CreateTextFile "FolderActions_01.txt", True, True
    If Not oFileSystem.FolderExists("D:\VBA\FolderActions_01") Then
        Set oFolder = oFileSystem.CreateFolder("D:\VBA\FolderActions_01")
        Debug.Print oFolder.Path
        oFolder.CreateTextFile "FolderActions_01.txt", True, True
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Работа с текущей папкой
Sub FolderActions_02()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oFolder As Folder
    Dim strFolder As String
    Debug.Print "-----------------------------------------------------------------------"
    Debug.Print "Определение текущей папки"
    Debug.Print oFileSystem.GetFolder("."), CurDir
    'Смена текущей папки
    ChDir "d:\vba"
    Debug.Print "Определение текущей папки"
    Debug.Print oFileSystem.GetFolder("."), CurDir
    'Смена текущего диска
    ChDrive "d"
    Debug.Print "Определение текущей папки"
    Debug.Print oFileSystem.GetFolder("."), CurDir
    'Смена одной строкой
    strFolder = "C:\Windows"
    ChDrive Left(strFolder, 1): ChDir strFolder
    Debug.Print "Определение текущей папки"
    Debug.Print oFileSystem.GetFolder("."), CurDir
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Создание объектов в текущей папке
Sub FolderActions_03()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim strFolder As String
    Debug.Print "-----------------------------------------------------------------------"
    strFolder = "D:\VBA"
    ChDrive Left(strFolder, 1): ChDir strFolder
    Debug.Print "Определение текущей папки"
    Debug.Print oFileSystem.GetFolder("."), CurDir
    If Not oFileSystem.FolderExists("Папка CreateFolder") Then
        oFileSystem.CreateFolder "Папка CreateFolder"
    End If
    If Not oFileSystem.FolderExists("Папка MkDir") Then
        MkDir "Папка MkDir"
    End If
    If Not oFileSystem.FileExists("Файл CreateTextFile.txt") Then
        oFileSystem.CreateTextFile ("Файл CreateTextFile.txt")
    End If
    If Not oFileSystem.FileExists("Файл для создания и удаления 1.txt") Then
        oFileSystem.CreateTextFile ("Файл для создания и удаления 1.txt")
    End If
    If Not oFileSystem.FileExists("Файл для создания и удаления 2.txt") Then
        oFileSystem.CreateTextFile ("Файл для создания и удаления 2.txt")
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Удаление объектов
Sub FolderActions_04()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oFolder As Folder
    Dim strFolder As String
    Debug.Print "-----------------------------------------------------------------------"
    If oFileSystem.FolderExists("D:\VBA\Папка CreateFolder") Then
        'oFileSystem.GetFolder("D:\VBA\Папка CreateFolder").Delete True
        oFileSystem.DeleteFolder "D:\VBA\Папка CreateFolder", True
    End If
    If oFileSystem.FolderExists("D:\VBA\Папка MkDir") Then
        RmDir "D:\VBA\Папка MkDir"
    End If
    If oFileSystem.FileExists("D:\VBA\Файл для создания и удаления 1.txt") Then
        oFileSystem.DeleteFile "D:\VBA\Файл для создания и удаления 1.txt", True
    End If
    If oFileSystem.FileExists("D:\VBA\Файл для создания и удаления 2.txt") Then
        Kill "D:\VBA\Файл для создания и удаления 2.txt"
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Подготовка к демонстрации
Sub FolderActions_05()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject, oFolder As Folder
    Dim strFolder As String
    Dim byteCounter As Byte
    Dim strFileNameTmp As String, strFileNameBin As String, strFileNameDB As String
    Dim oTextStream As TextStream
    strFolder = "D:\VBA\Файлы для обработки"
    If oFileSystem.FolderExists(strFolder) Then
        oFileSystem.DeleteFolder strFolder, Force:=True
    End If
    MkDir strFolder
    For byteCounter = 1 To 20
        strFileNameBin = "Данные приложения " & Format(byteCounter, "000") & ".bin"
        strFileNameDB = "Данные приложения " & Format(byteCounter, "000") & ".txt"
        strFileNameTmp = "Временный файл " & Format(byteCounter, "000") & ".tmp"
        Set oTextStream = oFileSystem.CreateTextFile _
            (oFileSystem.BuildPath(strFolder, strFileNameBin), True)
        oTextStream.Close
        Set oTextStream = oFileSystem.CreateTextFile _
            (oFileSystem.BuildPath(strFolder, strFileNameDB), True)
        oTextStream.Close
        Set oTextStream = oFileSystem.CreateTextFile _
            (oFileSystem.BuildPath(strFolder, strFileNameTmp), True)
        oTextStream.Close
    Next byteCounter
    strFolder = "D:\VBA\Назначение для копирования"
    If oFileSystem.FolderExists(strFolder) Then
        oFileSystem.DeleteFolder strFolder, Force:=True
    End If
    MkDir strFolder
    strFolder = "D:\VBA\Назначение для перемещения"
    If oFileSystem.FolderExists(strFolder) Then
        oFileSystem.DeleteFolder strFolder, Force:=True
    End If
    MkDir strFolder
    Debug.Print "Успешное завершение"
Finalization:
    If Not (oTextStream Is Nothing) Then oTextStream.Close
    Set oTextStream = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'Удаление по шаблону
Sub FolderActions_06()
    On Error GoTo ErrorHandler
    Dim strFileTemplate As String
    strFileTemplate = "D:\VBA\Файлы для обработки\*.tmp"

    Kill strFileTemplate
    
    Debug.Print "Успешное удаление файлов"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'Копирование файлов
Sub FolderActions_07()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim strFolder As String
    strFolder = "D:\VBA\Файлы для обработки"
    oFileSystem.CopyFolder strFolder, _
        "D:\VBA\Назначение для копирования\Файлы приложения", True
    oFileSystem.CopyFile "D:\VBA\ANSI.txt", _
        "D:\VBA\Назначение для копирования\ANSI.txt", True
    FileCopy "D:\VBA\utf-8.txt", "D:\VBA\Назначение для копирования\utf-8.txt"
    Debug.Print "Успешное копирование файлов"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description & " / " & Err.Source
End Sub


'Перемещение файлов
Sub FolderActions_08()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim strFolder As String
    strFolder = "D:\VBA\Файлы для обработки"
    oFileSystem.MoveFile "D:\VBA\Файлы для обработки\Данные приложения 001.bin", _
        "D:\VBA\Назначение для перемещения\Data001.bin"
    Debug.Print "Успешное перемещение файлов"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description & " / " & Err.Source
End Sub

'Информация о Файле
Sub FileInfo_01()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oFile As File
    Debug.Print "-----------------------------------------------------------------------"
    Set oFile = oFileSystem.GetFile("C:\Windows\Cursors\aero_arrow.cur")
        Debug.Print "Drive", , oFile.Drive
        Debug.Print "Attributes", , oFile.Attributes
        Debug.Print "DateCreated", , oFile.DateCreated
        Debug.Print "DateLastAccessed", oFile.DateLastAccessed
        Debug.Print "DateLastModified", oFile.DateLastModified
        Debug.Print "ParentFolder", , oFile.ParentFolder
        Debug.Print "ShortName", , oFile.ShortName
        Debug.Print "Type", , oFile.Type
        Debug.Print "Size", , oFile.Size
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'Информация о Файле. Атрибуты
Sub FileInfo_02()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oFile As Object
    Debug.Print "-----------------------------------------------------------------------"
    Set oFile = oFileSystem.GetFile("C:\Windows\Cursors\aero_arrow.cur")
    Set oFile = oFileSystem.GetFile("D:\VBA\Attributes\readonly.txt")
    Set oFile = oFileSystem.GetFile("D:\VBA\Attributes\hidden.txt")
    Set oFile = oFileSystem.GetFolder("D:\VBA\Attributes\Folder")
    Set oFile = oFileSystem.GetFile("D:\VBA\Attributes\compressed.txt")
    Debug.Print "ATTR_READONLY", oFile.Attributes And ATTR_READONLY
    Debug.Print "ATTR_HIDDEN", oFile.Attributes And ATTR_HIDDEN
    Debug.Print "ATTR_SYSTEM", oFile.Attributes And ATTR_SYSTEM
    Debug.Print "ATTR_DIRECTORY", oFile.Attributes And ATTR_DIRECTORY
    Debug.Print "ATTR_ARCHIVE", oFile.Attributes And ATTR_ARCHIVE
    Debug.Print "ATTR_COMPRESSED", oFile.Attributes And ATTR_COMPRESSED
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


