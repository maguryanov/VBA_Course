Attribute VB_Name = "m_01_17_FilesManipulation"
Option Explicit
' ��������� ��������� ��� ������ � �����
Private Const ATTR_READONLY = 1
Private Const ATTR_HIDDEN = 2
Private Const ATTR_SYSTEM = 4
Private Const ATTR_DIRECTORY = 16
Private Const ATTR_ARCHIVE = 32
Private Const ATTR_COMPRESSED = 2048
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'��������� ���������� o ����� (Drive)
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
            Debug.Print oDrive.Path, "���� �� �����"
        End If
    Next
    If oFileSystem.DriveExists("D") Then
        Debug.Print "D ����������"
        Debug.Print oFileSystem.Drives("D").FreeSpace
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'��������� Folder
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

'���������� � ����� (Folder)
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



'�������� ����� � �����, ����� ����� � ����� � ���
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


'������ � ������� ������
Sub FolderActions_02()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oFolder As Folder
    Dim strFolder As String
    Debug.Print "-----------------------------------------------------------------------"
    Debug.Print "����������� ������� �����"
    Debug.Print oFileSystem.GetFolder("."), CurDir
    '����� ������� �����
    ChDir "d:\vba"
    Debug.Print "����������� ������� �����"
    Debug.Print oFileSystem.GetFolder("."), CurDir
    '����� �������� �����
    ChDrive "d"
    Debug.Print "����������� ������� �����"
    Debug.Print oFileSystem.GetFolder("."), CurDir
    '����� ����� �������
    strFolder = "C:\Windows"
    ChDrive Left(strFolder, 1): ChDir strFolder
    Debug.Print "����������� ������� �����"
    Debug.Print oFileSystem.GetFolder("."), CurDir
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'�������� �������� � ������� �����
Sub FolderActions_03()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim strFolder As String
    Debug.Print "-----------------------------------------------------------------------"
    strFolder = "D:\VBA"
    ChDrive Left(strFolder, 1): ChDir strFolder
    Debug.Print "����������� ������� �����"
    Debug.Print oFileSystem.GetFolder("."), CurDir
    If Not oFileSystem.FolderExists("����� CreateFolder") Then
        oFileSystem.CreateFolder "����� CreateFolder"
    End If
    If Not oFileSystem.FolderExists("����� MkDir") Then
        MkDir "����� MkDir"
    End If
    If Not oFileSystem.FileExists("���� CreateTextFile.txt") Then
        oFileSystem.CreateTextFile ("���� CreateTextFile.txt")
    End If
    If Not oFileSystem.FileExists("���� ��� �������� � �������� 1.txt") Then
        oFileSystem.CreateTextFile ("���� ��� �������� � �������� 1.txt")
    End If
    If Not oFileSystem.FileExists("���� ��� �������� � �������� 2.txt") Then
        oFileSystem.CreateTextFile ("���� ��� �������� � �������� 2.txt")
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'�������� ��������
Sub FolderActions_04()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim oFolder As Folder
    Dim strFolder As String
    Debug.Print "-----------------------------------------------------------------------"
    If oFileSystem.FolderExists("D:\VBA\����� CreateFolder") Then
        'oFileSystem.GetFolder("D:\VBA\����� CreateFolder").Delete True
        oFileSystem.DeleteFolder "D:\VBA\����� CreateFolder", True
    End If
    If oFileSystem.FolderExists("D:\VBA\����� MkDir") Then
        RmDir "D:\VBA\����� MkDir"
    End If
    If oFileSystem.FileExists("D:\VBA\���� ��� �������� � �������� 1.txt") Then
        oFileSystem.DeleteFile "D:\VBA\���� ��� �������� � �������� 1.txt", True
    End If
    If oFileSystem.FileExists("D:\VBA\���� ��� �������� � �������� 2.txt") Then
        Kill "D:\VBA\���� ��� �������� � �������� 2.txt"
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub


'���������� � ������������
Sub FolderActions_05()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject, oFolder As Folder
    Dim strFolder As String
    Dim byteCounter As Byte
    Dim strFileNameTmp As String, strFileNameBin As String, strFileNameDB As String
    Dim oTextStream As TextStream
    strFolder = "D:\VBA\����� ��� ���������"
    If oFileSystem.FolderExists(strFolder) Then
        oFileSystem.DeleteFolder strFolder, Force:=True
    End If
    MkDir strFolder
    For byteCounter = 1 To 20
        strFileNameBin = "������ ���������� " & Format(byteCounter, "000") & ".bin"
        strFileNameDB = "������ ���������� " & Format(byteCounter, "000") & ".txt"
        strFileNameTmp = "��������� ���� " & Format(byteCounter, "000") & ".tmp"
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
    strFolder = "D:\VBA\���������� ��� �����������"
    If oFileSystem.FolderExists(strFolder) Then
        oFileSystem.DeleteFolder strFolder, Force:=True
    End If
    MkDir strFolder
    strFolder = "D:\VBA\���������� ��� �����������"
    If oFileSystem.FolderExists(strFolder) Then
        oFileSystem.DeleteFolder strFolder, Force:=True
    End If
    MkDir strFolder
    Debug.Print "�������� ����������"
Finalization:
    If Not (oTextStream Is Nothing) Then oTextStream.Close
    Set oTextStream = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
    GoTo Finalization
End Sub


'�������� �� �������
Sub FolderActions_06()
    On Error GoTo ErrorHandler
    Dim strFileTemplate As String
    strFileTemplate = "D:\VBA\����� ��� ���������\*.tmp"

    Kill strFileTemplate
    
    Debug.Print "�������� �������� ������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description
End Sub

'����������� ������
Sub FolderActions_07()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim strFolder As String
    strFolder = "D:\VBA\����� ��� ���������"
    oFileSystem.CopyFolder strFolder, _
        "D:\VBA\���������� ��� �����������\����� ����������", True
    oFileSystem.CopyFile "D:\VBA\ANSI.txt", _
        "D:\VBA\���������� ��� �����������\ANSI.txt", True
    FileCopy "D:\VBA\utf-8.txt", "D:\VBA\���������� ��� �����������\utf-8.txt"
    Debug.Print "�������� ����������� ������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description & " / " & Err.Source
End Sub


'����������� ������
Sub FolderActions_08()
    On Error GoTo ErrorHandler
    Dim oFileSystem As New FileSystemObject
    Dim strFolder As String
    strFolder = "D:\VBA\����� ��� ���������"
    oFileSystem.MoveFile "D:\VBA\����� ��� ���������\������ ���������� 001.bin", _
        "D:\VBA\���������� ��� �����������\Data001.bin"
    Debug.Print "�������� ����������� ������"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & " / " & Err.Description & " / " & Err.Source
End Sub

'���������� � �����
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


'���������� � �����. ��������
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


