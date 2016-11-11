Option Explicit
' General utility class for doing file checks.
Private Function GetDirnameBasenameParts(fullPath As String) As String()
    Dim pathFileNameParts(1) As String
    Dim fullPathSplitElements() As String
    Dim dirname As String
    Dim basename As String
    If InStr(1, fullPath, "\", vbTextCompare) = 0 Then
        pathFileNameParts(0) = ""
        pathFileNameParts(1) = fullPath
        GetDirnameBasenameParts = pathFileNameParts
        Exit Function
    End If
    fullPathSplitElements = Split(fullPath, "\")
    basename = fullPathSplitElements(UBound(fullPathSplitElements))
    dirname = Replace(fullPath, "\" & basename, "")
    pathFileNameParts(0) = dirname
    pathFileNameParts(1) = basename
    GetDirnameBasenameParts = pathFileNameParts
End Function
Public Function GetDirname(fullPath As String) As String
    Dim dirnameBasenameParts() As String
    Dim dirname As String
    dirnameBasenameParts = GetDirnameBasenameParts(fullPath)
    dirname = dirnameBasenameParts(0)
    GetDirname = dirname
End Function
Public Function GetBasename(fullPath As String) As String
    Dim dirnameBasenameParts() As String
    Dim basename As String
    dirnameBasenameParts = GetDirnameBasenameParts(fullPath)
    basename = dirnameBasenameParts(1)
    GetBasename = basename
End Function
Public Function DirectoryExists(dirPath As String) As Boolean
    If Len(Dir(dirPath, vbDirectory)) > 0 Then
        DirectoryExists = True
        Exit Function
    End If
    DirectoryExists = False
End Function

Public Function PathIsOk(dbPath As String) As Boolean
    Dim dirnameBasenameParts() As String
    Dim dirname As String
    Dim basename As String
    dirnameBasenameParts = GetDirnameBasenameParts(dbPath)
    dirname = dirnameBasenameParts(0)
    basename = dirnameBasenameParts(1)
    If DirectoryExists(dbPath) Then
        PathIsOk = True
    Else
        PathIsOk = False
    End If
End Function
