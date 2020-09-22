Attribute VB_Name = "Module2"
Option Explicit
Private Declare Function FindFirstFile Lib "Kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "Kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "Kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "Kernel32" (ByVal hFindFile As Long) As Long
Public qt_files As Integer
Public Min_size(0 To 2) As Long

Public abortou As Integer
Public Looking As Integer

Const MAX_PATH = 260 ':( As Integer ?':( Missing Scope
Const MAXDWORD = &HFFFF ':( As Integer ?':( Missing Scope
Const INVALID_HANDLE_VALUE = -1 ':( As Integer ?':( Missing Scope
Const FILE_ATTRIBUTE_ARCHIVE = &H20 ':( As Integer ?':( Missing Scope
Const FILE_ATTRIBUTE_DIRECTORY = &H10 ':( As Integer ?':( Missing Scope
Const FILE_ATTRIBUTE_HIDDEN = &H2 ':( As Integer ?':( Missing Scope
Const FILE_ATTRIBUTE_NORMAL = &H80 ':( As Integer ?':( Missing Scope
Const FILE_ATTRIBUTE_READONLY = &H1 ':( As Integer ?':( Missing Scope
Const FILE_ATTRIBUTE_SYSTEM = &H4 ':( As Integer ?':( Missing Scope
Const FILE_ATTRIBUTE_TEMPORARY = &H100 ':( As Integer ?':( Missing Scope

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Function StripNulls(OriginalStr As String) As String

    If (InStr(OriginalStr, Chr$(0)) > 0) Then
        OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr$(0)) - 1)
    End If
    StripNulls = OriginalStr

End Function

Public Function FindFilesAPI(xxx As String, SearchStr As String, FileCount As Integer, DirCount As Integer)

  Dim LenFile As Long
  Dim filename As String ' Walking filename variable...
  Dim DirName As String ' SubDirectory Name
  Dim dirNames() As String ' Buffer for directory name entries
  Dim nDir As Integer ' Number of directories in this path
  Dim i As Integer ' For-loop counter...
  Dim hSearch As Long ' Search Handle
  Dim WFD As WIN32_FIND_DATA
  Dim Cont As Integer

    If Right$(xxx, 1) <> "\" Then
        xxx = xxx & "\"
    End If
        
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    Cont = True
    hSearch = FindFirstFile(xxx & "*", WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
            DoEvents
            If abortou Then
                Cont = FindClose(hSearch)
                Exit Function
            End If
            DirName = StripNulls(WFD.cFileName)
            ' Ignore the current and encompassing directories.
            If (DirName <> ".") And (DirName <> "..") Then
                ' Check for directory with bitwise comparison.
                If GetFileAttributes(xxx & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                    dirNames(nDir) = DirName
                    DirCount = DirCount + 1
                    nDir = nDir + 1
                    ReDim Preserve dirNames(nDir)
                End If
            End If
            Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
        Loop
        Cont = FindClose(hSearch)
    End If
    ' Walk through this directory and sum file sizes.
    hSearch = FindFirstFile(xxx & SearchStr, WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            If abortou Then
                Cont = FindClose(hSearch)
                Exit Function
            End If
            filename = StripNulls(WFD.cFileName)
            If (filename <> ".") And (filename <> "..") Then
                LenFile = (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
                
                If (LenFile > Min_size(0) And SearchStr = "*.BMP") _
                   Or (LenFile > Min_size(1) And SearchStr = "*.JPG") _
                   Or (LenFile > Min_size(2)) And SearchStr = "*.GIF" Then
                   
                    FindFilesAPI = FindFilesAPI + LenFile
                    FileCount = FileCount + 1
                    If Check_valid(xxx & filename) Then
                        qt_files = qt_files + 1
                        Form5.qtf(Looking) = qt_files
                        Form5.qtf(Looking).Refresh
                        Form3.List1.AddItem filename
                        path(Form3.List1.NewIndex) = xxx
                      Else
                        'Stop
                    End If
                End If
            End If
            Cont = FindNextFile(hSearch, WFD) ' Get next file
        Wend
        Cont = FindClose(hSearch)
    End If
    ' If there are sub-directories...
    If nDir > 0 Then
        ' Recursively walk into them...
        For i = 0 To nDir - 1
            If abortou Then
                Exit Function
                Cont = FindClose(hSearch)
            End If
            FindFilesAPI = FindFilesAPI + FindFilesAPI(xxx & dirNames(i) & "\", SearchStr, FileCount, DirCount)
        Next i
    End If

End Function


