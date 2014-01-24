Attribute VB_Name = "modFiles"
Option Explicit
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Declare Function SHFileExists Lib "shell32" Alias "#45" (ByVal szPath As String) As Long

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long

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
    cFileName As String * 260
    cAlternate As String * 14
End Type

Private Const DRIVE_FIXED = 3
Private Const DRIVE_RAMDISK = 6

Public Function EnumFiles$(sFolder$)
    Dim hFind&, sFile$, uWFD As WIN32_FIND_DATA, sList$
    If Not FileExists(sFolder) Then Exit Function
    hFind = FindFirstFile(BuildPath(sFolder, "*.*"), uWFD)
    If hFind <= 0 Then Exit Function
    Do
        sFile = TrimNull(uWFD.cFileName)
        If sFile <> "." And sFile <> ".." Then
            sList = sList & "|" & sFile
        End If
        If bAbort Then
            FindClose hFind
            Exit Function
        End If
    Loop Until FindNextFile(hFind, uWFD) = 0
    FindClose hFind
    If sList <> vbNullString Then EnumFiles = Mid$(sList, 2)
End Function

Public Function FileExists(sFile$) As Boolean
    If bIsWinNT Then
        FileExists = IIf(SHFileExists(StrConv(sFile, vbUnicode)) = 1, True, False)
    Else
        FileExists = IIf(SHFileExists(sFile) = 1, True, False)
    End If
End Function

Public Function GetLocalDisks$()
    Dim lDrives&, i&, sDrive$, sLocalDrives$
    lDrives = GetLogicalDrives()
    For i = 0 To 26
        If (lDrives And 2 ^ i) Then
            sDrive = Chr$(Asc("A") + i) & ":\"
            Select Case GetDriveType(sDrive)
                Case DRIVE_FIXED, DRIVE_RAMDISK: sLocalDrives = sLocalDrives & Chr$(Asc("A") + i) & " "
            End Select
        End If
    Next i
    GetLocalDisks = Trim$(sLocalDrives)
End Function
