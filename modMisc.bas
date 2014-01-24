Attribute VB_Name = "modMisc"
Option Explicit
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long    'this file needs to be on the comp

Private Declare Function GetUsername Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long

Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long
'Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function WinVerifyTrust Lib "wintrust.dll" (ByVal hwnd As Long, ByRef pgActionID As GUID, ByRef pWVTData As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As GUID) As Long

Private Declare Function WinVerifyFile Lib "istrusted.dll" Alias "Checkfile" (ByVal sFilename As String) As Boolean

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type WINTRUST_FILE_INFO
    cbStruct As Long
    pcwszFilePath As String
    hFile As Long
    pgKnownSubject As GUID
End Type

Private Type WINTRUST_DATA
    cbStruct As Long
    pPolicyCallbackData As Long
    pSIPClientData As Long
    dwUIChoice As Long
    fdwRevocationChecks As Long
    dwUnionChoice As Long
    pFile As Long 'WINTRUST_FILE_INFO
    pCatalog As Long
    pBlob As Long
    pSgnr As Long
    pCert As Long
    dwStateAction As Long
    hWVTStateData As Long
    pwszURLReference As String
    dwProvFlags As Long
    dwUIContext As Long
End Type

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    '  Optional fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Private Type KEYBDINPUT
    wVk As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Type GENERALINPUT
    dwType As Long
    xi(0 To 23) As Byte
End Type

Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Const WTD_UI_NONE = 2
Private Const WTD_REVOKE_NONE = 0
Private Const WTD_CHOICE_FILE = 1
Private Const WTD_CHOICE_CATALOG = 2
Private Const WTD_SAFER_FLAG = &H100
Private Const WTD_STATEACTION_VERIFY = 1
Private Const WTD_UICONTEXT_EXECUTE = 0

Private Const TRUST_E_PROVIDER_UNKNOWN = -2146762751
Private Const TRUST_E_ACTION_UNKNOWN = -2146762750
Private Const TRUST_E_SUBJECT_FORM_UNKNOWN = -2146762749
Private Const TRUST_E_SUBJECT_NOT_TRUSTED = -2146762748

Private Const WM_KEYDOWN = &H100
Private Const WM_CHAR = &H102
'Private Const WM_SETTEXT = &HC

'Private Const VK_SHIFT = &H10
Private Const VK_HOME = &H24
Private Const VK_RIGHT = &H27
Private Const VK_LEFT = &H25
'Private Const VK_OEM_MINUS = &HBD
'Private Const VK_OEM_5 = &HDC
'Private Const KEYEVENTF_KEYUP = &H2
'Private Const INPUT_MOUSE = 0
'Private Const INPUT_KEYBOARD = 1
'Private Const INPUT_HARDWARE = 2

Private Const SW_SHOWNORMAL = 1

Private Const SEE_MASK_DOENVSUBST = &H200
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_INVOKEIDLIST = &HC

Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800

Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_NT_WORKSTATION = 1

Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3

Public bShowEmpty As Boolean
Public bShowCLSIDs As Boolean
Public bShowCmts As Boolean
Public bShowPrivacy As Boolean
Public bAutoSave As Boolean
Public sAutoSavePath$

Public bShowUsers As Boolean
Public bShowHardware As Boolean

Public sWinVersion$, sWinDir$, sSysDir$
Public bIsWinNT As Boolean, bIsWinNT4 As Boolean
Public lEnumBufLen&
Public sUsernames$(), sHardwareCfgs$()

Private lTicks&
Public bDebug As Boolean
Public bAbort As Boolean

Public bShowLargeHosts As Boolean, bShowLargeZones As Boolean

Public Const SEC_RUNNINGPROCESSES = "Running processes"
Public Const SEC_AUTOSTARTFOLDERS = "Autostart folders"
Public Const SEC_TASKSCHEDULER = "Task Scheduler jobs"
Public Const SEC_INIFILE = ".Ini file values"
Public Const SEC_AUTORUNINF = "Autorun.inf files"
Public Const SEC_BATFILES = "Autostarting batch files"
Public Const SEC_EXPLORERCLONES = "Explorer.exe clones"

Public Const SEC_BHOS = "Browser Helper Objects"
Public Const SEC_IETOOLBARS = "IE Toolbars"
Public Const SEC_IEEXTENSIONS = "IE Buttons / Tools"
Public Const SEC_IEBARS = "IE Bars"
Public Const SEC_IEMENUEXT = "IE menu extensions"
Public Const SEC_IEBANDS = "IE Bands"
Public Const SEC_DPFS = "Downloaded Program Files"
Public Const SEC_ACTIVEX = "ActiveX objects"
Public Const SEC_DESKTOPCOMPONENTS = "Desktop Components"
Public Const SEC_URLSEARCHHOOKS = "URL Search Hooks"

Public Const SEC_APPPATHS = "Application Paths"
Public Const SEC_SHELLEXT = "Approved Shell Extensions"
Public Const SEC_COLUMNHANDLERS = "Column Handlers"
Public Const SEC_CMDPROC = "Command processor autostart"
Public Const SEC_CONTEXTMENUHANDLERS = "Contextmenu Handlers"
Public Const SEC_DRIVERFILTERS = "Driver Filters"
Public Const SEC_DRIVERS32 = "Drivers32 libraries"
Public Const SEC_IMAGEFILEEXECUTION = "Image File Execution"
Public Const SEC_LSAPACKAGES = "LSA packages"
Public Const SEC_MOUNTPOINTS = "Mountpoints"
Public Const SEC_MPRSERVICES = "MPR Services"
Public Const SEC_ONREBOOT = "On-reboot actions"
Public Const SEC_POLICIES = "Policies"
Public Const SEC_PRINTMONITORS = "Print Monitors"
Public Const SEC_PROTOCOLS = "Protocol/Filter handlers"
Public Const SEC_INIMAPPING = "Registry-mapped .ini files"
Public Const SEC_REGRUNKEYS = "Registry 'Run' keys"
Public Const SEC_REGRUNEXKEYS = "Registry 'Run' subkeys"
Public Const SEC_SECURITYPROVIDERS = "Security Providers"
Public Const SEC_SERVICES = "Services"
Public Const SEC_SHAREDTASKSCHEDULER = "Shared Task Scheduler"
Public Const SEC_SHELLCOMMANDS = "Shell commands"
Public Const SEC_SHELLEXECUTEHOOKS = "Shell Execute Hooks"
Public Const SEC_SSODL = "ShellServiceObjectDelayLoad"
Public Const SEC_UTILMANAGER = "Utility Manager autostarts"
Public Const SEC_WINLOGON = "Winlogon autostarts"
Public Const SEC_SCRIPTPOLICIES = "WinNT script policies"
Public Const SEC_WINSOCKLSP = "Winsock LSPs"
Public Const SEC_WOW = "WOW compatibility"
Public Const SEC_3RDPARTY = "3rd party program autostarts"

Public Const SEC_RESETWEBSETTINGS = "Reset Web Settings URLs"
Public Const SEC_IEURLS = "Internet Explorer URLs"
Public Const SEC_URLPREFIX = "Default URL prefixes"
Public Const SEC_HOSTSFILEPATH = "Hosts file path"

Public Const SEC_HOSTSFILE = "Hosts file items"
Public Const SEC_KILLBITS = "ActiveX kill bits"
Public Const SEC_ZONES = "IE Security Zones"
Public Const SEC_MSCONFIG9X = "Msconfig 9x/ME disabled items"
Public Const SEC_MSCONFIGXP = "Msconfig XP disabled items"
Public Const SEC_STOPPEDSERVICES = "Stopped/disabled services"
Public Const SEC_XPSECURITY = "Windows XP Security Center"

Public Sub Status(s$)
    frmMain.stbStatus.SimpleText = s
    DoEvents
End Sub

Public Function TrimNull$(s$)
    If InStr(s, Chr$(0)) > 0 Then
        TrimNull = Left$(s, InStr(s, Chr$(0)) - 1)
    Else
        TrimNull = s
    End If
End Function

Public Function GetWindowsVersion$()
    Dim uOVI As OSVERSIONINFO, uOVI2 As OSVERSIONINFOEX
    Dim sFriendlyVer$, sCSD$
    With uOVI
        .dwOSVersionInfoSize = Len(uOVI)
        GetVersionEx uOVI
        If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
            bIsWinNT = True
            'this is the buffer size for RegEnumValueEx()
            lEnumBufLen = 16400
            uOVI2.dwOSVersionInfoSize = Len(uOVI2)
            GetVersionEx uOVI2
            sWinVersion = "WinNT " & .dwMajorVersion & "." & _
                        Format$(.dwMinorVersion, "00") & "." & _
                        Format$(.dwBuildNumber, "0000")
        Else
            'this is the buffer size for RegEnumValueEx()
            lEnumBufLen = 260
            sWinVersion = "Win9x " & .dwMajorVersion & "." & _
                        Format$(.dwMinorVersion, "00") & "." & _
                        Format$(.dwBuildNumber And &HFFF, "0000")
        End If
        sCSD = Trim$(TrimNull(.szCSDVersion))
        If InStr(1, sCSD, "Service Pack ", vbTextCompare) > 0 Then sCSD = Replace(sCSD, "Service Pack ", "SP")
        If InStr(1, sCSD, "Service Pack", vbTextCompare) > 0 Then sCSD = Replace(sCSD, "Service Pack", "SP")
    
        Select Case .dwPlatformId
            Case VER_PLATFORM_WIN32_NT
                Select Case .dwMajorVersion
                    Case 3
                        sFriendlyVer = Trim$("Windows NT3." & Format$(.dwMinorVersion, "00") & " " & sCSD)
                    Case 4
                        sFriendlyVer = Trim$("Windows NT4 " & sCSD)
                        bIsWinNT4 = True
                    Case 5
                        Select Case .dwMinorVersion
                            Case 0
                                sFriendlyVer = Trim$("Windows 2000 " & sCSD)
                            Case 1
                                sFriendlyVer = Trim$("Windows XP " & sCSD)
                            Case 2
                                If (uOVI2.wProductType And VER_NT_WORKSTATION) Then
                                    'this is bullshit
                                    'sFriendlyVer = Trim$("Windows XP 64bit " & sCSD)
                                    sFriendlyVer = Trim$("Windows 2003 Small Business Server " & sCSD)
                                Else
                                    sFriendlyVer = Trim$("Windows 2003 " & sCSD)
                                End If
                        End Select
                    Case 6
                        sFriendlyVer = Trim$("Windows Vista " & sCSD)
                End Select
            Case VER_PLATFORM_WIN32_WINDOWS
                Select Case .dwMajorVersion
                    Case 4
                        Select Case .dwMinorVersion
                            Case 0
                                Select Case sCSD
                                    Case "B", "C"
                                        sFriendlyVer = "Windows 95 OSR2"
                                    Case Else
                                        sFriendlyVer = "Windows 95 Gold"
                                End Select
                            Case 10
                                Select Case sCSD
                                    Case "A"
                                        sFriendlyVer = "Windows 98 SE"
                                    Case Else
                                        sFriendlyVer = "Windows 98 Gold"
                                End Select
                            Case 90
                                sFriendlyVer = "Windows ME"
                        End Select
                End Select
        End Select
    End With
    GetWindowsVersion = sFriendlyVer & " (" & sWinVersion & ")"
    
    sWinDir = String$(260, 0)
    GetWindowsDirectory sWinDir, Len(sWinDir)
    sWinDir = TrimNull(sWinDir)
    sSysDir = String$(260, 0)
    GetSystemDirectory sSysDir, Len(sSysDir)
    sSysDir = TrimNull(sSysDir)
End Function

Public Function BuildPath$(sFolder$, sFile$)
    If Right$(sFolder, 1) = "\" Then
        BuildPath = sFolder & sFile
    Else
        BuildPath = sFolder & "\" & sFile
    End If
End Function

Public Function GetUser$()
    Dim sBuf$
    sBuf = String$(260, 0)
    GetUsername sBuf, Len(sBuf)
    GetUser = TrimNull(sBuf)
End Function

Public Function GetComputer$()
    Dim sBuf$
    sBuf = String$(260, 0)
    GetComputerName sBuf, Len(sBuf)
    GetComputer = TrimNull(sBuf)
End Function

Public Function InputFile$(sFile$)
    'this uses APIs instead of Input(), which is ~3x slower and doesn't cache :P
    Dim hFile&, uBuffer() As Byte, lFileSize&, lBytesRead&
    hFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0, OPEN_EXISTING, 0, 0)
    If hFile = -1 Then Exit Function
    
    'second parameter is dwSizeHigh, we ignore that
    'since it's only used if the file is >2GB
    lFileSize = GetFileSize(hFile, 0)
    If lFileSize = -1 Or lFileSize = 0 Then
        CloseHandle hFile
        Exit Function
    End If
    
    ReDim uBuffer(lFileSize - 1)
    If ReadFile(hFile, uBuffer(0), lFileSize, lBytesRead, ByVal 0) > 0 Then
        If lBytesRead <> lFileSize Then
            ReDim Preserve uBuffer(lBytesRead)
        End If
        InputFile = StrConv(uBuffer, vbUnicode)
    End If
    CloseHandle hFile
End Function

Public Function CmnDialogSave$(sTitle$, sDefFile$, sFilter$)
    Dim uOFN As OPENFILENAME
    With uOFN
        .lStructSize = Len(uOFN)
        .hwndOwner = frmMain.hwnd
        .lpstrFile = sDefFile & String$(260 - Len(sDefFile), 0)
        .lpstrFilter = Replace(sFilter, "|", Chr$(0)) & Chr$(0) & Chr$(0)
        .lpstrInitialDir = App.Path
        .lpstrTitle = sTitle
        .nMaxFile = Len(.lpstrFile)
        .flags = OFN_HIDEREADONLY Or OFN_NONETWORKBUTTON Or OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT
        GetSaveFileName uOFN
        CmnDialogSave = TrimNull(.lpstrFile)
    End With
End Function

Public Sub ShowFile(sFile$)
    Dim sSEI As SHELLEXECUTEINFO
    If Not FileExists(sFile) Then Exit Sub
    With sSEI
        .cbSize = Len(sSEI)
        .hwnd = frmMain.hwnd
        .lpFile = sWinDir & "\explorer.exe"
        .lpParameters = "/select," & sFile
        .lpVerb = "open"
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .nShow = 1
    End With
    ShellExecuteEx sSEI
End Sub

Public Sub ShowFileProp(sFile$)
    Dim sSEI As SHELLEXECUTEINFO
    With sSEI
        .cbSize = Len(sSEI)
        .hwnd = frmMain.hwnd
        .lpFile = sFile
        .lpVerb = "properties"
        .fMask = SEE_MASK_DOENVSUBST Or SEE_MASK_FLAG_NO_UI Or _
                 SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST
    End With
    ShellExecuteEx sSEI
End Sub

Public Sub SendToNotepad(sFile$)
    If Not FileExists(sFile) Then Exit Sub
    Dim sNotepad$
    sNotepad = RegGetString(HKEY_CLASSES_ROOT, ".txt", vbNullString)
    sNotepad = ExpandEnvironmentVars(RegGetString(HKEY_CLASSES_ROOT, sNotepad & "\shell\open\command", vbNullString))
    If sNotepad <> vbNullString Then
        sNotepad = Left$(sNotepad, InStr(1, sNotepad, ".exe", vbTextCompare) + 3)
        If Not FileExists(sNotepad) Then sNotepad = sWinDir & "\notepad.exe"
    End If
    
    Dim sSEI As SHELLEXECUTEINFO
    With sSEI
        .cbSize = Len(sSEI)
        .hwnd = frmMain.hwnd
        '.lpFile = sWinDir & "\notepad.exe"
        .lpFile = sNotepad
        .lpVerb = "open"
        .lpParameters = sFile
        .fMask = SEE_MASK_DOENVSUBST Or SEE_MASK_FLAG_NO_UI Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_NOCLOSEPROCESS
        .nShow = 1
    End With
    ShellExecuteEx sSEI
End Sub

Public Function GuessFullpathFromAutorun$(sAutorunFile$)
    Dim sFile$
    If Trim$(sAutorunFile) = vbNullString Then Exit Function
    sFile = sAutorunFile
    
    'already full path? return
    If InStr(sFile, "\") > 0 And FileExists(sFile) Then
        GuessFullpathFromAutorun = sFile
        Exit Function
    End If
    'if enclosed in quotes, assume that's the full path and return
    If InStr(sFile, """") > 0 Then
        sFile = Mid$(sFile, 2)
        sFile = Left$(sFile, InStr(sFile, """") - 1)
    ElseIf InStr(sFile, "\") > 0 And InStr(sFile, " ") > 0 And InStr(1, sFile, ".exe", vbTextCompare) < Len(sFile) - 3 Then
        'cut off everything after .exe if it's a full path
        sFile = Left$(sFile, InStr(1, sFile, ".exe", vbTextCompare) + 3)
    Else
        'strip everything after the first space (parameters)
        If InStr(sFile, " ") > 0 Then sFile = Mid$(sFile, 1, InStr(sFile, " ") - 1)
        'add extension if not there
        If InStr(sFile, ".") = 0 Then sFile = sFile & ".exe"
        'try a few common paths to find the file
        If Not FileExists(sFile) Then
            'windir
            If FileExists(BuildPath(sWinDir, sFile)) Then
                sFile = BuildPath(sWinDir, sFile)
            Else
                'sysdir
                If FileExists(BuildPath(sSysDir, sFile)) Then
                    sFile = BuildPath(sSysDir, sFile)
                Else
                    'root
                    If FileExists(BuildPath(Left$(sWinDir, 3), sFile)) Then
                        sFile = BuildPath(Left$(sWinDir, 3), sFile)
                    End If
                End If
            End If
        End If
    End If
    If FileExists(sFile) Then
        GuessFullpathFromAutorun = sFile
    Else
        GuessFullpathFromAutorun = sAutorunFile
    End If
End Function

Public Sub GetUsernames()
    ReDim sUsernames(0)
    Dim sKeys$(), i%
    sKeys = Split(RegEnumSubKeys(HKEY_USERS, vbNullString), "|")
    For i = 0 To UBound(sKeys)
        If InStr(1, sKeys(i), "_Classes", vbTextCompare) = 0 Then
            ReDim Preserve sUsernames(UBound(sUsernames) + 1)
            sUsernames(UBound(sUsernames) - 1) = sKeys(i)
        End If
    Next i
    ReDim Preserve sUsernames(UBound(sUsernames) - 1)
End Sub

Public Function MapSIDToUsername$(sSID$)
    If UCase$(sSID) = ".DEFAULT" Then
        MapSIDToUsername = "Default user"
        Exit Function
    End If

    'dirty dirty WMI function
    Dim objWMI As Object, objSID As Object
    On Error Resume Next
    Set objWMI = GetObject("winmgmts:{impersonationLevel=Impersonate}")
    If InStr(sSID, "_Classes") = 0 Then
        Set objSID = objWMI.Get("Win32_SID.SID='" & sSID & "'")
        MapSIDToUsername = CStr(objSID.AccountName)
        ' & " (" & CStr(objSID.ReferencedDomainName) & ")"
        If MapSIDToUsername = vbNullString Then MapSIDToUsername = sSID
        Set objSID = Nothing
    End If
    Set objWMI = Nothing
End Function

Public Function ExpandEnvironmentVars$(s$)
    Dim lLen&, sDummy$
    If LenB(s) = 0 Then
        ExpandEnvironmentVars = s
        Exit Function
    End If
    If InStr(s, "%") = 0 Then
        ExpandEnvironmentVars = s
        Exit Function
    End If
    If InStr(s, "%systemroot%") > 0 Then
        s = Replace(s, "%systemroot%", sWinDir, , , vbTextCompare)
    End If
    If InStr(s, "%windir%") > 0 Then
        s = Replace(s, "%windir%", sWinDir, , , vbTextCompare)
    End If
    
    If InStr(s, "%") = 0 Then
        ExpandEnvironmentVars = s
        Exit Function
    End If
    lLen = ExpandEnvironmentStrings(s, ByVal 0, 0)
    If lLen > 0 Then
        sDummy = String$(lLen, 0)
        ExpandEnvironmentStrings s, sDummy, Len(sDummy)
        sDummy = TrimNull(sDummy)
    Else
        sDummy = s
    End If
    ExpandEnvironmentVars = sDummy
End Function

Public Sub GetHardwareCfgs()
    Dim lDefault&, lCurrent&, lLastKnownGood&, lFailed&
    lDefault = RegGetDword(HKEY_LOCAL_MACHINE, "System\Select", "Default")
    lCurrent = RegGetDword(HKEY_LOCAL_MACHINE, "System\Select", "Current")
    lLastKnownGood = RegGetDword(HKEY_LOCAL_MACHINE, "System\Select", "LastKnownGood")
    lFailed = RegGetDword(HKEY_LOCAL_MACHINE, "System\Select", "Failed")
    
    ReDim sHardwareCfgs(0)
    sHardwareCfgs(0) = "ControlSet" & Format$(lCurrent, "000")
    If lDefault <> lCurrent And lDefault > 0 Then
        sHardwareCfgs(UBound(sHardwareCfgs)) = "ControlSet" & Format$(lDefault, "000")
    End If
    If lLastKnownGood <> lCurrent And lLastKnownGood > 0 Then
        ReDim Preserve sHardwareCfgs(UBound(sHardwareCfgs) + 1)
        sHardwareCfgs(UBound(sHardwareCfgs)) = "ControlSet" & Format$(lLastKnownGood, "000")
    End If
    If lFailed <> lCurrent And lFailed > 0 Then
        ReDim Preserve sHardwareCfgs(UBound(sHardwareCfgs) + 1)
        sHardwareCfgs(UBound(sHardwareCfgs)) = "ControlSet" & Format$(lFailed, "000")
    End If
    'MsgBox Join(sHardwareCfgs, vbCrLf)
End Sub

Public Function MapControlSetToHardwareCfg$(sControlSet$)
    Dim lThisCS&, lDefault&, lCurrent&, lFailed&, lLKG&
    lThisCS = Val(Right$(sControlSet, 3))
    
    lDefault = RegGetDword(HKEY_LOCAL_MACHINE, "System\Select", "Default")
    lCurrent = RegGetDword(HKEY_LOCAL_MACHINE, "System\Select", "Current")
    lFailed = RegGetDword(HKEY_LOCAL_MACHINE, "System\Select", "Failed")
    lLKG = RegGetDword(HKEY_LOCAL_MACHINE, "System\Select", "LastKnownGood")
    
    Select Case lThisCS
        Case lDefault: MapControlSetToHardwareCfg = "Default"
        Case lCurrent: MapControlSetToHardwareCfg = "Current"
        Case lFailed:  MapControlSetToHardwareCfg = "Failed"
        Case lLKG:     MapControlSetToHardwareCfg = "Last known good"
    End Select
End Function

Public Sub RegistryJump(sRegKey$)
    Dim i&, lHive&, sKey$, uSEI As SHELLEXECUTEINFO
    Dim hwndRegedit&, hwndTreeView&, hwndListView&
    'verify the key actually exists
    Select Case UCase$(Left$(sRegKey, InStr(sRegKey, "\") - 1))
        Case "HKEY_CLASSES_ROOT": lHive = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_USER": lHive = HKEY_CURRENT_USER
        Case "HKEY_LOCAL_MACHINE": lHive = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS": lHive = HKEY_USERS
        Case Else: Exit Sub
    End Select
    sKey = Mid$(sRegKey, InStr(sRegKey, "\") + 1)
    If Not RegKeyExists(lHive, sKey) Then Exit Sub
    
    'start regedit and wait until it's done loading
    With uSEI
        .cbSize = Len(uSEI)
        .lpVerb = "open"
        .lpFile = BuildPath(sWinDir, "regedit.exe")
        .fMask = SEE_MASK_NOCLOSEPROCESS
        .nShow = SW_SHOWNORMAL
    End With
    If ShellExecuteEx(uSEI) = 0 Then
        Status "Unable to start Regedit."
        Exit Sub
    End If
    WaitForInputIdle uSEI.hProcess, 10000
    
    'find the regedit window and it's components
    hwndRegedit = FindWindow("RegEdit_RegEdit", vbNullString)
    hwndTreeView = FindWindowEx(hwndRegedit, 0, "SysTreeView32", vbNullString)
    hwndListView = FindWindowEx(hwndRegedit, 0, "SysListView32", vbNullString)
    If hwndTreeView = 0 Or hwndListView = 0 Then
        Status "Unable to start Regedit."
        Exit Sub
    End If
    SetForegroundWindow hwndRegedit
    
    'if regedit was already open, collapse any open keys
    For i = 0 To 20
        SendMessage hwndTreeView, WM_KEYDOWN, VK_LEFT, 0
    Next i
    SendMessage hwndTreeView, WM_KEYDOWN, VK_HOME, 0
    SendMessage hwndTreeView, WM_KEYDOWN, VK_RIGHT, 0
    
    'type out the key we want to jump to
    For i = 1 To Len(sRegKey)
        Select Case Mid$(sRegKey, i, 1)
            Case "\": SendMessage hwndTreeView, WM_KEYDOWN, VK_RIGHT, 0
            Case Else: SendMessage hwndTreeView, WM_CHAR, Asc(UCase$(Mid$(sRegKey, i, 1))), 0
        End Select
        DoEvents
        Sleep 50
    Next i
    SendMessage hwndTreeView, WM_KEYDOWN, VK_RIGHT, 0
End Sub

Public Sub RegistryJump_(sRegKey$)
    'this sub has a bug! if regkeys exist similar to the target that
    'contain spaces, things are screwed up. e.g. a jump to any
    '"Internet Explorer" key will fail if there is also a key
    '"Internet Account Manager" present. the space is somehow to blame.
    Dim i&, lHive&, sKey$, sKeyStrokes$
    'verify the key actually exists
    Select Case UCase$(Left$(sRegKey, InStr(sRegKey, "\") - 1))
        Case "HKEY_CLASSES_ROOT": lHive = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_USER": lHive = HKEY_CURRENT_USER
        Case "HKEY_LOCAL_MACHINE": lHive = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS": lHive = HKEY_USERS
        Case Else: Exit Sub
    End Select
    sKey = Mid$(sRegKey, InStr(sRegKey, "\") + 1)
    If Not RegKeyExists(lHive, sKey) Then Exit Sub
    
    Shell BuildPath(sWinDir, "regedit.exe"), vbNormalFocus
    'Shell "notepad.exe", vbNormalFocus
    
    sKeyStrokes = sRegKey
    sKeyStrokes = Replace(sKeyStrokes, "{", "{{}")
    sKeyStrokes = Replace(sKeyStrokes, "}", "{}}")
    sKeyStrokes = Replace(sKeyStrokes, "{{{}}", "{{}")
    sKeyStrokes = Replace(sKeyStrokes, "~", "{~}")
    sKeyStrokes = Replace(sKeyStrokes, "%", "{%}")
    sKeyStrokes = Replace(sKeyStrokes, "^", "{^}")
    sKeyStrokes = Replace(sKeyStrokes, "(", "{(}")
    sKeyStrokes = Replace(sKeyStrokes, ")", "{)}")
    sKeyStrokes = Replace(sKeyStrokes, "+", "{+}")
    sKeyStrokes = Replace(sKeyStrokes, "[", "{[}")
    sKeyStrokes = Replace(sKeyStrokes, "]", "{]}")
    sKeyStrokes = Replace(sKeyStrokes, "\", "{RIGHT}")
    
    sKeyStrokes = Replace(sKeyStrokes, " ", vbNullString)
    
    SendKeys "{HOME}", True
    SendKeys sKeyStrokes, True
    SendKeys "{RIGHT}", True
    
'    For i = 1 To Len(sRegKey)
'        Select Case Mid$(sRegKey, i, 1)
'            Case "\" 'send right arrow to expand branch
'                SendKeys "{RIGHT}"
'            'these are special characters and need curly braces
'            Case "~": SendKeys "{~}"
'            Case "%": SendKeys "{%}"
'            Case "^": SendKeys "{^}"
'            Case "(": SendKeys "{(}"
'            Case ")": SendKeys "{)}"
'            Case "+": SendKeys "{+}"
'            Case "{": SendKeys "{{}"
'            Case "}": SendKeys "{}}"
'            'the ONLY character not allowed in a regkey (apart from
'            'high-ascii crap, I suppose) is the BACKSLASH :)
'            Case Else: SendKeys Mid$(sRegKey, i, 1)
'        End Select
'        DoEvents
'    Next i
'    SendKeys "{RIGHT}"
End Sub

'Public Sub RegistryJump(ByVal sRegKey$)
'    Dim uSEI As SHELLEXECUTEINFO, i&, sChr$, lVkey&, bShift As Boolean
'    Dim uGInput() As GENERALINPUT, uKInput As KEYBDINPUT
'    Dim lHive&
'    Select Case Left$(sRegKey, InStr(sRegKey, "\") - 1)
'        Case "HKEY_LOCAL_MACHINE": lHive = HKEY_LOCAL_MACHINE
'        Case "HKEY_CURRENT_USER": lHive = HKEY_CURRENT_USER
'        Case "HKEY_CLASSES_ROOT": lHive = HKEY_CLASSES_ROOT
'        Case "HKEY_USERS": lHive = HKEY_USERS
'        Case Else: Exit Sub
'    End Select
'    If Not RegKeyExists(lHive, Mid$(sRegKey, InStr(sRegKey, "\") + 1)) Then Exit Sub
'
'    With uSEI
'        .cbSize = Len(uSEI)
'        .lpFile = sWinDir & "\regedit.exe"
'        '.lpFile = sWinDir & "\notepad.exe"
'        .nShow = 1
'    End With
'    ShellExecuteEx uSEI
'    DoEvents
'    Sleep 1000
'
'    sRegKey = "\" & Replace(sRegKey, "_", "^-^") & "\"
'    ReDim uGInput(Len(sRegKey) * 2)
'    For i = 1 To Len(sRegKey)
'        sChr = Mid$(sRegKey, i, 1)
'        Select Case sChr
'            Case "a" To "z": lVkey = Asc(sChr) - 32
'            Case "\":
'                If i = 1 Then
'                    lVkey = VK_HOME
'                Else
'                    lVkey = VK_RIGHT
'                End If
'            Case "^":
'                lVkey = VK_SHIFT
'                bShift = Not bShift
'            Case "-": lVkey = VK_OEM_MINUS
'            Case Else: lVkey = Asc(sChr)
'        End Select
'
'        If Not (lVkey = VK_SHIFT And Not bShift) Then
'            uKInput.dwFlags = 0
'            uKInput.wVk = lVkey
'            uGInput(i * 2 - 1).dwType = INPUT_KEYBOARD
'            CopyMemory uGInput(i * 2 - 1).xi(0), uKInput, Len(uKInput)
'        End If
'
'        If Not (lVkey = VK_SHIFT And bShift) Then
'            uKInput.dwFlags = KEYEVENTF_KEYUP
'            uKInput.wVk = lVkey
'            uGInput(i * 2).dwType = INPUT_KEYBOARD
'            CopyMemory uGInput(i * 2).xi(0), uKInput, Len(uKInput)
'        End If
'    Next i
'    SendInput UBound(uGInput) + 1, uGInput(1), Len(uGInput(1))
'End Sub

Public Function IsRunningInIDE() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err Then IsRunningInIDE = True
End Function

Public Sub DoTicks(tvwMain As TreeView, Optional sNode$)
    If Not bDebug Then Exit Sub
    If sNode = vbNullString Then
        'start
        lTicks = GetTickCount
    Else
        'stop + display
        lTicks = GetTickCount - lTicks
        On Error Resume Next
        tvwMain.Nodes.Add sNode, tvwChild, sNode & "Ticks", " Time: " & lTicks & " ms", "clock"
    End If
End Sub

Public Function IsCLSID(sCLSID$) As Boolean
    If sCLSID Like "{????????-????-????-????-????????????}" Then IsCLSID = True
End Function

Public Function GetStringResFromDLL$(sFile$, iResID%)
    Dim hMod&, lLen&, sBuf$
    If FileExists(sFile) Then
        hMod = LoadLibrary(sFile)
        If hMod > 0 Then
            sBuf = String$(260, 0)
            lLen = LoadString(hMod, Abs(iResID), sBuf, Len(sBuf))
            If lLen > 0 Then GetStringResFromDLL = TrimNull(sBuf)
            FreeLibrary hMod
        End If
    End If
End Function

Public Sub ShellRun(sFile$, Optional bHidden As Boolean = False)
    Dim uSEI As SHELLEXECUTEINFO
    With uSEI
        .cbSize = Len(uSEI)
        .lpFile = sFile
        .lpVerb = "open"
        .nShow = Not Abs(CLng(bHidden))
    End With
    ShellExecuteEx uSEI
End Sub

Public Function VerifyFileSignature(sFile$) As Integer
    If Not FileExists(App.Path & "\istrusted.dll") Then
        If MsgBox("To verify file signatures, StartupList needs to " & _
                  "download an external library from www.merijn.org. " & _
                  vbCrLf & vbCrLf & "Continue?", vbYesNo + vbQuestion) = vbYes Then
            If DownloadFile("http://www.merijn.org/files/istrusted.dll", App.Path & "\istrusted.dll") Then
                'file downloaded ok, continue
            Else
                'file download failed
                bAbort = True
                VerifyFileSignature = -1
                Exit Function
            End If
        Else
            'user aborted download
            bAbort = True
            VerifyFileSignature = -1
            Exit Function
        End If
    End If
    
    If WinVerifyFile(sFile) Then
        VerifyFileSignature = 1
    Else
        VerifyFileSignature = 0
    End If
End Function
'    If Not bIsWinNT Then Exit Function
'    If Not FileExists(sFile) Then Exit Function
'
'    Dim uVerifyV2 As GUID, uWTfileinfo As WINTRUST_FILE_INFO
'    Dim uWTdata As WINTRUST_DATA, lRet&
'    With uWTfileinfo
'        .cbStruct = Len(uWTfileinfo)
'        .pcwszFilePath = sFile
'    End With
'    With uWTdata
'        .pPolicyCallbackData = 0
'        .pSIPClientData = 0
'        .dwStateAction = WTD_STATEACTION_VERIFY
'        .hWVTStateData = 0
'        .pwszURLReference = 0
'        .dwProvFlags = 0
'        .dwUIContext = WTD_UICONTEXT_EXECUTE
'
'        .cbStruct = Len(uWTdata)
'        .dwUIChoice = WTD_UI_NONE
'        .fdwRevocationChecks = WTD_REVOKE_NONE
'        .dwUnionChoice = WTD_CHOICE_FILE
'        .dwProvFlags = 0 'WTD_SAFER_FLAG
'        .pFile = VarPtr(uWTfileinfo)
'    End With
'    If CLSIDFromString(StrPtr("{00AAC56B-CD44-11d0-8CC2-00C04FC295EE}"), uVerifyV2) = 0 Then
'        lRet = WinVerifyTrust(0, uVerifyV2, uWTdata)
'    End If
'    If lRet = 0 Then
'        VerifyFileSignature = True
'    Else
'        Select Case lRet
'            Case TRUST_E_ACTION_UNKNOWN: MsgBox "TRUST_E_ACTION_UNKNOWN"
'            Case TRUST_E_PROVIDER_UNKNOWN: MsgBox "TRUST_E_PROVIDER_UNKNOWN"
'            Case TRUST_E_SUBJECT_FORM_UNKNOWN: MsgBox "TRUST_E_SUBJECT_FORM_UNKNOWN"
'            Case TRUST_E_SUBJECT_NOT_TRUSTED: MsgBox "TRUST_E_SUBJECT_FORM_UNKNOWN"
'        End Select
'    End If

Public Sub RunScannerGetMD5(sFile$, sKey$)
    Dim sMD5$, sAppVer$, sSection$
    sMD5 = GetFileMD5(sFile)
    sAppVer = "StartupList" & App.Major & "." & Format$(App.Minor, "00") & "." & App.Revision
    sSection = GetRunScannerItem(GetSectionFromKey(sKey), sKey)
        
    ShellRun "http://www.runscanner.net/getMD5.aspx?" & _
      "MD5=" & sMD5 & _
      "&source=" & sAppVer & _
      "&item=" & sSection
End Sub

Public Sub RunScannerGetCLSID(sCLSID$, sKey$)
    Dim sAppVer$, sSection$
    sAppVer = "StartupList" & App.Major & "." & Format$(App.Minor, "00") & "." & App.Revision
    sSection = GetRunScannerItem(GetSectionFromKey(sKey), sKey)
    
    ShellRun "http://www.runscanner.net/getGUID.aspx?GUID=" & sCLSID & _
          "&source=StartupList" & App.Major & "." & Format$(App.Minor, "00") & "." & App.Revision
End Sub

Private Function GetRunScannerItem$(sSection$, sKey$)
    Select Case sSection
        Case "RunningProcesses"
            GetRunScannerItem = "001"
        Case "RunRegkeys"
            If InStr(sKey, "System") > 0 Then
                If InStr(sKey, "Once") > 0 Then
                    GetRunScannerItem = "136"
                Else
                    GetRunScannerItem = "002" 'system registry autorun
                End If
            End If
            If InStr(sKey, "User") > 0 Then
                If InStr(sKey, "Once") > 0 Then
                    GetRunScannerItem = "135"
                Else
                    GetRunScannerItem = "003" 'user registry autorun
                End If
            End If
        Case "AutoStartFoldersCommon Startup", "AutoStartFoldersUser Common Startup", "Windows Vista common Startup"
            GetRunScannerItem = "004" 'all users startup
        Case "AutoStartFoldersStartup", "AutoStartFoldersUser Startup"
            GetRunScannerItem = "005" 'user startup
        Case "Windows Vista roaming profile Startup", "Windows Vista roaming profile Startup 2"
            GetRunScannerItem = "007" 'roaming user startup
        Case "NTServices", "VxDServices"
            GetRunScannerItem = "010" 'installed services
        Case "ProtocolsFilter"
            GetRunScannerItem = "030" 'installed protocol filters
        Case "ProtocolsHandler"
            GetRunScannerItem = "031" 'installed protocol handlers
        Case "WinLogonL"
            If InStr(sKey, "WinLogonL0") > 0 Then
                GetRunScannerItem = "033" 'winlogon userinit
            End If
        Case "IniMapping"
            If sKey = "IniMapping0" Then
                GetRunScannerItem = "034"
            Else
                If CInt(Right(sKey, 1)) Mod 2 = 0 Then
                    GetRunScannerItem = "140"
                Else
                    GetRunScannerItem = "139"
                End If
            End If
        Case "ActiveX"
            GetRunScannerItem = "035"
        Case "WinLogonL1"
            GetRunScannerItem = "037"
        Case "WinLogonL3"
            GetRunScannerItem = "038"
        Case "URLSearchHooks"
            GetRunScannerItem = "040"
        Case "IEToolbars"
            If InStr(sKey, "IEToolbarsUserShell") > 0 Then
                GetRunScannerItem = "045"
            ElseIf InStr(sKey, "IEToolbarsUserWeb") > 0 Then
                GetRunScannerItem = "046"
            Else
                GetRunScannerItem = "041"
            End If
        Case "IEExtensions"
            GetRunScannerItem = "042"
        Case "ShellExecuteHooks"
            GetRunScannerItem = "050"
        Case "SharedTaskScheduler"
            GetRunScannerItem = "051"
        Case "BHO"
            GetRunScannerItem = "052"
        Case "SSODL"
            GetRunScannerItem = "060"
        Case "ShellExts"
            GetRunScannerItem = "061"
        Case "ColumnHandlers"
            GetRunScannerItem = "062"
        Case "OnRebootActionsBootExecute"
            GetRunScannerItem = "063"
        Case "WOWKnownDlls", "WOWKnownDlls32b"
            GetRunScannerItem = "064"
        Case "ImageFileExecution"
            GetRunScannerItem = "065"
        Case "WinLogonL4"
            GetRunScannerItem = "066"
        Case "WinLogonNotify"
            GetRunScannerItem = "067"
        Case "WinsockLSPProtocols"
            GetRunScannerItem = "068"
        Case "PrintMonitors"
            GetRunScannerItem = "069"
        Case "TaskSchedulerJobs"
            If InStr(sKey, "System") = 0 Then
                GetRunScannerItem = "073"
            Else
                GetRunScannerItem = "074"
            End If
        Case "IEURLs"
            GetRunScannerItem = "100"
        Case "IEExplBars"
            GetRunScannerItem = "102"
        Case "DPF"
            GetRunScannerItem = "104"
        Case "WinsockLSPNamespaces"
            GetRunScannerItem = "107"
        Case "WinLogonW"
            If InStr(sKey, "WinLogonW0") > 0 Then
                GetRunScannerItem = "121"
            End If
        Case "WinLogonGinaDLL"
            GetRunScannerItem = "122"
        Case "RunExRegkeys"
            If InStr(sKey, "System") > 0 Then
                If InStr(sKey, "Ex") > 0 Then
                    GetRunScannerItem = "138"
                Else
                    GetRunScannerItem = "136"
                End If
            ElseIf InStr(sKey, "User") > 0 Then
                If InStr(sKey, "Ex") > 0 Then
                    GetRunScannerItem = "137"
                Else
                    GetRunScannerItem = "135"
                End If
            End If
        Case "DriverFiltersClass", "DriverFiltersDevice"
            If InStr(sKey, "Upper") > 0 Then
                GetRunScannerItem = "145"
            End If
        Case "SafeBootAltShell"
            GetRunScannerItem = "146"
        Case "SecurityProviders"
            GetRunScannerItem = "147"
        Case "WOW"
            If sKey = "WOW1" Then
                GetRunScannerItem = "148"
            ElseIf sKey = "WOW2" Then
                GetRunScannerItem = "149"
            End If
        Case "XPSecurityRestore"
            GetRunScannerItem = "150"
        Case "Policies"
            If InStr(sKey, "System") > 0 Then
                GetRunScannerItem = "161"
            ElseIf InStr(sKey, "User") > 0 Then
                GetRunScannerItem = "160"
            End If
        Case "MountPoints", "MountPoints2"
            GetRunScannerItem = "170"
        Case "IniFiles"
            If InStr(sKey, "IniFilessystem.ini3") > 0 Then
                GetRunScannerItem = "171"
            End If
        Case "ContextMenuHandlers"
            GetRunScannerItem = "173"
        Case "ShellCommandsbat", "ShellCommandscmd", "ShellCommandscom", "ShellCommandsexe", "ShellCommandshta", "ShellCommandsjs", "ShellCommandsjse", "ShellCommandspif", "ShellCommandsscr", "ShellCommandstxt", "ShellCommandsvbe", "ShellCommandsvbs", "ShellCommandswsf", "ShellCommandswsh"
            GetRunScannerItem = "180"
        
    End Select
End Function

Public Function GetLongFilename$(sFilename$)
    Dim sLongFilename$
    If InStr(sFilename, "~") = 0 Then
        GetLongFilename = sFilename
        Exit Function
    End If
    sLongFilename = String(512, 0)
    GetLongPathName sFilename, sLongFilename, Len(sLongFilename)
    GetLongFilename = TrimNull(sLongFilename)
End Function

Public Sub WinTrustVerifyChildNodes(sKey$)
    If bAbort Then Exit Sub
    If Not NodeExists(sKey) Then Exit Sub
    Dim nodFirst As Node, nodCurr As Node
    Set nodFirst = frmMain.tvwMain.Nodes(sKey).Child
    Set nodCurr = nodFirst
    Do
        If nodCurr.Children > 0 Then WinTrustVerifyChildNodes nodCurr.Key
        
        WinTrustVerifyNode nodCurr.Key
        
        If nodCurr = nodFirst.LastSibling Then Exit Do
        Set nodCurr = nodCurr.Next
        If bAbort Then Exit Sub
    Loop
End Sub

Public Sub WinTrustVerifyNode(sKey$)
    If bAbort Then Exit Sub
    If Not NodeIsValidFile(frmMain.tvwMain.Nodes(sKey)) Then Exit Sub
        
    Dim sFile$, sMD5$, sIcon$
    sFile = frmMain.tvwMain.Nodes(sKey).Text
    If Not FileExists(sFile) Then
        sFile = frmMain.tvwMain.Nodes(sKey).Tag
        If Not FileExists(sFile) Then Exit Sub
    End If
    Status "Verifying file signature of: " & sFile
    'sMD5 = GetFileMD5(sFile)
    
    Select Case VerifyFileSignature(sFile)
        Case 1: sIcon = "wintrust1"
        Case 0: sIcon = "wintrust3"
        Case -1: Exit Sub
    End Select
    
    frmMain.tvwMain.Nodes(sKey).Image = sIcon
    frmMain.tvwMain.Nodes(sKey).SelectedImage = sIcon
End Sub

Public Function NodeIsValidFile(objNode As Node) As Boolean
    NodeIsValidFile = False
    If objNode.Tag <> vbNullString Then
        If FileExists(objNode.Tag) And Not IsFolder(objNode.Tag) Then
            NodeIsValidFile = True
        End If
    End If
End Function

Public Function NodeIsValidRegkey(objNode As Node) As Boolean
    NodeIsValidRegkey = False
    If InStr(1, objNode.Tag, "HKEY_") <> 1 Then
        'selected item is not a regkey but a file - climb up in the
        'tree until we find a regkey
        Dim MyNode As Node
        Set MyNode = objNode
        With frmMain.tvwMain
            Do Until MyNode = .Nodes("System") Or _
                     MyNode = .Nodes("Users") Or _
                     MyNode = .Nodes("Hardware")
                Set MyNode = MyNode.Parent
                If InStr(1, MyNode.Tag, "HKEY_") = 1 Then
                    NodeIsValidRegkey = True
                    Exit Function
                End If
            Loop
        End With
    Else
        NodeIsValidRegkey = True
    End If
End Function

Public Function NodeExists(sKey$) As Boolean
    Dim s$
    On Error Resume Next
    s = frmMain.tvwMain.Nodes(sKey).Text
    If Err Then
    'If s <> vbNullString Then
        NodeExists = False
    Else
        NodeExists = True
    End If
    Err.Clear
End Function

Public Function IsFolder(sFile$) As Boolean
    If GetFileAttributes(sFile) And FILE_ATTRIBUTE_DIRECTORY Then
        IsFolder = True
    Else
        IsFolder = False
    End If
End Function

Public Function DownloadFile(sURL$, sTarget$) As Boolean
    Dim hInternet&, hFile&, sBuffer$, sFile$, lBytesRead&
    Dim sUserAgent$
    DownloadFile = False
    If FileExists(sTarget) Then Exit Function
    sUserAgent = "StartupList v" & App.Major & "." & Format(App.Minor, "00") & "." & App.Revision
    
    Status "Downloading Wintrust library..."
    hInternet = InternetOpen(sUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    If hInternet Then
        hFile = InternetOpenUrl(hInternet, sURL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
        If hFile Then
            Do
                sBuffer = Space(16384)
                InternetReadFile hFile, sBuffer, Len(sBuffer), lBytesRead
                sFile = sFile & Left(sBuffer, lBytesRead)
            Loop Until lBytesRead = 0
            InternetCloseHandle hFile
            
            Open sTarget For Output As #1
                Print #1, sFile
            Close #1
            DownloadFile = True
        Else
            MsgBox "Unable to connect to the Internet.", vbCritical
        End If
        InternetCloseHandle hInternet
    End If
    Status "Done."
End Function
