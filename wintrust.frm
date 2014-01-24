VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "winverifytrust"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "c:\windows\explorer.exe"
      Top             =   1800
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As GUID) As Long

Private Declare Function SHFileExists Lib "shell32" Alias "#45" (ByVal szPath As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

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
    pWTFINFO As Long
    pFake As Long
    pFake1 As Long
    pFake2 As Long
    pFake3 As Long
    dwStateAction As Long
    hWVTStateData As Long
    pwszURLReference As String
    dwProvFlags As Long
    dwUIContext As Long
    pFile As WINTRUST_FILE_INFO
End Type

Private Type WINTRUST_CATALOG_INFO
    cbStruct As Long
    dwCatalogVersion As Long
    pcwszCatalogFilePath As String
    pcwszMemberTag As String
    pcwszMemberFilePath As String
    hMemberFile As Long
End Type

Private Type CATALOG_INFO
    cbStruct As Long
    sCatalogFile As String * 260
End Type

'.dwUnionChoice
Private Const WTD_CHOICE_FILE = 1
Private Const WTD_CHOICE_CATALOG = 2

'.dwStateAction
Private Const WTD_STATEACTION_IGNORE = 0
Private Const WTD_STATEACTION_VERIFY = 1

'WINTRUST_DATE
Private Const WTD_UI_NONE = 2

'return value
Private Const WTD_REVOKE_NONE = 0

'TrustProvider
Private Const WTD_SAFER_FLAG = 256

'WinTrust action GUID
Private WINTRUST_ACTION_GENERIC_VERIFY_V2 As GUID

Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const INVALID_HANDLE_VALUE = -1

Private hWinTrustDLL&
Private CryptCATAdminAcquireContext&
Private CryptCATAdminReleaseContext&
Private CryptCATAdminCalcHashFromFileHandle&
Private CryptCATAdminEnumCatalogFromHash&
Private CryptCATCatalogInfoFromContext&
Private CryptCATAdminReleaseCatalogContext&
Private WinVerifyTrust&

Private HCatAdmin As Long

Private Sub Command1_Click()
    Dim aByteHash(255) As Byte
    Dim iByteCount As Integer
    
    'Dim hCatAdminContact HCatAdmin
    Dim WTrustData As WINTRUST_DATA
    Dim WTDCatalogInfo As WINTRUST_CATALOG_INFO
    Dim WTDFileInfo As WINTRUST_FILE_INFO
    Dim CatalogInfo As CATALOG_INFO
    
    Dim hFile As Long
    Dim hCatalogContext As Long
    Dim swFilename As String
    Dim swMemberTag As String
    Dim ilRet As Long
    Dim x As Integer
    
    Dim Result As Boolean
    swFilename = Text1.Text
    
    If Not FileExists(swFilename) Then Exit Sub
    
    If CallWindowProc(CryptCATAdminAcquireContext, Me.hWnd, HCatAdmin, vbNull, vbNull) = 0 Then
        'error
    End If
    
    hFile = CreateFile(swFilename, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        'error
    End If
    
    iByteCount = UBound(aByteHash)
    
    Call CallWindowProc(CryptCATAdminCalcHashFromFileHandle, Me.hWnd, hFile, CLng(iByteCount), aByteHash(0))
    
    For x = 0 To iByteCount - 1
        swMemberTag = swMemberTag & Hex(aByteHash(x))
    Next x
    
    CloseHandle hFile
    
End Sub

Private Sub Form_Load()
    hWinTrustDLL = LoadLibrary("wintrust.dll")
    If hWinTrustDLL > 32 Then
        CryptCATAdminAcquireContext = GetProcAddress(hWinTrustDLL, "CryptCATAdminAcquireContext")
        CryptCATAdminReleaseContext = GetProcAddress(hWinTrustDLL, "CryptCATAdminReleaseContext")
        CryptCATAdminCalcHashFromFileHandle = GetProcAddress(hWinTrustDLL, "CryptCATAdminCalcHashFromFileHandle")
        CryptCATAdminEnumCatalogFromHash = GetProcAddress(hWinTrustDLL, "CryptCATAdminEnumCatalogFromHash")
        CryptCATCatalogInfoFromContext = GetProcAddress(hWinTrustDLL, "CryptCATCatalogInfoFromContext")
        CryptCATAdminReleaseCatalogContext = GetProcAddress(hWinTrustDLL, "CryptCATAdminReleaseCatalogContext")
        WinVerifyTrust = GetProcAddress(hWinTrustDLL, "WinVerifyTrust")
        
    End If
    If CLSIDFromString(StrPtr("{00AAC56B-CD44-11d0-8CC2-00C04FC295EE}"), WINTRUST_ACTION_GENERIC_VERIFY_V2) <> 0 Then
        'failed
    End If
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
     If hWinTrustDLL > 32 Then FreeLibrary hWinTrustDLL
End Sub

Public Function FileExists(sFile$) As Boolean
    FileExists = IIf(SHFileExists(StrConv(sFile, vbUnicode)) = 1, True, False)
End Function

