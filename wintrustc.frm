VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function WinVerifyTrust Lib "wintrust.dll" (ByVal hwnd As Long, ByRef pgActionID As GUID, ByRef pWVTData As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As GUID) As Long
         
Private Declare Function CryptQueryObject Lib "Crypt32.dll" (ByVal dwObjectType As Long, ByVal pvObject As Long, ByVal dwExpectedContentTypeFlags As Long, ByVal dwExpectedFormatTypeFlags As Long, ByVal dwFlags As Long, ByRef pdwMsgAndCertEncodingType As Long, ByRef pdwContentType As Long, ByRef pdwFormatType As Long, ByRef phCertStore As Long, ByRef phMsg As Long, ByRef ppvContext As Long) As Long

Private Declare Function GetLastError Lib "kernel32" () As Long

Private Const CERT_QUERY_OBJECT_FILE As Long = &H1
Private Const CERT_QUERY_CONTENT_PKCS7_SIGNED_EMBED As Long = 10
Private Const CERT_QUERY_CONTENT_FLAG_PKCS7_SIGNED_EMBED As Long = 2 ^ CERT_QUERY_CONTENT_PKCS7_SIGNED_EMBED
Private Const CERT_QUERY_FORMAT_BINARY As Long = &H1
Private Const CERT_QUERY_FORMAT_FLAG_BINARY As Long = 2 ^ CERT_QUERY_FORMAT_BINARY

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

Private Const WTD_UI_NONE = 2
Private Const WTD_REVOKE_NONE = 0
Private Const WTD_CHOICE_FILE = 1
Private Const WTD_CHOICE_CATALOG = 2
Private Const WTD_SAFER_FLAG = &H100
Private Const WTD_STATEACTION_VERIFY = 1
Private Const WTD_STATEACTION_IGNORE = 0
Private Const WTD_UICONTEXT_EXECUTE = 0

'error: -2146885629

    Private Const TRUST_E_PROVIDER_UNKNOWN = &H800B0001
      Private Const TRUST_E_ACTION_UNKNOWN = &H800B0002
Private Const TRUST_E_SUBJECT_FORM_UNKNOWN = &H800B0003
 Private Const TRUST_E_SUBJECT_NOT_TRUSTED = &H800B0004
         Private Const TRUST_E_NOSIGNATURE = &H800B0100
   Private Const TRUST_E_EXPLICIT_DISTRUST = &H800B0111
   Private Const CRYPT_E_SECURITY_SETTINGS = &H80092026
 
Private Sub Command1_Click()
    Call VerifyFileSignature(Text1.Text)
    'MsgBox CryptDing(Text1.Text)
End Sub

Private Sub Form_Load()
    'MsgBox VerifyFileSignature("c:\windows\0.log")
    Text1.Text = "c:\windows\notepad.exe"
End Sub

Public Function VerifyFileSignature(sFile$) As Boolean
    Dim uVerifyV2 As GUID, uWTfileinfo As WINTRUST_FILE_INFO
    Dim uWTdata As WINTRUST_DATA, lRet&
    With uWTfileinfo
        .cbStruct = Len(uWTfileinfo)
        .pcwszFilePath = sFile
    End With
    With uWTdata
        .cbStruct = Len(uWTdata)
        .dwUIChoice = WTD_UI_NONE
        .fdwRevocationChecks = WTD_REVOKE_NONE
        .dwUnionChoice = WTD_CHOICE_FILE
        .dwStateAction = 0 'WTD_STATEACTION_IGNORE
        .dwProvFlags = WTD_SAFER_FLAG '0
        .pFile = VarPtr(uWTfileinfo)
    End With
    If CLSIDFromString(StrPtr("{00AAC56B-CD44-11d0-8CC2-00C04FC295EE}"), uVerifyV2) = 0 Then
        lRet = WinVerifyTrust(0, uVerifyV2, uWTdata)
    End If
    If lRet = 0 Then
        MsgBox "The file is signed and the signature was verified.", vbInformation
        VerifyFileSignature = True
    Else
        Select Case lRet
            Case TRUST_E_NOSIGNATURE
                Select Case GetLastError
                    Case TRUST_E_NOSIGNATURE, TRUST_E_SUBJECT_FORM_UNKNOWN, TRUST_E_PROVIDER_UNKNOWN
                        MsgBox "The file is not signed.", vbExclamation
                    Case Else
                        MsgBox "An unknown error occurred trying to verify the signature.", vbExclamation
                End Select
            Case TRUST_E_SUBJECT_NOT_TRUSTED: MsgBox "The file is signed, but not trusted.", vbExclamation
            Case TRUST_E_EXPLICIT_DISTRUST: MsgBox "The file is signed, but specifically disallowed.", vbExclamation
            Case CRYPT_E_SECURITY_SETTINGS: MsgBox "User trust disabled, and the admin has not allowed the subject or publisher.", vbInformation
            Case Else: MsgBox "Error: " & lRet & vbCrLf & "GetLastError: " & GetLastError, vbInformation
        End Select
    End If
End Function

Private Function CryptDing(sFile$)
    Dim fResult As Long
    Dim szFileName As String
    Dim dwEncoding As Long
    Dim dwContentType As Long
    Dim dwFormatType As Long
    Dim hStore As Long
    Dim hMsg As Long
    
    szFileName = sFile
    
    fResult = CryptQueryObject(CERT_QUERY_OBJECT_FILE, ByVal StrPtr(szFileName), CERT_QUERY_CONTENT_FLAG_PKCS7_SIGNED_EMBED, CERT_QUERY_FORMAT_FLAG_BINARY, 0&, dwEncoding, dwContentType, dwFormatType, hStore, hMsg, ByVal 0&)
    MsgBox GetLastError
End Function
