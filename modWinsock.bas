Attribute VB_Name = "modWinsock"
Option Explicit
Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function WSAEnumProtocols Lib "ws2_32.dll" Alias "WSAEnumProtocolsA" (ByVal lpiProtocols As Long, lpProtocolBuffer As Any, lpdwBufferLength As Long) As Long
Private Declare Function WSAEnumNameSpaceProviders Lib "ws2_32.dll" Alias "WSAEnumNameSpaceProvidersA" (lpdwBufferLength As Long, lpnspBuffer As Any) As Long
Private Declare Function WSCGetProviderPath Lib "ws2_32.dll" (ByVal lpProviderId As Long, ByRef lpszProviderDllPath As Byte, ByRef lpProviderDllPathLen As Long, ByRef lpErrno As Long) As Long

Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As String, lpString2 As Any) As String
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpsz As String, ByVal cchMax As Long) As Long

Private Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * 257
    szSystemStatus As String * 129
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type WSANAMESPACE_INFO
    NSProviderId   As GUID
    dwNameSpace    As Long
    fActive        As Long
    dwVersion      As Long
    lpszIdentifier As Long
End Type

Private Type WSAPROTOCOLCHAIN
    ChainLen As Long
    ChainEntries(6) As Long
End Type

Private Type WSAPROTOCOL_INFO
    dwServiceFlags1 As Long
    dwServiceFlags2 As Long
    dwServiceFlags3 As Long
    dwServiceFlags4 As Long
    dwProviderFlags As Long
    ProviderId As GUID
    dwCatalogEntryId As Long
    ProtocolChain As WSAPROTOCOLCHAIN
    iVersion As Long
    iAddressFamily As Long
    iMaxSockAddr As Long
    iMinSockAddr As Long
    iSocketType As Long
    iProtocol As Long
    iProtocolMaxOffset As Long
    iNetworkByteOrder As Long
    iSecurityScheme As Long
    dwMessageSize As Long
    dwProviderReserved As Long
    szProtocol As String * 256
End Type

Public Function EnumWinsockProtocol$()
    On Error Resume Next
    Dim i%, sEnumProt$
    Dim uWSAData As WSAData, sGUID$, sFile$
    Dim uWSAProtInfo As WSAPROTOCOL_INFO
    Dim uBuffer() As Byte, lBufferSize&
    Dim lNumProtocols&, sLSPName$, lDummy&
    
    If WSAStartup(&H202, uWSAData) > 0 Then Exit Function
    
    ReDim uBuffer(1)
    WSAEnumProtocols 0, uBuffer(0), lBufferSize
    ReDim uBuffer(lBufferSize - 1)
    
    lNumProtocols = WSAEnumProtocols(0, uBuffer(0), lBufferSize)
    If lNumProtocols <> -1 Then
        For i = 0 To lNumProtocols - 1
            CopyMemory uWSAProtInfo, uBuffer(i * Len(uWSAProtInfo)), Len(uWSAProtInfo)
            sGUID = GuidToString(uWSAProtInfo.ProviderId)
            sFile = GetProviderFile(uWSAProtInfo.ProviderId)
            sLSPName = TrimNull(uWSAProtInfo.szProtocol)
            If bShowCLSIDs Then
                sEnumProt = sEnumProt & "|" & sLSPName & " - " & sGUID & " - " & sFile
            Else
                sEnumProt = sEnumProt & "|" & sLSPName & " - " & sFile
            End If
        Next i
    End If
    
    Do
    Loop Until WSACleanup() = -1
    If sEnumProt <> vbNullString Then EnumWinsockProtocol = Mid$(sEnumProt, 2)
End Function

Public Function EnumWinsockNameSpace$()
    Dim lNumNameSpace&, sLSPName$, sEnumNamespace$
    Dim uWSANameSpaceInfo As WSANAMESPACE_INFO
    Dim uWSAData As WSAData, i%, sGUID$, sFile$
    Dim uBuffer() As Byte, lBufferSize&
    
    If WSAStartup(&H202, uWSAData) > 0 Then Exit Function

    ReDim uBuffer(1)
    lBufferSize = 0
    WSAEnumNameSpaceProviders lBufferSize, ByVal 0
    ReDim uBuffer(lBufferSize - 1)
    
    lNumNameSpace = WSAEnumNameSpaceProviders(lBufferSize, uBuffer(0))
    If lNumNameSpace <> -1 Then
        For i = 0 To lNumNameSpace - 1
            CopyMemory uWSANameSpaceInfo, uBuffer(i * Len(uWSANameSpaceInfo)), Len(uWSANameSpaceInfo)
            sGUID = GuidToString(uWSANameSpaceInfo.NSProviderId)
            sLSPName = String$(255, 0)
            lstrcpy sLSPName, ByVal uWSANameSpaceInfo.lpszIdentifier
            sLSPName = TrimNull(sLSPName)
            sFile = GetNSProviderFile(sLSPName)
            If bShowCLSIDs Then
                sEnumNamespace = sEnumNamespace & "|" & sLSPName & " - " & sGUID & " - " & sFile
            Else
                sEnumNamespace = sEnumNamespace & "|" & sLSPName & " - " & sFile
            End If
        Next i
    End If

    Do
    Loop Until WSACleanup() = -1
    If sEnumNamespace <> vbNullString Then EnumWinsockNameSpace = Mid$(sEnumNamespace, 2)
End Function

Private Function GuidToString$(uGuid As GUID)
    'internal function
    Dim sGUID$
    sGUID = String$(80, 0)
    If StringFromGUID2(uGuid, sGUID, Len(sGUID)) > 0 Then
        GuidToString = StrConv(sGUID, vbFromUnicode)
        GuidToString = TrimNull(GuidToString)
    End If
End Function

Private Function GetProviderFile$(uProviderID As GUID)
    Dim sFile$, uFile() As Byte, lFileLen&, lErr&
    'this function works for GUIDs returned from WSCEnumProtocols,
    'but not for those from WSAEnumNameSpaceProviders (??)
    lFileLen = 260
    'Debug.Print "uProviderID = " & uProviderID.Data1 & "," & _
    '            uProviderID.Data2 & "," & uProviderID.Data3 & "," & _
    '            StrConv(uProviderID.Data4, vbUnicode)
    ReDim uFile(lFileLen)
    If WSCGetProviderPath(VarPtr(uProviderID), uFile(0), lFileLen, lErr) = 0 Then
        sFile = StrConv(uFile, vbUnicode)
        'wtf? but it seems to work
        sFile = StrConv(sFile, vbFromUnicode)
        sFile = ExpandEnvironmentVars(TrimNull(sFile))
        GetProviderFile = sFile
    End If
End Function

Private Function GetNSProviderFile$(sName$)
    Dim sWS2Key$, sKeys$(), i&, sFile$
    sWS2Key = "System\CurrentControlSet\Services\Winsock2\Parameters\NameSpace_Catalog5\Catalog_Entries"
    sKeys = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sWS2Key), "|")
    For i = 0 To UBound(sKeys)
        If sName = RegGetString(HKEY_LOCAL_MACHINE, sWS2Key & "\" & sKeys(i), "DisplayString") Then
            sFile = ExpandEnvironmentVars(RegGetString(HKEY_LOCAL_MACHINE, sWS2Key & "\" & sKeys(i), "LibraryPath"))
            GetNSProviderFile = sFile
            Exit For
        End If
    Next i
End Function
