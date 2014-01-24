Attribute VB_Name = "modTriage"
Option Explicit
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal InternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Long
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, sOptional As Any, ByVal lOptionalLength As Long) As Integer
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal pCryptHash As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pcbData As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long

Private Const ALG_TYPE_ANY As Long = 0
Private Const ALG_SID_MD5 As Long = 3
Private Const ALG_CLASS_HASH As Long = 32768

Private Const HP_HASHVAL As Long = 2
Private Const HP_HASHSIZE As Long = 4

Private Const CRYPT_VERIFYCONTEXT = &HF0000000

Private Const PROV_RSA_FULL As Long = 1
Private Const MS_ENHANCED_PROV As String = "Microsoft Enhanced Cryptographic Provider v1.0"

Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
'Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_SERVICE_HTTP = 3
Private Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000

Private sTriageObj$()

Public Sub AddTriageObj(sName$, sType$, sFile$, Optional sCLSID$, Optional sCodebase$)
    Dim sPath$, sFilename$, sFilesize$, sMD5$, sItem$()
    If Not FileExists(sFile) Then Exit Sub
    If InStr(sFile, "\") = 0 Then Exit Sub
    'sPath = Left$(sFile, InStrRev(sFile, "\") - 1)
    sFilename = Mid$(sFile, InStrRev(sFile, "\") + 1)
    sFilesize = CStr(FileLen(sFile))
    sMD5 = GetFileMD5(sFile)
    
    ReDim sItem(8)
    sItem(0) = sName     'id to item
    sItem(1) = sFilename 'name
    sItem(2) = sCLSID
    sItem(3) = sFile     'complete path+filename
    sItem(4) = sFilename 'filename
    sItem(5) = sFilesize
    sItem(6) = sMD5
    sItem(7) = sType
    sItem(8) = sCodebase 'Codebase, for DPF
    
    On Error Resume Next
    If UBound(sTriageObj) = -2 Then ReDim sTriageObj(0)
    If Err Then ReDim sTriageObj(0)
    On Error GoTo 0:
    
    ReDim Preserve sTriageObj(UBound(sTriageObj) + 1)
    sTriageObj(UBound(sTriageObj)) = "ITEM[]=" & Join(sItem, "|")
End Sub

Public Function GetTriage$()
    Dim hInternet&, hConnect&, sURL$, sUserAgent$, sPost$
    Dim hRequest&, sResponse$, sBuffer$, lBufferLen&, sHeaders$
    sURL = "http://www.spywareguide.com/report/triage.php"
    sUserAgent = "StartupList v" & App.Major & "." & Format$(App.Minor, "00")
    sPost = Mid$(URLEncode(Join(sTriageObj, "&")), 2)
    If sPost = vbNullString Then Exit Function
    sHeaders = "Accept: text/html,text/plain" & vbCrLf & _
               "Accept-Charset: ISO-8859-1,utf-8" & vbCrLf & _
               "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
               "Content-Length: " & Len(sPost)
    
    hInternet = InternetOpen(sUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    If hInternet = 0 Then Exit Function

    hConnect = InternetConnect(hInternet, "www.spywareguide.com", 80, vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
    If hConnect > 0 Then
        hRequest = HttpOpenRequest(hConnect, "POST", "/report/triage.php", "HTTP/1.1", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
        If hRequest > 0 Then
            HttpSendRequest hRequest, sHeaders, Len(sHeaders), ByVal sPost, Len(sPost)
            sResponse = vbNullString
            Do
                sBuffer = Space$(1024)
                InternetReadFile hRequest, sBuffer, Len(sBuffer), lBufferLen
                sBuffer = Left$(sBuffer, lBufferLen)
                sResponse = sResponse & sBuffer
            Loop Until lBufferLen = 0
            GetTriage = sResponse
            InternetCloseHandle hRequest
        End If
        InternetCloseHandle hConnect
    End If
    InternetCloseHandle hInternet
End Function

Public Function GetFileMD5$(sFile$)
    'note: this needs at least Win95 /w OSR2 (IE3)
    If Not FileExists(sFile) Then Exit Function
    Dim lFileLen&, sFileContents$
    On Error Resume Next
    lFileLen = FileLen(sFile)
    If lFileLen = 0 Then
        GetFileMD5 = "D41D8CD98F00B204E9800998ECF8427E"
        Exit Function
    End If
    sFileContents = InputFile(sFile)
    If Err Then Exit Function
    
    Dim hCrypt&, hHash&, uMD5(255) As Byte, lMD5Len&, i%, sMD5$
    If CryptAcquireContext(hCrypt, vbNullString, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) > 0 Then
        If CryptCreateHash(hCrypt, ALG_TYPE_ANY Or ALG_CLASS_HASH Or ALG_SID_MD5, 0, 0, hHash) > 0 Then
            If CryptHashData(hHash, sFileContents, Len(sFileContents), 0) > 0 Then
                If CryptGetHashParam(hHash, HP_HASHSIZE, uMD5(0), UBound(uMD5) + 1, 0) > 0 Then
                    lMD5Len = uMD5(0)
                    If CryptGetHashParam(hHash, HP_HASHVAL, uMD5(0), UBound(uMD5) + 1, 0) > 0 Then
                        For i = 0 To lMD5Len - 1
                            sMD5 = sMD5 & Right$("0" & Hex$(uMD5(i)), 2)
                        Next i
                    End If
                End If
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hCrypt, 0
    Else
        Exit Function
    End If
    If sMD5 <> vbNullString Then GetFileMD5 = sMD5
End Function

Private Function URLEncode$(sURL$)
    Dim sDummy$, sReplace$(), i&
    sDummy = sURL
    ReDim sReplace(7)
    sReplace(0) = "|"
    sReplace(1) = "\"
    sReplace(2) = "/"
    sReplace(3) = "["
    sReplace(4) = "]"
    sReplace(5) = ":"
    sReplace(6) = "("
    sReplace(7) = ")"
    For i = 0 To UBound(sReplace)
        sDummy = Replace(sDummy, sReplace(i), "%" & UCase$(Hex$(Asc(sReplace(i)))))
    Next i
    
    URLEncode = sDummy
End Function
