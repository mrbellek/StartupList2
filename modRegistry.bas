Attribute VB_Name = "modRegistry"
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
'Public Const HKEY_PERFORMANCE_DATA = &H80000004
'Public Const HKEY_CURRENT_CONFIG = &H80000005
'Public Const HKEY_DYN_DATA = &H80000006

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Const REG_NONE = 0
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7

Public Function RegGetString$(lHive&, sKey$, sVal$, Optional bTrimNull As Boolean = True)
    Dim hKey&, uData() As Byte, lDataLen&, sData$
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        RegQueryValueEx hKey, sVal, 0, 0, ByVal 0, lDataLen
        ReDim uData(lDataLen)
        If RegQueryValueEx(hKey, sVal, 0, 0, uData(0), lDataLen) = 0 Then
            If bTrimNull Then
                sData = StrConv(uData, vbUnicode)
                sData = TrimNull(sData)
            Else
                If lDataLen > 2 Then
                    ReDim Preserve uData(lDataLen - 2)
                    sData = StrConv(uData, vbUnicode)
                End If
            End If
            RegGetString = sData
        End If
        RegCloseKey hKey
    End If
End Function

Public Function RegGetDword&(lHive$, sKey$, sVal$)
    Dim hKey&, lData&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        If RegQueryValueEx(hKey, sVal, 0, 0, lData, 4) = 0 Then
            RegGetDword = lData
        End If
        RegCloseKey hKey
    End If
End Function

Public Function RegKeyExists(lHive&, sKey$) As Boolean
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        RegKeyExists = True
        RegCloseKey hKey
    End If
End Function

Public Function RegValExists(lHive&, sKey$, sVal$) As Boolean
    Dim hKey&, lDataLen&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        If RegQueryValueEx(hKey, sVal, 0, 0, ByVal 0, lDataLen) = 0 Then
            RegValExists = True
        End If
        RegCloseKey hKey
    End If
End Function

Public Function RegEnumSubKeys$(lHive&, sKey$)
    Dim hKey&, i&, sName$, sList$
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        sName = String$(260, 0)
        Do Until RegEnumKeyEx(hKey, i, sName, Len(sName), 0, vbNullString, 0, ByVal 0) <> 0
            sName = TrimNull(sName)
            sList = sList & "|" & sName
            i = i + 1
            sName = String$(260, 0)
            If bAbort Then
                RegCloseKey hKey
                Exit Function
            End If
        Loop
        RegCloseKey hKey
    End If
    If sList <> vbNullString Then RegEnumSubKeys = Mid$(sList, 2)
End Function

Public Function RegEnumValues$(lHive&, sKey$, Optional bNullSep As Boolean = False, Optional bIgnoreBinaries As Boolean = True, Optional bIgnoreDwords As Boolean = True)
    Dim hKey&, i&, sName$, uData() As Byte, lDataLen&
    Dim lType&, sData$, sList$
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        sName = String$(lEnumBufLen, 0)
        ReDim uData(32768)
        lDataLen = UBound(uData)
        Do Until RegEnumValue(hKey, i, sName, Len(sName), 0, lType, uData(0), lDataLen) <> 0
            
            sName = TrimNull(sName)
            If sName = vbNullString Then sName = "@"
            
            If lType = REG_SZ Then
                ReDim Preserve uData(lDataLen)
                sData = TrimNull(StrConv(uData, vbUnicode))
                If bNullSep Then
                    sList = sList & Chr$(0) & sName & " = " & sData
                Else
                    sList = sList & "|" & sName & " = " & sData
                End If
            End If
            
            If lType = REG_BINARY And Not bIgnoreBinaries Then
                sList = sList & "|" & sName & " (binary)"
            End If
            
            If lType = REG_DWORD And Not bIgnoreDwords Then
                'look at me! I'm haxxoring word values from binary!
                'sData = "dword: " & Hex$(uData(0)) & "." & Hex$(uData(1)) & "." & Hex$(uData(2)) & "." & Hex$(uData(3))
                'sData = "dword: " & Val("&H" & Hex$(uData(3)) & Hex$(uData(2)) & Hex$(uData(1)) & Hex$(uData(0)))
                sData = "dword: " & CStr(16 ^ 6 * uData(3) + 16 ^ 4 * uData(2) + 16 ^ 2 * uData(1) + uData(0))
                sList = sList & "|" & sName & " = " & sData
            End If
            sName = String$(lEnumBufLen, 0)
            ReDim uData(32768)
            lDataLen = UBound(uData)
            i = i + 1
            
            If bAbort Then
                RegCloseKey hKey
                Exit Function
            End If
        Loop
        RegCloseKey hKey
    End If
    If sList <> vbNullString Then RegEnumValues = Mid$(sList, 2)
End Function

Public Function RegEnumDwordValues$(lHive&, sKey$)
    Dim hKey&, i&, sName$, uData() As Byte, lDataLen&
    Dim lType&, lData&, sList$
    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
        sName = String$(lEnumBufLen, 0)
        ReDim uData(32768)
        lDataLen = UBound(uData)
        Do Until RegEnumValue(hKey, i, sName, Len(sName), 0, lType, uData(0), lDataLen) <> 0
            If lType = REG_DWORD And lDataLen = 4 Then
                sName = TrimNull(sName)
                If sName = vbNullString Then sName = "@"
                ReDim Preserve uData(4)
                CopyMemory lData, uData(0), 4
                sList = sList & "|" & sName & " = " & CStr(lData)
            End If
            sName = String$(lEnumBufLen, 0)
            ReDim uData(32768)
            lDataLen = UBound(uData)
            i = i + 1
        
            If bAbort Then
                RegCloseKey hKey
                Exit Function
            End If
        Loop
        RegCloseKey hKey
    End If
    If sList <> vbNullString Then RegEnumDwordValues = Mid$(sList, 2)
End Function

'Public Function RegGetNumOfSubkeys&(lHive&, sKey$)
'    Dim hKey&, lSubkeys&
'    If RegOpenKeyEx(lHive, sKey, 0, KEY_READ, hKey) = 0 Then
'        RegQueryInfoKey hKey, vbNullString, 0, 0, lSubkeys, 0, 0, 0, 0, 0, 0, ByVal 0
'        RegGetNumOfSubkeys = lSubkeys
'        RegCloseKey hKey
'    End If
'End Function

Public Sub RegEnumIEBands(tvwMain As TreeView)
    If bAbort Then Exit Sub
    Status "Loading... Internet Explorer Bands"
    'HKCR\CLSID\*\Implemented Categories\{00021493-0000-0000-C000-000000000046}
    'HKCR\CLSID\*\Implemented Categories\{00021494-0000-0000-C000-000000000046}
    tvwMain.Nodes.Add "System", tvwChild, "IEBands", SEC_IEBANDS, "msie"
    tvwMain.Nodes("IEBands").Tag = "HKEY_CLASSES_ROOT\CLSID"
    
    Dim hKey&, i&, lNumItems&, sCLSID$, sName$, sFile$
    If RegOpenKeyEx(HKEY_CLASSES_ROOT, "CLSID", 0, KEY_READ, hKey) = 0 Then
        RegQueryInfoKey hKey, vbNullString, 0, 0, lNumItems, 0, 0, 0, 0, 0, 0, ByVal 0
        
        sCLSID = String$(260, 0)
        Do Until RegEnumKeyEx(hKey, i, sCLSID, Len(sCLSID), 0, vbNullString, 0, ByVal 0) <> 0
            sCLSID = TrimNull(sCLSID)
    
            If RegKeyExists(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\Implemented Categories\{00021493-0000-0000-C000-000000000046}") Or _
               RegKeyExists(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\Implemented Categories\{00021494-0000-0000-C000-000000000046}") Then
                sName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString)
                sFile = ExpandEnvironmentVars(RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
                sFile = GetLongFilename(sFile)
                If bShowCLSIDs Then
                    tvwMain.Nodes.Add "IEBands", tvwChild, "IEBands" & i, sName & " - " & sCLSID & " - " & sFile, "dll"
                Else
                    tvwMain.Nodes.Add "IEBands", tvwChild, "IEBands" & i, sName & " - " & sFile, "dll"
                End If
                tvwMain.Nodes("IEBands" & i).Tag = GuessFullpathFromAutorun(sFile)
            End If
    
            sCLSID = String$(260, 0)
            i = i + 1
            If i Mod 100 = 0 And lNumItems > 0 Then
                Status "Loading... Internet Explorer Bands (" & CInt(i * 100 / lNumItems) & "%, " & i & " CLSIDs)"
            End If
        
            If bAbort Then
                RegCloseKey hKey
                Exit Sub
            End If
        Loop
        RegCloseKey hKey
    End If
    
    If tvwMain.Nodes("IEBands").Children > 0 Then
        tvwMain.Nodes("IEBands").Text = tvwMain.Nodes("IEBands").Text & " (" & tvwMain.Nodes("IEBands").Children & ")"
    Else
        If Not bShowEmpty Then
            tvwMain.Nodes.Remove "IEBands"
        End If
    End If
End Sub

Public Sub RegEnumKillBits(tvwMain As TreeView)
    If bAbort Then Exit Sub
    Status "Loading... ActiveX killbits"
    'HKLM\Software\Microsoft\Internet Explorer\ActiveXCompatibility
    'note: this sub will not show all set Killbits - only those that
    'are actually blocking a CLSID+File that exists on the system.
    Dim sKey$, lNumItems&
    sKey = "Software\Microsoft\Internet Explorer\ActiveX Compatibility"
    tvwMain.Nodes.Add "DisabledEnums", tvwChild, "Killbits", SEC_KILLBITS, "msie"
    tvwMain.Nodes("Killbits").Tag = "HKEY_LOCAL_MACHINE\" & sKey
    tvwMain.Nodes("Killbits").Sorted = True
    Dim hKey&, sCLSID$, i&, sName$, sFile$, lKill&
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKey, 0, KEY_READ, hKey) = 0 Then
        'RegQueryInfoKey hKey, vbNullString, 0, 0, lNumItems, 0, 0, 0, 0, 0, 0, ByVal 0
        
        sCLSID = String$(260, 0)
        Do Until RegEnumKeyEx(hKey, i, sCLSID, Len(sCLSID), 0, vbNullString, 0, ByVal 0) <> 0
            sCLSID = TrimNull(sCLSID)
        
            lKill = RegGetDword(HKEY_LOCAL_MACHINE, sKey & "\" & sCLSID, "Compatibility Flags")
            If lKill = 1024 Then
                sName = RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, vbNullString)
                sFile = ExpandEnvironmentVars(RegGetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID & "\InprocServer32", vbNullString))
                sFile = GetLongFilename(sFile)
                If sFile <> vbNullString Then
                    If sName = vbNullString Then sName = "(no name)"
                    If Not bShowCLSIDs Then
                        tvwMain.Nodes.Add "Killbits", tvwChild, "Killbits" & i, sName & " - " & sFile, "dll"
                    Else
                        tvwMain.Nodes.Add "Killbits", tvwChild, "Killbits" & i, sName & " - " & sCLSID & " - " & sFile, "dll"
                    End If
                    tvwMain.Nodes("Killbits" & i).Tag = GuessFullpathFromAutorun(sFile)
                End If
            End If
            
            sCLSID = String$(260, 0)
            i = i + 1
            'If i Mod 100 = 0 And lNumItems<> 0Then
            '    Status "Loading... ActiveX killbits (" & CInt(CLng(i) * 100 / lNumItems) & "%, " & i & " CLSIDs)"
            'End If
            If bAbort Then
                RegCloseKey hKey
                Exit Sub
            End If
        Loop
        RegCloseKey hKey
        
        tvwMain.Nodes("Killbits").Text = tvwMain.Nodes("Killbits").Text & " (" & tvwMain.Nodes("Killbits").Children & ")"
        If tvwMain.Nodes("Killbits").Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove "Killbits"
        End If
    End If

    '----------------------------------------------------------------
    'nothing - this is system-wide
End Sub

Public Sub RegEnumZones(tvwMain As TreeView)
    Dim sKey$, sZoneNames$(), i&, lNumItems&
    Dim hKey&, sDomain$, lZone&, sIcon$, sSubkeys$(), j&, sRange$
    If bAbort Then Exit Sub
    Status "Loading... Trusted sites & Restricted sites"
    tvwMain.Nodes.Add "DisabledEnums", tvwChild, "Zones", SEC_ZONES, "internet"
    sKey = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"
    
    Status "Loading... Trusted sites & Restricted sites (this user)"
    sZoneNames = Split(RegEnumSubKeys(HKEY_CURRENT_USER, sKey & "\Zones"), "|")
    For i = 0 To UBound(sZoneNames)
        sZoneNames(i) = RegGetString(HKEY_CURRENT_USER, sKey & "\Zones\" & sZoneNames(i), "DisplayName")
    Next i
    tvwMain.Nodes.Add "Zones", tvwChild, "ZonesUser", "This user", "user"
    'add root keys for zones
    For i = 0 To UBound(sZoneNames)
        tvwMain.Nodes.Add "ZonesUser", tvwChild, "ZonesUser" & i, sZoneNames(i), "internet"
        tvwMain.Nodes("ZonesUser" & i).Tag = "HKEY_CURRENT_USER\" & sKey & "\ZoneMap\Domains"
    Next i
    If RegOpenKeyEx(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains", 0, KEY_READ, hKey) = 0 Then
        RegQueryInfoKey hKey, vbNullString, 0, 0, lNumItems, 0, 0, 0, 0, 0, 0, ByVal 0
        If lNumItems > 1000 And Not bShowLargeZones Then
            frmMain.ShowError "Skipping Zones for this user, since there are over 1000 domains in them. (" & lNumItems & " to be exact)"
            RegCloseKey hKey
            GoTo CheckHKCURanges:
        End If
        sDomain = String$(260, 0)
        i = 0
        'loop through subkeys and add them to proper zone
        Do Until RegEnumKeyEx(hKey, i, sDomain, Len(sDomain), 0, vbNullString, 0, ByVal 0) <> 0
            If RegValExists(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain, "http") Then
                lZone = RegGetDword(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain, "http")
            Else
                If RegValExists(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain, "*") Then
                    lZone = RegGetDword(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain, "*")
                End If
            End If
            
            If RegValExists(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain, "http") Or _
               RegValExists(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain, "*") Then
                Select Case lZone
                    Case 0, 1: sIcon = "system"
                    Case 2: sIcon = "good"
                    Case 3: sIcon = "internet"
                    Case 4: sIcon = "bad"
                    Case Else: sIcon = "internet"
                End Select
                tvwMain.Nodes.Add "ZonesUser" & CStr(lZone), tvwChild, "ZonesUser" & CStr(lZone) & i, sDomain, sIcon
            End If
            'check for subdomains
            sSubkeys = Split(RegEnumSubKeys(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain), "|")
            If UBound(sSubkeys) > -1 Then
                For j = 0 To UBound(sSubkeys)
                    If RegValExists(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http") Then
                        lZone = RegGetDword(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http")
                    Else
                        If RegValExists(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*") Then
                            lZone = RegGetDword(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*")
                        End If
                    End If
                    
                    If RegValExists(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http") Or _
                       RegValExists(HKEY_CURRENT_USER, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*") Then
                        Select Case lZone
                            Case 0, 1: sIcon = "system"
                            Case 2: sIcon = "good"
                            Case 3: sIcon = "internet"
                            Case 4: sIcon = "bad"
                            Case Else: sIcon = "internet"
                        End Select
                        tvwMain.Nodes.Add "ZonesUser" & CStr(lZone), tvwChild, "ZonesUser" & CStr(lZone) & i & "s" & j, sSubkeys(j) & "." & sDomain, sIcon
                    End If
                Next j
            End If
            sDomain = String$(260, 0)
            i = i + 1
            If bShowLargeZones And i Mod 100 = 0 And lNumItems > 0 Then
                Status "Loading... Trusted sites & Restricted sites (this user, " & CInt(CLng(i) * 100 / lNumItems) & "%, " & i & " domains)"
            End If
            If bAbort Then
                RegCloseKey hKey
                Exit Sub
            End If
        Loop
        RegCloseKey hKey
    End If
    
CheckHKCURanges:
    If RegOpenKeyEx(HKEY_CURRENT_USER, sKey & "\ZoneMap\Ranges", 0, KEY_READ, hKey) = 0 Then
        'RegQueryInfoKey hKey, vbNullString, 0, 0, lNumItems, 0, 0, 0, 0, 0, 0, ByVal 0
        sDomain = String$(260, 0)
        i = 0
        Do Until RegEnumKeyEx(hKey, i, sDomain, Len(sDomain), 0, vbNullString, 0, ByVal 0) <> 0
            sDomain = TrimNull(sDomain)
            sRange = RegGetString(HKEY_CURRENT_USER, sKey & "\ZoneMap\Ranges\" & sDomain, ":Range")
            lZone = RegGetDword(HKEY_CURRENT_USER, sKey & "\ZoneMap\Ranges\" & sDomain, "*")
            
            If Trim$(sRange) <> vbNullString Then
                Select Case lZone
                    Case 0, 1: sIcon = "system"
                    Case 2: sIcon = "good"
                    Case 3: sIcon = "internet"
                    Case 4: sIcon = "bad"
                    Case Else: sIcon = "internet"
                End Select
                tvwMain.Nodes.Add "ZonesUser" & CStr(lZone), tvwChild, "ZonesUser" & CStr(lZone) & "Range" & i, sRange, sIcon
            End If
            
            sDomain = String$(260, 0)
            i = i + 1
            If bShowLargeZones And i Mod 100 = 0 And lNumItems > 0 Then
                Status "Loading... Trusted sites & Restricted sites (this user, " & CInt(CLng(i) * 100 / lNumItems) & "%, " & i & " IP)"
            End If
            If bAbort Then
                RegCloseKey hKey
                Exit Sub
            End If
        Loop
        RegCloseKey hKey
    End If
    
    For i = 0 To UBound(sZoneNames)
        If tvwMain.Nodes("ZonesUser" & i).Children > 0 Then
            tvwMain.Nodes("ZonesUser" & i).Text = tvwMain.Nodes("ZonesUser" & i).Text & " (" & tvwMain.Nodes("ZonesUser" & i).Children & ")"
            tvwMain.Nodes("ZonesUser" & i).Sorted = True
        Else
            If Not bShowEmpty Then
                tvwMain.Nodes.Remove "ZonesUser" & i
            End If
        End If
    Next i
    If tvwMain.Nodes("ZonesUser").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "ZonesUser"
    End If
    
    '---------------------------------
    
    Status "Loading... Trusted sites & Restricted sites (all users)"
    sZoneNames = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sKey & "\Zones"), "|")
    For i = 0 To UBound(sZoneNames)
        sZoneNames(i) = RegGetString(HKEY_LOCAL_MACHINE, sKey & "\Zones\" & sZoneNames(i), "DisplayName")
    Next i
    tvwMain.Nodes.Add "Zones", tvwChild, "ZonesSystem", "All users", "users"
    For i = 0 To UBound(sZoneNames)
        tvwMain.Nodes.Add "ZonesSystem", tvwChild, "ZonesSystem" & i, sZoneNames(i), "internet"
        tvwMain.Nodes("ZonesSystem" & i).Tag = "HKEY_LOCAL_MACHINE\" & sKey & "\ZoneMap\Domains"
    Next i
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains", 0, KEY_READ, hKey) = 0 Then
        RegQueryInfoKey hKey, vbNullString, 0, 0, lNumItems, 0, 0, 0, 0, 0, 0, ByVal 0
        If lNumItems > 1000 And Not bShowLargeZones Then
            frmMain.ShowError "Skipping Zones for all users, since there are over 1000 domains in them. (" & lNumItems & " to be exact)"
            RegCloseKey hKey
            GoTo CheckHKLMRanges:
        End If
        
        sDomain = String$(260, 0)
        i = 0
        Do Until RegEnumKeyEx(hKey, i, sDomain, Len(sDomain), 0, vbNullString, 0, ByVal 0) <> 0
            If RegValExists(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain, "http") Then
                lZone = RegGetDword(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain, "http")
            Else
                If RegValExists(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain, "*") Then
                    lZone = RegGetDword(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain, "*")
                End If
            End If
            
            If RegValExists(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain, "http") Or _
               RegValExists(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain, "*") Then
                Select Case lZone
                    Case 0, 1: sIcon = "system"
                    Case 2: sIcon = "good"
                    Case 3: sIcon = "internet"
                    Case 4: sIcon = "bad"
                    Case Else: sIcon = "internet"
                End Select
                tvwMain.Nodes.Add "ZonesSystem" & CStr(lZone), tvwChild, "ZonesSystem" & CStr(lZone) & i, sDomain, sIcon
            End If
            'check for subdomains
            sSubkeys = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain), "|")
            If UBound(sSubkeys) > -1 Then
                For j = 0 To UBound(sSubkeys)
                    If RegValExists(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http") Then
                        lZone = RegGetDword(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http")
                    Else
                        If RegValExists(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*") Then
                            lZone = RegGetDword(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*")
                        End If
                    End If
                    
                    If RegValExists(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http") Or _
                       RegValExists(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*") Then
                        Select Case lZone
                            Case 0, 1: sIcon = "system"
                            Case 2: sIcon = "good"
                            Case 3: sIcon = "internet"
                            Case 4: sIcon = "bad"
                            Case Else: sIcon = "internet"
                        End Select
                        tvwMain.Nodes.Add "ZonesSystem" & CStr(lZone), tvwChild, "ZonesUser" & CStr(lZone) & i & "s" & j, sSubkeys(j) & "." & sDomain, sIcon
                    End If
                Next j
            End If
            
            sDomain = String$(260, 0)
            i = i + 1
            If bShowLargeZones And i Mod 100 = 0 And lNumItems > 0 Then
                Status "Loading... Trusted sites & Restricted sites (all users, " & CInt(CLng(i) * 100 / lNumItems) & "%, " & i & " domains)"
            End If
            If bAbort Then
                RegCloseKey hKey
                Exit Sub
            End If
        Loop
        RegCloseKey hKey
    End If
    
CheckHKLMRanges:
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Ranges", 0, KEY_READ, hKey) = 0 Then
        'RegQueryInfoKey hKey, vbNullString, 0, 0, lNumItems, 0, 0, 0, 0, 0, 0, ByVal 0
        sDomain = String$(260, 0)
        i = 0
        Do Until RegEnumKeyEx(hKey, i, sDomain, Len(sDomain), 0, vbNullString, 0, ByVal 0) <> 0
            sDomain = TrimNull(sDomain)
            sRange = RegGetString(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Ranges\" & sDomain, ":Range")
            lZone = RegGetDword(HKEY_LOCAL_MACHINE, sKey & "\ZoneMap\Ranges\" & sDomain, "*")
            
            If Trim$(sRange) <> vbNullString Then
                Select Case lZone
                    Case 0, 1: sIcon = "system"
                    Case 2: sIcon = "good"
                    Case 3: sIcon = "internet"
                    Case 4: sIcon = "bad"
                    Case Else: sIcon = "internet"
                End Select
                tvwMain.Nodes.Add "ZonesSystem" & CStr(lZone), tvwChild, "ZonesSystem" & CStr(lZone) & "Range" & i, sRange, sIcon
            End If
            
            sDomain = String$(260, 0)
            i = i + 1
            If bShowLargeZones And i Mod 100 = 0 And lNumItems > 0 Then
                Status "Loading... Trusted sites & Restricted sites (all users, " & CInt(CLng(i) * 100 / lNumItems) & "%, " & i & " IPs)"
            End If
            If bAbort Then
                RegCloseKey hKey
                Exit Sub
            End If
        Loop
        RegCloseKey hKey
    End If
    For i = 0 To UBound(sZoneNames)
        If tvwMain.Nodes("ZonesSystem" & i).Children > 0 Then
            tvwMain.Nodes("ZonesSystem" & i).Text = tvwMain.Nodes("ZonesSystem" & i).Text & " (" & tvwMain.Nodes("ZonesSystem" & i).Children & ")"
            tvwMain.Nodes("ZonesSystem" & i).Sorted = True
        Else
            If Not bShowEmpty Then
                tvwMain.Nodes.Remove "ZonesSystem" & i
            End If
        End If
    Next i
    If tvwMain.Nodes("ZonesSystem").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "ZonesSystem"
    End If
        
    If tvwMain.Nodes("Zones").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "Zones"
    End If

    If Not bShowUsers Then Exit Sub
    '----------------------------------------------------------------
    Dim sUsername$, l&
    For l = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(l))
        If sUsername <> GetUser() And sUsername <> vbNullString Then
            Status "Loading... Trusted sites & Restricted sites (" & sUsername & ")"
            tvwMain.Nodes.Add sUsernames(l) & "DisabledEnums", tvwChild, sUsernames(l) & "Zones", SEC_ZONES, "internet"
            
            For i = 0 To UBound(sZoneNames)
                tvwMain.Nodes.Add sUsernames(l) & "Zones", tvwChild, sUsernames(l) & "ZonesUser" & i, sZoneNames(i), "internet"
                tvwMain.Nodes(sUsernames(l) & "ZonesUser" & i).Tag = "HKEY_USERS\" & sUsernames(l) & "\" & sKey & "\ZoneMap\Domains"
            Next i
            If RegOpenKeyEx(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Domains", 0, KEY_READ, hKey) = 0 Then
                RegQueryInfoKey hKey, vbNullString, 0, 0, lNumItems, 0, 0, 0, 0, 0, 0, ByVal 0
                If lNumItems > 1000 And Not bShowLargeZones Then
                    frmMain.ShowError "Skipping Zones for user " & sUsername & ", since there are over 1000 domains in them. (" & lNumItems & " to be exact)"
                    RegCloseKey hKey
                    GoTo CheckUserRanges:
                End If
                
                'loop through subkeys and add them to proper zone
                sDomain = String$(260, 0)
                i = 0
                Do Until RegEnumKeyEx(hKey, i, sDomain, Len(sDomain), 0, vbNullString, 0, ByVal 0) <> 0
                    If RegValExists(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Domains\" & sDomain, "http") Then
                        lZone = RegGetDword(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Domains\" & sDomain, "http")
                    Else
                        If RegValExists(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Domains\" & sDomain, "*") Then
                            lZone = RegGetDword(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Domains\" & sDomain, "*")
                        End If
                    End If
                    
                    If RegValExists(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Domains\" & sDomain, "http") Or _
                       RegValExists(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Domains\" & sDomain, "*") Then
                        Select Case lZone
                            Case 0, 1: sIcon = "system"
                            Case 2: sIcon = "good"
                            Case 3: sIcon = "internet"
                            Case 4: sIcon = "bad"
                            Case Else: sIcon = "internet"
                        End Select
                        tvwMain.Nodes.Add sUsernames(l) & "ZonesUser" & CStr(lZone), tvwChild, sUsernames(l) & "ZonesUser" & CStr(lZone) & i, sDomain, sIcon
                    End If
                    'check for subdomains
                    sSubkeys = Split(RegEnumSubKeys(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Domains\" & sDomain), "|")
                    If UBound(sSubkeys) > -1 Then
                        For j = 0 To UBound(sSubkeys)
                            
                            If RegValExists(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http") Then
                                lZone = RegGetDword(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http")
                            Else
                                If RegValExists(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*") Then
                                    lZone = RegGetDword(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*")
                                End If
                            End If
                            
                            If RegValExists(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "http") Or _
                               RegValExists(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Domains\" & sDomain & "\" & sSubkeys(j), "*") Then
                                Select Case lZone
                                    Case 0, 1: sIcon = "system"
                                    Case 2: sIcon = "good"
                                    Case 3: sIcon = "internet"
                                    Case 4: sIcon = "bad"
                                    Case Else: sIcon = "internet"
                                End Select
                                tvwMain.Nodes.Add sUsernames(l) & "ZonesUser" & CStr(lZone), tvwChild, sUsernames(l) & "ZonesUser" & CStr(lZone) & i & "s" & j, sSubkeys(j) & "." & sDomain, sIcon
                            End If
                        Next j
                    End If
                    
                    i = i + 1
                    sDomain = String$(260, 0)
                    If bShowLargeZones And i Mod 100 = 0 And lNumItems > 0 Then
                        Status "Loading... Trusted sites & Restricted sites (" & sUsername & ", " & CInt(CLng(i) * 100 / lNumItems) & "%, " & i & " domains)"
                    End If
                    If bAbort Then
                        RegCloseKey hKey
                        Exit Sub
                    End If
                Loop
                RegCloseKey hKey
            End If
            
CheckUserRanges:
            If RegOpenKeyEx(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Ranges", 0, KEY_READ, hKey) = 0 Then
                'RegQueryInfoKey hKey, vbNullString, 0, 0, lNumItems, 0, 0, 0, 0, 0, 0, ByVal 0
                sDomain = String$(260, 0)
                i = 0
                Do Until RegEnumKeyEx(hKey, i, sDomain, Len(sDomain), 0, vbNullString, 0, ByVal 0) <> 0
                    sDomain = TrimNull(sDomain)
                    sRange = RegGetString(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Ranges\" & sDomain, ":Range")
                    lZone = RegGetDword(HKEY_USERS, sUsernames(l) & "\" & sKey & "\ZoneMap\Ranges\" & sDomain, "*")
                    
                    If lZone > 0 And Trim$(sRange) <> vbNullString Then
                        Select Case lZone
                            Case 0, 1: sIcon = "system"
                            Case 2: sIcon = "good"
                            Case 3: sIcon = "internet"
                            Case 4: sIcon = "bad"
                            Case Else: sIcon = "internet"
                        End Select
                        tvwMain.Nodes.Add sUsernames(l) & "ZonesUser" & CStr(lZone), tvwChild, sUsernames(l) & "ZonesUser" & CStr(lZone) & "Range" & i, sRange, sIcon
                    End If
                    
                    sDomain = String$(260, 0)
                    i = i + 1
                    If bShowLargeZones And i Mod 100 = 0 And lNumItems > 0 Then
                        Status "Loading... Trusted sites & Restricted sites (" & sUsername & ", " & CInt(CLng(i) * 100 / lNumItems) & "%, " & i & " IPs)"
                    End If
                    If bAbort Then
                        RegCloseKey hKey
                        Exit Sub
                    End If
                Loop
                RegCloseKey hKey
            End If
            
            For i = 0 To UBound(sZoneNames)
                If tvwMain.Nodes(sUsernames(l) & "ZonesUser" & i).Children > 0 Then
                    tvwMain.Nodes(sUsernames(l) & "ZonesUser" & i).Text = tvwMain.Nodes(sUsernames(l) & "ZonesUser" & i).Text & " (" & tvwMain.Nodes(sUsernames(l) & "ZonesUser" & i).Children & ")"
                    tvwMain.Nodes(sUsernames(l) & "ZonesUser" & i).Sorted = True
                Else
                    If Not bShowEmpty Then
                        tvwMain.Nodes.Remove sUsernames(l) & "ZonesUser" & i
                    End If
                End If
            Next i
            
            If tvwMain.Nodes(sUsernames(l) & "Zones").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove sUsernames(l) & "Zones"
            End If
        End If
    Next l
End Sub

Public Sub RegEnumDriverFilters(tvwMain As TreeView)
    'enumerate UpperFilters, LowerFilters on:
    'HKLM\System\CCS\Control\Class\* (Class Lower/Upper Filters)
    'HKLM\System\CCS\Enum\*\*\*      (Device Lower/Upper Filters)
    'HKLM\System\CS?\..etc..
    If bAbort Then Exit Sub
    tvwMain.Nodes.Add "System", tvwChild, "DriverFilters", SEC_DRIVERFILTERS, "dll"
    
    Dim hKey&, i&, j&, sKey$, sName$, sLFilters$(), sUFilters$()
    Dim sClassKey$, sDeviceKey$
    sClassKey = "System\CurrentControlSet\Control\Class"
    sDeviceKey = "System\CurrentControlSet\Enum"
    
    tvwMain.Nodes.Add "DriverFilters", tvwChild, "DriverFiltersClass", "Class filters", "dll"
    tvwMain.Nodes("DriverFiltersClass").Tag = "HKEY_LOCAL_MACHINE\" & sClassKey
    tvwMain.Nodes("DriverFiltersClass").Sorted = True
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sClassKey, 0, KEY_READ, hKey) = 0 Then
        sKey = String$(260, 0)
        Do Until RegEnumKeyEx(hKey, i, sKey, Len(sKey), 0, vbNullString, 0, ByVal 0) <> 0
            sKey = TrimNull(sKey)
            sName = RegGetString(HKEY_LOCAL_MACHINE, sClassKey & "\" & sKey, vbNullString)
            If sName = vbNullString Then sName = "(no name)"
            sLFilters = Split(RegGetString(HKEY_LOCAL_MACHINE, sClassKey & "\" & sKey, "LowerFilters", False), Chr$(0))
            sUFilters = Split(RegGetString(HKEY_LOCAL_MACHINE, sClassKey & "\" & sKey, "UpperFilters", False), Chr$(0))
            'root key for device
            If UBound(sLFilters) > 0 Or UBound(sUFilters) > 0 Then
                tvwMain.Nodes.Add "DriverFiltersClass", tvwChild, "DriverFiltersClass" & i, sName, "hardware"
                tvwMain.Nodes("DriverFiltersClass" & i).Tag = "HKEY_LOCAL_MACHINE\" & sClassKey & "\" & sKey
            End If
            'upper filters
            If UBound(sUFilters) > 0 Then
                tvwMain.Nodes.Add "DriverFiltersClass" & i, tvwChild, "DriverFiltersClass" & i & "Upper", "Upper filters", "dll"
                tvwMain.Nodes("DriverFiltersClass" & i & "Upper").Tag = "HKEY_LOCAL_MACHINE\" & sClassKey & "\" & sKey
                For j = 0 To UBound(sUFilters)
                    If Trim$(sUFilters(j)) <> vbNullString Then
                        sName = sUFilters(j) & ".sys"
                        If FileExists(sSysDir & "\drivers\" & sName) Then
                            sName = BuildPath(sSysDir & "\drivers\", sName)
                        Else
                            sName = GuessFullpathFromAutorun(sName)
                        End If
                        tvwMain.Nodes.Add "DriverFiltersClass" & i & "Upper", tvwChild, "DriverFiltersClass" & i & "Upper" & j, sUFilters(j) & ".sys", "dll"
                        tvwMain.Nodes("DriverFiltersClass" & i & "Upper" & j).Tag = sName
                    End If
                Next j
            End If
            'lower filters
            If UBound(sLFilters) > 0 Then
                tvwMain.Nodes.Add "DriverFiltersClass" & i, tvwChild, "DriverFiltersClass" & i & "Lower", "Lower filters", "dll"
                tvwMain.Nodes("DriverFiltersClass" & i & "Lower").Tag = "HKEY_LOCAL_MACHINE\" & sClassKey & "\" & sKey
                For j = 0 To UBound(sLFilters)
                    If Trim$(sLFilters(j)) <> vbNullString Then
                        sName = sLFilters(j) & ".sys"
                        If FileExists(sSysDir & "\drivers\" & sName) Then
                            sName = BuildPath(sSysDir & "\drivers\", sName)
                        Else
                            sName = GuessFullpathFromAutorun(sName)
                        End If
                        tvwMain.Nodes.Add "DriverFiltersClass" & i & "Lower", tvwChild, "DriverFiltersClass" & i & "Lower" & j, sLFilters(j) & ".sys", "dll"
                        tvwMain.Nodes("DriverFiltersClass" & i & "Lower" & j).Tag = sName
                    End If
                Next j
            End If
            
            
            sKey = String$(260, 0)
            i = i + 1
            If bAbort Then
                RegCloseKey hKey
                Exit Sub
            End If
        Loop
        RegCloseKey hKey
    End If
    If tvwMain.Nodes("DriverFiltersClass").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "DriverFiltersClass"
    End If
    '---------------------
    
    tvwMain.Nodes.Add "DriverFilters", tvwChild, "DriverFiltersDevice", "Device filters", "dll"
    tvwMain.Nodes("DriverFiltersDevice").Tag = "HKEY_LOCAL_MACHINE\" & sDeviceKey
    tvwMain.Nodes("DriverFiltersDevice").Sorted = True
    Dim sSections$(), sDevices$(), sSubkeys$(), k&, m&
    'this fucking sucks
    sSections = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sDeviceKey), "|")
    For i = 0 To UBound(sSections)
        sDevices = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i)), "|")
        For j = 0 To UBound(sDevices)
            sSubkeys = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i) & "\" & sDevices(j)), "|")
            For k = 0 To UBound(sSubkeys)
                sName = RegGetString(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i) & "\" & sDevices(j) & "\" & sSubkeys(k), "DeviceDesc")
                If sName = vbNullString Then sName = "(no name)"
                sUFilters = Split(RegGetString(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i) & "\" & sDevices(j) & "\" & sSubkeys(k), "UpperFilters", False), Chr$(0))
                sLFilters = Split(RegGetString(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i) & "\" & sDevices(j) & "\" & sSubkeys(k), "LowerFilters", False), Chr$(0))
                If UBound(sUFilters) > 0 Or UBound(sLFilters) > 0 Then
                    tvwMain.Nodes.Add "DriverFiltersDevice", tvwChild, "DriverFiltersDevice" & i & "." & j & "." & k, sName, "hardware"
                End If
                If UBound(sUFilters) > 0 Then
                    tvwMain.Nodes.Add "DriverFiltersDevice" & i & "." & j & "." & k, tvwChild, "DriverFiltersDevice" & i & "." & j & "." & k & "Upper", "Upper filters", "dll"
                    For m = 0 To UBound(sUFilters)
                        If Trim$(sUFilters(m)) <> vbNullString Then
                            sName = sUFilters(m) & ".sys"
                            If FileExists(sSysDir & "\drivers\" & sName) Then
                                sName = BuildPath(sSysDir & "\drivers\", sName)
                            Else
                                sName = GuessFullpathFromAutorun(sName)
                            End If
                            tvwMain.Nodes.Add "DriverFiltersDevice" & i & "." & j & "." & k & "Upper", tvwChild, "DriverFiltersDevice" & i & "." & j & "." & k & "Upper" & m, sUFilters(m) & ".sys", "dll"
                            tvwMain.Nodes("DriverFiltersDevice" & i & "." & j & "." & k & "Upper" & m).Tag = sName
                        End If
                    Next m
                End If
                If UBound(sLFilters) > 0 Then
                    tvwMain.Nodes.Add "DriverFiltersDevice" & i & "." & j & "." & k, tvwChild, "DriverFiltersDevice" & i & "." & j & "." & k & "Lower", "Lower filters", "dll"
                    For m = 0 To UBound(sLFilters)
                        If Trim$(sLFilters(m)) <> vbNullString Then
                            sName = sLFilters(m) & ".sys"
                            If FileExists(sSysDir & "\drivers\" & sName) Then
                                sName = BuildPath(sSysDir & "\drivers\", sName)
                            Else
                                sName = GuessFullpathFromAutorun(sName)
                            End If
                            tvwMain.Nodes.Add "DriverFiltersDevice" & i & "." & j & "." & k & "Lower", tvwChild, "DriverFiltersDevice" & i & "." & j & "." & k & "Lower" & m, sLFilters(m) & ".sys", "dll"
                            tvwMain.Nodes("DriverFiltersDevice" & i & "." & j & "." & k & "Lower" & m).Tag = sName
                        End If
                    Next m
                End If
                If bAbort Then Exit Sub
            Next k
        Next j
    Next i
    If tvwMain.Nodes("DriverFiltersDevice").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "DriverFiltersDevice"
    End If
    
    If tvwMain.Nodes("DriverFilters").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "DriverFilters"
    End If
    
    If Not bShowHardware Then Exit Sub
    '-------------------------------------------------------------------------
    Dim l&
    For l = 1 To UBound(sHardwareCfgs)
        sClassKey = "System\" & sHardwareCfgs(l) & "\Control\Class"
        sDeviceKey = "System\" & sHardwareCfgs(l) & "\Enum"
        tvwMain.Nodes.Add "Hardware" & sHardwareCfgs(l), tvwChild, sHardwareCfgs(l) & "DriverFilters", SEC_DRIVERFILTERS, "dll"
        
        tvwMain.Nodes.Add sHardwareCfgs(l) & "DriverFilters", tvwChild, sHardwareCfgs(l) & "DriverFiltersClass", "Class filters", "dll"
        tvwMain.Nodes(sHardwareCfgs(l) & "DriverFiltersClass").Tag = "HKEY_LOCAL_MACHINE\" & sClassKey
        tvwMain.Nodes(sHardwareCfgs(l) & "DriverFiltersClass").Sorted = True
        If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sClassKey, 0, KEY_READ, hKey) = 0 Then
            sKey = String$(260, 0)
            Do Until RegEnumKeyEx(hKey, i, sKey, Len(sKey), 0, vbNullString, 0, ByVal 0) <> 0
                sKey = TrimNull(sKey)
                sName = RegGetString(HKEY_LOCAL_MACHINE, sClassKey & "\" & sKey, vbNullString)
                If sName = vbNullString Then sName = "(no name)"
                sLFilters = Split(RegGetString(HKEY_LOCAL_MACHINE, sClassKey & "\" & sKey, "LowerFilters", False), Chr$(0))
                sUFilters = Split(RegGetString(HKEY_LOCAL_MACHINE, sClassKey & "\" & sKey, "UpperFilters", False), Chr$(0))
                'root key for device
                If UBound(sLFilters) > 0 Or UBound(sUFilters) > 0 Then
                    tvwMain.Nodes.Add sHardwareCfgs(l) & "DriverFiltersClass", tvwChild, sHardwareCfgs(l) & "DriverFiltersClass" & i, sName, "hardware"
                    tvwMain.Nodes(sHardwareCfgs(l) & "DriverFiltersClass" & i).Tag = "HKEY_LOCAL_MACHINE\" & sClassKey & "\" & sKey
                End If
                'upper filters
                If UBound(sUFilters) > 0 Then
                    tvwMain.Nodes.Add sHardwareCfgs(l) & "DriverFiltersClass" & i, tvwChild, sHardwareCfgs(l) & "DriverFiltersClass" & i & "Upper", "Upper filters", "dll"
                    tvwMain.Nodes(sHardwareCfgs(l) & "DriverFiltersClass" & i & "Upper").Tag = "HKEY_LOCAL_MACHINE\" & sClassKey & "\" & sKey
                    For j = 0 To UBound(sUFilters)
                        If Trim$(sUFilters(j)) <> vbNullString Then
                            sName = sUFilters(j) & ".sys"
                            If FileExists(sSysDir & "\drivers\" & sName) Then
                                sName = BuildPath(sSysDir & "\drivers\", sName)
                            Else
                                sName = GuessFullpathFromAutorun(sName)
                            End If
                            tvwMain.Nodes.Add sHardwareCfgs(l) & "DriverFiltersClass" & i & "Upper", tvwChild, sHardwareCfgs(l) & "DriverFiltersClass" & i & "Upper" & j, sUFilters(j) & ".sys", "dll"
                            tvwMain.Nodes(sHardwareCfgs(l) & "DriverFiltersClass" & i & "Upper" & j).Tag = sName
                        End If
                    Next j
                End If
                'lower filters
                If UBound(sLFilters) > 0 Then
                    tvwMain.Nodes.Add sHardwareCfgs(l) & "DriverFiltersClass" & i, tvwChild, sHardwareCfgs(l) & "DriverFiltersClass" & i & "Lower", "Lower filters", "dll"
                    tvwMain.Nodes(sHardwareCfgs(l) & "DriverFiltersClass" & i & "Lower").Tag = "HKEY_LOCAL_MACHINE\" & sClassKey & "\" & sKey
                    For j = 0 To UBound(sLFilters)
                        If Trim$(sLFilters(j)) <> vbNullString Then
                            sName = sLFilters(j) & ".sys"
                            If FileExists(sSysDir & "\drivers\" & sName) Then
                                sName = BuildPath(sSysDir & "\drivers\", sName)
                            Else
                                sName = GuessFullpathFromAutorun(sName)
                            End If
                            tvwMain.Nodes.Add sHardwareCfgs(l) & "DriverFiltersClass" & i & "Lower", tvwChild, sHardwareCfgs(l) & "DriverFiltersClass" & i & "Lower" & j, sLFilters(j) & ".sys", "dll"
                            tvwMain.Nodes(sHardwareCfgs(l) & "DriverFiltersClass" & i & "Lower" & j).Tag = sName
                        End If
                    Next j
                End If
                
                
                sKey = String$(260, 0)
                i = i + 1
                If bAbort Then
                    RegCloseKey hKey
                    Exit Sub
                End If
            Loop
            RegCloseKey hKey
        End If
        If tvwMain.Nodes(sHardwareCfgs(l) & "DriverFiltersClass").Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove sHardwareCfgs(l) & "DriverFiltersClass"
        End If
    
        tvwMain.Nodes.Add sHardwareCfgs(l) & "DriverFilters", tvwChild, sHardwareCfgs(l) & "DriverFiltersDevice", "Device filters", "dll"
        tvwMain.Nodes(sHardwareCfgs(l) & "DriverFiltersDevice").Tag = "HKEY_LOCAL_MACHINE\" & sDeviceKey
        tvwMain.Nodes(sHardwareCfgs(l) & "DriverFiltersDevice").Sorted = True
        'this fucking sucks - again
        sSections = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sDeviceKey), "|")
        For i = 0 To UBound(sSections)
            sDevices = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i)), "|")
            For j = 0 To UBound(sDevices)
                sSubkeys = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i) & "\" & sDevices(j)), "|")
                For k = 0 To UBound(sSubkeys)
                    sName = RegGetString(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i) & "\" & sDevices(j) & "\" & sSubkeys(k), "DeviceDesc")
                    If sName = vbNullString Then sName = "(no name)"
                    sUFilters = Split(RegGetString(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i) & "\" & sDevices(j) & "\" & sSubkeys(k), "UpperFilters", False), Chr$(0))
                    sLFilters = Split(RegGetString(HKEY_LOCAL_MACHINE, sDeviceKey & "\" & sSections(i) & "\" & sDevices(j) & "\" & sSubkeys(k), "LowerFilters", False), Chr$(0))
                    If UBound(sUFilters) > 0 Or UBound(sLFilters) > 0 Then
                        tvwMain.Nodes.Add sHardwareCfgs(l) & "DriverFiltersDevice", tvwChild, sHardwareCfgs(l) & "DriverFiltersDevice" & i & "." & j & "." & k, sName, "hardware"
                    End If
                    If UBound(sUFilters) > 0 Then
                        tvwMain.Nodes.Add sHardwareCfgs(l) & "DriverFiltersDevice" & i & "." & j & "." & k, tvwChild, sHardwareCfgs(l) & "DriverFiltersDevice" & i & "." & j & "." & k & "Upper", "Upper filters", "dll"
                        For m = 0 To UBound(sUFilters)
                            If Trim$(sUFilters(m)) <> vbNullString Then
                                sName = sUFilters(m) & ".sys"
                                If FileExists(sSysDir & "\drivers\" & sName) Then
                                    sName = BuildPath(sSysDir & "\drivers\", sName)
                                Else
                                    sName = GuessFullpathFromAutorun(sName)
                                End If
                                tvwMain.Nodes.Add sHardwareCfgs(l) & "DriverFiltersDevice" & i & "." & j & "." & k & "Upper", tvwChild, sHardwareCfgs(l) & "DriverFiltersDevice" & i & "." & j & "." & k & "Upper" & m, sUFilters(m) & ".sys", "dll"
                                tvwMain.Nodes(sHardwareCfgs(l) & "DriverFiltersDevice" & i & "." & j & "." & k & "Upper" & m).Tag = sName
                            End If
                        Next m
                    End If
                    If UBound(sLFilters) > 0 Then
                        tvwMain.Nodes.Add sHardwareCfgs(l) & "DriverFiltersDevice" & i & "." & j & "." & k, tvwChild, sHardwareCfgs(l) & "DriverFiltersDevice" & i & "." & j & "." & k & "Lower", "Lower filters", "dll"
                        For m = 0 To UBound(sLFilters)
                            If Trim$(sLFilters(m)) <> vbNullString Then
                                sName = sLFilters(m) & ".sys"
                                If FileExists(sSysDir & "\drivers\" & sName) Then
                                    sName = BuildPath(sSysDir & "\drivers\", sName)
                                Else
                                    sName = GuessFullpathFromAutorun(sName)
                                End If
                                tvwMain.Nodes.Add sHardwareCfgs(l) & "DriverFiltersDevice" & i & "." & j & "." & k & "Lower", tvwChild, sHardwareCfgs(l) & "DriverFiltersDevice" & i & "." & j & "." & k & "Lower" & m, sLFilters(m) & ".sys", "dll"
                                tvwMain.Nodes(sHardwareCfgs(l) & "DriverFiltersDevice" & i & "." & j & "." & k & "Lower" & m).Tag = sName
                            End If
                        Next m
                    End If
                    If bAbort Then Exit Sub
                Next k
            Next j
        Next i
        If tvwMain.Nodes(sHardwareCfgs(l) & "DriverFiltersDevice").Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove sHardwareCfgs(l) & "DriverFiltersDevice"
        End If
    Next l
End Sub

Public Sub RegEnumPolicies(tvwMain As TreeView)
    If bAbort Then Exit Sub
    'policies - EVERYTHING
    tvwMain.Nodes.Add "System", tvwChild, "Policies", SEC_POLICIES, "policy"
    tvwMain.Nodes.Add "Policies", tvwChild, "PoliciesUser", "This user", "user"
    tvwMain.Nodes.Add "Policies", tvwChild, "PoliciesSystem", "All users", "users"
    'enum the tree structures below:
    ' Software\Policies
    ' Software\Microsoft\Windows\CurrentVersion\policies
    ' SOFTWARE\Microsoft\Security Center
    'and then enum all values (REG_SZ, REG_DWORD) in there
    
    Dim sPolicyKeys$(), sPolicyNames$(), k&
    ReDim sPolicyNames(1)
    sPolicyNames(0) = "Primary policies"
    sPolicyNames(1) = "Alternate policies"
    'sPolicyNames(2) = "Security Center policies" - moved to XPSecurityCenter
    ReDim sPolicyKeys(1)
    sPolicyKeys(0) = "Software\Policies"
    sPolicyKeys(1) = "Software\Microsoft\Windows\CurrentVersion\policies"
    'sPolicyKeys(2) = "Software\Microsoft\Security Center" - moved to XPSecurityCenter
    
    Dim sRegKeysUser$(), sRegKeysSystem$(), sValues$(), i&, j&
    
    For k = 0 To UBound(sPolicyKeys)
        tvwMain.Nodes.Add "PoliciesUser", tvwChild, "Policies" & k & "User", sPolicyNames(k), "winlogon"
        tvwMain.Nodes.Add "PoliciesSystem", tvwChild, "Policies" & k & "System", sPolicyNames(k), "winlogon"
        tvwMain.Nodes("Policies" & k & "User").Tag = "HKEY_CURRENT_USER\" & sPolicyKeys(k)
        tvwMain.Nodes("Policies" & k & "System").Tag = "HKEY_LOCAL_MACHINE\" & sPolicyKeys(k)
        
        sValues = Split(RegEnumValues(HKEY_CURRENT_USER, sPolicyKeys(k), , , False), "|")
        For j = 0 To UBound(sValues)
            tvwMain.Nodes.Add "Policies" & k & "User", tvwChild, "Policies" & k & "User" & j, sValues(j), "reg"
        Next j
        
        sValues = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sPolicyKeys(k), , , False), "|")
        For j = 0 To UBound(sValues)
            tvwMain.Nodes.Add "Policies" & k & "System", tvwChild, "Policies" & k & "System" & j, sValues(j), "reg"
        Next j
        
        sRegKeysUser = Split(RegEnumSubKeysTree(HKEY_CURRENT_USER, sPolicyKeys(k)), "|")
        sRegKeysSystem = Split(RegEnumSubKeysTree(HKEY_LOCAL_MACHINE, sPolicyKeys(k)), "|")
        
        For i = 0 To UBound(sRegKeysUser)
            sValues = Split(RegEnumValues(HKEY_CURRENT_USER, sRegKeysUser(i), , , False), "|")
            If UBound(sValues) > -1 Then
                tvwMain.Nodes.Add "Policies" & k & "User", tvwChild, "Policies" & k & "User" & i, sRegKeysUser(i), "registry"
                tvwMain.Nodes("Policies" & k & "User" & i).Tag = "HKEY_CURRENT_USER\" & sRegKeysUser(i)
                For j = 0 To UBound(sValues)
                    tvwMain.Nodes.Add "Policies" & k & "User" & i, tvwChild, "Policies" & k & "User" & i & "." & j, sValues(j), "reg"
                Next j
                tvwMain.Nodes("Policies" & k & "User" & i).Text = tvwMain.Nodes("Policies" & k & "User" & i).Text & " (" & tvwMain.Nodes("Policies" & k & "User" & i).Children & ")"
            End If
            If bAbort Then Exit Sub
        Next i
        For i = 0 To UBound(sRegKeysSystem)
            sValues = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sRegKeysSystem(i), , , False), "|")
            If UBound(sValues) > -1 Then
                tvwMain.Nodes.Add "Policies" & k & "System", tvwChild, "Policies" & k & "System" & i, sRegKeysSystem(i), "registry"
                tvwMain.Nodes("Policies" & k & "System" & i).Tag = "HKEY_LOCAL_MACHINE\" & sRegKeysSystem(i)
                For j = 0 To UBound(sValues)
                    tvwMain.Nodes.Add "Policies" & k & "System" & i, tvwChild, "Policies" & k & "System" & i & "." & j, sValues(j), "reg"
                Next j
                tvwMain.Nodes("Policies" & k & "System" & i).Text = tvwMain.Nodes("Policies" & k & "System" & i).Text & " (" & tvwMain.Nodes("Policies" & k & "System" & i).Children & ")"
            End If
            If bAbort Then Exit Sub
        Next i
        
        If tvwMain.Nodes("Policies" & k & "User").Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove "Policies" & k & "User"
        End If
        If tvwMain.Nodes("Policies" & k & "System").Children = 0 And Not bShowEmpty Then
            tvwMain.Nodes.Remove "Policies" & k & "System"
        End If
        If bAbort Then Exit Sub
    Next k
    
    If tvwMain.Nodes("PoliciesUser").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "PoliciesUser"
    End If
    If tvwMain.Nodes("PoliciesSystem").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "PoliciesSystem"
    End If
    
    If tvwMain.Nodes("Policies").Children = 0 And Not bShowEmpty Then
        tvwMain.Nodes.Remove "Policies"
    End If

    If Not bShowUsers Then Exit Sub
    '-----------------------------------------------------------------------
    Dim sUsername$, l&
    For l = 0 To UBound(sUsernames)
        sUsername = MapSIDToUsername(sUsernames(l))
        If sUsername <> GetUser() And sUsername <> vbNullString Then
            tvwMain.Nodes.Add "Users" & sUsernames(l), tvwChild, sUsernames(l) & "PoliciesUser", SEC_POLICIES, "policy"

            For k = 0 To UBound(sPolicyKeys)
                tvwMain.Nodes.Add sUsernames(l) & "PoliciesUser", tvwChild, sUsernames(l) & "Policies" & k & "User", sPolicyNames(k), "winlogon"
                tvwMain.Nodes(sUsernames(l) & "Policies" & k & "User").Tag = "HKEY_USERS\" & sUsernames(l) & "\" & sPolicyKeys(k)
                
                sValues = Split(RegEnumValues(HKEY_USERS, sUsernames(l) & "\" & sPolicyKeys(k), , , False), "|")
                For j = 0 To UBound(sValues)
                    tvwMain.Nodes.Add sUsernames(l) & "Policies" & k & "User", tvwChild, sUsernames(l) & "Policies" & k & "User" & j, sValues(j), "reg"
                Next j
                
                sRegKeysUser = Split(RegEnumSubKeysTree(HKEY_USERS, sUsernames(l) & "\" & sPolicyKeys(k)), "|")
                
                For i = 0 To UBound(sRegKeysUser)
                    sValues = Split(RegEnumValues(HKEY_USERS, sRegKeysUser(i), , , False), "|")
                    If UBound(sValues) > -1 Then
                        tvwMain.Nodes.Add sUsernames(l) & "Policies" & k & "User", tvwChild, sUsernames(l) & "Policies" & k & "User" & i, Mid$(sRegKeysUser(i), Len(sUsernames(l)) + 2), "registry"
                        tvwMain.Nodes(sUsernames(l) & "Policies" & k & "User" & i).Tag = "HKEY_USERS\" & sRegKeysUser(i)
                        For j = 0 To UBound(sValues)
                            tvwMain.Nodes.Add sUsernames(l) & "Policies" & k & "User" & i, tvwChild, sUsernames(l) & "Policies" & k & "User" & i & "." & j, sValues(j), "reg"
                        Next j
                        tvwMain.Nodes(sUsernames(l) & "Policies" & k & "User" & i).Text = tvwMain.Nodes(sUsernames(l) & "Policies" & k & "User" & i).Text & " (" & tvwMain.Nodes(sUsernames(l) & "Policies" & k & "User" & i).Children & ")"
                    End If
                    If bAbort Then Exit Sub
                Next i
                
                If tvwMain.Nodes(sUsernames(l) & "Policies" & k & "User").Children = 0 And Not bShowEmpty Then
                    tvwMain.Nodes.Remove sUsernames(l) & "Policies" & k & "User"
                End If
            Next k
            
            If tvwMain.Nodes(sUsernames(l) & "PoliciesUser").Children = 0 And Not bShowEmpty Then
                tvwMain.Nodes.Remove sUsernames(l) & "PoliciesUser"
            End If
        End If
        If bAbort Then Exit Sub
    Next l
End Sub

Public Sub RegEnumDrivers32(tvwMain As TreeView)
    If bAbort Then Exit Sub
    Const sDrivers$ = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Drivers32"
    
    tvwMain.Nodes.Add "System", tvwChild, "Drivers32", SEC_DRIVERS32, "dll"
    tvwMain.Nodes("Drivers32").Tag = "HKEY_LOCAL_MACHINE\" & sDrivers
    tvwMain.Nodes("Drivers32").Sorted = True
    Dim i&, sDriverKeys$()
    sDriverKeys = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sDrivers), "|")
    For i = 0 To UBound(sDriverKeys)
        tvwMain.Nodes.Add "Drivers32", tvwChild, "Drivers32" & i, sDriverKeys(i), "dll", "dll"
        tvwMain.Nodes("Drivers32" & i).Tag = GuessFullpathFromAutorun(Mid(sDriverKeys(i), InStrRev(sDriverKeys(i), " = ") + 3))
        If bAbort Then Exit Sub
    Next i
    
    tvwMain.Nodes.Add "Drivers32", tvwChild, "Drivers32RDP", " Terminal Services", "internet", "internet"
    tvwMain.Nodes("Drivers32RDP").Tag = "HKEY_LOCAL_MACHINE\" & sDrivers & "\Terminal Server\RDP"
    tvwMain.Nodes("Drivers32RDP").Sorted = True
    sDriverKeys = Split(RegEnumValues(HKEY_LOCAL_MACHINE, sDrivers & "\Terminal Server\RDP"), "|")
    For i = 0 To UBound(sDriverKeys)
        tvwMain.Nodes.Add "Drivers32RDP", tvwChild, "Drivers32RDP" & i, sDriverKeys(i), "dll", "dll"
        tvwMain.Nodes("Drivers32RDP" & i).Tag = GuessFullpathFromAutorun(Mid(sDriverKeys(i), InStrRev(sDriverKeys(i), " = ") + 3))
    Next i
    
    If tvwMain.Nodes("Drivers32RDP").Children > 0 Then
        tvwMain.Nodes("Drivers32RDP").Text = tvwMain.Nodes("Drivers32RDP").Text & " (" & tvwMain.Nodes("Drivers32RDP").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "Drivers32RDP"
    End If
    If tvwMain.Nodes("Drivers32").Children > 0 Then
        tvwMain.Nodes("Drivers32").Text = tvwMain.Nodes("Drivers32").Text & " (" & tvwMain.Nodes("Drivers32").Children & ")"
    Else
        If Not bShowEmpty Then tvwMain.Nodes.Remove "Drivers32"
    End If
End Sub

Public Function RegEnumSubKeysTree$(lHive&, sRootKey$)
    Dim hKey&, i&, sName$, sList$
    If bAbort Then Exit Function
    If RegOpenKeyEx(lHive, sRootKey, 0, KEY_READ, hKey) = 0 Then
        sName = String$(260, 0)
        Do Until RegEnumKeyEx(hKey, i, sName, Len(sName), 0, vbNullString, 0, ByVal 0) <> 0
            sName = TrimNull(sName)
            
            sList = sList & "|" & sRootKey & "\" & sName
            sList = sList & "|" & RegEnumSubKeysTree(lHive, sRootKey & "\" & sName)
            
            i = i + 1
            sName = String$(260, 0)
            If bAbort Then
                RegCloseKey hKey
                Exit Function
            End If
        Loop
        RegCloseKey hKey
    End If
    If sList <> vbNullString Then RegEnumSubKeysTree = Mid$(Replace(sList, "||", "|"), 2)
End Function

