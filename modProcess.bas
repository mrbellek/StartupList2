Attribute VB_Name = "modProcess"
Option Explicit
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long

Private Declare Function EnumProcesses Lib "PSAPI.DLL" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Private Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    modBaseSize As Long
    hModule As Long
    szModule As String * 256
    szExePath As String * 260
End Type

Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPMODULE = &H8
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16

Public Function GetRunningProcesses$()
    Dim hSnap&, uPE32 As PROCESSENTRY32, sList$, i&, hProc&
    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&, sProcessName$, lModules&(1 To 1024)
    
    If Not bIsWinNT Then
        'windows 9x/me method
        hSnap = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)
        If hSnap > 0 Then
            uPE32.dwSize = Len(uPE32)
            If ProcessFirst(hSnap, uPE32) = 0 Then
                CloseHandle hSnap
                Exit Function
            End If
            
            Do
                sList = sList & "|" & uPE32.th32ProcessID & "=" & TrimNull(uPE32.szExeFile)
                If bAbort Then Exit Function
            Loop Until ProcessNext(hSnap, uPE32) = 0
            CloseHandle hSnap
        End If
    Else
        'windows nt/2k/xp/2003/etc method
        On Error Resume Next
        If EnumProcesses(lProcesses(1), CLng(1024) * 4, lNeeded) = 0 Then
            frmMain.ShowError "PSAPI.DLL not found, unable to list running processes."
            Exit Function
        End If
        lNumProcesses = lNeeded / 4
        For i = 1 To lNumProcesses
            hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(i))
            If hProc > 0 Then
                lNeeded = 0
                sProcessName = String$(260, 0)
                If EnumProcessModules(hProc, lModules(i), CLng(1024) * 4, lNeeded) > 0 Then
                    GetModuleFileNameExA hProc, lModules(1), sProcessName, Len(sProcessName)
                    sProcessName = TrimNull(sProcessName)
                    If sProcessName <> vbNullString Then
                        If Left$(sProcessName, 1) = "\" Then sProcessName = Mid$(sProcessName, 2)
                        If Left$(sProcessName, 3) = "??\" Then sProcessName = Mid$(sProcessName, 4)
                        sProcessName = ExpandEnvironmentVars(sProcessName)
                        If InStr(1, sProcessName, "SystemRoot", vbTextCompare) > 0 Then sProcessName = Replace(sProcessName, "SystemRoot", sWinDir, , , vbTextCompare)
                        
                        sList = sList & "|" & lProcesses(i) & "=" & sProcessName
                    End If
                End If
                CloseHandle hProc
            End If
            If bAbort Then Exit Function
        Next i
    End If
    If sList <> vbNullString Then GetRunningProcesses = Mid$(sList, 2)
End Function

Public Function GetLoadedModules$(lPID&, sProcess$)
    Dim sModuleList$
    Dim hProc&, lNeeded&, i&, lNumProcesses&, sModuleName$, lModules&(1 To 1024)
    Dim hSnap&, uME32 As MODULEENTRY32

    If Not bIsWinNT Then
        hSnap = CreateToolhelpSnapshot(TH32CS_SNAPMODULE, lPID)
        uME32.dwSize = Len(uME32)
        If Module32First(hSnap, uME32) = 0 Then
            CloseHandle hSnap
            Exit Function
        End If
        Do
            sModuleName = TrimNull(uME32.szExePath)
            If InStr(1, sProcess, sModuleName, vbTextCompare) = 0 Then
                sModuleList = sModuleList & "|" & sModuleName
            End If
            If bAbort Then Exit Function
        Loop Until Module32Next(hSnap, uME32) = 0
        CloseHandle hSnap
    Else
        hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lPID)
        If hProc > 0 Then
            lNeeded = 0
            If EnumProcessModules(hProc, lModules(1), CLng(1024) * 4, lNeeded) > 0 Then
                For i = 2 To 1024
                    If lModules(i) = 0 Then Exit For
                    sModuleName = String$(260, 0)
                    GetModuleFileNameExA hProc, lModules(i), sModuleName, Len(sModuleName)
                    sModuleName = TrimNull(sModuleName)
                    If sModuleName <> vbNullString And sModuleName <> "?" Then
                        sModuleList = sModuleList & "|" & sModuleName
                    End If
                    If bAbort Then Exit Function
                Next i
            End If
            CloseHandle hProc
        End If
    End If
    If sModuleList <> vbNullString Then GetLoadedModules = Mid$(sModuleList, 2)
End Function
