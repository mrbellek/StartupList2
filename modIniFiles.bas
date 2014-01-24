Attribute VB_Name = "modIniFiles"
Option Explicit

Public Function IniGetString$(ByVal sFile$, sSection$, sVal$, Optional bMultiple As Boolean = False)
    If Not FileExists(sFile) Then
        If FileExists(BuildPath(sWinDir, sFile)) Then
            sFile = BuildPath(sWinDir, sFile)
        Else
            If FileExists(BuildPath(sSysDir, sFile)) Then
                sFile = BuildPath(sSysDir, sFile)
            Else
                If FileExists(BuildPath(Left$(sWinDir, 3), sFile)) Then
                    sFile = BuildPath(Left$(sWinDir, 3), sFile)
                Else
                    Exit Function
                End If
            End If
        End If
    End If
        
    Dim sContents$(), i&, sData$
    sContents = Split(InputFile(sFile), vbCrLf)
    Do Until InStr(1, sContents(i), "[" & sSection & "]", vbTextCompare) = 1
        i = i + 1
        If i > UBound(sContents) Then Exit Function
    Loop
    i = i + 1
    Do Until Left$(sContents(i), 1) = "["
        If InStr(1, sContents(i), sVal, vbTextCompare) = 1 Then
            sData = sData & "|" & sContents(i)
            If Not bMultiple Then Exit Do
        End If
        i = i + 1
        If i > UBound(sContents) Then Exit Do
    Loop
    'IniGetString = Mid$(sContents(i), InStr(sContents(i), "=") + 1)
    If sData <> vbNullString Then
        If Not bMultiple Then
            IniGetString = Mid$(sData, InStr(sData, "=") + 1)
        Else
            IniGetString = Replace(Mid$(sData, 2), "=", " = ")
        End If
    End If
End Function
