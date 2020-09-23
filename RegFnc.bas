Attribute VB_Name = "RegFnc"
Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function SHDeleteKey Lib "shlwapi.dll" Alias "SHDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long


' Registry value type definitions
Private Const REG_NONE                  As Long = 0
Private Const REG_SZ                    As Long = 1
Private Const REG_EXPAND_SZ             As Long = 2
Private Const REG_BINARY                As Long = 3
Private Const REG_DWORD                 As Long = 4
Private Const REG_LINK                  As Long = 6
Private Const REG_MULTI_SZ              As Long = 7
Private Const REG_RESOURCE_LIST         As Long = 8
Private Const MAX_PATH                  As Long = 260
Private Const HKEY_CLASSES_ROOT         As Long = &H80000000
Private Const HKEY_CURRENT_USER         As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE        As Long = &H80000002
Private Const HKEY_USERS                As Long = &H80000003
Private Const HKEY_PERFORMANCE_DATA     As Long = &H80000004

Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_NOTIFY = &H10

Private Const csDescription = "==============================================" & vbCrLf & _
        "Registry Cleaner Utility Command Line Parameters" & vbCrLf & _
        "==============================================" & vbCrLf & _
        "Remove references to *.dll file from Registry" & vbCrLf & vbCrLf & _
        "/?                 Display this dialog" & vbCrLf & _
        "/s                 Disable all notifications (silent mode)" & vbCrLf & _
        "[filename]    Name of *.dll file" & vbCrLf
Private msLogFile   As String
Public Sub Main()
    Dim sDLLName    As String
    Dim sCommand    As String
    Dim bSilentMode As Boolean
    sCommand = Command()
    msLogFile = App.Path & "\DLLClean.log"
    EnsurePath msLogFile
    If Len(sCommand) = o Then
        sDLLName = InputBox("Please enter the DLL name", "Registry Cleaner", "Sample.dll")
    Else
        sDLLName = sCommand
        If InStr(sDLLName, "/s") > 0 Then
            bSilentMode = True
            sDLLName = Replace(sDLLName, " /s", "")
        End If
        sDLLName = Trim(sDLLName)
        If InStr(sDLLName, " ") > 0 Or InStr(sDLLName, Chr(9)) > 0 Or InStr(sDLLName, "/?") > 0 Then
            ShowDescription
            Exit Sub
        End If
        If InStr(sDLLName, ".dll") = 0 Then
            ShowDescription
            Exit Sub
        End If
    End If
    sDLLName = LCase(sDLLName)
    If Len(sDLLName) = 0 Or InStr(sDLLName, "sample.dll") > 0 Then Exit Sub
    If InStr(sDLLName, ".dll") = 0 Then
        MsgBox "You can clean references to *.dll file only!", vbOKOnly, "Registry Cleaner"
        Exit Sub
    End If
    WriteLog msLogFile, "Cleaning references to " & sDLLName
    Call CleanUpReg(sDLLName)
    If Not bSilentMode Then MsgBox "Cleanup completed." & vbCrLf & vbCrLf & "Please refer to log file for details:" & vbCrLf & msLogFile, vbOKOnly + vbInformation, "Registry Cleaner"
End Sub
Private Sub ShowDescription()
    MsgBox csDescription, vbInformation + vbOKOnly, "Registry Cleaner"
End Sub
Public Function CleanUpReg(sDLLName As String)
'   search through two sections:
'   HKEY_CLASSES_ROOT\TypeLib\
'   HKEY_CLASSES_ROOT\CLSID\
'   can have AppID - should be deleted in HKEY_CLASSES_ROOT\AppID\
    Dim lKey        As Long
    Dim lSubKey     As Long
    Dim lChildKey   As Long
    Dim lResult     As Long
    Dim lIndex      As Long
    Dim sName       As String
    Dim sClass      As String
    Dim lNameLen    As Long
    Dim lClassLen   As Long
    Dim ft          As FILETIME
    sName = Space(MAX_PATH + 1)
    sClass = Space(MAX_PATH + 1)
    lNameLen = MAX_PATH + 1
    lClassLen = MAX_PATH + 1
    lResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, "TypeLib", 0&, _
    (KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_NOTIFY Or KEY_ENUMERATE_SUB_KEYS), lKey)
    lIndex = 0
    lResult = RegEnumKeyEx(lKey, lIndex, sName, lNameLen, 0&, sClass, lClassLen, ft)
    Do Until lResult <> 0
        sName = Left(sName, lNameLen)
        sClass = Left(sName, lClassLen)
        lResult = GetSubKeyInfo(lKey, sName, sDLLName)
        If lResult = -1 Then
            WriteLog msLogFile, "... HKEY_CLASSES_ROOT\TypeLib\" & sName & " deleted"
            lResult = SHDeleteKey(lKey, sName)
        End If
        lIndex = lIndex + 1
        sName = Space(MAX_PATH + 1)
        sClass = Space(MAX_PATH + 1)
        lNameLen = MAX_PATH + 1
        lClassLen = MAX_PATH + 1
        lResult = RegEnumKeyEx(lKey, lIndex, sName, lNameLen, 0&, sClass, lClassLen, ft)
    Loop
    RegCloseKey lKey
    sName = Space(MAX_PATH + 1)
    sClass = Space(MAX_PATH + 1)
    lNameLen = MAX_PATH + 1
    lClassLen = MAX_PATH + 1
    lResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, "CLSID", 0&, _
    (KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_NOTIFY Or KEY_ENUMERATE_SUB_KEYS), lKey)
    lIndex = 0
    lResult = RegEnumKeyEx(lKey, lIndex, sName, lNameLen, 0&, sClass, lClassLen, ft)
    Do Until lResult <> 0
        sName = Left(sName, lNameLen)
        sClass = Left(sName, lClassLen)
        lResult = GetSubKeyInfo(lKey, sName, sDLLName)
        If lResult = -1 Then
            WriteLog msLogFile, "... HKEY_CLASSES_ROOT\CLSID\" & sName & " deleted"
            lResult = SHDeleteKey(lKey, sName)
        End If
        lIndex = lIndex + 1
        sName = Space(MAX_PATH + 1)
        sClass = Space(MAX_PATH + 1)
        lNameLen = MAX_PATH + 1
        lClassLen = MAX_PATH + 1
        lResult = RegEnumKeyEx(lKey, lIndex, sName, lNameLen, 0&, sClass, lClassLen, ft)
    Loop
    RegCloseKey lKey
ReSetLoop:
    sName = Space(MAX_PATH + 1)
    sClass = Space(MAX_PATH + 1)
    lNameLen = MAX_PATH + 1
    lClassLen = MAX_PATH + 1
    lResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, "", 0&, _
    (KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_NOTIFY Or KEY_ENUMERATE_SUB_KEYS), lKey)
    lIndex = 0
    lResult = RegEnumKeyEx(lKey, lIndex, sName, lNameLen, 0&, sClass, lClassLen, ft)
    Do Until lResult <> 0
        sName = Left(sName, lNameLen)
        sClass = Left(sName, lClassLen)
        lResult = GetSubKeyInfo(lKey, sName, sDLLName)
        If lResult = -1 Then
            If UCase(sName) <> "CLSID" Then
                WriteLog msLogFile, "... HKEY_CLASSES_ROOT\" & sName & " deleted"
                lResult = RegOpenKeyEx(lKey, sName, 0&, KEY_SET, lSubKey)
                lResult = RegOpenKeyEx(lSubKey, "CLSID", 0&, KEY_SET, lChildKey)
                lResult = RegDeleteKey(lSubKey, "CLSID")
                lResult = RegDeleteKey(lKey, sName)
                GoTo ReSetLoop
            End If
        End If
        lIndex = lIndex + 1
        sName = Space(MAX_PATH + 1)
        sClass = Space(MAX_PATH + 1)
        lNameLen = MAX_PATH + 1
        lClassLen = MAX_PATH + 1
        lResult = RegEnumKeyEx(lKey, lIndex, sName, lNameLen, 0&, sClass, lClassLen, ft)
    Loop
    RegCloseKey lKey
    
End Function
Private Function GetSubKeyInfo(lKeyHdl As Long, sSubKeyName As String, sMatch As String) As Long
    Dim lKey        As Long
    Dim lSubKey     As Long
    Dim lResult     As Long
    Dim lIndex      As Long
    Dim sName       As String
    Dim sClass      As String
    Dim sValue      As String
    Dim lNameLen    As Long
    Dim lClassLen   As Long
    Dim lValueLen   As Long
    Dim ft          As FILETIME
    sName = Space(MAX_PATH + 1)
    sClass = Space(MAX_PATH + 1)
    sValue = Space(MAX_PATH + 1)
    lNameLen = MAX_PATH + 1
    lClassLen = MAX_PATH + 1
    lValueLen = MAX_PATH + 1
    lResult = RegOpenKeyEx(lKeyHdl, sSubKeyName, 0&, KEY_SET, lSubKey)
    lIndex = 0
    lResult = RegQueryValueEx(lSubKey, "", 0&, 0&, sValue, lValueLen)
    sValue = LCase(Left(sValue, lValueLen))
    If InStr(sValue, sMatch) > 0 Then
        RegCloseKey lSubKey
        GetSubKeyInfo = -1
        Exit Function
    End If
    lResult = RegEnumKeyEx(lSubKey, lIndex, sName, lNameLen, 0&, sClass, lClassLen, ft)
    Do Until lResult <> 0
        DoEvents
        sName = LCase(Left(sName, lNameLen))
        sClass = LCase(Left(sClass, lClassLen))
        If InStr(sName, sMatch) > 0 Or InStr(sClass, sMatch) > 0 Then
            RegCloseKey lSubKey
            GetSubKeyInfo = -1
            Exit Function
        End If
        lResult = RegQueryValueEx(lSubKey, "", 0&, 0&, sValue, lValueLen)
        sValue = LCase(Left(sValue, lValueLen))
        If InStr(sValue, sMatch) > 0 Then
            RegCloseKey lSubKey
            GetSubKeyInfo = -1
            Exit Function
        End If
        lResult = GetSubKeyInfo(lSubKey, sName, sMatch)
        If lResult = -1 Then Exit Do
        lIndex = lIndex + 1
        sName = Space(MAX_PATH + 1)
        sClass = Space(MAX_PATH + 1)
        lNameLen = MAX_PATH + 1
        lClassLen = MAX_PATH + 1
        lResult = RegEnumKeyEx(lSubKey, lIndex, sName, lNameLen, 0&, sClass, lClassLen, ft)
    Loop
    RegCloseKey lSubKey
    GetSubKeyInfo = lResult
End Function
Private Sub WriteLog(ByVal sFileName As String, ByRef sMsgText As String)
    Dim iFile   As Integer
On Error GoTo ErrorHandler
    If sFileName = vbNullString Or sMsgText = vbNullString Then Exit Sub
    iFile = FreeFile
    Open sFileName For Append As #iFile
    Print #iFile, , vbCrLf & Format$(Now, "YYYY-MM-DD hh:mm:ss") & " " & sMsgText
ExitHere:
    Close #iFile
    Exit Sub
ErrorHandler:
    If Err.Number = 76 Then
        EnsurePath sFileName
        Resume
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitHere
    Resume
End Sub
Private Sub EnsurePath(sLocation As String)
    Dim nPos0   As Long
    Dim sFolder As String
    If sLocation = vbNullString Then Exit Sub
    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    Do
        nPos0 = InStr(nPos0 + 1, sLocation, "\")
        sFolder = Left(sLocation, nPos0)
        If Not oFileSystem.FolderExists(sFolder) And Len(sFolder) > 0 Then
            Call oFileSystem.CreateFolder(sFolder)
        End If
    Loop While nPos0 > 0
    Set oFileSystem = Nothing
End Sub
