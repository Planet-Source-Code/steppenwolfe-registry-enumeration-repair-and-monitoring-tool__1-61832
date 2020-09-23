Attribute VB_Name = "mRecovery"
Option Explicit

Private Type SMGRSTATUS
    nStatus                                   As Long
    llSequenceNumber                          As Currency
End Type

Private Const REG_SZ                      As Long = &H1
Private Const REG_DWORD                   As Long = &H4
Private Const HKEY_CLASSES_ROOT           As Long = &H80000000
Private Const HKEY_CURRENT_USER           As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE          As Long = &H80000002
Private Const HKEY_USERS                  As Long = &H80000003

Private Const TOKEN_QUERY                 As Long = &H8&
Private Const TOKEN_ADJUST_PRIVILEGES     As Long = &H20&
Private Const SE_PRIVILEGE_ENABLED        As Long = &H2
Private Const SE_RESTORE_NAME             As String = "SeRestorePrivilege"
Private Const SE_BACKUP_NAME              As String = "SeBackupPrivilege"
Private Const REG_FORCE_RESTORE           As Long = 8&
Private Const READ_CONTROL                As Long = &H20000
Private Const SYNCHRONIZE                 As Long = &H100000
Private Const STANDARD_RIGHTS_ALL         As Long = &H1F0000
Private Const KEY_QUERY_VALUE             As Long = &H1
Private Const KEY_SET_VALUE               As Long = &H2
Private Const KEY_CREATE_SUB_KEY          As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS      As Long = &H8
Private Const KEY_NOTIFY                  As Long = &H10
Private Const KEY_CREATE_LINK             As Long = &H20
Private Const KEY_ALL_ACCESS              As Double = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Type LUID
    lowpart                                   As Long
    highpart                                  As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid                                     As LUID
    Attributes                                As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount                            As Long
    Privileges                                As LUID_AND_ATTRIBUTES
End Type

Private Declare Function GetCurrentProcess Lib "Kernel32" () As Long
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, _
                                                                            ByVal lpFile As String, _
                                                                            lpSecurityAttributes As Any) As Long
                                                                            
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long
                                                                                
Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, _
                                                                                  ByVal lpFile As String, _
                                                                                  ByVal dwFlags As Long) As Long
                                                                                  
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, _
                                                                   ByVal DisableAllPriv As Long, _
                                                                   NewState As TOKEN_PRIVILEGES, _
                                                                   ByVal BufferLength As Long, _
                                                                   PreviousState As TOKEN_PRIVILEGES, _
                                                                   ReturnLength As Long) As Long
                                                                   
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, _
                                                                                                ByVal lpName As String, _
                                                                                                lpLuid As LUID) As Long
                                                                                                
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, _
                                                              ByVal DesiredAccess As Long, _
                                                              TokenHandle As Long) As Long


Private Sub Make_Directory(ByVal sFolder As String)

    MakeSureDirectoryPathExists sFolder

End Sub

Public Sub Backup_Key()

Dim sPath As String
Dim hKey  As Long

        '//backup used by monitor sub
    sPath = App.Path & "\Recovery\"
    Make_Directory sPath
    sPath = sPath & sModmname & ".kbs"

    If EnablePrivilege(SE_BACKUP_NAME) Then
        RegOpenKeyEx lModlhkey, sModskey, 0&, KEY_ALL_ACCESS, hKey
        If LenB(Dir(sPath)) Then
            Kill sPath
        End If
        RegSaveKey hKey, sPath, ByVal 0&
        RegCloseKey hKey
    End If

End Sub

Private Function EnablePrivilege(seName As String) As Boolean

Dim p_lngRtn           As Long
Dim p_lngToken         As Long
Dim p_lngBufferLen     As Long
Dim p_typLUID          As LUID
Dim p_typTokenPriv     As TOKEN_PRIVILEGES
Dim p_typPrevTokenPriv As TOKEN_PRIVILEGES

    p_lngRtn = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_lngToken)
    If p_lngRtn = 0 Then
        Exit Function
    ElseIf Err.LastDllError <> 0 Then
        Exit Function
    End If
    p_lngRtn = LookupPrivilegeValue(0&, seName, p_typLUID)
    If p_lngRtn = 0 Then
        Exit Function
    End If
    With p_typTokenPriv
        .PrivilegeCount = 1
        .Privileges.Attributes = SE_PRIVILEGE_ENABLED
        .Privileges.pLuid = p_typLUID
    End With

    EnablePrivilege = (AdjustTokenPrivileges(p_lngToken, False, p_typTokenPriv, Len(p_typPrevTokenPriv), p_typPrevTokenPriv, p_lngBufferLen) <> 0)

End Function

Public Function Save_Key(ByVal sKey As String, _
                         lKey As Long, _
                         ByVal sName As String) As Boolean

Dim hKey     As Long
Dim sPath    As String
Dim sCurDate As String
Dim RetVal   As Long

    sCurDate = Format$(Now, ("dd_mmm_yy"))
    sPath = App.Path & "\Recovery\Backup-" & sCurDate & Chr$(92)
    Make_Directory sPath
    sName = Trim$(String_Convert(sName)) & ".kbs"
    sPath = sPath & sName
    'Debug.Print sPath
    '//needed user proofing, installing to
    '//wrong key should fail, but you never know..
    '//conversion needed to save file name
    '//but must be a better way, if you know one, post it and email me
    If EnablePrivilege(SE_BACKUP_NAME) Then
        RegOpenKeyEx lKey, sKey, 0&, KEY_ALL_ACCESS, hKey
        If LenB(Dir(sPath)) Then
            Kill sPath
        End If
        RetVal = RegSaveKey(hKey, sPath, ByVal 0&)
        RegCloseKey hKey
    End If

    If RetVal = 0 Then
        Save_Key = True
    End If

End Function

Public Function Deploy_Key(ByVal sImage As String) As Boolean

Dim hKey   As Long
Dim sKey   As String
Dim lHKey  As Long
Dim sHkey  As String
Dim sPath  As String
Dim RetVal As Long

    sPath = sImage
    sImage = Mid$(sImage, InStrRev(sImage, Chr$(92)) + 1)
    sImage = String_Convert(sImage)
    sHkey = Left$(sImage, InStr(sImage, Chr$(92)) - 1)
    sKey = Mid$(sImage, InStr(sImage, Chr$(92)) + 1)
    sKey = Left$(sKey, InStr(sKey, Chr$(46)) - 1)
    '//a lot of work just to user proof
    '//a better way would be to give the
    '//file a generic name, and write the path
    '//in the file header info, or to the file
    '//itself.. just a quick fix here though
    Select Case sHkey
    Case "HKEY_CLASSES_ROOT"
        lHKey = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
        lHKey = HKEY_CURRENT_USER
    Case "HKEY_LOCAL_MACHINE"
        lHKey = HKEY_LOCAL_MACHINE
    Case "HKEY_USERS"
        lHKey = HKEY_USERS
    Case "HKEY_CURRENT_CONFIG"
        lHKey = HKEY_CURRENT_CONFIG
    Case Else
        MsgBox "Invalid Root Key!" & vbNewLine & _
       "Ex. HKEY_CURRENT_USER\AppEvents", vbExclamation, "Invalid Key!"
    End Select

    If EnablePrivilege(SE_RESTORE_NAME) Then
        RegOpenKeyEx lHKey, sKey, 0&, KEY_ALL_ACCESS, hKey
        RetVal = RegRestoreKey(hKey, sPath, REG_FORCE_RESTORE)
        RegCloseKey hKey
    End If

    If RetVal = 0 Then
        Deploy_Key = True
    End If

End Function

Private Function String_Convert(ByVal sName As String) As String

'//found this on psc, but you should make something better
Dim LookUpTable(0 To 255) As Byte
Dim i                     As Integer
Dim c                     As Long
Dim newstr()              As Byte

    For i = 0 To 255
        LookUpTable(i) = i
    Next i
    LookUpTable(92) = 45
    LookUpTable(45) = 92
    newstr() = sName

    For i = 0 To UBound(newstr) Step 2
        If LookUpTable(newstr(i)) <> 0 Then
            newstr(c) = LookUpTable(newstr(i))
            c = c + 2
        End If
    Next i
    ReDim Preserve newstr(c)
    String_Convert = newstr()

End Function

Public Sub Restore_Key()

Dim sPath As String
Dim hKey  As Long

    '//restore key via monitor sub
    sPath = App.Path & "\Recovery\"
    sPath = sPath & sModmname & ".kbs"

    If Not FileExists(sPath) Then
        MsgBox "Can Not Restore Key!" & vbNewLine & "Backup Key File is Missing!", vbExclamation, "No .kbs File!"
    Else
        If EnablePrivilege(SE_RESTORE_NAME) Then
            RegOpenKeyEx lModlhkey, sModskey, 0&, KEY_ALL_ACCESS, hKey
            RegRestoreKey hKey, sPath, REG_FORCE_RESTORE
            RegCloseKey hKey
        End If
    End If

End Sub
