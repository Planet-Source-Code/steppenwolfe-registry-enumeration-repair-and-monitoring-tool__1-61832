Attribute VB_Name = "mMonitor"
Option Explicit
'//if you re-use the reg routines, add the proper error handling, (I pulled this from an activex control
'//with global handler), so to save time, I put the resume next statements, but look up event
'//specific errors and add select case to manage errors properly ;o)
'//I don't have time to comment all this, for explainations of api, go to allapi.com..
'//John Underhill 23-07-2005

Private Type FILETIME
    dwLowDateTime                              As Long
    dwHighDateTime                             As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength                                    As Long
    lpSecurityDescriptor                       As Long
    bInheritHandle                             As Boolean
End Type

Private Type cRegValue
    Key                                        As String
    DataType                                   As Reg_Type
    Value                                      As Variant
    sName                                      As Variant
End Type

Private Enum HKEY_Type
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

#If False Then
Private HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USERS, HKEY_PERFORMANCE_DATA, HKEY_CURRENT_CONFIG
Private HKEY_DYN_DATA
#End If

Private Enum Reg_Type
    REG_NONE = 0
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_DWORD_LITTLE_ENDIAN = 4
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7
    REG_RESOURCE_LIST = 8
End Enum

#If False Then
Private REG_NONE, REG_SZ, REG_EXPAND_SZ, REG_BINARY, REG_DWORD, REG_DWORD_LITTLE_ENDIAN, REG_DWORD_BIG_ENDIAN, REG_LINK, REG_MULTI_SZ
Private REG_RESOURCE_LIST
#End If

Private Const KEY_ALL_ACCESS               As Long = &HF003F
Private Const KEY_CREATE_LINK              As Long = &H20
Private Const KEY_CREATE_SUB_KEY           As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS       As Long = &H8
Private Const KEY_NOTIFY                   As Long = &H10
Private Const KEY_QUERY_VALUE              As Long = &H1
Private Const KEY_SET_VALUE                As Long = &H2
Private Const KEY_WRITE                    As Long = &H20006
Private Const ERROR_NONE                   As Integer = 0
Private Const ERROR_MORE_DATA              As Integer = 234
Private Const ERROR_NO_MORE_ITEMS          As Integer = 259

Public bMonitor                            As Boolean
Public sModmstr                            As String
Public sModmname                           As String
Public lModlhkey                           As Long
Public sModskey                            As String
Public sModApp                             As String
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                      ByVal lpValueName As String, _
                                                                                      ByVal lpReserved As Long, _
                                                                                      lpType As Long, _
                                                                                      lpData As Any, _
                                                                                      lpcbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, _
                                                                                    ByVal lpValueName As String) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, _
                                                                                    ByVal lpSubKey As String, _
                                                                                    ByVal Reserved As Long, _
                                                                                    ByVal lpClass As String, _
                                                                                    ByVal dwOptions As Long, _
                                                                                    ByVal samDesired As Long, _
                                                                                    lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                                                                                    phkResult As Long, _
                                                                                    lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                                                  ByVal lpValueName As String, _
                                                                                  ByVal Reserved As Long, _
                                                                                  ByVal dwType As Long, _
                                                                                  lpData As Any, _
                                                                                  ByVal cbData As Long) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                     Source As Any, _
                                                                     ByVal length As Long)

Private Function KeyExist(Key As HKEY_Type, _
                          sSubKey As String) As Boolean

Dim hKey        As Long
Dim RetVal      As Long

    RetVal = RegOpenKeyEx(Key, sSubKey, 0, KEY_QUERY_VALUE, hKey)

    If RetVal = ERROR_NONE Then
        KeyExist = True
    Else
        KeyExist = False
    End If
    RegCloseKey hKey

End Function

Private Function DeleteKey(Key As HKEY_Type, _
                          sSubKey As String) As Boolean

Dim RetVal  As Long

    RetVal = RegDeleteKey(Key, sSubKey)
    If RetVal = ERROR_NONE Then
        DeleteKey = True
    Else
        DeleteKey = False
    End If

End Function

Private Function ReadMulti(Key As HKEY_Type, _
                           Subkey As String, _
                           sName As String) As String

Dim hKey        As Long
Dim RetVal      As Long
Dim sBuffer     As String
Dim length      As Long
Dim resBinary() As Byte
Dim resString   As String

On Error Resume Next

    RetVal = RegOpenKeyEx(Key, Subkey, 0, KEY_ALL_ACCESS, hKey)
    If RetVal <> ERROR_NONE Then
        RegCloseKey (hKey)
    Else
        length = 1024
        ReDim resBinary(0 To length - 1) As Byte

        RetVal = RegQueryValueEx(hKey, sName, 0, REG_MULTI_SZ, resBinary(0), length)

        If RetVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To length - 1) As Byte
            RetVal = RegQueryValueEx(hKey, sName, 0, REG_MULTI_SZ, resBinary(0), length)
        End If

        If RetVal = ERROR_NONE Then
            resString = Space$(length - 2)
            CopyMemory ByVal resString, resBinary(0), length - 2
            sBuffer = resString
            If Len(TrimNull(sBuffer)) > 0 Then
                ReadMulti = resString
            End If
        End If

        RetVal = RegCloseKey(hKey)

On Error GoTo 0

    End If

End Function

Private Function WriteMulti(Key As HKEY_Type, _
                           Subkey As String, _
                           sName As String, _
                           sData As String) As Boolean

Dim hKey        As Long
Dim RetVal      As Long
Dim deposit     As Long
Dim secattr     As SECURITY_ATTRIBUTES

On Error Resume Next

    secattr.nLength = Len(secattr)
    secattr.lpSecurityDescriptor = 0
    secattr.bInheritHandle = 1
    
    RetVal = RegCreateKeyEx(Key, Subkey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If RetVal <> ERROR_NONE Then
        WriteMulti = False
        Exit Function
    End If

    RetVal = RegSetValueEx(hKey, sName, 0, REG_MULTI_SZ, ByVal sData, Len(sData))
    
    If RetVal <> ERROR_NONE Then
        WriteMulti = False
        Exit Function
    End If
    
    RetVal = RegCloseKey(hKey)
    WriteMulti = True

On Error GoTo 0

End Function

Public Function TrimNull(Item As String) As String

Dim pos As Integer

        pos = InStr(Item, Chr$(0))
        If pos Then Item = Left$(Item, pos - 1)
        TrimNull = Item
        
End Function


                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    '>               START MONITORING ENGINE              <
                    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Public Sub Start_Monitor(mStr As String, _
                         ByVal mName As String, _
                         ByVal lHKey As Long, _
                         ByVal sKey As String)

    sModmstr = mStr                                         '//place the vars in module memory
    sModmname = mName                                       '//so we needn't pass it between subs
    lModlhkey = lHKey
    sModskey = sKey

    sModApp = "Software\" & App.FileDescription & Chr$(92) & mName

    Backup_Key                                              '//create a binary image of key

    If Not LenB(mStr) > 2500 Then                           '//check size of string
        If Not KeyExist(HKEY_CURRENT_USER, sModApp) Then    '//check if key already exists
            If Not WriteMulti(HKEY_CURRENT_USER, sModApp, mName, mStr) Then
                MsgBox "App Key could not be created!" & vbNewLine & _
       "Check your User Rights!", vbExclamation, "Check Permissions!"
                Exit Sub
            End If
        End If
    End If

    bMonitor = True
    Start_Timer

End Sub

Public Sub Start_Timer()

Dim iInterval As Integer
Dim sComp     As String
Dim sBase     As String
Dim aMBase()  As String
Dim aMComp()  As String
Dim lBase     As Long
Dim lComp     As Long
Dim lLow      As Long
Dim lHigh     As Long
Dim lMax      As Long
Dim lResInd   As Long
Dim bMatch    As Boolean
Dim mLen      As Long

bMatch = False
lResInd = 5

On Error Resume Next

    With frmTest
        If Not .txtInterval.Text = vbNullString Then        '//default polling interval
            iInterval = CInt(.txtInterval.Text)
        Else
            iInterval = 5
        End If
    End With

    Do While bMonitor
        sComp = ReadMulti(lModlhkey, sModApp, sModmname)    '//get vals and add to array
        sBase = Return_Values(lModlhkey, sModskey)
        sComp = Left$(sComp, Len(sComp) - 1)                '//trim the null char

        If LenB(sComp) = 0 Then
                                                            '//half ass user proofing, needs much more
            MsgBox "Comparison key has no values!" & vbNewLine & _
       "Aborting Monitor!", vbExclamation, "Check Path!"
            bMonitor = False
            Exit Sub
        ElseIf LenB(sBase) = 0 Then
            MsgBox "Could Not Read the key values specified!" & vbNewLine & _
       "Check Path and if Key contains Values!", vbExclamation, "Check Values!"
            bMonitor = False
            Exit Sub
        End If

        aMComp = Split(sComp, vbNewLine)
        aMBase = Split(sBase, vbNewLine)

        If UBound(aMComp) <> (UBound(aMBase) - 1) Then      '//react if the value count changes
            MsgBox "A New Value has been Added to the Sub Key!" & vbNewLine & _
       "Restarting Monitor to Compensate!", vbExclamation, "New Value!"
            bMonitor = False
            Start_Monitor Return_Values(lModlhkey, sModskey), sModmname, lModlhkey, sModskey
        End If

        With frmTest
            If .chkDifferential.Value Then                  '//add 1 sec for every 10 entries
                mLen = UBound(aMBase)
                If Not mLen < 10 Then
                    iInterval = (mLen / 10) + 2
                Else
                    iInterval = 5
                End If
            End If
        End With

        Wait_Timer iInterval                                '//cheapie wait timer

        lMax = UBound(aMBase)

        For lComp = 0 To UBound(aMComp)                     '//comparison file
            '//set lower search boundry
            If Not lComp < (lResInd) Then
                lLow = lComp - (lResInd)
            Else
                lLow = 0
            End If
            '//set upper search boundry
            If Not (lComp + lResInd) > lMax Then
                lHigh = lComp + lResInd
            Else
                lHigh = lMax
            End If
            '//start comparing arrays
            For lBase = lLow To lHigh                           '//base file
                If aMBase(lBase) = aMComp(lComp) Then
                    bMatch = True
                    Exit For
                End If
            Next lBase
            If Not bMatch Then
                If Not LenB(aMComp(lComp)) = 0 Then             '//filter blanks
                    bMonitor = False
                    User_Notify aMBase(lComp), aMComp(lComp)    '//call user notify and stop
                End If
            End If
            bMatch = False
            DoEvents
        Next lComp
    Loop

On Error GoTo 0

End Sub

Private Sub User_Notify(ByVal sNewVal As String, _
                        ByVal sOldVal As String)

Dim iChoice  As Integer
Dim sNval    As String
Dim sNewData As String
Dim sOldData As String

On Error GoTo Skip

    sNval = Left$(sNewVal, InStr(sNewVal, DL_MK) - 1)
    sNewData = Mid$(sNewVal, InStr(sNewVal, DL_MK) + 1)
    sOldData = Mid$(sOldVal, InStr(sOldVal, DL_MK) + 1)

    iChoice = MsgBox("The Registry Value: " & sNval & " has changed!" & vbNewLine _
    & "The New Value is: " & sNewData & vbNewLine & "The Original Value was: " & _
    sOldData & vbNewLine & "Click YES to Accept this Change, or NO TO Revert.", vbYesNo, "Value has Changed!!")

    If iChoice = 6 Then             '//if accept change, reset image and restart
        If Not DeleteKey(HKEY_CURRENT_USER, sModApp) Then
            MsgBox "Could Not Reset App Key!" & vbNewLine & _
       "Check your User Rights!", vbExclamation, "Check Permissions!"
            Exit Sub
        End If
        Start_Monitor Return_Values(lModlhkey, sModskey), sModmname, lModlhkey, sModskey
        bMonitor = True
    ElseIf iChoice = 7 Then         '//if deny change, restore the original key
        Restore_Key
    End If
    bMonitor = True

Skip:

End Sub

Public Sub Wait_Timer(ByVal lSecs As Long)

Dim l As Long

    '//wait timer
    For l = 1 To lSecs * 10
        Sleep 100
        DoEvents
    Next l

End Sub
