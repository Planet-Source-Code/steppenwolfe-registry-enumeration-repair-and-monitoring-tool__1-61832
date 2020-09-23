Attribute VB_Name = "mRegEnum"
Option Explicit

'***cudos out to http://vb-helper.com/ for the enum example upon which getkeyinfo was based
'***(but entirely rewritten top to bottom - as you should do with this example ;o)
'***added some new features that demonstrate a simple monitoring technique, if you
'***plan to develop this, consider writing the back end in c/c++ as a library, as the
'***speed will be much improved. Alternatively, you could write all the reg functions
'***into an activex library, this would also greatly improve on performance
'***ToDo's (for you) might include a better timer, testing for optimal search tolerance spec
'***better UI, a lot more features, and maybe integration into tray/explorer toolbar
'***anyways - feedback is always appreciated, and REMEMBER TO VOTE!!! (I need the software :o}

Private Const ERROR_SUCCESS                       As Long = 0
Public Const HKEY_CLASSES_ROOT                    As Long = &H80000000
Public Const HKEY_CURRENT_CONFIG                  As Long = &H80000005
Public Const HKEY_CURRENT_USER                    As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE                   As Long = &H80000002
Public Const HKEY_USERS                           As Long = &H80000003
Private Const STANDARD_RIGHTS_ALL                 As Long = &H1F0000
Private Const KEY_QUERY_VALUE                     As Long = &H1
Private Const KEY_SET_VALUE                       As Long = &H2
Private Const KEY_CREATE_SUB_KEY                  As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS              As Long = &H8
Private Const KEY_NOTIFY                          As Long = &H10
Private Const KEY_CREATE_LINK                     As Long = &H20
Private Const SYNCHRONIZE                         As Long = &H100000
Private Const KEY_ALL_ACCESS                      As Double = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const ERROR_NO_MORE_ITEMS                 As Long = 259
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_EXPAND_SZ = 2
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_NONE = 0
Private Const REG_RESOURCE_LIST = 8
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Private Const REG_SZ = 1

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long
                                                                                
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, _
                                                                            ByVal dwIndex As Long, _
                                                                            ByVal lpName As String, _
                                                                            ByVal cbName As Long) As Long
                                                                            
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, _
                                                                                ByVal dwIndex As Long, _
                                                                                ByVal lpValueName As String, _
                                                                                lpcbValueName As Long, _
                                                                                ByVal lpReserved As Long, _
                                                                                lpType As Long, _
                                                                                lpData As Byte, _
                                                                                lpcbData As Long) As Long
                                                                                
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, _
                                                                                  ByVal lpSubKey As String, _
                                                                                  ByVal lpValue As String, _
                                                                                  lpcbValue As Long) As Long

Public lCount                   As Long
Public aKeyArr()                As String
Public bDimn                    As Boolean
Public aBase()                  As String
Public aComp()                  As String
Public aDiff()                  As String
Private lmCount                 As Long
Public cValues                  As Boolean
Private lValCount               As Long
Private lmValCount              As Long
Public aValArr()                As String
Private Const DB_MK             As String = "  "
Public Const DL_MK              As String = " = "

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Purpose    :Create an array of key/value names
'SubSet     :Snapshot and Compare Sub/Predator DCA
'Ref/Call   :advapi32
'Ret/Out    :aKeyArr()
'Author     :John Underhill 07/21/05
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Public Sub GetKeyInfo(ByVal lKey As Long, _
                      ByVal sKey As String)

Dim cSubKey                     As Collection
Dim lKeyNum                     As Long
Dim sKeyName                    As String
Dim hKey                        As Long
Dim sVal                        As String
'//start value vars
Dim cKeyVal                     As Collection
Dim lValNum                     As Long
Dim lValType                    As Long
Dim sValName                    As String
Dim sValStr                     As String
Dim bValData(1 To 1024)         As Byte
Dim lValRet                     As Long
Dim lRetData                    As Long
Dim l                           As Long
Const BF_SZ                     As Long = 256
Const NK_BF                     As Long = 1024

        '//create collection
    Set cSubKey = New Collection
    Set cKeyVal = New Collection
        '//initial dimn check
    If Not bDimn Then
        If Not cValues Then
            ReDim aKeyArr(0 To 1023)
            lCount = 0
            lmCount = 0
        Else
            ReDim aValArr(0 To 1023)
            lValCount = 0
            lmValCount = 0
        End If
        bDimn = True
    End If
        '//open key
    If RegOpenKeyEx(lKey, sKey, 0&, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then
        Exit Sub
    End If
            '//optionally recurse values and add to 2nd collection
        If cValues Then
            Erase bValData
            lValNum = 0
            sValStr = vbNullString
            sValName = Space$(BF_SZ)
            lValRet = BF_SZ
            lRetData = NK_BF
            cKeyVal.Add sKey
            'Debug.Print sKey
                '//enum values
                
            While RegEnumValue(hKey, lValNum, sValName, lValRet, 0, _
            lValType, bValData(1), lRetData) = ERROR_SUCCESS
 
                '//apply a filter to data collection. for app changes/additions, string and dword
                '//are all you should be concerned with. processing large binary entries is time
                '//consuming, and not really necessary. you could use a second filter to adjust
                '//string size, and filter on size specific entries. The first time I ran this
                '//with no filters on hklm branch, I had to kill it at 2 million plus entries,
                '//was only a minute of time but.. left the debugs you could see what I am checking
            If Not lRetData > 1024 Then        '//filter 1: limit data
                Select Case lValType            '//filter 2: string and dword only
                    Case REG_DWORD
                        sValStr = "&H" & _
                        Format$(Hex$(bValData(4)), "00") & _
                        Format$(Hex$(bValData(3)), "00") & _
                        Format$(Hex$(bValData(2)), "00") & _
                        Format$(Hex$(bValData(1)), "00")
                        Select Case 0
                            Case lRetData   '//filter 3: format based on status
                                cKeyVal.Add sKey & DB_MK & Left$(sValName, lValRet) '//empty value
                                'Debug.Print sKey & DB_MK & Left$(sValName, lValRet)
                            Case lValRet
                            '//empty key - do nothing
                            Case Else                                               '//key/val/data
                                cKeyVal.Add sKey & DB_MK & Left$(sValName, lValRet) + "=" + Left$(sValStr, lRetData - 1)
                                'Debug.Print sKey & DB_MK & Left$(sValName, lValRet) + "=" + Left$(sValStr, lRetData - 1)
                        End Select
                        
                    Case REG_SZ
                        For l = 1 To lRetData - 1
                            sValStr = sValStr & Chr$(bValData(l))
                        Next l
                        Select Case 0
                            Case lRetData
                                cKeyVal.Add sKey & DB_MK & Left$(sValName, lValRet) '//empty value
                                'Debug.Print sKey & DB_MK & Left$(sValName, lValRet)
                            Case lValRet
                            '//empty key - do nothing
                            Case Else
                                cKeyVal.Add sKey & DB_MK & Left$(sValName, lValRet) + "=" + Left$(sValStr, lRetData - 1)
                                'Debug.Print sKey & DB_MK & Left$(sValName, lValRet) + "=" + Left$(sValStr, lRetData - 1)
                        End Select
                            '//you can add these, format is the same
                            '//for string types, add the chr conversion
                            '//dword use hex conversion
                            '//for other resources, simply drop as is
                            '//into the comparison file, it will still detect
                            '//changes to the resource
                    'Case REG_BINARY                        '//bin
                    'Case REG_DWORD_BIG_ENDIAN              '//dwd
                    'Case REG_DWORD_LITTLE_ENDIAN           '//dwd
                    'Case REG_EXPAND_SZ                     '//str
                    'Case REG_FULL_RESOURCE_DESCRIPTOR      '//str
                    'Case REG_LINK                          '//str
                    'Case REG_MULTI_SZ                      '//str
                    'Case REG_NONE                          '//nll
                    'Case REG_RESOURCE_LIST                 '//str
                    'Case REG_RESOURCE_REQUIREMENTS_LIST    '//str
                End Select
            End If
                Erase bValData
                sValStr = vbNullString
                sValName = vbNullString
                lValNum = lValNum + 1
                'sValStr = Space$(NK_BF)
                sValName = Space$(BF_SZ)
                lValRet = BF_SZ
                lRetData = NK_BF
            DoEvents
            Wend            '//loop values
        End If
        '//start key enum
    lKeyNum = 0
    Do
        sKeyName = Space$(BF_SZ)
        If RegEnumKey(hKey, lKeyNum, sKeyName, BF_SZ) <> ERROR_SUCCESS Then
            Exit Do
        End If
        lKeyNum = lKeyNum + 1
        sKeyName = Left$(sKeyName, InStr(sKeyName, vbNullChar) - 1)
        cSubKey.Add sKeyName
            '//iterate through branch key values
        DoEvents
    Loop
        '//close
        RegCloseKey hKey
        
        '//recurse through keys
    For lKeyNum = 1 To cSubKey.Count
        '//dimensions check
        If Not cValues Then
            If lCount > UBound(aKeyArr()) Then
                ReDim Preserve aKeyArr(0 To UBound(aKeyArr()) + NK_BF)
            End If
        End If
            '//add to array
            '//filter for first pass
        If Not LenB(sKey) = 0 Then
            sVal = sKey & Chr(92) & cSubKey(lKeyNum)
        Else
            sVal = cSubKey(lKeyNum)
        End If
        If Not cValues Then
            aKeyArr(lCount) = sVal
        End If
            lCount = lCount + 1
            lmCount = lmCount + 1
        '//recurse
        GetKeyInfo lKey, sVal
        DoEvents
    Next lKeyNum
    
    If cValues Then
        For lValNum = 1 To cKeyVal.Count
                '//dimensions check
            If lValCount > UBound(aValArr()) Then
                ReDim Preserve aValArr(0 To UBound(aValArr()) + NK_BF)
            End If
                '//add values to array
            aValArr(lValCount) = cKeyVal(lValNum)
            lValCount = lValCount + 1
            lmValCount = lmValCount + 1
            DoEvents
        Next lValNum
    End If

End Sub

Public Sub Dyn_Compare(ByRef aBase() As String, _
                       ByRef aComp() As String, _
                       ByVal lResInd As Long)

Dim lBase               As Long
Dim lComp               As Long
Dim lIndex              As Long
Dim lLow                As Long
Dim lHigh               As Long
Dim lMax                As Long
Dim lgCount              As Long
Dim lReflex             As Long
Dim bMatch              As Boolean
Const RD_AD             As Long = 16

ReDim aDiff(0 To 0)
bMatch = False
lMax = UBound(aBase)
lReflex = 0

    '*** lResInd is the base tolerance for the search loop, no need iterating through
    '*** entire comparison array, this is processing excessive for two similar arrays
    '*** The comparison routine responds to new keys by increasing/decreasing
    '*** the search boundaries based on match frequency (lReflex); if a new key is found,
    '*** it increases the low/high search boundaries, if after iterating through 100 keys
    '*** with no new keys found, it narrows the search window..
    
    For lComp = 0 To UBound(aComp)                      '//comparison file
        '//set lower search boundry
        If Not lComp < (lResInd + lReflex) Then
            lLow = lComp - (lResInd + lReflex)
        Else
            lLow = 0
        End If
        '//set upper search boundry
        If Not lComp + (lResInd + lReflex) > lMax Then
            lHigh = lComp + (lResInd + lReflex)
        Else
            lHigh = lMax
        End If
        '//start comparing arrays
        For lBase = lLow To lHigh                       '//base file
            lgCount = lgCount + 1
            If lgCount > 100 And lReflex > 1 Then
                lReflex = lReflex - 1                   '//decrement search boundaries
            End If
            If aBase(lBase) = aComp(lComp) Then
                bMatch = True
                Exit For
            End If
        Next lBase
            If Not bMatch Then
                lgCount = 0
                If Not LenB(aComp(lComp)) = 0 Then      '//filter blanks
                    If lIndex > UBound(aDiff()) Then    '//dimn check
                        ReDim Preserve aDiff(0 To UBound(aDiff()) + RD_AD)
                    End If
                    aDiff(lIndex) = aComp(lComp)        '//add to diff array
                    'Debug.Print aDiff(lIndex)
                    lIndex = lIndex + 1
                    lReflex = lReflex + 1               '//increment search boundaries
                End If
            End If
            bMatch = False
            DoEvents
    Next lComp
    Arr_Cleanup
    ReDim Preserve aDiff(0 To lIndex)
    
End Sub

Private Sub Arr_Cleanup()
Erase aComp
Erase aBase
Erase aKeyArr
End Sub

Public Function FileExists(ByVal sName As String) As Boolean

    If LenB(Dir(sName)) Then
        FileExists = True
    End If

End Function

Public Sub Reset_Dimensions()

        '//remove blank entries
    If Not cValues Then
        ReDim Preserve aKeyArr(0 To lmCount)
    Else
        ReDim Preserve aValArr(0 To lmValCount)
    End If

End Sub

Public Function Auto_Set_Tolerance() As Long

Dim lStart As Long
Dim lPerc  As Double

    If UBound(aComp) > UBound(aBase) Then
        lStart = UBound(aComp) - UBound(aBase)  '//add the difference between arrays
        lPerc = UBound(aComp) * 0.0025          '//up tolerance factor by 25 per 1000 entries
        lStart = lStart * lPerc                 '//add the two
    End If

    If Not lStart < 100 Then                    '//sanity check: set minimum
        Auto_Set_Tolerance = lStart
    Else
        Auto_Set_Tolerance = 100
    End If

End Function

Public Function Return_Values(lKey As Long, _
                              ByVal sKey As String) As String
                              
Dim hKey                        As Long
Dim lValNum                     As Long
Dim lValType                    As Long
Dim sValName                    As String
Dim sValStr                     As String
Dim bValData(1 To 1024)         As Byte
Dim lValRet                     As Long
Dim lRetData                    As Long
Dim sRetStr                     As String
Dim l                           As Long
Const BF_SZ                     As Long = 256
Const NK_BF                     As Long = 1024

    If RegOpenKeyEx(lKey, sKey, 0&, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then
        Exit Function
    End If
    
    sValStr = vbNullString
    sValName = vbNullString
    lValNum = 0
    sValName = Space$(BF_SZ)
    lValRet = BF_SZ
    lRetData = NK_BF
                '//enum values
    While RegEnumValue(hKey, lValNum, sValName, lValRet, 0, _
    lValType, bValData(1), lRetData) = ERROR_SUCCESS

        If Not lRetData > 1024 Then     '//filter 1: data
            Select Case lValType        '//filter 2: string and dword
                Case REG_DWORD
                sValStr = "&H" & _
                Format$(Hex$(bValData(4)), "00") & _
                Format$(Hex$(bValData(3)), "00") & _
                Format$(Hex$(bValData(2)), "00") & _
                Format$(Hex$(bValData(1)), "00")
                Select Case 0
                    Case lRetData   '//add to string with line and value delimiters
                        sRetStr = sRetStr & Left$(sValName, lValRet) & vbNewLine  '//empty value
                        'Debug.Print sRetStr = sRetStr & Left$(sValName, lValRet) & vbNewLine
                    Case lValRet
                        '//empty key - do nothing
                    Case Else                                               '//key/val/data
                        sRetStr = sRetStr & Left$(sValName, lValRet) + DL_MK + Left$(sValStr, lRetData - 1) & vbNewLine
                        'Debug.Print sRetStr & Left$(sValName, lValRet) + DL_MK + Left$(sValStr, lRetData - 1) & vbNewLine
                End Select
                        
            Case REG_SZ
                For l = 1 To lRetData - 1
                    sValStr = sValStr & Chr$(bValData(l))
                Next l
                Select Case 0
                    Case lRetData
                        sRetStr = sRetStr & Left$(sValName, lValRet) & vbNewLine
                        'Debug.Print sRetStr = sRetStr & Left$(sValName, lValRet) & vbNewLine
                            '//no data
                    Case lValRet
                        '//do nothing
                    Case Else
                        sRetStr = sRetStr & Left$(sValName, lValRet) + DL_MK + Left$(sValStr, lRetData - 1) & vbNewLine
                        'Debug.Print sRetStr & Left$(sValName, lValRet) + DL_MK + Left$(sValStr, lRetData - 1) & vbNewLine
                End Select
            End Select
        End If

        Erase bValData
        sValStr = vbNullString
        sValName = vbNullString
        lValNum = lValNum + 1
        sValName = Space$(BF_SZ)
        lValRet = BF_SZ
        lRetData = NK_BF
        DoEvents
    Wend            '//loop values
    
    RegCloseKey hKey
    Return_Values = sRetStr
        
End Function

