VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Key Snapshot"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   8925
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmImage 
      Caption         =   "Save and Restore"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   37
      Top             =   7920
      Width           =   8685
      Begin VB.CommandButton cmdCreateRestore 
         Caption         =   "Create"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6180
         TabIndex        =   41
         Top             =   330
         Width           =   1125
      End
      Begin VB.CommandButton cmdDeployRestore 
         Caption         =   "Deploy"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7380
         TabIndex        =   40
         Top             =   330
         Width           =   1125
      End
      Begin VB.TextBox txtRestore 
         Height          =   285
         Left            =   150
         TabIndex        =   39
         Text            =   "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\CurrentVersion\Run"
         Top             =   420
         Width           =   5865
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Create a Binary Image of a Key and Restore"
         Height          =   195
         Left            =   150
         TabIndex        =   38
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame frmMonitor 
      Caption         =   "Key Monitor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   120
      TabIndex        =   24
      Top             =   6090
      Width           =   8685
      Begin VB.CheckBox chkDifferential 
         Caption         =   "Auto Calculate Poll Interval"
         Height          =   225
         Left            =   180
         TabIndex        =   35
         Top             =   1350
         Width           =   2745
      End
      Begin VB.TextBox txtInterval 
         Height          =   285
         Left            =   1650
         TabIndex        =   31
         Text            =   "5"
         Top             =   960
         Width           =   405
      End
      Begin VB.TextBox txtMonitor 
         Height          =   315
         Left            =   150
         TabIndex        =   27
         Text            =   "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main"
         Top             =   600
         Width           =   8325
      End
      Begin VB.CommandButton cmdStopMonitor 
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7380
         TabIndex        =   26
         Top             =   1260
         Width           =   1125
      End
      Begin VB.CommandButton cmdStartMonitor 
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6180
         TabIndex        =   25
         Top             =   1260
         Width           =   1125
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Seconds"
         Height          =   195
         Left            =   2100
         TabIndex        =   33
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Polling Interval:"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   1020
         Width           =   1365
      End
      Begin VB.Label lblMonitor 
         AutoSize        =   -1  'True
         Caption         =   "Monitor Values Changed in this SubKey"
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Top             =   420
         Width           =   3375
      End
   End
   Begin VB.Frame frmInstruct 
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5865
      Left            =   120
      TabIndex        =   12
      Top             =   210
      Width           =   2295
      Begin VB.Label lblVote 
         AutoSize        =   -1  'True
         Caption         =   "Click to Vote"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   450
         TabIndex        =   42
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Monitor/Restore"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   30
         Top             =   3630
         Width           =   1605
      End
      Begin VB.Label Label9 
         Caption         =   $"frmTest.frx":0000
         Height          =   1575
         Left            =   90
         TabIndex        =   29
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Start the Compare Routine, it should identify all the keys you have added, (and quickly!)"
         Height          =   1065
         Left            =   90
         TabIndex        =   18
         Top             =   2580
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Step 3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   2370
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Add several arbitrary keys to the selected registry branch using regedit, bury them as deep as you like.."
         Height          =   1035
         Left            =   90
         TabIndex        =   16
         Top             =   1350
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Step 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   1140
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Select the test branch, and create the Base Snapshot."
         Height          =   675
         Left            =   90
         TabIndex        =   14
         Top             =   510
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Step 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame fmData 
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   2580
      TabIndex        =   1
      Top             =   210
      Width           =   6225
      Begin MSComctlLib.ListView lstOutput 
         Height          =   2475
         Left            =   120
         TabIndex        =   5
         Top             =   270
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   4366
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Key"
            Object.Width           =   10585
         EndProperty
      End
      Begin VB.Label lblOperation 
         AutoSize        =   -1  'True
         Caption         =   "Current Operation: Idle.."
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   2790
         Width           =   2145
      End
   End
   Begin VB.Frame fmControls 
      Caption         =   "Enumeration Controls"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   2580
      TabIndex        =   0
      Top             =   3360
      Width           =   6225
      Begin VB.CheckBox ChkAutoTolerance 
         Caption         =   "Auto Calculate Ideal Tolerance"
         Height          =   225
         Left            =   2100
         TabIndex        =   36
         Top             =   1530
         Width           =   3045
      End
      Begin VB.CheckBox chkValues 
         Caption         =   "Enumerate Sub Key Values"
         Height          =   225
         Left            =   180
         TabIndex        =   34
         Top             =   2340
         Width           =   2745
      End
      Begin VB.CheckBox chkSubkey 
         Caption         =   "Search SubKey"
         Height          =   225
         Left            =   1380
         TabIndex        =   22
         Top             =   930
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.TextBox txtTolerance 
         Height          =   285
         Left            =   1380
         TabIndex        =   21
         Text            =   "10"
         Top             =   1470
         Width           =   495
      End
      Begin VB.TextBox txtSubkey 
         Height          =   255
         Left            =   1380
         TabIndex        =   20
         Text            =   "SOFTWARE\Microsoft\Windows\CurrentVersion"
         Top             =   600
         Width           =   4665
      End
      Begin VB.CheckBox chkSaveResults 
         Caption         =   "Save Results to File"
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   2040
         Width           =   2085
      End
      Begin VB.OptionButton Option1 
         Caption         =   "HKCC"
         Height          =   225
         Index           =   4
         Left            =   150
         TabIndex        =   10
         Top             =   1740
         Width           =   885
      End
      Begin VB.OptionButton Option1 
         Caption         =   "HKU"
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   9
         Top             =   1470
         Width           =   765
      End
      Begin VB.OptionButton Option1 
         Caption         =   "HKLM"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   8
         Top             =   1170
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "HKCU"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   870
         Width           =   915
      End
      Begin VB.OptionButton Option1 
         Caption         =   "HKCR"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   570
         Width           =   885
      End
      Begin VB.CommandButton cmdSnapshot 
         Caption         =   "Snap Shot"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3720
         TabIndex        =   3
         Top             =   2190
         Width           =   1125
      End
      Begin VB.CommandButton cmdCompare 
         Caption         =   "Compare"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4920
         TabIndex        =   2
         Top             =   2190
         Width           =   1125
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Set Search Tolerance (min 10)"
         Height          =   195
         Left            =   1380
         TabIndex        =   23
         Top             =   1260
         Width           =   2640
      End
      Begin VB.Label Label7 
         Caption         =   "Select a Test Branch:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   19
         Top             =   300
         Width           =   2145
      End
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'The goal of this program is to detect relatively small changes
'to a file or database. The search routine has a tolerance factor
'that could be incremented if searching for a larger number of
'changes. I have used the registry for this example, but the
'comparison routine could be applied to any two similar ordered files.

Option Explicit

Private lHKey  As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                                                               ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, _
                                                                               ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, _
                                                                               ByVal nShowCmd As Long) As Long


Private Sub chkValues_Click()

    cValues = CBool(chkValues.Value)
    If cValues Then
        txtTolerance.Text = "100"
        With lstOutput
            .ListItems.Clear
            .AllowColumnReorder = True
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Key and Values", 5915
        End With
    Else
        txtTolerance.Text = "10"
        With lstOutput
            .ListItems.Clear
            .AllowColumnReorder = True
            .ColumnHeaders.Clear
            .ColumnHeaders.Add 1, , "Key", 5915
        End With
    End If

End Sub

Private Sub cmdCompare_Click()

Dim sPath    As String
Dim sData    As String
Dim l        As Long
Dim Item     As ListItem
Dim lRes     As Long
Dim sSubText As String
Dim sSubKey  As String

    If chkSubkey Then                               '//check for subkey
        sSubKey = txtSubkey.Text                    '//subkey string
        If Not InStrB(txtSubkey.Text, "\") Then     '//extract for filename check from solidus
            sSubText = Mid$(txtSubkey.Text, InStrRev(txtSubkey.Text, "\") + 1)
        Else
            sSubText = txtSubkey.Text               '//if there is no solidus
        End If
    End If

    If chkValues Then                               '//check for val enum option
        sSubText = sSubText & "Val"
    End If

    Select Case True                                '//select key and name the path
    Case Option1(0)
        lHKey = HKEY_CLASSES_ROOT
        sPath = App.Path & "\hkcrsnap" & sSubText & ".txt"
    Case Option1(1)
        lHKey = HKEY_CURRENT_USER
        sPath = App.Path & "\hkcusnap" & sSubText & ".txt"
    Case Option1(2)
        lHKey = HKEY_LOCAL_MACHINE
        sPath = App.Path & "\hklmsnap" & sSubText & ".txt"
    Case Option1(3)
        lHKey = HKEY_USERS
        sPath = App.Path & "\hkusnap" & sSubText & ".txt"
    Case Option1(4)
        lHKey = HKEY_CURRENT_CONFIG
        sPath = App.Path & "\hkccsnap" & sSubText & ".txt"
    End Select

    If Not FileExists(sPath) Then                   '//check for base file
        lblOperation.Caption = "Comparison Aborted.."
        MsgBox "Please take a Snapshot of this Key First!", vbExclamation, "No SnapShot!"
        Exit Sub
    Else
        lblOperation.Caption = "Preparing Comparison Snapshot.."
        Open sPath For Binary As #1                 '//load base file
        sData = Input$(LOF(1), 1)
        Close #1
    End If

    If Not txtTolerance.Text = vbNullString Then    '//get tolerance num
        lRes = CLng(txtTolerance.Text)
    Else
        lRes = 10
    End If

    bDimn = False                                   '//reset aKeyArr
    GetKeyInfo lHKey, sSubKey                       '//get the key array
    Reset_Dimensions

    aBase = Split(sData, vbNewLine)                 '//assign base and compare
    ReDim aComp(0 To 0)
    If Not cValues Then
        aComp() = aKeyArr()
    Else
        aComp() = aValArr()
    End If

    lblOperation.Caption = "Comparing Images.."

    If ChkAutoTolerance Then                        '//call to tolerance calc
        lRes = Auto_Set_Tolerance
        txtTolerance.Text = lRes
    End If

    Dyn_Compare aBase(), aComp(), lRes              '//send to comparison engine

    lstOutput.ListItems.Clear

    For l = 0 To UBound(aDiff)
        Set Item = lstOutput.ListItems.Add(, , aDiff(l))    '//user notify
    Next l

    lblOperation.Caption = "Scan Complete, Saving Diff File.."
    If chkSaveResults Then
        Open App.Path & "\snapout.txt" For Append As #1     '//save results
        For l = 0 To UBound(aDiff)
            Print #1, aDiff(l)
        Next l
        Close #1
    End If

    lblOperation.Caption = "Comparison Complete!"

End Sub

Private Sub cmdCreateRestore_Click()

Dim sHkey As String
Dim lHKey As Long
Dim sKey  As String
Dim sName As String

    sName = txtRestore.Text
    If Not Len(txtRestore.Text) = 0 Then
        sHkey = Left$(txtRestore.Text, InStr(txtRestore.Text, "\") - 1)
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
            Exit Sub
        End Select

        sKey = Mid$(txtRestore.Text, InStr(txtRestore.Text, "\") + 1)
    Else
        MsgBox "Please specify a fully qualified Key!" & vbNewLine & _
       "No Key Specified!", vbExclamation, "Backup Aborted!"
        Exit Sub
    End If

    If Not Save_Key(sKey, lHKey, sName) Then        '//notify on status
        MsgBox "Key Backup has Failed!" & vbNewLine & _
       "Please Check the Key Path!", vbExclamation, "Backup Failed!"
    Else
        MsgBox "Backup of Key Successful!!" & vbNewLine & _
       "The Key has been saved in the Recovery Folder!", vbExclamation, "Backup Success!"
    End If

End Sub

Private Sub cmdDeployRestore_Click()

    With cDialog
        .DialogTitle = "Please select a recovery file"
        .Filter = "Registry Restore (*.kbs)|*.kbs"
        .ShowOpen
        .InitDir = App.Path & "Recovery\"
        txtRestore.Text = .FileName
    End With

    If Not Deploy_Key(txtRestore.Text) Then
        MsgBox "Key Restore has Failed!" & vbNewLine & _
       "Please Check the Key Path!", vbExclamation, "Backup Failed!"
    Else
        MsgBox "Key Restore Successful!!" & vbNewLine & _
       "The Key has been Restored to its original state!", vbExclamation, "Backup Success!"
    End If

End Sub

Private Sub cmdSnapshot_Click()

Dim l        As Long
Dim sPath    As String
Dim sSubText As String
Dim sSubKey  As String

    If chkSubkey Then
        sSubKey = txtSubkey.Text
        If Not InStrB(txtSubkey.Text, "\") Then
            sSubText = Mid$(txtSubkey.Text, InStrRev(txtSubkey.Text, "\") + 1)
        Else
            sSubText = txtSubkey.Text
        End If
    End If

    If chkValues Then
        sSubText = sSubText & "Val"
        cValues = True
    End If

    Select Case True
    Case Option1(0)
        lHKey = HKEY_CLASSES_ROOT
        sPath = App.Path & "\hkcrsnap" & sSubText & ".txt"
    Case Option1(1)
        lHKey = HKEY_CURRENT_USER
        sPath = App.Path & "\hkcusnap" & sSubText & ".txt"
    Case Option1(2)
        lHKey = HKEY_LOCAL_MACHINE
        sPath = App.Path & "\hklmsnap" & sSubText & ".txt"
    Case Option1(3)
        lHKey = HKEY_USERS
        sPath = App.Path & "\hkusnap" & sSubText & ".txt"
    Case Option1(4)
        lHKey = HKEY_CURRENT_CONFIG
        sPath = App.Path & "\hkccsnap" & sSubText & ".txt"
    End Select

    lblOperation.Caption = "Creating Snap Shot.."
    bDimn = False
    GetKeyInfo lHKey, sSubKey                       '//build the base array
    Reset_Dimensions

    If Not cValues Then
        Open sPath For Output As #1                 '//send to file
        For l = 0 To UBound(aKeyArr)
            Print #1, aKeyArr(l)
            lblOperation.Caption = "Printing: " & l
            DoEvents
        Next l
        Close #1
    Else
        Open sPath For Output As #1
        For l = 0 To UBound(aValArr)
            Print #1, aValArr(l)
            lblOperation.Caption = "Printing: " & l
            DoEvents
        Next l
        Close #1
    End If

    lblOperation.Caption = "Snapshot Complete.."

End Sub

Private Sub cmdStartMonitor_Click()

Dim sHkey As String
Dim lHKey As Long
Dim sKey  As String
Dim mStr  As String
Dim mName As String

    If Not txtMonitor.Text = vbNullString Then
        sHkey = Left$(txtMonitor.Text, InStr(txtMonitor.Text, "\") - 1)
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
            Exit Sub
        End Select

        sKey = Mid$(txtMonitor.Text, InStr(txtMonitor.Text, "\") + 1)
        If sKey = vbNullString Then
            MsgBox "No Sub Key Specified!" & vbNewLine & _
       "Ex. HKEY_CURRENT_USER\AppEvents", vbExclamation, "No Sub Key!"
            Exit Sub
        End If
    Else
        MsgBox "Please specify a Branch and Subkey!" & vbNewLine & _
       "Ex. HKEY_CURRENT_USER\AppEvents", vbExclamation, "No Key!"
        Exit Sub
    End If
    mName = Mid$(txtMonitor.Text, InStrRev(txtMonitor.Text, "\") + 1)
    mStr = Return_Values(lHKey, sKey)
    Start_Monitor mStr, mName, lHKey, sKey

End Sub

Private Sub cmdStopMonitor_Click()

        '//stop and reset variables
    bMonitor = False
    sModmstr = vbNullString
    sModmname = vbNullString
    lModlhkey = 0
    sModskey = vbNullString
    sModApp = vbNullString

End Sub

Private Sub lblVote_Click()

    ShellExecute Me.hWnd, "open", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=61832&lngWId=1", "", "", 1

End Sub

Private Sub Option1_Click(Index As Integer)

    txtSubkey.Text = vbNullString                   '//user proofing
    chkSubkey.Value = 0

End Sub
