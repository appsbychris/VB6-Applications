VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4170
   ClientLeft      =   2145
   ClientTop       =   3675
   ClientWidth     =   8625
   FillColor       =   &H00008000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8625
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      ScaleHeight     =   375
      ScaleWidth      =   1815
      TabIndex        =   19
      Top             =   3480
      Width           =   1815
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Send a server message-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   1785
      End
   End
   Begin VB.TextBox ServerMessage 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   375
      Left            =   2280
      TabIndex        =   15
      Top             =   3480
      Width           =   4095
   End
   Begin DoDMudServer.UltraBox lstEvents 
      Height          =   1095
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1931
      Style           =   3
      Fill            =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mult            =   0   'False
      Sort            =   0   'False
      SELECTSTYLE     =   0
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   160
      ScaleHeight     =   255
      ScaleWidth      =   8295
      TabIndex        =   1
      Top             =   160
      Width           =   8295
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Users online-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   930
      End
      Begin VB.Label Online 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   6
         Top             =   0
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your IP is:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5640
         TabIndex        =   5
         Top             =   0
         Width           =   735
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Online-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   4
         Top             =   0
         Width           =   885
      End
      Begin VB.Label lblTimeSeen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 Hours, 0 Minutes, 0 Seconds"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   3
         Top             =   0
         Width           =   2190
      End
      Begin VB.Label lblIP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000.000.000.000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6480
         TabIndex        =   2
         Top             =   0
         Width           =   1260
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2160
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer timTimer 
      Interval        =   1000
      Left            =   240
      Top             =   3720
   End
   Begin VB.Timer timCombat 
      Interval        =   2500
      Left            =   1200
      Top             =   3720
   End
   Begin VB.Timer timPRest 
      Interval        =   8000
      Left            =   720
      Top             =   3720
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   0
      Left            =   1680
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   23
   End
   Begin DoDMudServer.Raise Raise 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   661
      Style           =   2
      Color           =   14737632
   End
   Begin DoDMudServer.Raise Raise 
      Height          =   1335
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2355
      Style           =   2
      Color           =   14737632
   End
   Begin DoDMudServer.UltraBox lstUsers 
      Height          =   930
      Left            =   360
      TabIndex        =   13
      Top             =   2240
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1640
      Style           =   3
      Color           =   0
      Fill            =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mult            =   0   'False
      Sort            =   0   'False
      SELECTSTYLE     =   0
   End
   Begin DoDMudServer.eButton cmdSend 
      Height          =   375
      Left            =   6480
      TabIndex        =   14
      Top             =   3480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Style           =   2
      Cap             =   "&Send"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      hCol            =   12632256
      bCol            =   12632256
      CA              =   2
   End
   Begin DoDMudServer.eButton cmdBoot 
      Height          =   375
      Left            =   6480
      TabIndex        =   17
      Top             =   2760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Style           =   2
      Cap             =   "&Boot selected user"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      hCol            =   12632256
      bCol            =   12632256
      CA              =   2
   End
   Begin DoDMudServer.Raise Raise 
      Height          =   1095
      Index           =   4
      Left            =   240
      TabIndex        =   18
      Top             =   2160
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1931
      Style           =   2
      Color           =   14737632
   End
   Begin DoDMudServer.Raise Raise 
      Height          =   615
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   3360
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1085
      Style           =   2
      Color           =   14737632
   End
   Begin DoDMudServer.Raise Raise 
      Height          =   2055
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   3625
      Style           =   2
      Color           =   14737632
   End
   Begin DoDMudServer.Raise cmdBack 
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Style           =   2
      Color           =   14737632
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Width           =   105
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuUserDefined 
         Caption         =   "&User Defined Settings"
      End
      Begin VB.Menu mnuDash000001 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuDash000002 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOPtions 
         Caption         =   "&Options"
         Begin VB.Menu mnuChoose 
            Caption         =   "Choose Random Port If Necessary"
         End
         Begin VB.Menu mnuDash0028 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTimeOnline 
            Caption         =   "&Enable Time Online"
            Checked         =   -1  'True
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuDash000003 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLogFile 
            Caption         =   "Write Events to Log File"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuDash000004 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReload 
         Caption         =   "&Reload Server"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Shutdown Server"
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "&Data"
      Begin VB.Menu mnuImport 
         Caption         =   "Import Database..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash00009 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepair 
         Caption         =   "&Compact and Repair Database"
      End
      Begin VB.Menu mnuDash000005 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameMaint 
         Caption         =   "&Game Maintence"
         Begin VB.Menu mnuFloorSweep 
            Caption         =   "&Floor Sweep"
         End
         Begin VB.Menu mnuDash000006 
            Caption         =   "-"
         End
         Begin VB.Menu mnuReset 
            Caption         =   "&Reset Game"
         End
      End
      Begin VB.Menu mnuDash000007 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShellAndShutdown 
         Caption         =   "&Shell Editor and Shutdown Server"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : frmMain
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'

Private AdIsShown   As Boolean
Private lSaveTime   As Long
Private EvilTime    As Long
Private TickTime    As Long

Private Sub cmdBoot_Click()
On Error Resume Next
If lstUsers.ListIndex > 0 Then
    If ws(lstUsers.ListIndex).State = sckConnected Then
        WrapAndSend lstUsers.ListIndex, BRIGHTBLUE & "Sysop discontected you." & WHITE, False
        ws(lstUsers.ListIndex).Close
    End If
    lstUsers.SetItemText lstUsers.ListIndex, "[Line " & CStr(lstUsers.ListIndex) & " - Open]"
    If Val(Online.Caption) > 0 Then Online.Caption = Val(Online.Caption) - 1
    dbPlayers(GetPlayerIndexNumber(lstUsers.ListIndex)).iIndex = 0
    X(lstUsers.ListIndex) = ""
    PNAME(lstUsers.ListIndex) = ""
    pPoint(lstUsers.ListIndex) = 0
    UpdateList "}bLine " & (lstUsers.ListIndex) & " has been booted from the server. }b(}n}i" & Time & "}n}b)"
End If
End Sub

Private Sub cmdSend_Click()
On Error Resume Next
SendToAll GREEN & "[SERVER MESSAGE]: " & ServerMessage.Text & vbCrLf & WHITE
ServerMessage.Text = ""
End Sub

Private Sub ServerMessage_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdSend_Click
End Sub

Private Sub Form_Load()
On Error GoTo Form_Load_Error
MousePointer = vbHourglass

Randomize
CheckLogFileSize
MonRgnCounter = 0

mnuTimeOnline.Checked = GetINI("ShowTime")
mnuLogFile.Checked = GetINI("LogFile")
mnuChoose.Checked = GetINI("ChoosePort")

lDeath = Val(GetINI("DeathLevel"))
lIsPvP = Val(GetINI("PvPE"))
lPvPLevel = Val(GetINI("PvPL"))

MousePointer = vbDefault

With cmdBack
    .Left = 0
    .Top = 0
    .Width = Me.ScaleWidth
    .Height = Me.ScaleHeight
End With
On Error GoTo 0
Exit Sub
Form_Load_Error:
MousePointer = vbDefault
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_Load in Form, frmMain"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If Val(Online.Caption) > 0 Then
    If MsgBox("WARNING: " & vbCrLf & "Shutting down the server will cause all players to be disconected!" & vbCrLf & "Do you wish to continue?", vbCritical + vbOKCancel, "Warning") = vbCancel Then
        Cancel = 1
        Exit Sub
    End If
End If
UpdateList "}bServer shutdown. }b(}n}i" & Time & "}n}b)"
ShutDownServer
SaveMemoryToDatabase 0
SaveMemoryToDatabase 1
SaveMemoryToDatabase 2
SaveMemoryToDatabase 3
SaveMemoryToDatabase 4
CloseDatabase
FileCopy App.Path & "\data.mud", App.Path & "\BACKUP.BAK"
Unload frmGraphics
Unload frmAbout
Unload frmImport
Unload frmUserDefined
End
End Sub

Private Sub mnuAbout_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub mnuChoose_Click()
Select Case mnuChoose.Checked
    Case True
        mnuChoose.Checked = False
    Case False
        mnuChoose.Checked = True
End Select
WriteINI "ChoosePort", mnuChoose.Checked
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuFloorSweep_Click()
Dim Items As String
Dim i As Long, j As Long
Dim tArr() As String
If MsgBox("By continuing with this action, your server must be put offline." & vbCrLf & "Anyone currently connected will be disconected." & vbCrLf & "All items and gold on the floor in the game will be removed." & vbCrLf & "Do you wish to continue?", vbQuestion + vbOKCancel, "Floor Sweep") = vbOK Then
    MousePointer = vbHourglass
    UpdateList "}bServer shutdown. }b(}n}i" & Time & "}n}b)"
    ShutDownServer
    UpdateList "Sweeping Floor. }b(}n}i" & Time & "}n}b)"
    For i = LBound(dbMap) To UBound(dbMap)
        If dbMap(i).sItems <> "0" Then Items = Items & Left$(dbMap(i).sItems, Len(dbMap(i).sItems) - 1)
        dbMap(i).sItems = "0"
        dbMap(i).dGold = 0
    Next
    UpdateList "Floor Sweep complete. }b(}n}i" & Time & "}n}b)"
    UpdateList "Check limited items. }b(}n}i" & Time & "}n}b)"
    If modSC.FastStringComp(Items, "") Then GoTo SkipLimitCheck
    SplitFast Left$(Items, Len(Items) - 1), tArr, ";"
    For j = LBound(tArr) To UBound(tArr)
        For i = LBound(dbItems) To UBound(dbItems)
            If modItemManip.GetItemIDFromUnFormattedString(tArr(j)) = dbItems(i).iID Then
                If dbItems(i).iLimit <> 0 Then
                    dbItems(i).iInGame = dbItems(i).iInGame - 1
                End If
            End If
        Next
    Next
SkipLimitCheck:
    UpdateList "Limit Check Complete. }b(}n}i" & Time & "}n}b)"
    SaveMemoryToDatabase 0
    SaveMemoryToDatabase 1
    SaveMemoryToDatabase 2
    SaveMemoryToDatabase 3
    SaveMemoryToDatabase 4
    modDatabase.CloseRecordsets
    modDatabase.CloseDatabase
    modDatabase.OpenDatabaseConnection
    modDatabase.InitRecordsets
    CheckSpecialItems
    UpdateList "Loading Depths of Despair MUD Server... }b(}n}i" & Time & "}n}b)"
    ReloadServer
    MonRgnCounter = 0
    FillType
    ws(0).Listen
    If mnuTimeOnline.Checked = True Then timTimer.Enabled = True
    UpdateList "}b}iReady to service users. (}n}i" & Time & "}n}b}i)"
    MousePointer = vbDefault
    Unload frmSplash
    UpdateList "Floor Sweep was successful. }b(}n}i" & Time & "}n}b)"
End If
End Sub

Private Sub mnuImport_Click()
Load frmImport
frmImport.Show
End Sub

Private Sub mnuLogFile_Click()
Select Case mnuLogFile.Checked
    Case True
        mnuLogFile.Checked = False
    Case False
        mnuLogFile.Checked = True
End Select
WriteINI "LogFile", mnuLogFile.Checked
End Sub

Private Sub mnuReload_Click()
If MsgBox("Are you sure you wish to shutdown, and reload the server?", vbQuestion + vbOKCancel, "Restart") = vbOK Then
    MousePointer = vbHourglass
    UpdateList "}bServer shutdown. }b(}n}i" & Time & "}n}b)"
    ShutDownServer
    SaveMemoryToDatabase 0
    SaveMemoryToDatabase 1
    SaveMemoryToDatabase 2
    SaveMemoryToDatabase 3
    SaveMemoryToDatabase 4
    modDatabase.CloseRecordsets
    modDatabase.CloseDatabase
    modDatabase.OpenDatabaseConnection
    modDatabase.InitRecordsets
    UpdateList "Loading Depths of Despair MUD Server... }b(}n}i" & Time & "}n}b)"
    ReloadServer
    MonRgnCounter = 0
    FillType
    lstUsers.Clear
    For i = 1 To MaxUsers
        lstUsers.AddItem "[Line " & CStr(i) & " - Open]"
    Next
    ws(0).Listen
    frmMain.lblIP.Caption = GetIPAddress & " (" & ws(0).LocalPort & ")"
    UpdateList "}b}iReady to service users. (}n}i" & Time & "}n}b}i)"
    Unload frmSplash
    UpdateList "Server reloaded. }b(}n}i" & Time & "}n}b)"
    MousePointer = vbDefault
End If
End Sub

Sub ReloadServer()
MaxUsers = CLng(GetINI("MaxPlayers"))
MaxMonsters = CLng(GetINI("MaxMonsters"))
ReDim SpellCombat(MaxUsers) As Boolean
ReDim X(MaxUsers) As String
ReDim PNAME(MaxUsers) As String
ReDim pLogOn(MaxUsers) As Boolean
ReDim pLogOnPW(MaxUsers) As Boolean
ReDim pPoint(MaxUsers) As Long
ReDim aMons(MaxMonsters) As Monster
modDatabase.OpenDatabaseConnection
modDatabase.InitRecordsets
CheckSpecialItems
modTime.SetNameArrays
modTime.SetTimeOfDay
modTime.SetDayOfWeek
modTime.LoadMonths
modTime.SetMonthOfYear
modTime.SetYear
LoadDatabaseIntoMemory
timTimer.Enabled = True
End Sub

Sub ShutDownServer()
On Error Resume Next
Dim a As Long
timTimer.Enabled = False
lblTime.Caption = "0"
For a = 1 To lstUsers.ListCount
    lstUsers.SetItemText a, "[Line " & CStr(a) & " - Open]"
Next
Online.Caption = "0"
For a = LBound(dbPlayers) To UBound(dbPlayers)
    dbPlayers(a).iIndex = 0
    dbPlayers(a).sParty = "0"
    dbPlayers(a).dMonsterID = 99999
    dbPlayers(a).iResting = 0
    dbPlayers(a).iMeditating = 0
    dbPlayers(a).iInvitedBy = 0
    dbPlayers(a).iPartyLeader = 0
    dbPlayers(a).iLeadingParty = 0
Next
For a = 0 To ws.UBound
    ws(a).Close
Next a
For a = LBound(aMons) To UBound(aMons)
    With aMons(a)
        If .mLoc <> -1 And .mLoc <> 0 Then
            If .mPEQ <> "" Or .mPMoney <> 0 Then
                On Error Resume Next
                With dbMap(.mdbMapID)
                    If .sItems = "0" Then .sItems = ""
                    .sItems = .sItems & aMons(a).mPEQ
                    .dGold = .dGold + aMons(a).mPMoney
                End With
            End If
        End If
    End With
    If DE Then DoEvents
Next
Erase aMons
modTime.WriteTimeToFile
modTime.WriteDayToFile
modTime.WriteMonthOfYearToFile
modTime.WriteYearToFile
End Sub

Private Sub mnuRepair_Click()
On Error GoTo mnuRepair_Click_Error
If MsgBox("By continuing with this action, your server must be put offline, and reset." & vbCrLf & "Anyone currently connected will be disconected." & vbCrLf & "Do you wish to continue?", vbQuestion + vbOKCancel, "Compact and Repair Database") = vbOK Then
    MousePointer = vbHourglass
    UpdateList "}bServer shutdown. }b(}n}i" & Time & "}n}b)"
    ShutDownServer
    SaveMemoryToDatabase 0
    SaveMemoryToDatabase 1
    SaveMemoryToDatabase 2
    SaveMemoryToDatabase 3
    SaveMemoryToDatabase 4
    modDatabase.CloseRecordsets
    modDatabase.CloseDatabase
    UpdateList "Compacting... }b(}n}i" & Time & "}n}b)"
    FileCopy App.Path & "\data.mud", App.Path & "\BACKUP.BAK"
    modDatabase.ValidateDatabase
    CompactDatabase App.Path & "\data.mud", App.Path & "\temp.file", False, False, ";pwd=" & uJunkIt(sValue)
    FileCopy App.Path & "\temp.file", App.Path & "\data.mud"
    Kill App.Path & "\temp.file"
    UpdateList "Loading Depths of Despair MUD Server... }b(}n}i" & Time & "}n}b)"
    ReloadServer
    TickTime = 0
    LoadDatabaseIntoMemory
    FillType
    ws(0).Listen
    UpdateList "}b}iReady to service users. (}n}i" & Time & "}n}b}i)"
    Unload frmSplash
    UpdateList "Compact and Repair complete. }b(}n}i" & Time & "}n}b)"
    MousePointer = vbDefault
End If
On Error GoTo 0
Exit Sub
mnuRepair_Click_Error:
MousePointer = vbDefault
FileCopy "BACKUP.BAK", "data.mud"
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: mnuRepair_Click in Form, frmMain"
End Sub

Private Sub mnuReset_Click()
On Error GoTo mnuReset_Click_Error
If MsgBox("Reseting will delete all players, restore all items and monsters in the game. Do you wish to continue?", vbCritical + vbOKCancel, "Reset Game") = vbOK Then
    If MsgBox("CAUTION: " & vbCrLf & "None of this data will be recoverable. Are you sure you wish to continue?", vbCritical + vbOKCancel, "Caution") = vbOK Then
        MousePointer = vbHourglass
        UpdateList "}bServer shutdown. }b(}n}i" & Time & "}n}b)"
        For i = 1 To UBound(dbMap)
            With dbMap(i)
                .dGold = 0
            End With
            If DE Then DoEvents
        Next
        ShutDownServer
        SaveMemoryToDatabase 0
        SaveMemoryToDatabase 1
        SaveMemoryToDatabase 2
        SaveMemoryToDatabase 3
        SaveMemoryToDatabase 4
        modDatabase.CloseRecordsets
        modDatabase.CloseDatabase
        modDatabase.OpenDatabaseConnection
        modDatabase.InitRecordsets
        With MRS
            .MoveFirst
            .MoveNext
            If .EOF Then GoTo SkipMe
            Do
                .Delete
                .MoveNext
            Loop Until .EOF
        End With
SkipMe:
        UpdateList "Players Deleted. }b(}n}i" & Time & "}n}b)"
        With MRSMAP
            .MoveFirst
            Do
                .Edit
                !Items = "0"
                !Monsters = "0"
                .Update
                .MoveNext
            Loop Until .EOF
        End With
        UpdateList "Rooms Cleaned. }b(}n}i" & Time & "}n}b)"
        With MRSMONSTER
            .MoveFirst
            Do
                .Edit
                !RegenTimeLeft = "0"
                .Update
                .MoveNext
            Loop Until .EOF
        End With
        UpdateList "Monsters Cleaned. }b(}n}i" & Time & "}n}b)"
        With MRSITEM
            .MoveFirst
            Do
                .Edit
                !InGame = "0"
                .Update
                .MoveNext
            Loop Until .EOF
        End With
        UpdateList "Items Cleaned. }b(}n}i" & Time & "}n}b)"
        UpdateList "Closeing Database Connection. }b(}n}i" & Time & "}n}b)"
        modDatabase.CloseRecordsets
        modDatabase.CloseDatabase
        UpdateList "Loading Depths of Despair MUD Server... }b(}n}i" & Time & "}n}b)"
        modDatabase.OpenDatabaseConnection
        modDatabase.InitRecordsets
        TickTime = 0
        CheckSpecialItems
        ReloadServer
        FillType
        ws(0).Listen
        UpdateList "}b}iReady to service users. (}n}i" & Time & "}n}b}i)"
        Unload frmSplash
        UpdateList "Game successfully reset. }b(}n}i" & Time & "}n}b)"
        MousePointer = vbDefault
    End If
End If
On Error GoTo 0
Exit Sub
mnuReset_Click_Error:
MousePointer = vbDefault

MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: mnuRepair_Click in Form, frmMain"
End Sub

Private Sub mnuShellAndShutdown_Click()
On Error GoTo mnuShellAndShutdown_Click_Error
If MsgBox("By continuing with this action, the server will be shutdown." _
    & vbCrLf & "Do you wish to continue?" _
    , vbOKCancel + vbQuestion, "Shell Editor") = vbOK Then
    MousePointer = vbHourglass
    UpdateList "}bServer shutdown. }b(}n}i" & Time & "}n}b)"
    ShutDownServer
    CloseDatabase
    If modSC.FastStringComp(GetINI("EditorPath"), "Error") Or modSC.FastStringComp(GetINI("EditorPath"), "") Then
        MsgBox "There is no set path to the editor." & vbCrLf & "On the next prompt, locate the the editor.exe file.", vbCritical + vbOKOnly, "Locate Editor"
        CD1.Filter = "Exe Files|*.exe"
        CD1.DefaultExt = ".exe"
        CD1.DialogTitle = "Find editor.exe"
        CD1.InitDir = App.Path
        CD1.ShowOpen
        If LCaseFast(Right$(CD1.FileName, 10)) = "editor.exe" Then
            WriteINI "EditorPath", CD1.FileName
        Else
            MsgBox "Shell has determined that the selected file is not the correct file." & vbCrLf & "Therefore, Shell cannot complete.", vbCritical, "Error"
            WriteINI "EditorPath", ""
            Exit Sub
        End If
    End If
    Shell GetINI("EditorPath"), vbNormalFocus
    MousePointer = vbDefault
    Unload Me
End If
On Error GoTo 0
Exit Sub
mnuShellAndShutdown_Click_Error:
WriteINI "EditorPath", ""
MousePointer = vbDefault
MsgBox "There is no set path to the editor or the path is no longer valid." & vbCrLf & "Please re-select this option to reset the location of the editor.", vbCritical + vbOKOnly, "Error"
End Sub

Private Sub mnuTimeOnline_Click()
Select Case mnuTimeOnline.Checked
    Case True
        lblTime.Caption = "0"
        lblTimeSeen.Caption = "Disabled"
        mnuTimeOnline.Checked = False
    Case False
        lblTime.Caption = "0"
        lblTimeSeen.Caption = "0 Seconds"
        mnuTimeOnline.Checked = True
End Select
WriteINI "ShowTime", mnuTimeOnline.Checked
End Sub

Private Sub mnuUserDefined_Click()
MousePointer = vbHourglass
Load frmUserDefined
frmUserDefined.Show
MousePointer = vbDefault
End Sub

Private Sub timCombat_Timer()
Dim i As Long
GenCombat
MoveTime
For i = LBound(dbPlayers) To UBound(dbPlayers)
    With dbPlayers(i)
        .lHasCasted = 0
    End With
    If DE Then DoEvents
Next
End Sub

Private Sub timPRest_Timer()
On Error GoTo timPRest_Timer_Error
Dim bDoneResting    As Boolean
Dim tArr()          As String
Dim tArr1()         As String
Dim tArr2()         As String
Dim aFlgs()         As String
Dim sItem           As String
Dim s               As String
Dim sExits          As String
Dim i               As Long
Dim j               As Long
Dim n               As Long
Dim m               As Long
Dim RndRoom         As Long
EvilTime = EvilTime + 1
TickTime = TickTime + 1
For i = LBound(dbPlayers) To UBound(dbPlayers)
    With dbPlayers(i)
        If .iIndex <> 0 And pPoint(.iIndex) = 0 Then
            If ws(.iIndex).State = 0 Then ws_Close (.iIndex)
            If .lQueryTimer = 1 Then WrapAndSend .iIndex, BRIGHTWHITE & "Timer Tick; TickTime=" & CStr(TickTime) & " EvilTime=" & CStr(EvilTime) & WHITE & vbCrLf
            If TickTime = 3 Or TickTime = 6 Then
                With dbMap(.lDBLocation)
                    If .lNorth <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lNorth
                    If .lSouth <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lSouth
                    If .lEast <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lEast
                    If .lWest <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lWest
                    If .lNorthEast <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lNorthEast
                    If .lNorthWest <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lNorthWest
                    If .lSouthEast <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lSouthEast
                    If .lSouthWest <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lSouthWest
                    If .lUp <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lUp
                    If .lDown <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lDown
                End With
                bDoneResting = False
                If .lHP <= 0 Then CheckDeath .iIndex, dbIndex:=i
                If .lRegain = 0 Then
                    If RndNumber(1, 100) > 82 Then .dClassPoints = .dClassPoints + 0.01
                    If .iResting = 1 Then
                        
                        n = (.iDex \ 3) + (.iStr \ 3) + (.lMaxHP \ 20) + (.iCha \ 18)
                        If n < 10 Then n = 10
                        .lHP = .lHP + n
                        If .lHP > .lMaxHP Then .lHP = .lMaxHP
                        
                        m = modMiscFlag.GetStatsPlusTotal(i, [Spell Casting])
                        n = (.iInt \ 3) + (.iCha \ 3) + (.lMaxMana \ 30) + (.iAgil \ 17) + (m \ 84)
                        If n < 10 Then n = 10
                        If n > m Then n = m
                        .lMana = .lMana + n
                        If .lMana > .lMaxMana Then .lMana = .lMaxMana
                        
                        n = (.iDex \ 2) + (.iStr \ 2) + (.lMaxHP \ 40) + (.iCha \ 35)
                        If n < 10 Then n = 10
                        .dStamina = .dStamina + RndNumber(CDbl(n \ 2), CDbl(n) * 2)
                        
                        n = (.iDex \ 3) + (.iStr \ 3) + (.lFamMHP \ 20) + (.iCha \ 18)
                        If n < 5 Then n = 5
                        
                        If .lFamID <> 0 Then
                            .lFamCHP = .lFamCHP + n
                            If .lFamCHP > .lFamMHP Then .lFamCHP = .lFamMHP
                        End If
                        
                        .dHunger = .dHunger - RndNumber(0, 1)
                    Else
                        n = (.iDex \ 12) + (.iStr \ 12) + (.lMaxHP \ 40) + (.iCha \ 36)
                        If n < 3 Then n = 3
                        .lHP = .lHP + n
                        If .lHP > .lMaxHP Then .lHP = .lMaxHP
                        
                        m = modMiscFlag.GetStatsPlusTotal(i, [Spell Casting])
                        n = (.iInt \ 18) + (.iCha \ 18) + (.lMaxMana \ 60) + (.iAgil \ 34) + (m \ 96)
                        If n < 2 Then n = 2
                        If n > m Then n = m - (m \ 2)
                        .lMana = .lMana + n
                        If .lMana > .lMaxMana Then .lMana = .lMaxMana
                        
                        n = (.iDex \ 6) + (.iStr \ 6) + (.lMaxHP \ 60) + (.iCha \ 70)
                        If n < 2 Then n = 2
                        .dStamina = .dStamina + RndNumber(1, CDbl(n))
                        
                        n = (.iDex \ 6) + (.iStr \ 6) + (.lFamMHP \ 80) + (.iCha \ 45)
                        If n < 3 Then n = 3
                        
                        If .lFamID <> 0 Then
                            .lFamCHP = .lFamCHP + n
                            If .lFamCHP > .lFamMHP Then .lFamCHP = .lFamMHP
                        End If
                        
                        .dHunger = .dHunger - RndNumber(0, 1)
                    
                    End If
                    
                    If .iMeditating = 1 Then
                        
                        n = (.iDex \ 3) + (.iStr \ 3) + (.lMaxHP \ 20) + (.iCha \ 18)
                        If n < 4 Then n = 4
                        .lHP = .lHP + n
                        If .lHP > .lMaxHP Then .lHP = .lMaxHP
                        
                        m = modMiscFlag.GetStatsPlusTotal(i, [Spell Casting])
                        n = (.iInt \ 3) + (.iCha \ 3) + (.lMaxMana \ 30) + (.iAgil \ 17) + (m \ 84)
                        If n < 20 Then n = 20
                        If n > m Then n = m
                        .lMana = .lMana + n
                        If .lMana > .lMaxMana Then .lMana = .lMaxMana
                        
                        n = (.iDex \ 2) + (.iStr \ 2) + (.lMaxHP \ 40) + (.iCha \ 35)
                        If n < 15 Then n = 15
                        .dStamina = .dStamina + RndNumber(CDbl(n \ 2), CDbl(n) * 2)
                        
                        n = (.iDex \ 3) + (.iStr \ 3) + (.lFamMHP \ 20) + (.iCha \ 18)
                        If n < 5 Then n = 5
                        
                        If .lFamID <> 0 Then
                            .lFamCHP = .lFamCHP + n
                            If .lFamCHP > .lFamMHP Then .lFamCHP = .lFamMHP
                        End If
                        
                    End If
                    
                    If .dHunger < 1 Then
                        WrapAndSend .iIndex, BRIGHTRED & "You are starving!" & WHITE & vbCrLf
                        .lHP = .lHP + RndNumber(.dHunger - 10, .dHunger)
                        CheckDeath .iIndex, dbIndex:=i
                    End If
                    
                    If (.lHP = .lMaxHP) And (.lMana = .lMaxMana) And (.dStamina > 99) Then bDoneResting = True
                    If .iResting = 1 Then WrapAndSend .iIndex, ""
                    
                ElseIf .lRegain = -1 Then
                    .lHP = .lHP - 1
                    WrapAndSend .iIndex, ""
                Else
                    .lHP = .lHP + 1
                    WrapAndSend .iIndex, ""
                End If
                If bDoneResting Then .iResting = 0: .iMeditating = 0
            End If
            If EvilTime >= 180 Then
                If .iEvil < 32000 And .iEvil > -32000 Then
                    .iEvil = .iEvil - 1
                End If
                EvilTime = 0
            End If
            If .sBlessSpells <> "0" Then
                SplitFast Left$(.sBlessSpells, Len(.sBlessSpells) - 1), tArr1, "Œ"
                'Timeout~roll~spellname~dbspellid
                For j = LBound(tArr1) To UBound(tArr1)
                    Erase tArr2
                    SplitFast tArr1(j), tArr2, "~"
                    tArr2(0) = CLng(tArr2(0)) - 1
                    If CLng(tArr2(0)) <= 0 Then
                        If Not modSC.FastStringComp(dbSpells(Val(tArr2(3))).sFlags, "0") Then
                        modUseItems.DoFlags i, dbSpells(Val(tArr2(3))).sFlags, lRoll:=CLng(tArr2(1)), Inverse:=True
                    End If
                        .sBlessSpells = ReplaceFast(.sBlessSpells, CStr((Val(tArr2(0)) + 1)) & "~" & tArr2(1) & "~" & tArr2(2) & "~" & tArr2(3) & "Œ", "", 1, 1)
                        If modSC.FastStringComp(.sBlessSpells, "") Then .sBlessSpells = "0"
                        sSend .iIndex, LIGHTBLUE & dbSpells(Val(tArr2(3))).sRunOutMessage
                    Else
                        s = CStr((Val(tArr2(0)) + 1)) & "~" & tArr2(1) & "~" & tArr2(2) & "~" & tArr2(3) & "Œ"
                        
                        .sBlessSpells = ReplaceFast(.sBlessSpells, s, Join(tArr2, "~") & "Œ", 1, 1)
                    End If
                    If DE Then DoEvents
                Next j
            End If
            If .sKillDurItems <> "0" Then
                Erase tArr1
                SplitFast .sKillDurItems, tArr1, ";"
                For j = LBound(tArr1) To UBound(tArr1)
                    If tArr1(j) <> "" Then
                        Select Case Left$(tArr1(j), InStr(1, tArr1(j), "/") - 1)
                            Case "shield"
                                sItem = dbItems(GetItemID(, _
                                    modItemManip.GetItemIDFromUnFormattedString( _
                                    .sShield))).sItemName
                                    
                                modItemManip.SubtractNFromItemDUR CLng(i), _
                                    CLng(Right$(tArr1(j), Len(tArr1(j)) - InStr(1, _
                                    tArr1(j), "/"))), _
                                    modItemManip.GetItemIDFromUnFormattedString( _
                                    .sShield), _
                                    modItemManip.GetItemUsesFromUnFormattedString( _
                                    .sShield), _
                                    modItemManip.GetItemDurFromUnFormattedString( _
                                    .sShield)
                                    
                                If .sShield = "0" Then
                                    
                                    .sKillDurItems = ReplaceFast(.sKillDurItems, tArr1(j) & ";", "")
                                    If .sKillDurItems = "" Then .sKillDurItems = "0"
                                End If
                            Case "weapon"
                                sItem = dbItems(GetItemID(, _
                                    modItemManip.GetItemIDFromUnFormattedString( _
                                    .sWeapon))).sItemName
                                modItemManip.SubtractNFromItemDUR CLng(i), _
                                    CLng(Right$(tArr1(j), Len(tArr1(j)) - InStr(1, _
                                    tArr1(j), "/"))), _
                                    modItemManip.GetItemIDFromUnFormattedString( _
                                    .sWeapon), _
                                    modItemManip.GetItemUsesFromUnFormattedString( _
                                    .sWeapon), _
                                    modItemManip.GetItemDurFromUnFormattedString( _
                                    .sWeapon)
                                If .sWeapon = "0" Then
                                    
                                    .sKillDurItems = ReplaceFast(.sKillDurItems, tArr1(j) & ";", "")
                                    If .sKillDurItems = "" Then .sKillDurItems = "0"
                                End If
                        End Select
                    End If
                    If DE Then DoEvents
                Next
            End If
        End If
    End With
    If DE Then DoEvents
Next

If TickTime = 6 Then
    If bUpdate Then Exit Sub
    For i = LBound(dbMonsters) To UBound(dbMonsters)
        If bUpdate Then Exit For
        With dbMonsters(i)
            If .iType = 2 Then
                If .lRegenTimeLeft > 0 Then
                    .lRegenTimeLeft = .lRegenTimeLeft - 1
                End If
            End If
        End With
        If DE Then DoEvents
    Next
    CloseDoorsTimer
End If

If TickTime = 2 Or TickTime = 4 Or TickTime = 6 Then ArenaMon

If TickTime = 2 Or TickTime = 6 Then
    n = 0
    For i = LBound(aMons) To UBound(aMons)
        If n > AmountMons + 10 Then Exit For
        With aMons(i)
            If .mRoams <> 0 And .mLoc <> 0 And .mLoc <> -1 Then
                n = n + 1
                If RndNumber(0, 100) > 80 Then
                    If .mIs_Being_Attacked = False And .mIsAttacking = False Then
                        sExits = modGetData.GetRoomExits2(aMons(i).mdbMapID)
                        If Not modSC.FastStringComp(sExits, "") Then
                            SplitFast sExits, tArr, ","
                            RndRoom = RndNumber(LBound(tArr), UBound(tArr))
                            sExits = modGetData.GetRoomExitFrom2Points(aMons(i).mdbMapID, CLng(tArr(RndRoom)))
                            If Not modSC.FastStringComp(sExits, "") Then
                                RoamMonsters i, CLng(tArr(RndRoom)), sExits
                            End If
                        End If
                    End If
                End If
            End If
        End With
        If DE Then DoEvents
    Next
End If
If TickTime = 1 Or TickTime = 3 Or TickTime = 5 Then
    For i = 1 To UBound(dbMBTimer)
        With dbMBTimer(i)
            .lTimePassed = .lTimePassed + 1
            If .lTimePassed >= .lInterval Then
                .lTimePassed = 0
                For j = LBound(dbPlayers) To UBound(dbPlayers)
                    If .lRoomID = dbPlayers(j).lLocation Then
                        sScripting dbPlayers(j).iIndex, IgnoreTimer:=True, UseThisScript:=.sScript
                    End If
                    If DE Then DoEvents
                Next
            End If
        End With
        If DE Then DoEvents
    Next
End If
If TickTime = 6 Then
    For i = LBound(dbMap) To UBound(dbMap)
        If RndNumber(1, 100) > 25 Then
            GenAMonster dbMap(i).lRoomID, , , , i
        End If
        If DE Then DoEvents
    Next
End If
If TickTime = 6 Then TickTime = 0
On Error GoTo 0
Exit Sub
timPRest_Timer_Error:
End Sub

Private Sub timTimer_Timer()
Dim tVar As Double
Dim i As Long
Dim dMinutes As Double, dHours As Double
On Error Resume Next
For i = LBound(dbPlayers) To UBound(dbPlayers)
    With dbPlayers(i)
        If .iStun <> 0 Then .iStun = .iStun - 1
    End With
    If DE Then DoEvents
Next
lSaveTime = lSaveTime + 1
Select Case lSaveTime
    Case 125
        SaveMemoryToDatabase 0
    Case 150
        SaveMemoryToDatabase 1
    Case 195
        SaveMemoryToDatabase 2
    Case 235
        SaveMemoryToDatabase 3
    Case 260
        SaveMemoryToDatabase 4
        lSaveTime = 0
End Select
If Me.mnuTimeOnline.Checked = True Then
    tVar = CDbl(lblTime.Caption)
    tVar = tVar + 1
    lblTime.Caption = tVar
    dMinutes = 0
    dHours = 0
    While tVar >= 60
        tVar = tVar - 60
        dMinutes = dMinutes + 1
    Wend
    While dMinutes >= 60
        dMinutes = dMinutes - 60
        dHours = dHours + 1
    Wend
    If dHours = 0 And dMinutes = 0 Then
        lblTimeSeen.Caption = tVar & " seconds."
    ElseIf dHours = 0 And dMinutes <> 0 Then
        lblTimeSeen.Caption = dMinutes & " minutes, " & tVar & " seconds."
    Else
        lblTimeSeen.Caption = dHours & " hours, " & dMinutes & " minutes, " & tVar & " seconds."
    End If
End If
End Sub

Private Sub ws_Close(Index As Integer)
Dim lngIndex As Long
lngIndex = Index
On Error GoTo ws_Close_Error
lstUsers.SetItemText lngIndex, "[Line " & CStr(Index) & " - Open]"
If Val(Online.Caption) > 0 Then Online.Caption = Val(Online.Caption) - 1
UpdateList "Line " & CStr(Index) & " signed off. }b(}n}i" & Time & "}n}b)"
RemoveFromParty lngIndex
If GetPlayerIndexNumber(lngIndex) <> 0 Then
With dbPlayers(GetPlayerIndexNumber(lngIndex))
    .iIndex = 0
    .sParty = "0"
    .dMonsterID = 99999
    .iResting = 0
    .iMeditating = 0
    .iInvitedBy = 0
    .iPartyLeader = 0
    .iLeadingParty = 0
    If .iPlayerAttacking <> 0 Then dbPlayers(GetPlayerIndexNumber(.iPlayerAttacking)).iPlayerAttacking = 0
    .iPlayerAttacking = 0
    UpdateList .sPlayerName & " has signed off. }b(}n}i" & Time & "}n}b)"
    SendToAll BRIGHTGREEN & .sPlayerName & " has just left the world." & WHITE & vbCrLf
End With
End If
ws(Index).Close
pPoint(Index) = 0
PNAME(Index) = ""
X(Index) = ""
On Error GoTo 0
Exit Sub
ws_Close_Error:
End Sub

Private Sub ws_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim Dragon$
Dim i As Long
Dim a As Long
Dim j As Long
Dim m As Long
Dim n As Long
Dim s As String
If Index = 0 Then
    For a = 1 To MaxUsers
        If modSC.FastStringComp(Right$(lstUsers.List(a), 5), "Open]") Then
            ws(a).Close
            ws(a).Accept requestID
            UpdateList "A player has connected on line " & a & " at " & Time & "."
            lstUsers.SetItemText a, "[Line " & CStr(a) & " - In Use (" & ws(a).RemoteHostIP & ")]"   'update stuff
            If CheckBIP(ws(a).RemoteHostIP) Then
                UpdateList "Players IP Blocked [" & ws(a).RemoteHostIP & "] }b(}n}i" & Time & "}n}b)"
                WrapAndSend CLng(a), BRIGHTRED & "Sorry, but your IP address has been blocked." & vbCrLf & BRIGHTBLUE & "DISCONTECTING", False
                ws_Close CLng(a)
                ws(a).Close
                Exit Sub
            End If
            j = 0
            s = ws(a).RemoteHostIP
            For i = 0 To lstUsers.ListCount - 1
                With lstUsers
                    m = InStr(1, .List(i), "(")
                    If m <> 0 Then
                        n = InStr(1, .List(i), ")")
                        If modSC.FastStringComp(s, Mid$(.List(i), m + 1, n - m - 1)) Then
                            j = j + 1
                        End If
                    End If
                End With
                If DE Then DoEvents
            Next
            If j > CLng(GetINI("Logons")) Then
                UpdateList "Too many connections from 1 IP [" & ws(a).RemoteHostIP & "] }b(}n}i" & Time & "}n}b)"
                WrapAndSend CLng(a), BRIGHTRED & "Sorry, but there are too many connections coming from your IP address." & vbCrLf & BRIGHTBLUE & "DISCONTECTING", False
                ws_Close CLng(a)
                ws(a).Close
                Exit Sub
            End If
            pLogOn(a) = True
            Online.Caption = Val(Online.Caption) + 1
            Dragon$ = sGraphic
            ws(a).SendData "]2J" & sDeCode(sBuild & " & newline", "") & Dragon & sDeCode("color.brightyellow & ;Welcome to the ; & color.brightred & ;Depths of Despair. ; & color.brightyellow & newline & ;Insert thy name to continue, or if; & newline & ;you have yet to join, type ""; & color.brightgreen & ;new; & color.brightyellow & ;"": ;", "")
            Exit Sub
        End If
    Next a
End If
End Sub

Private Sub ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Msg As String
Dim dbIndex As Long
Dim lngIndex As Long
lngIndex = CLng(Index)
'On Error GoTo ws_DataArrival_Error
ws(lngIndex).GetData Msg, vbString
pWeapon(lngIndex).lCol = 0
If pPoint(lngIndex) = 15 Then
    'TRAIN MODE!
    Select Case Msg
        Case UP_ARROW, DOWN_ARROW, RIGHT_ARROW, LEFT_ARROW
            If X(lngIndex) <> "" Then modTrain.ValidateAndAdjust lngIndex
            X(lngIndex) = Msg
            modTrain.MOVECURSOR lngIndex
        Case Else
            If (InStr(1, Msg, vbCrLf) Or InStr(1, Msg, vbLf) Or InStr(1, Msg, vbCr)) Then
                X(lngIndex) = Msg
                modTrain.ValidateAndAdjust lngIndex
            Else
                If Asc(Msg) = vbKeyBack Then
                    If Not modSC.FastStringComp(X(lngIndex), "") Then
                        X(lngIndex) = Left$(X(lngIndex), Len(X(lngIndex)) - 1)
                    End If
                Else
                    X(lngIndex) = X(lngIndex) & Msg
                End If
            End If
    End Select
    If PlayerHasEcho(lngIndex, dbIndex) = True Then
        If Msg <> Chr$(vbKeyBack) And Msg <> vbCrLf Then
            ws(lngIndex).SendData Msg
        End If
    End If
    Exit Sub
ElseIf pPoint(lngIndex) >= 56 And pPoint(lngIndex) < 70 Then
    If pPoint(lngIndex) = 56 Then
        Select Case Left$(LCaseFast(X(Index)), 1)
            Case "m", "i", "f"
                modNewChar.ChooseGender lngIndex
                sSend lngIndex, ";" & ANSICLS & "; & color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your hair length; & newline & " & ";" & HairLen(0) & ";", ""
                X(Index) = ""
                Msg = ""
                pPoint(lngIndex) = 57
            Case Else
                GoTo SkipAll
        End Select
    End If
    Select Case Msg
        Case UP_ARROW, DOWN_ARROW, RIGHT_ARROW, LEFT_ARROW
            modAppearance.MoveAppearance lngIndex, Msg
        Case Else
            If (InStr(1, Msg, vbCrLf) Or InStr(1, Msg, vbLf) Or InStr(1, Msg, vbCr)) Then
                Select Case pPoint(lngIndex)
                    Case 57
                        With dbPlayers(GetPlayerIndexNumber(lngIndex))
                            .sAppearance = .lAppStep & ":"
                            .lAppStep = 0
                        End With
                        pPoint(lngIndex) = 58
                        sSend lngIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your hair color; & newline & " & ";" & ColorLst(0) & ";", ""
                    Case 58
                        With dbPlayers(GetPlayerIndexNumber(lngIndex))
                            .sAppearance = .sAppearance & .lAppStep & ":"
                            .lAppStep = 0
                        End With
                        pPoint(lngIndex) = 59
                        sSend lngIndex, ";" & ANSICLS & "; & color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your hair style; & newline & " & ";" & HairStyle(0) & ";", ""
                    Case 59
                        With dbPlayers(GetPlayerIndexNumber(lngIndex))
                            .sAppearance = .sAppearance & .lAppStep & ":"
                            .lAppStep = 0
                        End With
                        pPoint(lngIndex) = 60
                        sSend lngIndex, ";" & ANSICLS & "; & color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your eye color; & newline & " & ";" & ColorLst(0) & ";", ""
                    Case 60
                        With dbPlayers(GetPlayerIndexNumber(lngIndex))
                            .sAppearance = .sAppearance & .lAppStep & ":"
                            .lAppStep = 0
                        End With
                        pPoint(lngIndex) = 61
                        sSend lngIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your moustache style; & newline & color.bgred & ;None; & newline & color.white & ;Normal; & newline & ;Box Car; & newline & ;Bullet Heads; & newline & ;Horse Shoe; & newline & ;Regent; & newline & ;Shermanic" & SetMoveCursor(4, 1) & ";", ""
                    Case 61
                        With dbPlayers(GetPlayerIndexNumber(lngIndex))
                            .sAppearance = .sAppearance & .lAppStep & ":"
                            .lAppStep = 0
                        End With
                        pPoint(lngIndex) = 62
                        sSend lngIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your beard style; & newline & color.bgred & ;None; & newline & color.white & ;Short Stubble; & newline & ;Bushy; & newline & ;Medium Length and Straight; & newline & ;Long and Curly; & newline & ;Long and Raspy; & newline & ;Medium Length and Curly" & SetMoveCursor(4, 1) & ";", ""
                    Case 62
                        With dbPlayers(GetPlayerIndexNumber(lngIndex))
                            .sAppearance = .sAppearance & .lAppStep
                            .lAppStep = 0
                        End With
                        pPoint(lngIndex) = 2
                        modLogon.LogOnSequence lngIndex
                End Select
            Else
                Exit Sub
            End If
    End Select
    If PlayerHasEcho(lngIndex, dbIndex) = True Then
        If Msg <> Chr$(vbKeyBack) And Msg <> vbCrLf Then
            ws(lngIndex).SendData Msg
        End If
    End If
    Exit Sub
End If
If (InStr(1, Msg, vbCrLf) Or InStr(1, Msg, vbLf) Or InStr(1, Msg, vbCr)) Then
    If Msg <> X(lngIndex) Then X(lngIndex) = X(lngIndex) & Msg
    X(lngIndex) = ReplaceFast(X(lngIndex), vbCrLf, "")
    X(lngIndex) = ReplaceFast(X(lngIndex), vbCr, "")
    X(lngIndex) = ReplaceFast(X(lngIndex), vbLf, "")
    pWeapon(lngIndex).lCol = 1
    LogOnSequence lngIndex
    dbIndex = GetPlayerIndexNumber(lngIndex)
    If dbIndex <> 0 Then
        With dbPlayers(dbIndex)
            If .lQueryTimer = 1 Then StartTimer
            .lCanClear = 0
            If .lHP <= 0 Then
                .iDropped = 1
                .dMonsterID = 99999
                If .iHasSentDropped <> 1 Then
                    WrapAndSend lngIndex, BGRED & "You fall to the ground!" & WHITE & vbCrLf
                    SendToAllInRoom lngIndex, BGRED & .sPlayerName & BRIGHTRED & " falls to the ground!" & WHITE & vbCrLf, .lLocation
                    .iHasSentDropped = 1
                End If
            Else
                .iDropped = 0
                If .iHasSentDropped = 1 Then
                    WrapAndSend lngIndex, BRIGHTWHITE & "You wounds aren't so bad anymore." & WHITE & vbCrLf
                    SendToAllInRoom lngIndex, BRIGHTWHITE & .sPlayerName & " gets up off the ground!" & WHITE & vbCrLf, .lLocation
                End If
                .iHasSentDropped = 0
            End If
            If .iStun > 0 Then
                WrapAndSend lngIndex, BRIGHTYELLOW & "You are stunned!" & WHITE & vbCrLf
                SendToAllInRoom lngIndex, YELLOW & .sPlayerName & " appears to be stuned!" & WHITE & vbCrLf, .lLocation
                X(lngIndex) = ""
                Exit Sub
            End If
        End With
    End If
    If SetEcho(lngIndex) = True Then Exit Sub
    Dim b As Boolean
    Dim bRet As Boolean
    Dim TInt As Long
    If dbIndex > 0 Then
        bRet = sScripting(lngIndex, dbPlayers(dbIndex).lLocation, , , , True, TInt, b)
        If b Then
            If sScripting(lngIndex, dbPlayers(dbIndex).lLocation) = True Then
                If dbPlayers(dbIndex).lCanClear = 0 Then
                    X(lngIndex) = ""
                    Exit Sub
                End If
            End If
        End If
    End If
    If dbIndex = 0 Then
        If DoCommands(lngIndex) = True Then Exit Sub
    Else
        If dbPlayers(dbIndex).iDropped > 0 Then
            If DoCommands(lngIndex, True, dbIndex) = True Then Exit Sub
        Else
            If DoCommands(lngIndex, , dbIndex) = True Then Exit Sub
            
        End If
    End If
    
    X(lngIndex) = ""
ElseIf Left$(Msg, 1) = "" And Right$(Msg, 1) = "R" Then
    pWeapon(lngIndex).lCol = Val(Mid$(Msg, InStr(1, Msg, ";") + 1, InStr(1, Msg, "R")))
    Msg = ""
End If
SkipAll:
If Msg = "" Then Exit Sub
If Asc(Msg) = vbKeyBack Then
    If Not modSC.FastStringComp(X(lngIndex), "") Then
        X(lngIndex) = Left$(X(lngIndex), Len(X(lngIndex)) - 1)
    End If
Else
    X(lngIndex) = X(lngIndex) & Msg
End If
If X(lngIndex) = vbCrLf Then X(lngIndex) = ""
If (pLogOnPW(lngIndex) = True And X(lngIndex) <> "" And X(lngIndex) <> vbCrLf) Or (pPoint(lngIndex) = 55 And X(lngIndex) <> "" And X(lngIndex) <> vbCrLf) Then ws(lngIndex).SendData MOVELEFTONE & "*" 'MOVELEFT23 & ReplaceFast(MOVERIGHTNUM, "#", "23") & String(Len(x(lngIndex)), "*")
If PlayerHasEcho(lngIndex, dbIndex) = True Then
        If Msg <> Chr$(vbKeyBack) And Msg <> vbCrLf Then
            ws(lngIndex).SendData Msg
        End If
    End If

On Error GoTo 0
Exit Sub
ws_DataArrival_Error:
UpdateList "}b}u}An error occured on line " & lngIndex & " at " & Time & ".", True
UpdateList "   }bError: " & Err.Number & "-", True
UpdateList "        }i" & Err.Description, True
X(lngIndex) = ""
End Sub

Sub LoadWinSocks()
Dim a As Long
For a = 1 To MaxUsers
    Load ws(a)
    ws(a).LocalPort = 23
    ws(a).Protocol = sckTCPProtocol
    lstUsers.AddItem "[Line " & CStr(a) & " - Open]"
Next a
End Sub

Private Sub ws_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
ws(Index).Close
End Sub
