VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2805
   ClientLeft      =   3210
   ClientTop       =   2280
   ClientWidth     =   4440
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "frmSplash.frx":08CA
   ScaleHeight     =   2805
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblPer 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " Loading () [0%] ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   4380
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " Please wait... "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : frmSplash
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Sub LoadServer()
Dim i As Long
lblPer.Caption = " Loading (Server Window) [3%] ..."
Load frmMain

lblPer.Caption = " Loading (Setting defaults) [6%] ..."
With frmMain
    .timTimer.Enabled = False
    .timCombat.Enabled = False
    .timPRest.Enabled = False
    .Caption = "Loading Depths of Despair MUD Server..."
    .Show
    .Enabled = False
End With

lblPer.Caption = " Loading (Restoring Saves) [9%] ..."
modDatabase.InitRecordsets
modTime.SetNameArrays
modTime.SetTimeOfDay
modTime.SetDayOfWeek
modTime.LoadMonths
modTime.SetMonthOfYear
modTime.SetYear

lblPer.Caption = " Loading (Preparing Data) [12%] ..."
CheckSpecialItems
RemoveCorpses
RemoveOutDoorFood
RemoveNormalFood

lblPer.Caption = " Loading (Loading Database) [15%] ..."
LoadDatabaseIntoMemory

lblPer.Caption = " Loading (Setting Defaults) [57%] ..."
MaxUsers = CLng(GetINI("MaxPlayers"))
MaxMonsters = CLng(GetINI("MaxMonsters"))

lblPer.Caption = " Loading (Expanding Memory) [61%] ..."
ReDim SpellCombat(MaxUsers) As Boolean
ReDim X(MaxUsers) As String
ReDim PNAME(MaxUsers) As String

lblPer.Caption = " Loading (Resizing Arrays) [65%] ..."
ReDim pLogOn(MaxUsers) As Boolean
ReDim pLogOnPW(MaxUsers) As Boolean
ReDim pPoint(MaxUsers) As Long
ReDim pWeapon(MaxUsers) As Weapon
ReDim aMons(MaxMonsters) As Monster

lblPer.Caption = " Loading (Checking Port Status) [69%] ..."
On Error Resume Next
frmMain.ws(0).Listen
If Err.Number = 10048 Then
    If GetINI("ChoosePort") = "True" Then
        With frmMain
            .ws(0).LocalPort = lPort
            .ws(0).Listen
        End With
    Else
        AlwaysOnTop Me, False
        MsgBox "The default port (23) is already in use. Please specify another port in the next prompt.", vbExclamation + vbOKOnly, "Port In Use"
        frmPort.Show 1
    End If
End If

lblPer.Caption = " Loading (Getting IP) [73%] ..."
frmMain.lblIP.Caption = GetIPAddress

lblPer.Caption = " Loading (Loading Sockets) [77%] ..."
frmMain.LoadWinSocks

lblPer.Caption = " Loading (Defining Active Monsters) [84%] ..."
FillType

lblPer.Caption = " Loading (Setting Properties) [91%] ..."
With frmMain
    .lblIP.Caption = .lblIP.Caption & " (" & .ws(0).LocalPort & ")"
    .Caption = "Depths of Despair"
    .Enabled = True
    .timTimer.Enabled = True
    .timCombat.Enabled = True
    .timPRest.Enabled = True
End With

lblPer.Caption = " Loading (Configuring Player Statistics) [98%] ..."
For i = LBound(dbPlayers) To UBound(dbPlayers)
    modMiscFlag.RedoStatsPlus i
    'modGetData.GetPlayerSwings i
Next

lblPer.Caption = " Loading (Loading Entrance Graphic) [100%] ..."
Open App.Path & "\intgrp.ansi" For Binary As #1
    sGraphic = Input$(LOF(1), 1)
Close #1

AlwaysOnTop Me, False
UpdateList "}b}iReady to service users. (}n}i" & Time & "}n}b)"
Screen.MousePointer = vbDefault
Unload Me
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width \ 2) - (Me.Width \ 2), (Screen.Height \ 2) - (Me.Height \ 2)
AlwaysOnTop Me, True
End Sub
