VERSION 5.00
Begin VB.Form frmRoomJoin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Join To Room"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   835
      Left            =   3360
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   13
      Top             =   1680
      Width           =   1387
   End
   Begin ServerEditor.FlagOptions flgOpts 
      Height          =   315
      Left            =   2040
      TabIndex        =   12
      Top             =   360
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   556
      Style           =   3
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      Default         =   -1  'True
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   3840
      Width           =   975
   End
   Begin VB.OptionButton optUp 
      Caption         =   "U&p"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.OptionButton optDown 
      Caption         =   "&Down"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.OptionButton optEast 
      Caption         =   "&East"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton optWest 
      Caption         =   "&West"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.OptionButton optNortheast 
      Caption         =   "N&ortheast"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.OptionButton optNorthwest 
      Caption         =   "No&rthwest"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.OptionButton optSoutheast 
      Caption         =   "So&utheast"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.OptionButton optSouthwest 
      Caption         =   "Sou&thwest"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.OptionButton optSouth 
      Caption         =   "&South"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.OptionButton optNorth 
      Caption         =   "&North"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   1455
   End
End
Attribute VB_Name = "frmRoomJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sJoinDir As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim i As Long
With udtMapArea(frmMapEdit.prevSel)
    .lJoinRoom = flgOpts.GetCurVal
    .sJoinExit = sJoinDir
    .sExits = .sExits & ":" & sJoinDir & ";"
End With
Select Case sJoinDir
    Case "n"
        i = frmMapEdit.prevSel - 30
    Case "s"
        i = frmMapEdit.prevSel + 30
    Case "e"
        i = frmMapEdit.prevSel + 1
    Case "w"
        i = frmMapEdit.prevSel - 1
    Case "nw"
        i = frmMapEdit.prevSel - 31
    Case "sw"
        i = frmMapEdit.prevSel + 29
    Case "ne"
        i = frmMapEdit.prevSel - 29
    Case "se"
        i = frmMapEdit.prevSel + 31
End Select
frmMapEdit.DrawExit i, sJoinDir
With udtMapArea(i)
    .lJoinRoom = flgOpts.GetCurVal
    .lAlreadyExist = 1
    .sIsRoom = True
    Select Case sJoinDir
        Case "n"
            sJoinDir = "s"
        Case "s"
            sJoinDir = "n"
        Case "e"
            sJoinDir = "w"
        Case "w"
            sJoinDir = "e"
        Case "nw"
            sJoinDir = "se"
        Case "sw"
            sJoinDir = "ne"
        Case "ne"
            sJoinDir = "sw"
        Case "se"
            sJoinDir = "nw"
    End Select
    .sExits = .sExits & ":" & sJoinDir & ";"
End With
frmMapEdit.DrawRm i
Unload Me
End Sub

Private Sub Form_Load()
modMain.FeedAList flgOpts, "rooms"
sJoinDir = "n"
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMapDef.Enabled = True
frmMapEdit.Enabled = True
End Sub

Private Sub optDown_Click()
sJoinDir = "d"
End Sub

Private Sub optEast_Click()
sJoinDir = "e"
End Sub

Private Sub optNorth_Click()
sJoinDir = "n"
End Sub

Private Sub optNortheast_Click()
sJoinDir = "ne"
End Sub

Private Sub optNorthwest_Click()
sJoinDir = "nw"
End Sub

Private Sub optSouth_Click()
sJoinDir = "s"
End Sub

Private Sub optSoutheast_Click()
sJoinDir = "se"
End Sub

Private Sub optSouthwest_Click()
sJoinDir = "sw"
End Sub

Private Sub optUp_Click()
sJoinDir = "u"
End Sub

Private Sub optWest_Click()
sJoinDir = "w"
End Sub
