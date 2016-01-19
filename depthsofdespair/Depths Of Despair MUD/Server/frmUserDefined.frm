VERSION 5.00
Begin VB.Form frmUserDefined 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User Defined Settings"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserDefined.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic 
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Index           =   6
      Left            =   3360
      ScaleHeight     =   2235
      ScaleWidth      =   4995
      TabIndex        =   38
      Top             =   240
      Visible         =   0   'False
      Width           =   5055
      Begin VB.OptionButton optDeath 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Penalties"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   42
         Top             =   1560
         Width           =   1695
      End
      Begin VB.OptionButton optDeath 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Death"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   41
         Top             =   1920
         Width           =   1695
      End
      Begin VB.OptionButton optDeath 
         BackColor       =   &H00E0E0E0&
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   $"frmUserDefined.frx":08CA
         Height          =   975
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   4815
      End
   End
   Begin DoDMudServer.UltraBox lstSettings 
      Height          =   2775
      Left            =   240
      TabIndex        =   37
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4895
      Style           =   3
      Fill            =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mult            =   0   'False
      Sort            =   0   'False
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Index           =   5
      Left            =   3360
      ScaleHeight     =   2235
      ScaleWidth      =   4995
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   5055
      Begin DoDMudServer.NumOnlyText IPa 
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   21
         Top             =   1080
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         AllowNeg        =   0   'False
         Align           =   2
         MaxLength       =   0
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin DoDMudServer.NumOnlyText IPa 
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   22
         Top             =   1080
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         AllowNeg        =   0   'False
         Align           =   2
         MaxLength       =   0
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin DoDMudServer.NumOnlyText IPa 
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   23
         Top             =   1080
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         AllowNeg        =   0   'False
         Align           =   2
         MaxLength       =   0
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin DoDMudServer.NumOnlyText IPa 
         Height          =   375
         Index           =   3
         Left            =   3720
         TabIndex        =   24
         Top             =   1080
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         AllowNeg        =   0   'False
         Align           =   2
         MaxLength       =   0
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin DoDMudServer.eButton cmdAdd 
         Height          =   375
         Left            =   3720
         TabIndex        =   28
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Style           =   2
         Cap             =   "&Add"
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
      Begin DoDMudServer.eButton cmdSL 
         Height          =   375
         Left            =   2400
         TabIndex        =   29
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Style           =   2
         Cap             =   "&See List"
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
      Begin VB.Label lblDot 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   2
         Left            =   3480
         TabIndex        =   27
         Top             =   840
         Width           =   195
      End
      Begin VB.Label lblDot 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   1
         Left            =   2400
         TabIndex        =   26
         Top             =   840
         Width           =   195
      End
      Begin VB.Label lblDot 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   0
         Left            =   1320
         TabIndex        =   25
         Top             =   840
         Width           =   195
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enter the IP address of the person you do not want to access your MUD server."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Index           =   4
      Left            =   3360
      ScaleHeight     =   2235
      ScaleWidth      =   4995
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   5055
      Begin DoDMudServer.NumOnlyText txtLogons 
         Height          =   495
         Left            =   1320
         TabIndex        =   31
         Top             =   1440
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         AllowNeg        =   0   'False
         Align           =   2
         MaxLength       =   0
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "You may restrict the amount of times 1 IP address can connect to your server. The recommended amount is 1."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Index           =   3
      Left            =   3360
      ScaleHeight     =   2235
      ScaleWidth      =   4995
      TabIndex        =   13
      Top             =   240
      Width           =   5055
      Begin DoDMudServer.NumOnlyText txtDeathLevel 
         Height          =   495
         Left            =   1440
         TabIndex        =   15
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         AllowNeg        =   -1  'True
         Align           =   2
         MaxLength       =   0
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   $"frmUserDefined.frx":09D8
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   4815
      End
   End
   Begin DoDMudServer.Raise Raise1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5318
      Style           =   2
      Color           =   14737632
   End
   Begin DoDMudServer.eButton cmdOK 
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Style           =   2
      Cap             =   "&OK"
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
   Begin DoDMudServer.eButton cmdCancel 
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Style           =   2
      Cap             =   "&Cancel"
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
   Begin VB.PictureBox Pic 
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Index           =   2
      Left            =   3360
      ScaleHeight     =   2235
      ScaleWidth      =   4995
      TabIndex        =   32
      Top             =   240
      Width           =   5055
      Begin DoDMudServer.NumOnlyText txtPVP 
         Height          =   400
         Left            =   1800
         TabIndex        =   33
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         AllowNeg        =   0   'False
         Align           =   2
         MaxLength       =   0
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin VB.CheckBox chkPvP 
         BackColor       =   &H00E0E0E0&
         Caption         =   $"frmUserDefined.frx":0AAE
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   35
         Top             =   120
         Width           =   4095
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   0
         X2              =   5040
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   $"frmUserDefined.frx":0B38
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   0
         TabIndex        =   34
         Top             =   980
         Width           =   4935
      End
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Index           =   1
      Left            =   3360
      ScaleHeight     =   2235
      ScaleWidth      =   4995
      TabIndex        =   11
      Top             =   240
      Width           =   5055
      Begin DoDMudServer.NumOnlyText txtMonsters 
         Height          =   495
         Left            =   1440
         TabIndex        =   16
         Top             =   1560
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         AllowNeg        =   0   'False
         Align           =   2
         MaxLength       =   0
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   $"frmUserDefined.frx":0BEE
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Index           =   0
      Left            =   3360
      ScaleHeight     =   2235
      ScaleWidth      =   4995
      TabIndex        =   2
      Top             =   240
      Width           =   5055
      Begin DoDMudServer.NumOnlyText NumOnlyText1 
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         AllowNeg        =   0   'False
         Align           =   0
         MaxLength       =   0
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Custom:"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   1800
         Width           =   975
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   "20"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   "15"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   "10"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   "5"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Choose the maximum amount of players that can be connected to the game at one time."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4815
      End
   End
   Begin DoDMudServer.Raise Raise2 
      Height          =   2520
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4445
      Style           =   2
      Color           =   14737632
   End
   Begin DoDMudServer.Raise rasMain 
      Height          =   1215
      Left            =   0
      TabIndex        =   36
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2143
      Style           =   2
      Color           =   14737632
   End
End
Attribute VB_Name = "frmUserDefined"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : frmUserDefined
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'

Private lDeathOpt As Long

Private Sub chkPvP_Click()
Select Case chkPvP.Value
    Case 1
        txtPVP.Enabled = True
    Case Else
        txtPVP.Enabled = False
End Select
End Sub

Private Sub cmdAdd_Click()
Dim i As Long
AddIPToList IPa(0).Text & "." & IPa(1).Text & "." & IPa(2).Text & "." & IPa(3).Text
For i = 0 To 3
    IPa(i).Text = "0"
    If DE Then DoEvents
Next
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim i As Long
For i = Opt.LBound To Opt.ubound
    If Opt(i).Value = True Then
        Select Case i
            Case 0
                WriteINI "MaxPlayers", "5"
            Case 1
                WriteINI "MaxPlayers", "10"
            Case 2
                WriteINI "MaxPlayers", "15"
            Case 3
                WriteINI "MaxPlayers", "20"
            Case 4
                WriteINI "MaxPlayers", txtCustom.Text
        End Select
    End If
Next
WriteINI "MaxMonsters", txtMonsters.Text
WriteINI "DeathLevel", txtDeathLevel.Text
WriteINI "Logons", txtLogons.Text
WriteINI "PvPE", chkPvP.Value
WriteINI "PvPL", txtPVP.Text
WriteINI "Age", CStr(lDeathOpt)

lDeath = Val(GetINI("DeathLevel"))
lIsPvP = Val(GetINI("PvPE"))
lPvPLevel = Val(GetINI("PvPL"))
lAgeD = lDeathOpt
Unload Me
End Sub

Private Sub cmdSL_Click()
Load frmIPS
frmIPS.Show
End Sub

Private Sub Form_Load()
rasMain.Top = 0
rasMain.Left = 0
rasMain.Width = Me.ScaleWidth
rasMain.Height = Me.ScaleHeight
lstSettings.SetSelected 1, True
lstSettings_Click
Select Case GetINI("MaxPlayers")
    Case "5"
        Opt(0).Value = True
    Case "10"
        Opt(1).Value = True
    Case "15"
        Opt(2).Value = True
    Case "20"
        Opt(3).Value = True
    Case Else
        Opt(4).Value = True
        txtCustom.Text = GetINI("MaxPlayers")
End Select
txtMonsters.Text = GetINI("MaxMonsters")
txtDeathLevel.Text = GetINI("DeathLevel")
txtLogons.Text = GetINI("Logons")
chkPvP.Value = Val(GetINI("PvPE"))
chkPvP_Click
txtPVP.Text = GetINI("PvPL")
lDeathOpt = Val(GetINI("Age"))
optDeath(lDeathOpt).Value = True
With lstSettings
    .AddItem "}3»}r }nMax Players"
    .AddItem "}3»}r }nMax Monsters"
    .AddItem "}3»}r }nP. vs. P."
    .AddItem "}3»}r }nMax Death Level"
    .AddItem "}3»}r }nMax Logons Per IP"
    .AddItem "}3»}r }nBlocked IPs"
    .AddItem "}3»}r }nAge Settings"
    .AddItem "}3»}r }nLogon Graphics"
End With
lstSettings.SetSelected 1, True
lstSettings_Click
End Sub

Private Sub IPa_Change(Index As Integer)
If Len(IPa(Index).Text) = 3 Then
    If Index <> IPa.ubound Then
        IPa(Index + 1).SetFocus
    Else
        cmdAdd.SetFocus
    End If
End If
End Sub

Private Sub lstSettings_Click()
Dim i As Long
lstSettings.Paint = False
For i = 1 To lstSettings.ListCount
    lstSettings.SetItemText i, Replace$(lstSettings.List(i, True), "}b", "}n")
Next
lstSettings.SetItemText lstSettings.ListIndex, Replace$(lstSettings.List(lstSettings.ListIndex, True), "}n", "}b")
lstSettings.Paint = True
If lstSettings.ListIndex = 8 Then
    Load frmGraphics
    frmGraphics.Show
ElseIf lstSettings.ListIndex > 0 Then
    For i = Pic.LBound To Pic.ubound
        Pic(i).Visible = False
    Next
    Pic(lstSettings.ListIndex - 1).Visible = True
End If
End Sub

Sub AddIPToList(s As String)
    Dim FileNumber&
    FileNumber = FreeFile
    Open App.Path & "\ipb.list" For Append Shared As #FileNumber
        If DE Then DoEvents
        Print #FileNumber, s
    Close #FileNumber
End Sub


Private Sub optDeath_Click(Index As Integer)
lDeathOpt = Index
End Sub
