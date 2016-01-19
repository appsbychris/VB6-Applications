VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic Game Board"
   ClientHeight    =   8820
   ClientLeft      =   150
   ClientTop       =   165
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Ticboard1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Ticboard1.frx":08CA
   ScaleHeight     =   8820
   ScaleWidth      =   11820
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSubMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   6480
      Picture         =   "Ticboard1.frx":A074
      ScaleHeight     =   975
      ScaleWidth      =   2415
      TabIndex        =   29
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   0
         X2              =   2880
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   2
         X1              =   -1
         X2              =   2890
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lblSubMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "&Play 1 player"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         MouseIcon       =   "Ticboard1.frx":11B9A
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   600
         Width           =   2280
      End
      Begin VB.Label lblSubMenu 
         BackColor       =   &H0000FF00&
         Caption         =   "&Connect to remote server"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   10
         MouseIcon       =   "Ticboard1.frx":11EA4
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   120
         Width           =   2280
      End
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   5160
      Picture         =   "Ticboard1.frx":121AE
      ScaleHeight     =   2175
      ScaleWidth      =   2850
      TabIndex        =   20
      Top             =   1680
      Visible         =   0   'False
      Width           =   2850
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   0
         X2              =   2880
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   2890
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   0
         X2              =   2880
         Y1              =   1700
         Y2              =   1700
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   0
         X1              =   -15
         X2              =   2875
         Y1              =   1700
         Y2              =   1700
      End
      Begin VB.Label lblMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "E&xit (CTRL + X)"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   10
         MouseIcon       =   "Ticboard1.frx":265EC
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label lblMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "&About"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   10
         MouseIcon       =   "Ticboard1.frx":268F6
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label lblMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "&Help Index"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   10
         MouseIcon       =   "Ticboard1.frx":26C00
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label lblMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "&Rules of the game"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   10
         MouseIcon       =   "Ticboard1.frx":26F0A
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblMenu 
         BackColor       =   &H0000FF00&
         Caption         =   "&Connect                        ->"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   10
         MouseIcon       =   "Ticboard1.frx":27214
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   120
         Width           =   2775
      End
   End
   Begin VB.TextBox txtTalk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   19
      ToolTipText     =   "Type in here to chat."
      Top             =   2520
      Width           =   5295
   End
   Begin VB.PictureBox picTextPics 
      AutoRedraw      =   -1  'True
      Height          =   4095
      Left            =   0
      ScaleHeight     =   269
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   789
      TabIndex        =   18
      Top             =   8760
      Width           =   11895
      Begin VB.Image imgTxtTalk 
         Height          =   375
         Left            =   240
         Top             =   3000
         Width           =   10335
      End
      Begin VB.Image imgTxtChat 
         Height          =   2535
         Left            =   240
         Top             =   120
         Width           =   11535
      End
   End
   Begin VB.PictureBox Frame2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5955
      Left            =   1200
      Picture         =   "Ticboard1.frx":2751E
      ScaleHeight     =   397
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   789
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   11835
      Begin VB.ListBox lstDiscard 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         ForeColor       =   &H00000000&
         Height          =   495
         IntegralHeight  =   0   'False
         Left            =   9600
         MouseIcon       =   "Ticboard1.frx":3293F
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   3480
         Width           =   1815
      End
      Begin VB.ListBox lstCurrent 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         ForeColor       =   &H00000000&
         Height          =   1335
         IntegralHeight  =   0   'False
         ItemData        =   "Ticboard1.frx":32D81
         Left            =   9720
         List            =   "Ticboard1.frx":32D88
         MouseIcon       =   "Ticboard1.frx":32D98
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   720
         Width           =   1650
      End
      Begin VB.PictureBox imgTICSetup 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1440
         Index           =   0
         Left            =   240
         ScaleHeight     =   1440
         ScaleWidth      =   1065
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.ListBox lstDiscardMask 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         ForeColor       =   &H00000000&
         Height          =   735
         IntegralHeight  =   0   'False
         Left            =   4440
         MouseIcon       =   "Ticboard1.frx":331DA
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   3000
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ListBox lstCardNames 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         ForeColor       =   &H00000000&
         Height          =   1335
         IntegralHeight  =   0   'False
         Left            =   4440
         MouseIcon       =   "Ticboard1.frx":3361C
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   3840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   3015
         Left            =   0
         ScaleHeight     =   197
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   285
         TabIndex        =   12
         Top             =   3120
         Width           =   4335
         Begin VB.Image Image2 
            Height          =   1095
            Left            =   0
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Image Image1 
            Height          =   1695
            Left            =   0
            Top             =   0
            Width           =   2535
         End
      End
      Begin VB.ListBox lstDone 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4440
         TabIndex        =   17
         Top             =   5280
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label cmdCancel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "         C&ancel"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   8640
         MouseIcon       =   "Ticboard1.frx":33A5E
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Image imgHCancel 
         Height          =   1320
         Left            =   8520
         Picture         =   "Ticboard1.frx":33EA0
         Top             =   4440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label cmdDone 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "         &Done"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   10200
         MouseIcon       =   "Ticboard1.frx":34B92
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Image imgHDone 
         Height          =   1320
         Left            =   10080
         Picture         =   "Ticboard1.frx":34FD4
         Top             =   4440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label cmdRedo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "         &Redo"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   6960
         MouseIcon       =   "Ticboard1.frx":35CC6
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label cmdSet 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Aside"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   9600
         MouseIcon       =   "Ticboard1.frx":36108
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "                              Set for Group"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   645
         Index           =   0
         Left            =   7080
         TabIndex        =   6
         Top             =   240
         Width           =   2160
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "                                 Discard this card"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   720
         Index           =   1
         Left            =   7080
         TabIndex        =   9
         Top             =   3240
         Width           =   2115
      End
      Begin VB.Image imghSet 
         Height          =   615
         Left            =   9480
         Picture         =   "Ticboard1.frx":36412
         Top             =   2640
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Image imgHRedo 
         Height          =   1320
         Left            =   6840
         Picture         =   "Ticboard1.frx":36E99
         Top             =   4440
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.PictureBox Frame1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5955
      Left            =   0
      Picture         =   "Ticboard1.frx":37B8B
      ScaleHeight     =   5955
      ScaleWidth      =   11835
      TabIndex        =   2
      Top             =   2880
      Width           =   11835
      Begin VB.PictureBox imgHand 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1440
         Index           =   0
         Left            =   240
         ScaleHeight     =   1440
         ScaleWidth      =   1065
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label cmdTIC 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Setup for TIC "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         MouseIcon       =   "Ticboard1.frx":46DC9
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Discard, Drag into this shape-"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8880
         TabIndex        =   3
         Top             =   240
         Width           =   2745
      End
      Begin VB.Image imgTrash 
         Appearance      =   0  'Flat
         Height          =   5610
         Left            =   8760
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   120
         Width           =   3000
      End
      Begin VB.Image imgHSetup 
         Height          =   660
         Left            =   300
         Picture         =   "Ticboard1.frx":470D3
         Top             =   5000
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   4200
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   3220
   End
   Begin TicBoard.Deck Deck1 
      Left            =   4080
      Top             =   480
      _ExtentX        =   1429
      _ExtentY        =   2090
      Picture         =   "Ticboard1.frx":4B3C5
   End
   Begin VB.TextBox txtChat 
      Appearance      =   0  'Flat
      BackColor       =   &H0034314A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5160
      MultiLine       =   -1  'True
      TabIndex        =   26
      ToolTipText     =   "Chat Screen"
      Top             =   240
      Width           =   6015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deck of Cards         Discard Pile"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   33
      Top             =   1920
      Width           =   3300
   End
   Begin VB.Label lblRound 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11160
      TabIndex        =   32
      ToolTipText     =   "The current round."
      Top             =   480
      Width           =   615
   End
   Begin VB.Image imgFile 
      Appearance      =   0  'Flat
      Height          =   200
      Left            =   0
      Picture         =   "Ticboard1.frx":4B3E1
      Stretch         =   -1  'True
      ToolTipText     =   "Click for menu"
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      MouseIcon       =   "Ticboard1.frx":4EF83
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Click to send text."
      Top             =   2520
      Width           =   855
   End
   Begin VB.Image imgHold 
      Height          =   495
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblTurn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "Active Player"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Image imgDiscard 
      Enabled         =   0   'False
      Height          =   1440
      Left            =   2880
      MouseIcon       =   "Ticboard1.frx":4F28D
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      ToolTipText     =   "Top discarded card"
      Top             =   360
      Width           =   1065
   End
   Begin VB.Image imgDeck 
      Enabled         =   0   'False
      Height          =   1440
      Left            =   720
      MouseIcon       =   "Ticboard1.frx":4F3DF
      MousePointer    =   99  'Custom
      Picture         =   "Ticboard1.frx":4F531
      Stretch         =   -1  'True
      ToolTipText     =   "Draw a card"
      Top             =   360
      Width           =   1065
   End
   Begin VB.Image imgHSend 
      Height          =   450
      Left            =   10860
      Picture         =   "Ticboard1.frx":50144
      Top             =   2440
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgRoundPic 
      Height          =   1125
      Left            =   11100
      Picture         =   "Ticboard1.frx":506A8
      Top             =   0
      Width           =   750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************
'*************************************************************************************
'***************       Code create by Chris Van Hooser          **********************
'***************                  (c)2001                       **********************
'*************** You may use this code and freely distribute it **********************
'***************   If you have any questions, please email me   **********************
'***************          at theendorbunker@attbi.com.          **********************
'***************       Thanks for downloading my project        **********************
'***************        and i hope you can use it well.         **********************
'***************                TicBoard                        **********************
'***************                TicBoard.vbp                    **********************
'*************************************************************************************
'*************************************************************************************

'*************************************************************
'Make a variable for my class module**************************
'*************************************************************
Dim FU As New Functions
'*************************************************************

'*************************************************************
'Globals for moveing the cards
'*************************************************************
Dim orgX As Single
Dim orgY As Single

'*************************************************************
'Variables used for the basic game functions******************
'*************************************************************
Dim Round$ 'The current round of the game
Dim TempHand(15) As String 'hand use temperaraly to store your
                           'hand
'*************************************************************

'*************************************************************
'Basic counter************************************************
'*************************************************************
Dim i As Integer 'used mainly for the ubound of imghand
'*************************************************************

'*************************************************************
'Used for dragging cards around*******************************
'*************************************************************
Dim bDown As Boolean
Dim bMove As Boolean
Dim ADown As Boolean
Dim AMove As Boolean
'*************************************************************

'*************************************************************
'True/False variable to see if the game has ended*************
'*************************************************************
Dim GameEnd As Boolean
'*************************************************************

'*************************************************************
'Check either a straight or a threeofakind********************
'*************************************************************
Dim CheckWhat$
'*************************************************************

'*************************************************************
'True/False value to determine if it is the last turn of the**
'round********************************************************
'*************************************************************
Dim LastTurn As Boolean
'*************************************************************

'*************************************************************
'Used to store the card you are trying to cheat with**********
'*************************************************************
Dim BackUpBuffer As Integer
'*************************************************************

'*************************************************************
'True/false value for determining if the menu is drawing******
'or not
'*************************************************************
Dim DrawingMenu As Boolean
'*************************************************************

Private Sub cmdCancel_Click()
On Error GoTo cmdCancel_Click_Error
'*************************************************************
'Clear out everything in cardsinhand**************************
'*************************************************************
Erase CardsInHand
'*************************************************************

'*************************************************************
'Default whatever is in cardsinhand to -1 then set it*********
'to the temparary hand that we stored*************************
'*************************************************************
For i = 0 To UBound(CardsInHand)
    CardsInHand(i) = -1
    CardsInHand(i) = TempHand(i)
Next

'*************************************************************
'Call the RedoHand sub, and default the values back***********
'*************************************************************
Call RedoHand
lstDiscard.Clear
lstCurrent.Clear
lstDiscardMask.Clear
lstCardNames.Clear
lstDone.Clear
Frame2.Visible = False
'*************************************************************
On Error GoTo 0
Exit Sub
cmdCancel_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdCancel_Click in Form, Form1"
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo cmdCancel_MouseMove_Error
'**************************************************
'Call the make all highted buttons invisable*******
'sub, and highlight the cancel button**************
'**************************************************
Call HAllHButtons
imgHCancel.Visible = True
cmdCancel.ForeColor = vbWhite
'**************************************************
On Error GoTo 0
Exit Sub
cmdCancel_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdCancel_MouseMove in Form, Form1"
End Sub

Private Sub cmdDone_Click()
On Error GoTo cmdDone_Click_Error
'make sure lstcurrent and listdiscard have nothing and 1 thing
'in each
Dim Counter% 'counter
'see how many cards are in their hand
For i = 0 To UBound(TempHand)
    If TempHand(i) <> -1 Then
        Counter% = Counter% + 1
    End If
Next
If Counter% < Val(Round$) + 1 Then Exit Sub
If lstCurrent.ListCount = 0 And lstDiscardMask.ListCount = 1 Then
    'If someone else hasnt tic'd out yet
    'make sure they have a card to discard
    If Counter% < Val(Round$) + 1 Then Exit Sub
    If LastTurn = False Then
        Dim tCount% 'get a variable as integer
        'loop through the hand and make sure it
        'is empty
        For i = 0 To UBound(CardsInHand)
            If CardsInHand(i) <> -1 Then
                MsgBox "Place all cards in the correct places.", vbCritical, "Error"
                Exit Sub
            End If
        Next
        'make everything in the hand default to -1
        For i = 0 To UBound(CardsInHand)
            CardsInHand(i) = -1
        Next
        'put the original hand back into the main hand
        'array
        For i = 0 To UBound(CardsInHand)
            CardsInHand(i) = TempHand(i)
        Next
        'Find out witch card was the discarded one
        'and get rid of it
        For i = 0 To UBound(CardsInHand)
            If CardsInHand(i) = Val(lstDiscardMask.List(0)) Then
                CardsInHand(i) = -1
                Exit For
            End If
        Next
        Dim b As Integer, a As Integer 'more variables
        'Get rid of the gaps in the hand..
        'ex, if the array is 0,-1,12,11...
        'this will turn it into 0,12,11,-1
        b = 0
        For a = 0 To 14
            If CardsInHand(a) <> -1 Then
                If a <> b Then
                    CardsInHand(b) = CardsInHand(a)
                    CardsInHand(a) = -1
                End If
                b = b + 1
            End If
        Next
        'Redraw the hand
        Call RedoHand
        'Make everything the default value
        For i = 0 To UBound(CardsInHand)
            CardsInHand(i) = -1
            TempHand(i) = -1
        Next
        'Clear out and go back to the main screen
        Frame2.Visible = False
        lstCurrent.Clear
        lstDiscard.Clear
        lstDone.Clear
        lstCardNames.Clear
        Frame1.Enabled = False
        'Send the info to the server that you have
        'tic'd out
        ws.SendData "»ïñ" & YName$
        'pause the prog for approx 400 miliseconds
        FU.WaitFor 1
        'Send the card you discarded
        ws.SendData "öõô" & lstDiscardMask.List(0)
        'Clear up some more
        lstDiscardMask.Clear
        'make it the last turn
        LastTurn = True
        'send the turn
        Call EndTurn
    Else
        Dim X%, T% ' Counting variables
        'set em to equal 0
        T% = 0
        X% = 0
        'begin a loop, looping through
        'cardsinhand
        For i = 0 To UBound(CardsInHand)
            'make x equal what cards in hand of i is
            X% = CardsInHand(i)
            'if it isn't the default 'empty' value
            'and contrinue
            If X% <> -1 Then
                'reduce the cards in hand to be a 0-12 value
                Select Case CardsInHand(i)
                    Case 13 To 25:
                        CardsInHand(i) = (CardsInHand(i) - 13)
                    Case 26 To 38:
                        CardsInHand(i) = (CardsInHand(i) - (13 * 2))
                    Case 39 To 51:
                        CardsInHand(i) = (CardsInHand(i) - (13 * 3))
                End Select
                'set x equal to the new value
                X% = CardsInHand(i)
                'and now assign the points
                'if the card is a wild card, make it 30 points
                If X% = Val(Round$) - 1 Then X% = 30: GoTo AddToT
                'if the card is J, Q, K, or 10 then make it 10 points
                If X% >= 9 Then X% = 10: GoTo AddToT
                'if the card is A then make it 15 points
                If X% = 0 Then X% = 15: GoTo AddToT
                'if the card is 2-9 then make it 5 points
                If X% < 9 Then X% = 5: GoTo AddToT
AddToT:
                'add the total of x to our
                'current total which is
                'stored in T
                T% = T% + X%
            End If
        Next
        'If the round is Q, K, A, or 2 then multiply
        'by the round point multiplyer
        If Round$ = "12" Then T% = T% * 2
        If Round$ = "13" Then T% = T% * 3
        If Round$ = "14" Then T% = T% * 4
        If Round$ = "15" Then T% = T% * 5
        'now send the total to the server
        'and send the discarded card to the server
        ws.SendData "¶¸·" & YName$ & "||" & T% & "**" & lstDiscardMask.List(0)
        'wait for 400 miliseconds
        FU.WaitFor 1
        'set the hand to default value
        For i = 0 To UBound(CardsInHand)
            CardsInHand(i) = -1
        Next
        'reset the hand to the original
        For i = 0 To UBound(CardsInHand)
            CardsInHand(i) = TempHand(i)
        Next
        'remove the card they discard
        For i = 0 To UBound(CardsInHand)
            If CardsInHand(i) = Val(lstDiscardMask.List(0)) Then
                CardsInHand(i) = -1
                Exit For
            End If
        Next
        'redraw the hand
        Call RedoHand
        'and set the defaults
        Frame2.Visible = False
        lstCurrent.Clear
        lstDiscard.Clear
        lstDone.Clear
        lstDiscardMask.Clear
        lstCardNames.Clear
        Frame1.Enabled = False
    End If
Else
    'error message if they didnt do something right
    MsgBox "Set aside all cards in the 'Set for group' box, or Place a card in the 'Discard this card' box.", vbCritical, "Error"
End If
On Error GoTo 0
Exit Sub
cmdDone_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdDone_Click in Form, Form1"
End Sub

Private Sub cmdDone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo cmdDone_MouseMove_Error
'Unhighlight all the buttons
Call HAllHButtons
'and make the done highlight visible
imgHDone.Visible = True
cmdDone.ForeColor = vbWhite
On Error GoTo 0
Exit Sub
cmdDone_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdDone_MouseMove in Form, Form1"
End Sub

Private Sub cmdRedo_Click()
On Error GoTo cmdRedo_Click_Error
'Clear out cardsinhand array
Erase CardsInHand
'default cardsinhand array
For i = 0 To UBound(CardsInHand)
    CardsInHand(i) = -1
Next
'reset cardinhand array to the original
For i = 0 To UBound(TempHand)
    CardsInHand(i) = TempHand(i)
Next
'clear the boxes
lstDiscard.Clear
lstCurrent.Clear
lstDiscardMask.Clear
lstCardNames.Clear
lstDone.Clear
'redraw the hand
Call CheckTIC
On Error GoTo 0
Exit Sub
cmdRedo_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdRedo_Click in Form, Form1"
End Sub

Private Sub cmdRedo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo cmdRedo_MouseMove_Error
'unhighlight the buttons
Call HAllHButtons
'highlight redo button
cmdRedo.ForeColor = vbWhite
imgHRedo.Visible = True
On Error GoTo 0
Exit Sub
cmdRedo_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdRedo_MouseMove in Form, Form1"
End Sub

Private Sub cmdSet_Click()
Dim Temp() As Integer, Temp1() As Integer 'Temp Variable arrays
On Error GoTo cmdSet_Click_Error
ReDim Temp(lstCardNames.ListCount - 1) As Integer 'Redim it to the correct diminsions
ReDim Temp1(UBound(Temp)) As Integer 'redim to correct diminsions
For i = 0 To UBound(Temp) 'set the temp arrays to the cards in lstcardnames
                          'which is the numbers of the cards
    Temp(i) = Val(lstCardNames.List(i))
    Temp1(i) = Temp(i)
Next
'Reduce the cards in temp1 array
For i = LBound(Temp1()) To UBound(Temp1())
    Select Case Temp1(i)
        Case 13 To 25:
            Temp1(i) = (Temp1(i) - 13)
        Case 26 To 38:
            Temp1(i) = (Temp1(i) - (13 * 2))
        Case 39 To 51:
            Temp1(i) = (Temp1(i) - (13 * 3))
    End Select
Next
'variables for counting and checking
Dim intVal1 As Variant, intThis%, intVal2 As Variant
'begin a loop to check if we want to
'check for pairs
For Each intVal1 In Temp1
    intTest% = intVal1
    intThis% = 0
    For Each intVal2 In Temp1
        If intVal2 = intVal1 And intVal1 <> Val(Round$) - 1 And intVal2 <> Val(Round$) - 1 Then intThis% = intThis% + 1
        If intThis% > 1 Then
            'if we get 2 of the same number, set checkwhat to 'PAIRS'
            CheckWhat$ = "PAIRS"
            'and end the loop
            GoTo EndLoop
            Exit For
        End If
    Next intVal2
Next intVal1
'if there isnt 2 of the same card, make checkwhat equal 'STR' for straight
CheckWhat$ = "STR"
EndLoop:
'Check if the cards put in work through our class module
'and if its true, clear out the boxes and get ready
'for another set
If FU.IsCorrect(Temp(), Val(Round$), CheckWhat$) = True Then
    For i = 0 To lstCardNames.ListCount - 1
        lstDone.AddItem lstCardNames.List(i)
    Next
    lstCurrent.Clear
    lstCardNames.Clear
End If
On Error GoTo 0
Exit Sub
cmdSet_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdSet_Click in Form, Form1"

End Sub

Private Sub cmdSet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo cmdSet_MouseMove_Error
'make the set visiblae true for the highlight
imghSet.Visible = True
cmdSet.ForeColor = vbWhite
On Error GoTo 0
Exit Sub
cmdSet_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdSet_MouseMove in Form, Form1"
End Sub

Private Sub cmdTIC_Click()
On Error GoTo cmdTIC_Click_Error
'clear out the listboxes
lstCurrent.Clear
lstDiscard.Clear
lstDone.Clear
lstDiscardMask.Clear
lstCardNames.Clear
'backup cardsinhand to temphand
For i = 0 To UBound(CardsInHand)
    TempHand(i) = CardsInHand(i)
Next
'draw the hand on the next baord
Call CheckTIC
On Error GoTo 0
Exit Sub
cmdTIC_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdTIC_Click in Form, Form1"
End Sub

Private Sub cmdTIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo cmdTIC_MouseMove_Error
'make the highlighted setup button visible
imgHSetup.Visible = True
cmdTIC.ForeColor = vbWhite
On Error GoTo 0
Exit Sub
cmdTIC_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdTIC_MouseMove in Form, Form1"
End Sub

Private Sub Command1_Click()
On Error Resume Next
'See if there is a cheat code being put in
If CheckCheat(txtTalk.Text) = True Then Exit Sub
'if not, send the data to the server,
'so it can be displayed in the chat area
ws.SendData "¼½¾" & txtTalk.Text
'clear out the text box
txtTalk.Text = ""
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Command1_MouseMove_Error
'make the highlighted hsend visible
imgHSend.Visible = True
Command1.ForeColor = vbWhite
On Error GoTo 0
Exit Sub
Command1_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Command1_MouseMove in Form, Form1"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo Form_KeyPress_Error
If KeyAscii = 3 Then 'CTRL + C
    'shortcut for connecting
    Call mnuconnect_Click
ElseIf KeyAscii = 24 Then 'CTRL + X
    'shortcut for quiting
    Call mnuExit_Click
ElseIf KeyAscii = 16 Then 'CTRL + P
    'shortcut for 1 player
    Call lblSubMenu_Click(1)
ElseIf KeyAscii = 19 Then 'CTRL + S
    If imgHand.UBound > 1 Then 'if they have cards in their hand
        'Unload Form7 'unload the form
        Load Form7 'load the sort hand form
        Form7.Show  'show and keep on top
        FU.PutOnTop Form7, True
    End If
End If
On Error GoTo 0
Exit Sub
Form_KeyPress_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_KeyPress in Form, Form1"
End Sub

Private Sub Form_Load()
On Error GoTo Form_Load_Error
Dim ret As Long, rRect As RECT 'Variables for trimming border off listboxes
Dim formDC As Long 'forms DC
Dim frmWid As Long 'Forms width
Dim frmHgt As Long 'forms height
Dim aControl As Control 'a control
Dim lScreenX&, lScreenY& 'hold the screen width
'get the screen resolution
lScreenX& = FU.GetSysMetrics(SM_CXSCREEN)
lScreenY& = FU.GetSysMetrics(SM_CYSCREEN)
'minmum resolution is 800x600
If lScreenX& < 800 Or lScreenY& < 600 Then
    MsgBox "Error:  This program requires a minium resolution of 800x600 pixels." & vbCrLf & "Program will now close.", vbCritical, "Error"
    Unload Form3
    Unload Me
End If
Me.AutoRedraw = False 'set autoredraw to false
FU.PutOnTop Me, True 'make the form always on top
'set the caption of the form
Me.Caption = "Tic Game Board [ver " & App.Major & "." & App.Minor & "]"
imgTrash.MouseIcon = LoadResPicture(101, vbResCursor) 'load the custom cursor
'moves this off the screen until the game begins
imgHand(0).Top = -imgHand(0).Height
imgHand(0).Left = -imgHand(0).Width

'**************************************************
'Code to make our multiline textbox transparent****
'Thanks to The Hand from www.visualbasicforum.com**
'for most of the code.*****************************
'**************************************************
'make a temparary property for txtchat called
'DoRedraw, and set it to -1, we'll check
'this later while we are subclassing it
SetProp txtChat.hwnd, "DoRedraw", -1
Me.ScaleMode = 3 'Pixels
'Make all textboxs false visiblaly
For Each aControl In Me.Controls
    If TypeOf aControl Is TextBox Then aControl.Visible = False
Next aControl
'Make the menu item visable false
imgFile.Visible = False
'Show the form
Me.Show
'and refresh it
Me.Refresh
'Get the forms width and height in pixels
frmWid = Me.ScaleX(Me.Width, Me.ScaleMode, vbPixels)
frmHgt = Me.ScaleY(Me.Height, Me.ScaleMode, vbPixels)
'get the forms DC
formDC = GetDC(Me.hwnd)
'Cleat the pic box
picTextPics.Cls
'Get a picture of the forms picture
BitBlt picTextPics.hdc, 0, 0, frmWid, frmHgt, formDC, 0, 0, vbSrcCopy
'Set the image to use to a image box
Set imgTxtChat.Picture = picTextPics.Image
'Set our brush to the handle of the picture
txtBoxBrush1 = CreatePatternBrush(imgTxtChat.Picture.Handle)
'Release the DC we made
ReleaseDC Me.hwnd, formDC
'Make the visible property of the text boxes
'to true
For Each aControl In Me.Controls
    If TypeOf aControl Is TextBox Then aControl.Visible = True
Next aControl
'**************************************************
'set autoredraw back to true, and make it so its not always ontop
Me.AutoRedraw = True
FU.PutOnTop Me, False
'**************************************************
'*******Trim of lstCurrents Border*****************
'**************************************************
rRect.lTop = 1
rRect.lLeft = 1
rRect.lRight = lstCurrent.Width - 1
rRect.lBottom = lstCurrent.Height - 1
ret = CreateRectRgnIndirect(rRect)
SetWindowRgn lstCurrent.hwnd, ret, True
'**************************************************

'**************************************************
'*******Trim of lstDiscards Border*****************
'**************************************************
rRect.lTop = 1
rRect.lLeft = 1
rRect.lRight = lstDiscard.Width - 1
rRect.lBottom = lstDiscard.Height - 1
ret = CreateRectRgnIndirect(rRect)
SetWindowRgn lstDiscard.hwnd, ret, True
'**************************************************

'*********************************************************
'Now to get the pictures that are behind the list boxes***
'and save them to an image box****************************
'*********************************************************
Picture1.Cls
BitBlt Picture1.hdc, 0, 0, lstCurrent.Width + 1, lstCurrent.Height + 1, Frame2.hdc, lstCurrent.Left + 1, lstCurrent.Top + 1, vbSrcCopy
Image1.Picture = Picture1.Image

Picture1.Cls
BitBlt Picture1.hdc, 0, 0, lstDiscard.Width + 1, lstDiscard.Height + 1, Frame2.hdc, lstDiscard.Left + 1, lstDiscard.Top + 1, vbSrcCopy
Image2.Picture = Picture1.Image
'*********************************************************

'*********************************************************
'Now to get the text box picture*************************
'*********************************************************
picTextPics.Cls
BitBlt picTextPics.hdc, 0, 0, txtTalk.Width, txtTalk.Height, Form1.hdc, txtTalk.Left, txtTalk.Top, vbSrcCopy
imgTxtTalk.Picture = picTextPics.Image
'*********************************************************

'*********************************************************
'Reset the scalemode back to Twips************************
'*********************************************************
Frame2.ScaleMode = 1
Form1.ScaleMode = 1
'*********************************************************

'*************************************************************
'Make the pattern brushes so we can redaw the listboxes when**
'we subclass them*********************************************
'*************************************************************
gBGBrush = CreatePatternBrush(Image1.Picture.Handle)
gBGBrush2 = CreatePatternBrush(Image2.Picture.Handle)
txtBoxBrush2 = CreatePatternBrush(imgTxtTalk.Picture.Handle)
'*************************************************************

'*************************************************************
'Make the spots where we store the pictures invisible*********
'*************************************************************
Picture1.Visible = False
picTextPics.Visible = False
'*************************************************************

'*************************************************************
'Begin the subclassing****************************************
'*************************************************************
oldWindowProc2 = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf NewWindowProc2)
oldWindowProc = SetWindowLong(Frame2.hwnd, GWL_WNDPROC, AddressOf NewWindowProc)
oldLbx1Proc = SetWindowLong(lstCurrent.hwnd, GWL_WNDPROC, AddressOf NewLbxProc)
oldLbx2Proc = SetWindowLong(lstDiscard.hwnd, GWL_WNDPROC, AddressOf NewLbxProc2)
'Refresh all the textboxes on the form
For Each aControl In Me.Controls
    If TypeOf aControl Is TextBox Then aControl.Refresh
Next aControl
'*************************************************************

'*************************************************************
'Set default values for the Card Hands************************
'*************************************************************
For i = 0 To UBound(CardsInHand)
    CardsInHand(i) = -1
    TempHand(i) = -1
Next
'*************************************************************

'********************************************************************
'And set the borderstyples of the card images used for the hand to***
'0 - None************************************************************
'********************************************************************
imgFile.Visible = True
On Error GoTo 0
Exit Sub
Form_Load_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_Load in Form, Form1"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Form_MouseMove_Error
'unhighlight all the buttons
Call HAllHButtons
On Error GoTo 0
Exit Sub
Form_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_MouseMove in Form, Form1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Form_Unload_Error
'unsubclass everything
SetWindowLong Frame2.hwnd, GWL_WNDPROC, oldWindowProc
SetWindowLong lstCurrent.hwnd, GWL_WNDPROC, oldLbx1Proc
SetWindowLong lstDiscard.hwnd, GWL_WNDPROC, oldLbx2Proc
SetWindowLong Me.hwnd, GWL_WNDPROC, oldWindowProc2
'delete out brushes
DeleteObject gBGBrush
DeleteObject gBGBrush2
DeleteObject txtBoxBrush1
DeleteObject txtBoxBrush2
'close any winsock connections
ws.Close
'Unload all the forms just in case
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload frmAbout
Unload Me
On Error GoTo 0
Exit Sub
Form_Unload_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_Unload in Form, Form1"
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Frame1_MouseMove_Error
'unhighlight the buttons
Call HAllHButtons
imgHSetup.Visible = False
imgHSend.Visible = False
On Error GoTo 0
Exit Sub
Frame1_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Frame1_MouseMove in Form, Form1"
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Frame2_MouseMove_Error
'unhighlight the buttons
Call HAllHButtons
On Error GoTo 0
Exit Sub
Frame2_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Frame2_MouseMove in Form, Form1"
End Sub

Private Sub imgDeck_Click()
On Error Resume Next
If lblTurn.Caption = YName$ Then 'if the turn is yours
    'send data to the server
    ws.SendData "¶®Æ"
    'disable the images
    imgDeck.Enabled = False
    imgDiscard.Enabled = False
End If
End Sub

Private Sub imgDeck_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'unhighlight the buttons
Call HAllHButtons
End Sub

Private Sub imgDiscard_Click()
On Error Resume Next
If lblTurn.Caption = YName$ Then 'if the turn is yours
    'send the data tot he server
    ws.SendData "×ÿ¡"
    'disbale the images
    imgDiscard.Enabled = False
    imgDeck.Enabled = False
End If
End Sub

Private Sub imgDiscard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'unhighlight the buttons
Call HAllHButtons
End Sub

Private Sub imgFile_Click()
On Error GoTo imgFile_Click_Error
'set the drawingmenu true/false value true
DrawingMenu = True
Call GenerateMenu 'gen the menu
DrawingMenu = False 'set it to false
On Error GoTo 0
Exit Sub
imgFile_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: imgFile_Click in Form, Form1"
End Sub

Private Sub imgHand_DblClick(Index As Integer)
On Error GoTo imgHand_DblClick_Error
'If it is your turn, you can continue
If lblTurn.Caption = YName$ Then
    Dim Counter% 'counter
    'see how many cards are in their hand
    For i = 0 To UBound(CardsInHand)
        If CardsInHand(i) <> -1 Then
            Counter% = Counter% + 1
        End If
    Next
    'make sure they have a card to discard
    If Counter% < Val(Round$) + 1 Then Exit Sub
    'get verification from the user if
    'they really want to discard the card
    If MsgBox("Are you sure you wish to discard this card?", vbOKCancel + vbQuestion, "Discard?") = vbOK Then 'if its ok then
        Dim p%
        'reduce the card to see if it is
        'a wild card, so as for the user
        'not to make a mistake
        Select Case CardsInHand(Index)
            Case 0 To 12:
                p% = CardsInHand(Index)
            Case 13 To 25:
                p% = (CardsInHand(Index) - 13)
            Case 26 To 38:
                p% = (CardsInHand(Index) - (13 * 2))
            Case 39 To 51:
                p% = (CardsInHand(Index) - (13 * 3))
        End Select
        'if the value of the card is equal to the
        'round - 1 (minus 1 because the card values
        'are 0 to 12), and ask for more
        'verification
        If Round$ <> "14" Or Round$ <> "15" Then
            If p% = Val(Round$) - 1 Then
                If MsgBox("This card is wild, are you sure you want to discard it?" _
                    , vbOKCancel + vbQuestion, "Wild Card Detected") _
                    = vbCancel Then imgHand(Index).Left = _
                    imgTrash.Left - imgHand(Index).Width: Exit Sub 'if its ok
                    'then exit the sub
            End If
        Else
            If p% = Val(Round$) - 14 Then
                If MsgBox("This card is wild, are you sure you want to discard it?" _
                    , vbOKCancel + vbQuestion, "Wild Card Detected") _
                    = vbCancel Then imgHand(Index).Left = _
                    imgTrash.Left - imgHand(Index).Width: Exit Sub 'if its ok
                    'then exit the sub
            End If
        End If
        'If everything is ok, then send the
        'discard card to the server
        ws.SendData "öõô" & CardsInHand(Index)
        'remvoe it from the hand
        CardsInHand(Index) = -1
        'redraw the hand
        Call RedoHand
        'and end the turn
        Call EndTurn
        Exit Sub
    End If
End If
On Error GoTo 0
Exit Sub
imgHand_DblClick_Error:
'MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: imgHand_DblClick in Form, Form1"
End Sub

Private Sub imgHand_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo imgHand_MouseDown_Error
If Button = 1 Then 'if left-click
    bDown = True 'make the mvoeing variables true
    bMove = True
    orgX = X 'store the position of the mouse pointer
    orgY = Y
    imgHand(Index).ZOrder (0) 'and put the card on top
End If
On Error GoTo 0
Exit Sub
imgHand_MouseDown_Error:
End Sub

Private Sub imgHand_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo imgHand_MouseMove_Error
If Button = 1 Then 'if left-click
    If bDown Then 'if it is true
        If bMove Then 'if it is true
            bMove = False 'set it to false
        End If
        'move the card by keeping the the mouse cursor in the orginal click location
        '//////
        'Thanks to Banjo from the ExtremeVB forums www.visualbasicforum.com
        '/////
        imgHand(Index).Top = (imgHand(Index).Top + Y) - orgY 'and move the card
        imgHand(Index).Left = (imgHand(Index).Left + X) - orgX 'move the card
        
    End If
End If
On Error GoTo 0
Exit Sub
imgHand_MouseMove_Error:
End Sub

Private Sub imgHand_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo imgHand_MouseUp_Error
If Button = 1 Then 'if left click
    bDown = False 'set bdown to false
    'make sure the card stays in view
    If imgHand(Index).Top <= 0 Then imgHand(Index).Top = 0
    If imgHand(Index).Left <= 0 Then imgHand(Index).Left = 0
    If imgHand(Index).Top >= Frame1.Height - imgHand(Index).Height Then imgHand(Index).Top = Frame1.Height - imgHand(Index).Height
    If imgHand(Index).Left >= Frame1.Width - imgHand(Index).Width Then imgHand(Index).Left = Frame1.Width - imgHand(Index).Width
    If lblTurn.Caption = YName$ Then 'if it is your turn
        Dim Counter% 'counter
        'see how many cards are in their hand
        For i = 0 To UBound(CardsInHand)
            If CardsInHand(i) <> -1 Then
                Counter% = Counter% + 1
            End If
        Next
        'make sure they have a card to discard
        If Counter% < Val(Round$) + 1 Then Exit Sub
        If imgHand(Index).Left >= imgTrash.Left And LastTurn = False Then 'if it isnt that last
        'turn of the round, and the card is in the discard area, continue
            If MsgBox("Are you sure you wish to discard this card?", vbOKCancel + vbQuestion, "Discard?") = vbOK Then 'if they want to
            'discard the card
                Dim p%
                'reduce the card to check if it a wild card.
                Select Case CardsInHand(Index)
                    Case 0 To 12:
                        p% = CardsInHand(Index)
                    Case 13 To 25:
                        p% = (CardsInHand(Index) - 13)
                    Case 26 To 38:
                        p% = (CardsInHand(Index) - (13 * 2))
                    Case 39 To 51:
                        p% = (CardsInHand(Index) - (13 * 3))
                End Select
                'if it is, ask for more verification from the
                'user.
                If p% = Val(Round$) - 1 Then
                    If MsgBox("This card is wild, are you sure you want to discard it?" _
                        , vbOKCancel + vbQuestion, "Wild Card Detected") _
                        = vbCancel Then imgHand(Index).Left = _
                        imgTrash.Left - imgHand(Index).Width: Exit Sub 'if they
                        'donot want to discard it, exit the sub
                End If
                'send the discarded card to the server
                ws.SendData "öõô" & CardsInHand(Index)
                'and remove that card from the users hand
                CardsInHand(Index) = -1
                'redaw the hand
                Call RedoHand
                'and end the turn
                Call EndTurn
                Exit Sub
            Else
                'if they do not wish to discard it, move it to the
                'immediate left of the area.
                imgHand(Index).Left = imgTrash.Left - imgHand(Index).Width
            End If
        End If
    End If
End If
On Error GoTo 0
Exit Sub
imgHand_MouseUp_Error:
'MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: imgHand_MouseUp in Form, Form1"
End Sub

Private Sub imgTICSetup_DblClick(Index As Integer)
On Error GoTo imgTICSetup_DblClick_Error
'Add the # of the card to a listbox
lstCardNames.AddItem CardsInHand(Index)
'remove the card from the hand
CardsInHand(Index) = -1
'refresh the listboxes
Call RefreshlstCurrent
'refresh the hand
Call CheckTIC
On Error GoTo 0
Exit Sub
imgTICSetup_DblClick_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: imgTICSetup_DblClick in Form, Form1"
End Sub

Private Sub imgTICSetup_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo imgTICSetup_MouseDown_Error
If Button = 1 Then 'if left click
    ADown = True 'booleans true
    AMove = True
    orgX = X
    orgY = Y
    imgTICSetup(Index).ZOrder (0) 'make it on top
End If
On Error GoTo 0
Exit Sub
imgTICSetup_MouseDown_Error:
End Sub

Private Sub imgTICSetup_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo imgTICSetup_MouseMove_Error
If Button = 1 Then 'if left click
    If ADown Then 'if it is true
        If AMove Then 'if it is true
            AMove = False 'make it false
        End If
        'keep the card from going over the listboxes cause of
        'graphic glitch errors
        If imgTICSetup(Index).Left + imgTICSetup(Index).Width < Label3(0).Left + Label3(0).Width Then
            imgTICSetup(Index).Move imgTICSetup(Index).Left + X - orgX, imgTICSetup(Index).Top + Y - orgY
        ElseIf imgTICSetup(Index).Left + X - orgX + imgTICSetup(Index).Width < Label3(0).Left + Label3(0).Width Then
            imgTICSetup(Index).Move imgTICSetup(Index).Left + X - orgX, imgTICSetup(Index).Top + Y - orgY
        Else
            imgTICSetup(Index).Move imgTICSetup(Index).Left, imgTICSetup(Index).Top + Y - orgY
        End If
    End If
End If
On Error GoTo 0
Exit Sub
imgTICSetup_MouseMove_Error:
End Sub

Private Sub imgTICSetup_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo imgTICSetup_MouseUp_Error
If Button = 1 Then 'if left click
    ADown = False 'set false
    'make sure teh card stays in view
    If imgTICSetup(Index).Top <= 0 Then imgTICSetup(Index).Top = 0
    If imgTICSetup(Index).Left <= 0 Then imgTICSetup(Index).Left = 0
    If imgTICSetup(Index).Top >= Frame2.Height - imgTICSetup(Index).Height Then imgTICSetup(Index).Top = Frame2.Height - imgTICSetup(Index).Height
    If imgTICSetup(Index).Left >= Frame2.Width - imgTICSetup(Index).Width Then imgTICSetup(Index).Left = Frame2.Width - imgTICSetup(Index).Width
    'if the card is on top of list current
    If imgTICSetup(Index).Top > Label3(0).Top And imgTICSetup(Index).Left > Label3(0).Left Then
        If imgTICSetup(Index).Left < Label3(0).Left + Label3(0).Width And imgTICSetup(Index).Top < Label3(0).Top + Label3(0).Height Then
            lstCardNames.AddItem CardsInHand(Index) 'add the card # to a listbox
            imgTICSetup(Index).Visible = False 'make the card invisible
            CardsInHand(Index) = -1 'remove the card from hand
            Call RefreshlstCurrent 'refresh the listboxes
            Call CheckTIC 'refresh the hand
        End If
    End If
    'if the card is on top of lstdiscard
    If imgTICSetup(Index).Top > Label3(1).Top And imgTICSetup(Index).Left > Label3(1).Left Then
        If imgTICSetup(Index).Left < Label3(1).Left + Label3(1).Width And imgTICSetup(Index).Top < Label3(1).Top + Label3(1).Height Then
            'if there isn't anything in there
            If lstDiscardMask.ListCount < 1 Then
                'add the card to a listbox
                lstDiscardMask.AddItem CardsInHand(Index)
                'make the card invisible
                imgTICSetup(Index).Visible = False
                'remove the card
                CardsInHand(Index) = -1
                'refresh thelistboxes
                Call RefreshlstCurrent
                'refrehs the hand
                Call CheckTIC
            End If
        End If
    End If
End If
On Error GoTo 0
Exit Sub
imgTICSetup_MouseUp_Error:
End Sub

Private Sub imgTrash_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo imgTrash_MouseMove_Error
Call HAllHButtons
On Error GoTo 0
Exit Sub
imgTrash_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: imgTrash_MouseMove in Form, Form1"
End Sub

Private Sub lblMenu_Click(Index As Integer)
On Error GoTo lblMenu_Click_Error
If DrawingMenu = False Then 'make sure the menu isnt drawing
    If Index <> 0 Then
        picMenu.Visible = False 'make the menu disapear
        imgFile.Height = 200 'shrink the image
    End If
    Select Case Index
        Case 0:
            'Draw the sub menu for the connecting
            'options
            DrawingMenu = True
            Call GenerateSubMenu
            DrawingMenu = False
        Case 1:
            Call mnuRules_Click 'show the rules
        Case 2:
            Call mnuIndex_Click 'show the help index
        Case 3:
            Call mnuAbout_Click 'show the about box
        Case 4:
            Call mnuExit_Click 'quit the prog
    End Select
End If
On Error GoTo 0
Exit Sub
lblMenu_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: lblMenu_Click in Form, Form1"
End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
HMenu Index, False 'highlight the selected choice
End Sub

Private Sub lblSubMenu_Click(Index As Integer)
'set the menus to default
picSubMenu.Visible = False
picMenu.Visible = False
imgFile.Height = 200
'select the index
Select Case Index
    Case 0:
        'if its 0, then call mnuconnect's click
        Call mnuconnect_Click
    Case 1:
        'if its 1, open up the server minimized
        Shell App.Path & "\TicServer.exe", vbMinimizedNoFocus
        'wait for 1 second
        FU.WaitFor 1000
        'close any current connect
        ws.Close
        Dim tName$
        'get a nickname
        tName$ = InputBox("Insert your nickname:", "Nickname")
        'if they inserted something
        If tName <> "" Then
            'get this machines IP
            ws.RemoteHost = GetIPAddress
            'set the global Name varible to the players
            'nickname
            YName$ = tName$
            'connect
            ws.Connect
            'Send a message about the game starting
            txtChat.Text = txtChat.Text & "Game starting, please wait..." & vbCrLf
            'wait for 400 miliseconds
            FU.WaitFor 1
            'send 1 player info to server for autostart
            ws.SendData "þ|âÿË®¹"
        End If
End Select
End Sub

Private Sub lblSubMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'highlight the submenu options
HMenu Index, True
End Sub

Private Sub lblTurn_Change()
On Error GoTo lblTurn_Change_Error
If lblTurn.Caption = YName$ And GameEnd = False Then 'if the new turn is yours, and its not the end of the game
    'reenable all the cards, and game field
    imgDeck.Enabled = True
    imgDiscard.Enabled = True
    Frame1.Enabled = True
End If
On Error GoTo 0
Exit Sub
lblTurn_Change_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: lblTurn_Change in Form, Form1"
End Sub

Private Sub lblTurn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call HAllHButtons 'unhighlight all the buttons
End Sub

Private Sub lstCurrent_DblClick()
On Error Resume Next
For i = 0 To UBound(CardsInHand) 'loop through the hand
    If CardsInHand(i) = -1 Then 'if there is an empty spot
        CardsInHand(i) = Val(lstCardNames.List(lstCurrent.ListIndex)) 'make the card
            'in that spot
        'remove it from the listbox
        lstCardNames.RemoveItem lstCurrent.ListIndex
        'remove it from the listbox
        lstCurrent.RemoveItem lstCurrent.ListIndex
        'refresh the listboxes
        Call RefreshlstCurrent
        Exit For
    End If
Next
'refresh the hand
Call CheckTIC
End Sub

Private Sub lstDiscard_DblClick()
On Error GoTo lstDiscard_DblClick_Error
For i = 0 To UBound(CardsInHand) 'loop through the hand
    If CardsInHand(i) = -1 Then 'if there is an empty spot
        CardsInHand(i) = Val(lstDiscardMask.List(lstDiscard.ListIndex)) 'set the card
            'into the empty spot
        lstDiscard.Clear
        lstDiscardMask.Clear 'remove the card
        'refresh thelistbox
        Call RefreshlstCurrent
        Exit For
    End If
Next
'refresh the hand
Call CheckTIC
On Error GoTo 0
Exit Sub
lstDiscard_DblClick_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: lstDiscard_DblClick in Form, Form1"
End Sub

Private Sub mnuAbout_Click()
On Error GoTo mnuAbout_Click_Error
'load and show the about form
Load frmAbout
frmAbout.Show
On Error GoTo 0
Exit Sub
mnuAbout_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: mnuAbout_Click in Form, Form1"
End Sub

Private Sub mnuconnect_Click()
On Error GoTo mnuconnect_Click_Error
'load and show, and keep on top, the connect to
'game form
Load Form2
Form2.Show 1
On Error GoTo 0
Exit Sub
mnuconnect_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: mnuconnect_Click in Form, Form1"
End Sub

Private Sub mnuExit_Click()
On Error GoTo mnuExit_Click_Error
'unload me
Unload Me
On Error GoTo 0
Exit Sub
mnuExit_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: mnuExit_Click in Form, Form1"
End Sub

Private Sub mnuIndex_Click()
On Error GoTo mnuIndex_Click_Error
'load and show the help index
Load Form4
Form4.Show
On Error GoTo 0
Exit Sub
mnuIndex_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: mnuIndex_Click in Form, Form1"
End Sub

Private Sub mnuRules_Click()
On Error GoTo mnuRules_Click_Error
'load and show the rules of the game
Load Form5
Form5.Show
On Error GoTo 0
Exit Sub
mnuRules_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: mnuRules_Click in Form, Form1"
End Sub

Private Sub txtChat_Change()
On Error Resume Next
Dim NumberOfLinesYouWantShowing As Integer 'variable for amount of lines to show
Dim TextInDaBox As String 'string to hold the text
Dim Length As Integer 'length of the text
Dim NumberOfReturns As Integer 'amount of chr(13) in the box
Dim Location As Integer 'counter to hold a number
'replace any symbols the prog uses
txtChat.Text = Replace(txtChat.Text, "¼½¾", vbCrLf)
txtChat.Text = Replace(txtChat.Text, "¼", "")
txtChat.Text = Replace(txtChat.Text, "½", "")
txtChat.Text = Replace(txtChat.Text, "¾", "")
strData$ = Replace(strData$, "Î¶¬", " ")
'i want 8 lines showing
NumberOfLinesYouWantShowing = 8
TextInDaBox = txtChat.Text 'set the sting to txtchat's text
Length = Len(TextInDaBox) 'set the length to len of our sting variable
For i = 1 To Length 'loop through the string
    If Mid$(TextInDaBox, i, 1) = Chr(13) Then 'if there is an enter, add to the counter
        NumberOfReturns = NumberOfReturns + 1
    End If
Next i
'make sure there are more then 8 lines
If NumberOfReturns <= NumberOfLinesYouWantShowing Then Exit Sub
'set the location to the last enter
Location = InStrRev(TextInDaBox, vbCrLf, Length)
For Counter = 1 To NumberOfLinesYouWantShowing 'loop through the text
    Location = InStrRev(TextInDaBox, vbCrLf, Location) 'and set the location to
        'the enter previous of the last time
Next Counter
'set the text to be the 8 lines
TextInDaBox = Right$(TextInDaBox, Length - Location - 1)
'TextInDaBox = Left$(TextInDaBox, Len(TextInDaBox) - 1)
'and set txtchats text to the stirng
txtChat.Text = TextInDaBox
'////////////
'this stuff is to stop it from glitching up...
'it kept visually messing up on my copmuter, so i was
'messing around with stuff, and this seems to fix the
'problem
'////////////
'call txtchat's click sub
Call txtChat_Click
'make is visible and false
txtChat.Visible = False
txtChat.Visible = True
End Sub

Private Sub txtChat_Click()
'set focus to txttalk
txtTalk.SetFocus
End Sub

Private Sub txtChat_GotFocus()
'set focus to txttalk
txtTalk.SetFocus
End Sub

Private Sub txtChat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'set focus to txttalk
txtTalk.SetFocus
End Sub

Private Sub txtChat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'set focus to txttalk
'txtTalk.SetFocus
End Sub

Private Sub txtChat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'set focus to txttalktxt
txtTalk.SetFocus
End Sub

Private Sub txtTalk_KeyPress(KeyAscii As Integer)
On Error GoTo txtTalk_KeyPress_Error
'if the user presses enter, call command1's click
If KeyAscii = 13 Then Call Command1_Click
On Error GoTo 0
Exit Sub
txtTalk_KeyPress_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: txtTalk_KeyPress in Form, Form1"
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
Dim strData$
On Error GoTo ws_DataArrival_Error
'get the data
ws.GetData strData$, vbString
Debug.Print strData$
'send the data through these subs to
'check for correct data
SendNameBack strData$
CreateHand strData$
FlipFirstOver strData$
CheckChat strData$
GetTurn strData$
DrawCard strData$
GetTopCard strData$
InvokeCheat strData$
ServerIsRefreshing strData$
'make strdata equal nothing
strData$ = ""
On Error GoTo 0
Exit Sub
ws_DataArrival_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: ws_DataArrival in Form, Form1"
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error GoTo ws_Error_Error
'error handling for winsock errors
On Error GoTo 0
Exit Sub
ws_Error_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: ws_Error in Form, Form1"
End Sub

Sub CheckChat(strData$)
'Sub to check for the server
'for send out 'chat' messages
'from other users...or messages
'that need to be displayed in
'txtchat
Dim a As Integer 'counter
'On Error GoTo CheckChat_Error
If Left(strData$, 3) = "¼½¾" Then 'if we find the special code
'for chat
    If InStr(strData$, "Î¶¬has got a tic!") Then 'if theres a tic,
        LastTurn = True 'make the lastturn variable true
        strData$ = Replace(strData$, "Î¶¬", " ") 'and remove the code
    End If
    If InStr(strData$, "×©ºhas won the game with") Then  'if there is the special code
        'game over code
        strData$ = Replace(strData$, "×©º", " ") 'and remove the code
        'games over, and disable everything
        Frame1.Enabled = False
        imgDiscard.Enabled = False
        imgDeck.Enabled = False
        GameEnd = True
    End If
    If InStr(strData$, "ªÿ½picked up the top discarded car.") Then 'if they got the top discarded card
        strData$ = Replace(strData$, "ªÿ½", " ") 'remove the code
        Deck1.GetAnotherCard = 53 'get a blank card pic
        imgDiscard.Picture = Deck1.Picture 'and set the pic
    End If
    If InStr(strData$, " discards a card from their hand.ÑÐª") Then 'change the discard pic
        Dim CardO%
        'get the card to change to
        CardO% = Val(Right(strData$, (Len(strData$) - InStr(strData$, "ÑÐª") - 2)))
        'trim off the program data
        strData$ = Left(strData$, InStr(strData$, "ÑÐª") - 1)
        Deck1.GetAnotherCard = CardO% 'change the deck pic to the card #
        imgDiscard.Picture = Deck1.Picture 'and change the discard pic to the card #
    End If
    'remove the code
    strData$ = Right(strData$, Len(strData$) - 3)
    'and add it to txtchat
    txtChat.Text = txtChat.Text & strData$ & vbCrLf
End If
On Error GoTo 0
Exit Sub
CheckChat_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: CheckChat in Form, Form1"
End Sub

Function CheckCheat(Txt As String) As Boolean
On Error GoTo CheckCheat_Error
'for the cheaters..check if they want to cheat
'if the left 15 characters = 'me.cheat.spike.' and its there turn
'continue on with the cheat
If Left(Txt, 15) = "me.cheat.spike." And lblTurn.Caption = YName$ Then
    Dim a%, b%, c% 'integer variables
    Txt = Right(Txt, Len(Txt) - 15) 'get rid of the first 15 characters
    a% = InStr(1, Txt, ".") 'find the first '.'
    b% = Val(Mid$(Txt, 1, a%)) 'set b equal to the first a% characters
    Txt = Right(Txt, Len(Txt) - Len(b%)) 'trim off thr first a% characters
    Txt = Replace(Txt, ".", "") 'get rid of the '.'
    a% = 0 'set a equal 0
    c% = Val(Txt) 'make c what is left
    If CardsInHand(c%) <> -1 Then 'make sure the cards in the hand
            'is not nothing
        BackUpBuffer = CardsInHand(c%) 'save the card into memory
            'incase it isnt there
        CardsInHand(c%) = -1
            'clear out the hand of that card
    End If
    'send the cheat to the server
    ws.SendData "§ÞÏkë" & b%
    'clear txttalk
    txtTalk.Text = ""
    'set this function to be true
    CheckCheat = True
ElseIf Left(Txt, 24) = "me.cough.cheat.cough.now" Then
    If LastTurn = True And lblTurn.Caption = YName$ Then
        ws.SendData "¶¸·" & YName$ & "||0**" & CardsInHand(0)
        'wait for 400 miliseconds
        FU.WaitFor 1
        'Send the card you discarded
        'ws.SendData "öõô" & CardsInHand(0)
        CardsInHand(0) = -1
        'redraw the hand
        Call RedoHand
        'and set the defaults
        Frame2.Visible = False
        lstCurrent.Clear
        lstDiscard.Clear
        lstDone.Clear
        lstDiscardMask.Clear
        lstCardNames.Clear
        Frame1.Enabled = False
        CheckCheat = True
        txtTalk.Text = ""
    End If
End If
On Error GoTo 0
Exit Function
CheckCheat_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: CheckCheat in Form, Form1"
End Function

Sub CheckTIC()
On Error Resume Next
Dim a As Integer
'default values for frame2
If Frame2.Visible = False Then
    With Frame2
        .Top = Frame1.Top
        .Left = Frame1.Left
        .Height = Frame1.Height
        .Width = Frame1.Width
        .ScaleMode = 1 'twips
        .Visible = True
    End With
End If
imgTICSetup(0).Top = 120 'default position for first card
imgTICSetup(0).Left = 120
Dim First As Boolean
First = True 'check for the first card
Dim q%, r%, z As Integer, intW%, intL% 'counters/number storers
intW% = imgTICSetup(0).Width / 3 'set this to the width of a card/3
intL% = 120 'set this to 120, default left for a card
q% = intL% 'set q to the default left
r% = intL% + intW% 'set r to the default left play the width of the card/3
'unload all the cards except the first one
For z = 1 To imgTICSetup.UBound
    Unload imgTICSetup(z)
Next
Dim b As Integer 'counter
b = 0 'default to 0
For a = 0 To 14 'compress all together
        'to make sure there are no stray -1 in the hand
    If CardsInHand(a) <> -1 Then
        If a <> b Then
            CardsInHand(b) = CardsInHand(a)
            CardsInHand(a) = -1
        End If
        b = b + 1
    End If
Next
Dim iCount% 'counter
iCount% = 0 'default to 0
'find out how many cards there are
For i = 0 To UBound(CardsInHand)
    If CardsInHand(i) <> -1 Then
        iCount% = iCount% + 1
    End If
Next
'if therea re no cards, exit the sub
If iCount% = 0 Then
    imgTICSetup(0).Picture = LoadPicture()
    imgTICSetup(0).Visible = False
    Exit Sub
End If
For z = 0 To iCount% - 1 'if there are, then make a loop from
            '0 to how much icount is minus 1
    If First Then 'if its the first card...
        First = False 'false it
        Deck1.GetAnotherCard = CardsInHand(0) 'change the deck to the
                    'card
        imgTICSetup(0).Picture = Deck1.Picture 'set the picture
                    'to the cards picture
        GoTo nNext 'and contine on with the loop
    Else
        i = imgTICSetup.UBound + 1 'get the next index of imgticsetup
        Load imgTICSetup(i) 'load a new one
        'make sure the cards don't run off the edge
        If r% >= (Frame2.Width - lstCurrent.Width - Label3(0).Width) Then
            q% = q% + (imgTICSetup(0).Top / 2)
            r% = imgTICSetup(0).Left
        End If
        'set the diminsions of the card
        imgTICSetup(i).Top = q%
        imgTICSetup(i).Left = r%
        'put it on top
        imgTICSetup(i).ZOrder (0)
        'get the picture
        Deck1.GetAnotherCard = CardsInHand(z)
        imgTICSetup(i).Picture = Deck1.Picture 'set the picture
    End If
    'get a new left
    r% = r% + intW%
nNext:
Next
'make them visible
For z = 0 To imgTICSetup.UBound
    imgTICSetup(z).Visible = True
Next
End Sub

Sub CreateHand(strData$)
On Error Resume Next
Dim a As Integer 'counter
If Left(strData$, 3) = "§©¨" Then  'if it contains the special code
    If GameEnd = True Then GameEnd = False 'make sure to set the game end to false
    Frame1.Enabled = True 'enable frame 1
    strData$ = Right(strData$, Len(strData$) - 3) 'trim off code
    'get the round #
    Round$ = Val(Mid$(strData$, 1, InStr(1, strData$, "¤")))
    'make the round number visible to the user
    lblRound.Caption = Round$
    'set the last turn to false
    LastTurn = False
    Dim First As Boolean 'to check if its the first card
    First = True 'make it true
    Dim X%, Y%, p%, q%, r% 'counters
    p% = InStr(1, strData$, "£") 'find "£" in the string
    For a = 1 To p% 'trim to off all characters to that char
        strData$ = Mid$(strData$, 2)
    Next
    'and set p to  the length to "¤" from the begining of the string
    p% = InStr(1, strData$, "¤")
    Dim intW%, intL% 'width and left of the card
    imgHand(0).Top = 240 'set the top
    imgHand(0).Left = 240 'and left of the card
    intW% = imgHand(0).Width / 3 'get the width / 3
    intL% = 240 'set the left to 240
    q% = imgHand(0).Top 'make q the top of the cards
    r% = intL% + intW% 'and make the new left the left and width/3 of the card
    Dim z As Integer 'counter
    'unload all the cards
    For z = 1 To imgHand.UBound
        Unload imgHand(z)
    Next
    'clear out cardsinhand
    Erase CardsInHand
    'set it all to default
    For i = 0 To UBound(CardsInHand)
        CardsInHand(i) = -1
    Next
    'begin a loop through to make the hand
    For z = 1 To Val(Round$)
        'get the # of the card from the string
        Y% = Val(Mid$(strData$, 1, p%))
        If First Then 'if its the first card...
            First = False 'set first to false
            Deck1.GetAnotherCard = Y% 'change the decks pic to the new card
            imgHand(0).Picture = Deck1.Picture 'set the hand pic to deck1's pic
            CardsInHand(0) = Y% 'and set the cardsinhand(0) to the card #
            GoTo nNext: 'continue the loop
        Else
            'get the ubound of the hand
            i = imgHand.UBound + 1
            Load imgHand(i) 'load the next index
            If r% >= imgTrash.Left Then 'make sure the hand
                    'doesnt overlap the trash
                q% = q% + (imgHand(0).Top / 2)
                r% = imgHand(0).Left
            End If
            'set the top and left of the card
            imgHand(i).Top = q%
            imgHand(i).Left = r%
            'make it on top
            imgHand(i).ZOrder (0)
            'change deck1 to the pic
            Deck1.GetAnotherCard = Y%
            'set cardsinhand(of i) to the card #
            CardsInHand(i) = Val(Y%)
            'set the picture of imghand to deck1's picture
            imgHand(i).Picture = Deck1.Picture
        End If
        'Debug.Print Y%
        'make a new left
        r% = r% + intW%
nNext:
        'get to the next card #
        p% = InStr(1, strData$, "£")
        'trim
        For a = 1 To p%
            strData$ = Mid$(strData$, 2)
        Next
        'set p to a new value
        p% = InStr(1, strData$, "¤")
    Next
End If
'make em visible
For i = 0 To imgHand.UBound
    imgHand(i).Visible = True
Next
End Sub

Sub DrawCard(strData$)
On Error GoTo DrawCard_Error
If Left(strData$, 3) = "¶®Æ" Then 'if there is the special code
    Dim a As Integer 'counter
    strData$ = Right(strData$, Len(strData$) - 3) 'trim off code
    Deck1.GetAnotherCard = Val(strData$) 'change the picture
    i = imgHand.UBound + 1 'get the new index
    Load imgHand(i) 'load the new index
    'set the top and left
    imgHand(i).Top = imgHand(0).Top
    imgHand(i).Left = imgHand(i - 1).Left + 400
    'set the pciture
    imgHand(i).Picture = Deck1.Picture
    'make it ontop
    imgHand(Index).ZOrder (0)
    'make it visible
    imgHand(i).Visible = True
    'set the card to the hand
    CardsInHand(Val(i)) = Val(strData$)
    'redraw the hand
    Call RedoHand
    'set the data so others can see what has happened
    ws.SendData "©ÈÇ" & YName$ & " drew a card from the deck."
End If
On Error GoTo 0
Exit Sub
DrawCard_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: DrawCard in Form, Form1"
End Sub

Sub EndTurn()
On Error GoTo EndTurn_Error
If LastTurn = True Then 'if its the last turn then
    Frame1.Enabled = False 'disable frame1
    LastTurn = False 'and set last turn to false
End If
On Error GoTo 0
Exit Sub
EndTurn_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: EndTurn in Form, Form1"
End Sub

Sub FlipFirstOver(strData$)
Dim a As Integer
On Error GoTo FlipFirstOver_Error
If Left(strData$, 3) = "ÑÐª" Then  'if theres the special code
    'trim it off
    strData$ = Right(strData$, Len(strData$) - 3)
    Dim v% 'temp variable
    v% = Val(strData$) 'get the card #
    Deck1.GetAnotherCard = v% 'change the deck pic to the card #
    imgDiscard.Picture = Deck1.Picture 'and change the discard pic to the card #
End If
On Error GoTo 0
Exit Sub
FlipFirstOver_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: FlipFirstOver in Form, Form1"
End Sub

Sub GenerateMenu()
Dim X%, Y% 'counter variables
Y% = 2200 'set y to the default height of the menu
imgFile.Height = 200 'and set imgfile to 200
X% = picMenu.Width 'set x equal to the width of picmenu
For i = 0 To Y% Step 100 'loop through y 100 at a time
    imgFile.Height = i 'and make imgfile's height what i is
    DoEvents
Next
'set imgfile's height to picmenu's height
imgFile.Height = picMenu.Height
'set picmenus top to imgfiles top
picMenu.Top = imgFile.Top
'set picmenus left to the immeediate right of imgfile
picMenu.Left = imgFile.Left + imgFile.Width
'shrink picmenu to 0 width
picMenu.Width = 0
'make it visible
picMenu.Visible = True
For i = 0 To X% Step 80 'loop through x 80 at a time
    picMenu.Width = i 'set the width to i
    DoEvents
Next
End Sub

Sub GetTopCard(strData$)
On Error GoTo GetTopCard_Error
If Left(strData$, 3) = "×ÿ¡" Then 'if theres the special code
    Dim a As Integer 'counter
    'trim off code
    strData$ = Right(strData$, Len(strData$) - 3)
    'change the decks pic to the new card #
    Deck1.GetAnotherCard = Val(strData$)
    'get a new index
    i = imgHand.UBound + 1
    'load the new index
    Load imgHand(i)
    'set the top and left
    imgHand(i).Top = imgHand(0).Top
    imgHand(i).Left = imgHand(i - 1).Left + imgHand(0).Width / 3
    'set the picture
    imgHand(i).Picture = Deck1.Picture
    'make it ontop
    imgHand(Index).ZOrder (0)
    'set the card into the hand array
    CardsInHand(Val(i)) = Val(strData$)
    'redraw the hand
    Call RedoHand
End If
On Error GoTo 0
Exit Sub
GetTopCard_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: GetTopCard in Form, Form1"
End Sub

Sub GetTurn(strData$)
On Error GoTo GetTurn_Error
If Left(strData$, 3) = "µôÐ" Then 'if theres the special code
    Dim a As Integer 'counter
    strData$ = Right(strData$, Len(strData$) - 3) 'trim off code
    Dim s$ 'temp string vairbale
    s$ = strData$ 'set it to strdata
    If InStr(1, s$, "and now ") Then 'if it finds "and now " in there, continue
        a = InStr(1, s$, "and now ") 'get the position of the starting of
                '"and now "
        Dim c% 'temp integer
        c% = InStr(a, s$, "'s t") 'get the postition of "'s t"
        s$ = Mid$(s$, a + 8, c% - (a + 8)) 'and pull out the name
    End If
    lblTurn.Caption = "" 'clear lblturn
    lblTurn.Caption = s$ 'set it to the new turn
    'MsgBox strData$
    If InStr(1, strData$, "and now ") Then
        txtChat.Text = txtChat.Text & strData$ & vbCrLf 'and send strdata to the chat box
    Else
        'if it get just a name, its most likely the begining of the game or round,
        'so tell the user that
        txtChat.Text = txtChat.Text & "This round will start with " & strData$ & "." & vbCrLf
    End If
End If
On Error GoTo 0
Exit Sub
GetTurn_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: GetTurn in Form, Form1"
End Sub

Sub HAllHButtons()
On Error GoTo HAllHButtons_Error
'set all the highlighted buttons to invisible
imgHSend.Visible = False
Command1.ForeColor = vbBlack
cmdTIC.ForeColor = vbBlack
cmdRedo.ForeColor = vbBlack
cmdCancel.ForeColor = vbBlack
cmdDone.ForeColor = vbBlack
cmdSet.ForeColor = vbGreen
imghSet.Visible = False
imgHDone.Visible = False
imgHCancel.Visible = False
imgHRedo.Visible = False
'if its not drawing the menu, make the menu invisible
If DrawingMenu = False Then picMenu.Visible = False: imgFile.Height = 200: picSubMenu.Visible = False
On Error GoTo 0
Exit Sub
HAllHButtons_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: HAllHButtons in Form, Form1"
End Sub

Sub HMenu(Index%, IsSubMenu As Boolean)
If DrawingMenu = False Then 'if its not drawing the menu
    If IsSubMenu = False Then 'if issubmenu is false
        For i = 0 To lblMenu.UBound 'unhighlight all the options
            If i <> Index% Then lblMenu(i).BackStyle = 0
        Next
        'make sure the submenu should be there or not
        If picSubMenu.Visible = True Then
            If Index% <> 0 Then
                picSubMenu.Visible = False
            End If
        End If
        'set the backstyle to opaque
        lblMenu(Index%).BackStyle = 1
        'and the color to green
        lblMenu(Index%).BackColor = &HFF00&
    ElseIf IsSubMenu = True Then 'if they are on the submenu
        For i = 0 To lblSubMenu.UBound 'unhighlight everything
            If i <> Index% Then lblSubMenu(i).BackStyle = 0
        Next
        'set the backstyle to opaque
        lblSubMenu(Index%).BackStyle = 1
        'and make it green
        lblSubMenu(Index%).BackColor = &HFF00&
    End If
End If
End Sub

Sub InvokeCheat(strData$)
On Error GoTo InvokeCheat_Error
If Left(strData$, 5) = "§ÞÏkë" Then 'if theres the cheat data
    Dim c%, i As Integer 'counters
    c% = CInt(Right(strData$, Len(strData$) - 5)) 'get the card #
    If c% <> -1 Then 'if its not -1
        'add it to the hand
        For i = 0 To UBound(CardsInHand)
            If CardsInHand(i) = -1 Then
                CardsInHand(i) = c%
                Call RedoHand 'redraw the hand
                Exit For
            End If
        Next
    Else
        'if the card wasnt there, add the original card back
        For i = 0 To UBound(CardsInHand)
            If CardsInHand(i) = -1 Then
                CardsInHand(i) = BackUpBuffer
                Call RedoHand 'redraw the hand
                Exit For
            End If
        Next
    End If
End If
On Error GoTo 0
Exit Sub
InvokeCheat_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: InvokeCheat in Form, Form1"
End Sub

Sub RedoHand()
Dim a As Integer, q%, r% 'counters/value holders
Dim b As Integer 'counter
On Error GoTo RedoHand_Error
b = 0 'set to default
'simplify the hand, get rid of stray -1's
For a = 0 To UBound(CardsInHand)
    If CardsInHand(a) <> -1 Then
        If a <> b Then
            CardsInHand(b) = CardsInHand(a)
            CardsInHand(a) = -1
        End If
        b = b + 1
    End If
Next
'unload all the images but index 0
For a = 1 To imgHand.UBound
    Unload imgHand(a)
Next
'temp width and left
Dim intW%, intL%
intW% = imgHand(0).Width / 3 'make it the width divided by 3
intL% = 240 'default left is 240
imgHand(0).Top = 240 'set the first card to 240x240
imgHand(0).Left = 240
q% = imgHand(0).Top 'get the default top
r% = intL% + intW% 'get the new left
imgHand(0).Picture = LoadPicture() 'clear the picture in the first card
Dim Counter% 'counter
Counter% = 0 'default
For a = 0 To UBound(CardsInHand)
    If CardsInHand(a) <> -1 Then
        'get the amount of cards in the hand
        Counter% = Counter% + 1
    End If
Next
'set the picture for the first card in the hand
Deck1.GetAnotherCard = CardsInHand(0)
imgHand(0).Picture = Deck1.Picture
For a = 1 To Counter% - 1 'begin a loop
    i = imgHand.UBound + 1 'get the new index
    Load imgHand(i) 'load the new index
    If r% >= imgTrash.Left Then 'make sure it doesnt overlap the trash
        q% = q% + (imgHand(0).Top / 2)
        r% = imgHand(0).Left
    End If
    imgHand(i).Top = q% 'set the top and left
    imgHand(i).Left = r%
    Deck1.GetAnotherCard = CardsInHand(a) 'get the picture of the card
    imgHand(i).Picture = Deck1.Picture
    imgHand(i).ZOrder (0) 'make it ontop
    r% = r% + intW% 'get a new left
Next
'make all the pics visible
For a = 0 To imgHand.UBound
    imgHand(a).Visible = True
Next
On Error GoTo 0
Exit Sub
RedoHand_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: RedoHand in Form, Form1"
End Sub

Sub RefreshlstCurrent()
Dim a() As Integer, X As Integer
'this sub is so the listbox won't scramble to cards
'(by scramble i mean when you drag a card over it,
'it jumbles the picture, and messes things up
On Error GoTo RefreshlstCurrent_Error
'if theres something in the lsitbox
If lstCardNames.ListCount <> 0 Then
    ReDim a(lstCardNames.ListCount - 1) 'redim a to the amount of entries
    For X = 0 To lstCardNames.ListCount - 1 'save all the entries to a
        a(X) = Val(lstCardNames.List(X))
    Next
    lstCurrent.Clear 'clear the listbox
    lstCardNames.Clear 'clear the numbers
    lstCurrent.Refresh 'refresh the visible listbox
    For X = 0 To UBound(a) 'readd all the things into the listbox
        lstCardNames.AddItem a(X)
        lstCurrent.AddItem Deck1.GetCardName(a(X)) 'get the cardname of the card #
    Next
End If
'if there is something in lstdiscardmask
If lstDiscardMask.ListCount <> 0 Then
    Erase a() 'clear out a
    ReDim a(lstDiscardMask.ListCount - 1) 'get the right amount
    For X = 0 To UBound(a) 'save all contents to a
        a(X) = Val(lstDiscardMask.List(X))
    Next
    'clear the listboxes
    lstDiscard.Clear
    lstDiscardMask.Clear 'refresh them
    lstDiscard.Refresh
    'readd everything in
    For X = 0 To UBound(a)
        lstDiscardMask.AddItem a(X)
        lstDiscard.AddItem Deck1.GetCardName(a(X)) 'get the correct card name of the card #
    Next
    lstDiscard.Refresh 'refresh the listbox
End If
On Error GoTo 0
Exit Sub
RefreshlstCurrent_Error:
End Sub

Sub SendNameBack(strData$)
'send the nickname ot the server
On Error GoTo SendNameBack_Error
If strData$ = "÷êÅ" Then 'if theres the code
    ws.SendData "÷êÅ" & YName$ ' send back the name
    'set the caption of the form to show were connected
    Me.Caption = "Tic Game Board [ver " & App.Major & "." & App.Minor & "], Connected as " & YName$ & " (" & ws.RemoteHost & ")"
    txtChat.Text = txtChat.Text & "Connected to " & ws.RemoteHost & " as " & YName$ & " at " & Time & "." & vbCrLf
End If
On Error GoTo 0
Exit Sub
SendNameBack_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: SendNameBack in Form, Form1"
End Sub

Sub GenerateSubMenu()
Dim tWid%, i As Integer 'temp value/counter
With picSubMenu
    tWid% = .Width 'set the width to the temp value
    .Top = Val(picMenu.Top) + (Val(lblMenu(0).Height) / 2) 'get the top positioning
    .Left = Val(picMenu.Left) + Val(picMenu.Width) 'get the left positioning
    .Width = 0 'make the width 0
    .Visible = True 'make it visible
    For i = 0 To tWid% Step 80 'loop through x 80 at a time
        picSubMenu.Width = i 'set the width to i
        DoEvents
    Next
End With
End Sub

Sub ServerIsRefreshing(strData$)
On Error GoTo ServerIsRefreshing_Error
If Left(strData$, 8) = "rÉFRË§µ¸" Then 'if the code
    'tell the user
    txtChat.Text = txtChat & vbCrLf & vbCrLf & "Server is refreshing!!!" & vbCrLf & "Please wait until it finishes." & vbCrLf & "Game will restart in about 5 to 10 seconds."
    FU.WaitFor 4000 'wait for 7 seconds before reconecting
    ws.Close 'close any connection
    ws.Connect 'reconnect to the server
End If
On Error GoTo 0
Exit Sub
ServerIsRefreshing_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: ServerIsRefreshing in Form, Form1"
End Sub
