VERSION 5.00
Begin VB.Form frmQuickMap 
   BackColor       =   &H00404040&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Q-Map"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   585
   ClientWidth     =   12780
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   6
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   618
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   852
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picFinish 
      Height          =   8895
      Left            =   -840
      ScaleHeight     =   8835
      ScaleWidth      =   12675
      TabIndex        =   13
      Top             =   -240
      Visible         =   0   'False
      Width           =   12735
      Begin ServerEditor.NumOnlyText txtLight 
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   2760
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
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
      Begin VB.TextBox txtRoomDesc 
         Height          =   2175
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   480
         Width           =   8655
      End
      Begin VB.TextBox txtRoomTitle 
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   120
         Width           =   3375
      End
      Begin VB.CommandButton cmdDone 
         Caption         =   "Done"
         Height          =   255
         Left            =   11760
         TabIndex        =   14
         Top             =   8520
         Width           =   855
      End
      Begin ServerEditor.NumOnlyText txtDeathRoom 
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   3120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "232"
         AllowNeg        =   0   'False
         Align           =   0
         MaxLength       =   0
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin VB.Label lblLab 
         AutoSize        =   -1  'True
         Caption         =   "Death Room:"
         Height          =   120
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   3240
         Width           =   990
      End
      Begin VB.Label lblLab 
         AutoSize        =   -1  'True
         Caption         =   "Light in the room:"
         Height          =   120
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   2880
         Width           =   1620
      End
      Begin VB.Label lblLab 
         AutoSize        =   -1  'True
         Caption         =   "Generic Room Description:"
         Height          =   120
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   2250
      End
      Begin VB.Label lblLab 
         AutoSize        =   -1  'True
         Caption         =   "Generic Room Title:"
         Height          =   120
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1710
      End
   End
   Begin VB.PictureBox picJoin 
      Height          =   2535
      Left            =   3240
      ScaleHeight     =   2475
      ScaleWidth      =   6555
      TabIndex        =   42
      Top             =   2880
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ComboBox cboDirs 
         Height          =   240
         ItemData        =   "frmQuickMap.frx":0000
         Left            =   3720
         List            =   "frmQuickMap.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cboRooms 
         Height          =   240
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   480
         Width           =   3975
      End
      Begin VB.CommandButton cmdRoomOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   4920
         TabIndex        =   44
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdRoomCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5760
         TabIndex        =   43
         Top             =   2040
         Width           =   735
      End
      Begin ServerEditor.NumOnlyText txtRoom 
         Height          =   255
         Left            =   1680
         TabIndex        =   46
         Top             =   480
         Width           =   495
         _ExtentX        =   661
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
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
      Begin VB.Label lblLab 
         AutoSize        =   -1  'True
         Caption         =   "to get to the new area."
         Height          =   120
         Index           =   10
         Left            =   4440
         TabIndex        =   50
         Top             =   960
         Width           =   2070
      End
      Begin VB.Label lblLab 
         AutoSize        =   -1  'True
         Caption         =   "From the old room, you would have to go "
         Height          =   120
         Index           =   9
         Left            =   120
         TabIndex        =   49
         Top             =   960
         Width           =   3600
      End
      Begin VB.Label lblLab 
         AutoSize        =   -1  'True
         Caption         =   "Join To Room:"
         Height          =   120
         Index           =   8
         Left            =   360
         TabIndex        =   48
         Top             =   480
         Width           =   1170
      End
      Begin VB.Label lblLab 
         AutoSize        =   -1  'True
         Caption         =   "Joining to another room"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   11
         Left            =   1440
         TabIndex        =   45
         Top             =   120
         Width           =   3810
      End
   End
   Begin VB.PictureBox picPic 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
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
      Height          =   555
      Index           =   11
      Left            =   9360
      Picture         =   "frmQuickMap.frx":003C
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   41
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picPic 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
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
      Height          =   555
      Index           =   10
      Left            =   8640
      Picture         =   "frmQuickMap.frx":10AE
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   40
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picPic 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
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
      Height          =   555
      Index           =   9
      Left            =   7920
      Picture         =   "frmQuickMap.frx":2120
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   39
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picPic 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
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
      Height          =   555
      Index           =   8
      Left            =   7200
      Picture         =   "frmQuickMap.frx":3192
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   38
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picPic 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
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
      Height          =   555
      Index           =   7
      Left            =   6480
      Picture         =   "frmQuickMap.frx":4204
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   37
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picPic 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
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
      Height          =   555
      Index           =   6
      Left            =   5760
      Picture         =   "frmQuickMap.frx":5276
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   36
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picPic 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
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
      Height          =   555
      Index           =   5
      Left            =   5040
      Picture         =   "frmQuickMap.frx":62E8
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   35
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picPic 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
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
      Height          =   555
      Index           =   4
      Left            =   4320
      Picture         =   "frmQuickMap.frx":735A
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   34
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picPic 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
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
      Height          =   555
      Index           =   3
      Left            =   3600
      Picture         =   "frmQuickMap.frx":83CC
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   33
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picPic 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
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
      Height          =   555
      Index           =   2
      Left            =   2880
      Picture         =   "frmQuickMap.frx":943E
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   32
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picPic 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
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
      Height          =   555
      Index           =   1
      Left            =   2160
      Picture         =   "frmQuickMap.frx":A4B0
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   31
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picPic 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
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
      Height          =   555
      Index           =   0
      Left            =   1440
      Picture         =   "frmQuickMap.frx":B522
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   30
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picDoorDef 
      Height          =   2535
      Left            =   3240
      ScaleHeight     =   2475
      ScaleWidth      =   6555
      TabIndex        =   25
      Top             =   2880
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CheckBox chkUse 
         Caption         =   "Use same for all doors made after this one."
         Height          =   255
         Left            =   3240
         TabIndex        =   4
         Top             =   1680
         Width           =   3255
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5760
         TabIndex        =   6
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   2040
         Width           =   735
      End
      Begin ServerEditor.NumOnlyText txtBash 
         Height          =   255
         Left            =   3240
         TabIndex        =   2
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
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
      Begin ServerEditor.NumOnlyText txtKey 
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         Top             =   600
         Width           =   495
         _ExtentX        =   661
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
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
      Begin VB.ComboBox cboKeys 
         Height          =   240
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   2535
      End
      Begin ServerEditor.NumOnlyText txtPick 
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
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
      Begin VB.Label lblLab 
         AutoSize        =   -1  'True
         Caption         =   "Picklocks Req. To Pick Lock:"
         Height          =   120
         Index           =   7
         Left            =   530
         TabIndex        =   29
         Top             =   1320
         Width           =   2520
      End
      Begin VB.Label lblLab 
         AutoSize        =   -1  'True
         Caption         =   "Str. Req. To Bash:"
         Height          =   120
         Index           =   6
         Left            =   1440
         TabIndex        =   28
         Top             =   960
         Width           =   1620
      End
      Begin VB.Label lblLab 
         AutoSize        =   -1  'True
         Caption         =   "Key:"
         Height          =   120
         Index           =   5
         Left            =   2700
         TabIndex        =   27
         Top             =   600
         Width           =   360
      End
      Begin VB.Label lblLab 
         AutoSize        =   -1  'True
         Caption         =   "Door Definition"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   1800
         TabIndex        =   26
         Top             =   120
         Width           =   2490
      End
   End
   Begin VB.OptionButton optDoorDef 
      Caption         =   "Door Definition"
      Height          =   255
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   360
      Width           =   1575
   End
   Begin VB.CheckBox chkMake 
      BackColor       =   &H00404040&
      Caption         =   "If chosen room isn't a defined room, make it so it is."
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9120
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CheckBox chkKeep 
      BackColor       =   &H00404040&
      Caption         =   "Keep last chosen room selected"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OptionButton optExitDef 
      Caption         =   "Exit Definition"
      Height          =   255
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   1575
   End
   Begin VB.OptionButton optRoomDef 
      Caption         =   "Room Definition"
      Height          =   255
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.PictureBox picRoom 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
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
      Height          =   555
      Index           =   0
      Left            =   120
      Picture         =   "frmQuickMap.frx":C594
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdMapFinished 
      Caption         =   "Done"
      Height          =   255
      Left            =   11880
      TabIndex        =   23
      Top             =   9000
      Width           =   855
   End
   Begin VB.Line lnJoin 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      Index           =   0
      Visible         =   0   'False
      X1              =   312
      X2              =   328
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Line lnDoor 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Index           =   0
      Visible         =   0   'False
      X1              =   312
      X2              =   328
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line lnExit 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Index           =   0
      Visible         =   0   'False
      X1              =   312
      X2              =   328
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   120
      Left            =   3720
      TabIndex        =   10
      Top             =   15
      Width           =   90
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuEnv 
         Caption         =   "&Enviroment"
         Begin VB.Menu mnuNot 
            Caption         =   "&Not A Room"
         End
         Begin VB.Menu mnuDash00001 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIndoor 
            Caption         =   "&Indoor"
         End
         Begin VB.Menu mnuOutDoor 
            Caption         =   "&Outdoor"
         End
      End
      Begin VB.Menu mnuRoom 
         Caption         =   "&Room"
         Begin VB.Menu mnuTitle 
            Caption         =   "&Title"
         End
         Begin VB.Menu mnuDescrption 
            Caption         =   "&Description"
         End
      End
      Begin VB.Menu mnuJoin 
         Caption         =   "&Toggle Join Map"
      End
   End
End
Attribute VB_Name = "frmQuickMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type DoorDef
    lKey As Long
    lBash As Long
    lPick As Long
    bHasDoor As Boolean
End Type
Private Type RoomDef
    sExits As String
    lTag As Long
    lX As Long
    lY As Long
    lRID As Long
    tDoors(7) As DoorDef
End Type
Private Rooms() As RoomDef
Private lOldIn As Long
Private Cols As Long
Private Rows As Long
Private Type typLnDef
    lX1 As Long
    lX2 As Long
    lY1 As Long
    lY2 As Long
End Type
Private lnHold As typLnDef
Private lNewIn As Long
Private lDDir As Long

Private Sub cboKeys_Change()
cboKeys_Click
End Sub

Private Sub cboKeys_Click()
txtKey.Text = Mid$(cboKeys.list(cboKeys.ListIndex), 2, InStr(2, cboKeys.list(cboKeys.ListIndex), ")"))
End Sub

Private Sub cboRoom_Change()
txtRoom.Text = Mid$(cboRooms.list(cboRooms.ListIndex), 2, InStr(2, cboRooms.list(cboRooms.ListIndex), ")"))
End Sub

Private Sub cboRoom_Click()
cboRoom_Change
End Sub

Private Sub cmdCancel_Click()
lNewIn = 0
Set picRoom(lOldIn).Picture = picPic(Rooms(lOldIn).lTag).Picture
lOldIn = 0
picDoorDef.Visible = False
End Sub

Private Sub cmdDone_Click()
Dim lID As Long
Dim i As Long
Dim a As Long
Dim t As Long
Dim sD As String
Dim sK As String
Dim sB As String
Dim sP As String
Dim Arr() As String
lID = dbMap(UBound(dbMap)).lRoomID + 1
For i = LBound(Rooms) To UBound(Rooms)
    With Rooms(i)
        .lRID = lID
    End With
    lID = lID + 1
Next
modUpdateDatabase.UpdateMRSSets
For i = LBound(Rooms) To UBound(Rooms)
    If Rooms(i).lTag <> 0 Then
        With Rooms(i)
            If .sExits <> "" Then
                Arr = Split(Replace$(.sExits, ":", ""), ";")
                .sExits = ""
                For a = LBound(Arr) To UBound(Arr)
                    If Arr(a) <> "" Then
                        Select Case Arr(a)
                            Case "n"
                                .sExits = .sExits & "n*" & Rooms(i - 1).lRID & ";"
                            Case "s"
                                .sExits = .sExits & "s*" & Rooms(i + 1).lRID & ";"
                            Case "e"
                                .sExits = .sExits & "e*" & Rooms(i + Rows).lRID & ";"
                            Case "w"
                                .sExits = .sExits & "w*" & Rooms(i - Rows).lRID & ";"
                            Case "ne"
                                .sExits = .sExits & "ne" & Rooms(i + (Rows - 1)).lRID & ";"
                            Case "nw"
                                .sExits = .sExits & "nw" & Rooms(i - (Rows + 1)).lRID & ";"
                            Case "se"
                                .sExits = .sExits & "se" & Rooms(i + (Rows + 1)).lRID & ";"
                            Case "sw"
                                .sExits = .sExits & "sw" & Rooms(i - (Rows - 1)).lRID & ";"
                            Case Else
                                .sExits = .sExits & Arr(a) & ";"
                        End Select
                    End If
                Next
            End If
            For a = LBound(.tDoors) To UBound(.tDoors)
                With .tDoors(a)
                    If .bHasDoor = True Then
                        sD = sD & ":1;"
                        sK = sK & ":" & .lKey & ";"
                        sB = sB & ":" & .lBash & ";"
                        sP = sP & ":" & .lPick & ";"
                    Else
                        sD = sD & ":0;"
                        sK = sK & ":0;"
                        sB = sB & ":0;"
                        sP = sP & ":0;"
                    End If
                End With
            Next
        End With

        Erase Arr
        With MRSMAP
            .AddNew
            !RoomID = Rooms(i).lRID
            Arr = Split(Rooms(i).sExits, ";")
            For a = LBound(Arr) To UBound(Arr)
                If Arr(a) <> "" Then
                    t = CLng(Val(Mid$(Arr(a), 3)))
                    Select Case Left$(Arr(a), 2)
                        Case "n*"
                            !North = t
                        Case "s*"
                            !South = t
                        Case "e*"
                            !East = t
                        Case "w*"
                            !West = t
                        Case "ne"
                            !NorthEast = t
                        Case "nw"
                            !NorthWest = t
                        Case "se"
                            !SouthEast = t
                        Case "sw"
                            !SouthWest = t
                    End Select
                End If
            Next
            If IsNull(!North) Then !North = "0"
            If IsNull(!South) Then !South = "0"
            If IsNull(!East) Then !East = "0"
            If IsNull(!West) Then !West = "0"
            If IsNull(!NorthEast) Then !NorthEast = "0"
            If IsNull(!NorthWest) Then !NorthWest = "0"
            If IsNull(!SouthEast) Then !SouthEast = "0"
            If IsNull(!SouthWest) Then !SouthWest = "0"
            If IsNull(!UP) Then !UP = "0"
            If IsNull(!Down) Then !Down = "0"
            !Door = sD
            !key = sK
            !Pick = sP
            !Bash = sB
            !Items = "0"
            !RoomTitle = txtRoomTitle.Text
            !RoomDesc = txtRoomDesc.Text
            !Monsters = "0"
            !MaxRegen = "1"
            !Type = "0"
            !ShopItems = "0"
            !MobGroup = "0"
            !Gold = "0"
            !SpecialMon = "0"
            !SpecialItem = "0"
            !Light = txtLight.Text
            !Hidden = "0"
            !Scripting = "0"
            !SafeRoom = "0"
            !DeathRoom = txtDeathRoom.Text
            !InDoor = Rooms(i).lTag - 1
            !TrainClass = "0"
            .Update
        End With
    End If
Next
MsgBox "Done."
modUpdateDatabase.SaveMemoryToDatabase Map
End Sub

Private Sub cmdMapFinished_Click()
picFinish.Top = 0
picFinish.Left = 0
picFinish.Width = Me.ScaleWidth
picFinish.Height = Me.ScaleHeight
With cmdDone
    .Top = picFinish.ScaleHeight - .Height - 10
    .Left = picFinish.ScaleWidth - .Width - 10
End With
picFinish.Visible = True
End Sub

Private Sub cmdOK_Click()
picDoorDef.Visible = False
lDDir = -1
If lOldIn - Rows = lNewIn Then 'west
    If InStr(1, Rooms(lOldIn).sExits, ":w;") Then
        Load lnDoor(lnDoor.UBound + 1)
        With lnDoor(lnDoor.UBound)
            .X1 = picRoom(lOldIn).Left - 1
            .X2 = .X1
            .Y1 = picRoom(lOldIn).Top + ((picRoom(lOldIn).Height \ 2) \ 2) + 3
            .Y2 = picRoom(lNewIn).Top + (picRoom(lNewIn).Height) - ((picRoom(lOldIn).Height \ 2) \ 2) - 3
            .Visible = True
        End With
        lnDoor(lnDoor.UBound).ZOrder 0
        Load lnDoor(lnDoor.UBound + 1)
        With lnDoor(lnDoor.UBound)
            .X1 = picRoom(lNewIn).Width + picRoom(lNewIn).Left
            .X2 = .X1
            .Y1 = picRoom(lOldIn).Top + ((picRoom(lOldIn).Height \ 2) \ 2) + 3
            .Y2 = picRoom(lNewIn).Top + (picRoom(lNewIn).Height) - ((picRoom(lOldIn).Height \ 2) \ 2) - 3
            .Visible = True
        End With
        lnDoor(lnDoor.UBound).ZOrder 0
        lblLabel.Caption = "You made a door to the west."
    End If
    lDDir = 3
ElseIf lOldIn + Rows = lNewIn Then 'east
    If InStr(1, Rooms(lOldIn).sExits, ":e;") Then
        Load lnDoor(lnDoor.UBound + 1)
        With lnDoor(lnDoor.UBound)
            .X1 = picRoom(lOldIn).Width + picRoom(lOldIn).Left
            .X2 = .X1
            .Y1 = picRoom(lOldIn).Top + ((picRoom(lOldIn).Height \ 2) \ 2) + 3
            .Y2 = picRoom(lOldIn).Top + (picRoom(lNewIn).Height) - ((picRoom(lOldIn).Height \ 2) \ 2) - 3
            .Visible = True
        End With
        lnDoor(lnDoor.UBound).ZOrder 0
        Load lnDoor(lnDoor.UBound + 1)
        With lnDoor(lnDoor.UBound)
            .X1 = picRoom(lNewIn).Left - 1
            .X2 = .X1
            .Y1 = picRoom(lNewIn).Top + ((picRoom(lOldIn).Height \ 2) \ 2) + 3
            .Y2 = picRoom(lNewIn).Top + (picRoom(lNewIn).Height) - ((picRoom(lNewIn).Height \ 2) \ 2) - 3
            .Visible = True
        End With
        lnDoor(lnDoor.UBound).ZOrder 0
        lblLabel.Caption = "You made a door to the east."
    End If
    lDDir = 2
ElseIf lOldIn - 1 = lNewIn Then 'north
    If InStr(1, Rooms(lOldIn).sExits, ":n;") Then
        Load lnDoor(lnDoor.UBound + 1)
        With lnDoor(lnDoor.UBound)
            .X1 = picRoom(lOldIn).Left + ((picRoom(lOldIn).Width \ 2) \ 2) + 3
            .X2 = picRoom(lOldIn).Left + picRoom(lOldIn).Width - ((picRoom(lOldIn).Width \ 2) \ 2) - 3
            .Y1 = picRoom(lOldIn).Top - 1
            .Y2 = .Y1
            .Visible = True
        End With
        lnDoor(lnDoor.UBound).ZOrder 0
        Load lnDoor(lnDoor.UBound + 1)
        With lnDoor(lnDoor.UBound)
            .X1 = picRoom(lNewIn).Left + ((picRoom(lNewIn).Width \ 2) \ 2) + 3
            .X2 = picRoom(lNewIn).Left + picRoom(lNewIn).Width - ((picRoom(lNewIn).Width \ 2) \ 2) - 3
            .Y1 = picRoom(lNewIn).Top + picRoom(lNewIn).Height '+ 1
            .Y2 = .Y1
            .Visible = True
        End With
        lnDoor(lnDoor.UBound).ZOrder 0
        lblLabel.Caption = "You made a door to the north."
        lDDir = 0
    End If
ElseIf lOldIn + 1 = lNewIn Then 'south
    If InStr(1, Rooms(lOldIn).sExits, ":s;") Then
        Load lnDoor(lnDoor.UBound + 1)
        With lnDoor(lnDoor.UBound)
            .X1 = picRoom(lNewIn).Left + ((picRoom(lNewIn).Width \ 2) \ 2) + 3
            .X2 = picRoom(lNewIn).Left + picRoom(lNewIn).Width - ((picRoom(lNewIn).Width \ 2) \ 2) - 3
            .Y1 = picRoom(lNewIn).Top - 1
            .Y2 = .Y1
            .Visible = True
        End With
        lnDoor(lnDoor.UBound).ZOrder 0
        Load lnDoor(lnDoor.UBound + 1)
        With lnDoor(lnDoor.UBound)
            .X1 = picRoom(lOldIn).Left + ((picRoom(lOldIn).Width \ 2) \ 2) + 3
            .X2 = picRoom(lOldIn).Left + picRoom(lOldIn).Width - ((picRoom(lOldIn).Width \ 2) \ 2) - 3
            .Y1 = picRoom(lOldIn).Top + picRoom(lOldIn).Height '+ 1
            .Y2 = .Y1
            .Visible = True
        End With
        lnDoor(lnDoor.UBound).ZOrder 0
        lblLabel.Caption = "You made a door to the south."
        lDDir = 1
    End If
ElseIf lOldIn + (Rows + 1) = lNewIn Then 'south east
    If InStr(1, Rooms(lOldIn).sExits, ":se;") Then
        Load lnDoor(lnDoor.UBound + 1)
        With lnDoor(lnDoor.UBound)
            .X1 = picRoom(lOldIn).Left + picRoom(lOldIn).Width - ((picRoom(lOldIn).Width \ 2) \ 2) + 6
            .X2 = picRoom(lOldIn).Left + picRoom(lOldIn).Width + ((picRoom(lOldIn).Width \ 2) \ 2) - 6
            .Y1 = picRoom(lOldIn).Top + picRoom(lOldIn).Height + ((picRoom(lOldIn).Height \ 2) \ 2) - 6
            .Y2 = picRoom(lOldIn).Top + picRoom(lOldIn).Height - ((picRoom(lOldIn).Height \ 2) \ 2) + 6
            .Visible = True
        End With
        lnDoor(lnDoor.UBound).ZOrder 0
        Load lnDoor(lnDoor.UBound + 1)
        With lnDoor(lnDoor.UBound)
            .X1 = picRoom(lNewIn).Left + ((picRoom(lNewIn).Width \ 2) \ 2) - 6
            .X2 = picRoom(lNewIn).Left - ((picRoom(lNewIn).Width \ 2) \ 2) + 6
            .Y1 = picRoom(lNewIn).Top - ((picRoom(lNewIn).Height \ 2) \ 2) + 6
            .Y2 = picRoom(lNewIn).Top + ((picRoom(lNewIn).Height \ 2) \ 2) - 6
            .Visible = True
        End With
        lnDoor(lnDoor.UBound).ZOrder 0
        lblLabel.Caption = "You made a door to the southeast."
        lDDir = 7
    End If
ElseIf lOldIn - (Rows + 1) = lNewIn Then 'north west
    If InStr(1, Rooms(lOldIn).sExits, ":nw;") Then
        Load lnDoor(lnDoor.UBound + 1)
        With lnDoor(lnDoor.UBound)
            .X1 = picRoom(lNewIn).Left + picRoom(lNewIn).Width - ((picRoom(lNewIn).Width \ 2) \ 2) + 6
            .X2 = picRoom(lNewIn).Left + picRoom(lNewIn).Width + ((picRoom(lNewIn).Width \ 2) \ 2) - 6
            .Y1 = picRoom(lNewIn).Top + picRoom(lNewIn).Height + ((picRoom(lNewIn).Height \ 2) \ 2) - 6
            .Y2 = picRoom(lNewIn).Top + picRoom(lNewIn).Height - ((picRoom(lNewIn).Height \ 2) \ 2) + 6
            .Visible = True
        End With
        lnDoor(lnDoor.UBound).ZOrder 0
        Load lnDoor(lnDoor.UBound + 1)
        With lnDoor(lnDoor.UBound)
            .X1 = picRoom(lOldIn).Left + ((picRoom(lOldIn).Width \ 2) \ 2) - 6
            .X2 = picRoom(lOldIn).Left - ((picRoom(lOldIn).Width \ 2) \ 2) + 6
            .Y1 = picRoom(lOldIn).Top - ((picRoom(lOldIn).Height \ 2) \ 2) + 6
            .Y2 = picRoom(lOldIn).Top + ((picRoom(lOldIn).Height \ 2) \ 2) - 6
            .Visible = True
        End With
        lnDoor(lnDoor.UBound).ZOrder 0
        lblLabel.Caption = "You made a door to the northwest."
        lDDir = 4
    End If

ElseIf lOldIn + (Rows - 1) = lNewIn Then 'north east
    If InStr(1, Rooms(lOldIn).sExits, ":ne;") Then
        Load lnDoor(lnDoor.UBound + 1)
        With lnDoor(lnDoor.UBound)
            .X1 = picRoom(lOldIn).Left + picRoom(lOldIn).Width + ((picRoom(lOldIn).Width \ 2) \ 2) - 6
            .X2 = picRoom(lOldIn).Left + picRoom(lOldIn).Width - ((picRoom(lOldIn).Width \ 2) \ 2) + 6
            .Y1 = picRoom(lOldIn).Top + ((picRoom(lOldIn).Height \ 2) \ 2) - 6
            .Y2 = picRoom(lOldIn).Top - ((picRoom(lOldIn).Height \ 2) \ 2) + 6
            .Visible = True
        End With
        lnDoor(lnDoor.UBound).ZOrder 0
        Load lnDoor(lnDoor.UBound + 1)
        With lnDoor(lnDoor.UBound)
            .X1 = picRoom(lNewIn).Left + ((picRoom(lNewIn).Width \ 2) \ 2) - 6
            .X2 = picRoom(lNewIn).Left - ((picRoom(lNewIn).Width \ 2) \ 2) + 6
            .Y2 = picRoom(lNewIn).Top + picRoom(lNewIn).Height - ((picRoom(lNewIn).Height \ 2) \ 2) + 6
            .Y1 = picRoom(lNewIn).Top + picRoom(lNewIn).Height + ((picRoom(lNewIn).Height \ 2) \ 2) - 6
            .Visible = True
        End With
        lnDoor(lnDoor.UBound).ZOrder 0
        lblLabel.Caption = "You made a door to the northeast."
        lDDir = 5
    End If
ElseIf lOldIn - (Rows - 1) = lNewIn Then 'South West
    If InStr(1, Rooms(lOldIn).sExits, ":sw;") Then
        Load lnDoor(lnDoor.UBound + 1)
        With lnDoor(lnDoor.UBound)
            .X1 = picRoom(lNewIn).Left + picRoom(lNewIn).Width + ((picRoom(lNewIn).Width \ 2) \ 2) - 6
            .X2 = picRoom(lNewIn).Left + picRoom(lNewIn).Width - ((picRoom(lNewIn).Width \ 2) \ 2) + 6
            .Y1 = picRoom(lNewIn).Top + ((picRoom(lNewIn).Height \ 2) \ 2) - 6
            .Y2 = picRoom(lNewIn).Top - ((picRoom(lNewIn).Height \ 2) \ 2) + 6
            .Visible = True
        End With
        lnDoor(lnDoor.UBound).ZOrder 0
        Load lnDoor(lnDoor.UBound + 1)
        With lnDoor(lnDoor.UBound)
            .X1 = picRoom(lOldIn).Left + ((picRoom(lOldIn).Width \ 2) \ 2) - 6
            .X2 = picRoom(lOldIn).Left - ((picRoom(lOldIn).Width \ 2) \ 2) + 6
            .Y2 = picRoom(lOldIn).Top + picRoom(lOldIn).Height - ((picRoom(lOldIn).Height \ 2) \ 2) + 6
            .Y1 = picRoom(lOldIn).Top + picRoom(lOldIn).Height + ((picRoom(lOldIn).Height \ 2) \ 2) - 6
            .Visible = True
        End With
        lnDoor(lnDoor.UBound).ZOrder 0
        lblLabel.Caption = "You made a door to the southwest."
        lDDir = 6
    End If
Else
    lblLabel.Caption = "Invalid selection. (Too far away)"
    Exit Sub
End If
If lDDir <> -1 Then
    With Rooms(lOldIn).tDoors(lDDir)
        .lBash = txtBash.Text
        .lKey = txtKey.Text
        .lPick = txtPick.Text
        .bHasDoor = True
    End With
    With Rooms(lNewIn).tDoors(GetOppDoor(lDDir))
        .lBash = txtBash.Text
        .lKey = txtKey.Text
        .lPick = txtPick.Text
        .bHasDoor = True
    End With
End If
Set picRoom(lOldIn).Picture = picPic(Rooms(lOldIn).lTag).Picture
If chkKeep.Value = 1 Then
    lOldIn = lNewIn
    Set picRoom(lNewIn).Picture = picPic(Rooms(lNewIn).lTag + 9)
Else
    lOldIn = 0
End If
End Sub

Private Function GetOppDoor(lIndex As Long) As Long
Select Case lIndex
    Case 0
        GetOppDoor = 0
    Case 1
        GetOppDoor = 1
    Case 2
        GetOppDoor = 3
    Case 3
        GetOppDoor = 2
    Case 4
        GetOppDoor = 7
    Case 5
        GetOppDoor = 6
    Case 6
        GetOppDoor = 5
    Case 7
        GetOppDoor = 4
End Select
End Function

Private Sub cmdRoomOK_Click()
Select Case cboDirs.list(cboDirs.ListIndex)
    Case "W"
        If InStr(1, Rooms(lOldIn).sExits, ":w") = 0 Then
            Load lnJoin(lnJoin.UBound + 1)
            With lnJoin(lnJoin.UBound)
                .X1 = picRoom(lOldIn).Left
                .X2 = picRoom(Index).Left + picRoom(Index).Width
                .Y1 = picRoom(lOldIn).Top + (picRoom(lOldIn).Height \ 2)
                .Y2 = picRoom(Index).Top + (picRoom(Index).Height \ 2)
                .Visible = True
            End With
            With Rooms(lOldIn)
                .sExits = .sExits & ":w*" & txtRoom.Text & ";"
            End With
            With Rooms(Index)
                .sExits = Replace$(.sExits, ":e;", "")
                .sExits = .sExits & ":e;"
            End With
            lblLabel.Caption = "You made an exit to the west."
        End If
    Case "E"
        If InStr(1, Rooms(lOldIn).sExits, ":e;") = 0 Then
            Load lnJoin(lnJoin.UBound + 1)
            With lnJoin(lnJoin.UBound)
                .X1 = picRoom(lOldIn).Left + picRoom(lOldIn).Width
                .X2 = picRoom(Index).Left
                .Y1 = picRoom(lOldIn).Top + (picRoom(lOldIn).Height \ 2)
                .Y2 = picRoom(Index).Top + (picRoom(Index).Height \ 2)
                .Visible = True
            End With
            With Rooms(lOldIn)
                .sExits = Replace$(.sExits, ":e;", "")
                .sExits = .sExits & ":e;"
            End With
            With Rooms(Index)
                .sExits = Replace$(.sExits, ":w;", "")
                .sExits = .sExits & ":w;"
            End With
            lblLabel.Caption = "You made an exit to the east."
        'End If
            ElseIf lOldIn - 1 = Index Then 'north
                If InStr(1, Rooms(lOldIn).sExits, ":n;") = 0 Then
                    Load lnJoin(lnJoin.UBound + 1)
                    With lnJoin(lnJoin.UBound)
                        .X1 = picRoom(lOldIn).Left + (picRoom(lOldIn).Width \ 2)
                        .X2 = .X1
                        .Y1 = picRoom(lOldIn).Top
                        .Y2 = picRoom(Index).Top + picRoom(Index).Height
                        .Visible = True
                    End With
                    With Rooms(lOldIn)
                        .sExits = Replace$(.sExits, ":n;", "")
                        .sExits = .sExits & ":n;"
                    End With
                    With Rooms(Index)
                        .sExits = Replace$(.sExits, ":s;", "")
                        .sExits = .sExits & ":s;"
                    End With
                    lblLabel.Caption = "You made an exit to the north."
                End If
            ElseIf lOldIn + 1 = Index Then 'south
                If InStr(1, Rooms(lOldIn).sExits, ":s;") = 0 Then
                    Load lnJoin(lnJoin.UBound + 1)
                    With lnJoin(lnJoin.UBound)
                        .X1 = picRoom(lOldIn).Left + (picRoom(lOldIn).Width \ 2)
                        .X2 = .X1
                        .Y1 = picRoom(lOldIn).Top + picRoom(lOldIn).Height
                        .Y2 = picRoom(Index).Top
                        .Visible = True
                    End With
                    With Rooms(lOldIn)
                        .sExits = Replace$(.sExits, ":s;", "")
                        .sExits = .sExits & ":s;"
                    End With
                    With Rooms(Index)
                        .sExits = Replace$(.sExits, ":n;", "")
                        .sExits = .sExits & ":n;"
                    End With
                    lblLabel.Caption = "You made an exit to the south."
                End If
            ElseIf lOldIn + (Rows + 1) = Index Then 'south east
                If InStr(1, Rooms(lOldIn).sExits, ":se;") = 0 Then
                    Load lnJoin(lnJoin.UBound + 1)
                    With lnJoin(lnJoin.UBound)
                        .X1 = picRoom(lOldIn).Left + picRoom(lOldIn).Width
                        .X2 = picRoom(Index).Left
                        .Y1 = picRoom(lOldIn).Top + picRoom(lOldIn).Height
                        .Y2 = picRoom(Index).Top
                        .Visible = True
                    End With
                    With Rooms(lOldIn)
                        .sExits = Replace$(.sExits, ":se;", "")
                        .sExits = .sExits & ":se;"
                    End With
                    With Rooms(Index)
                        .sExits = Replace$(.sExits, ":nw;", "")
                        .sExits = .sExits & ":nw;"
                    End With
                    lblLabel.Caption = "You made an exit to the southeast."
                End If
            ElseIf lOldIn - (Rows + 1) = Index Then 'north west
                If InStr(1, Rooms(lOldIn).sExits, ":nw;") = 0 Then
                    Load lnJoin(lnJoin.UBound + 1)
                    With lnJoin(lnJoin.UBound)
                        .X1 = picRoom(lOldIn).Left
                        .X2 = picRoom(Index).Left + picRoom(Index).Width
                        .Y1 = picRoom(lOldIn).Top
                        .Y2 = picRoom(Index).Top + picRoom(Index).Height
                        .Visible = True
                    End With
                    With Rooms(lOldIn)
                        .sExits = Replace$(.sExits, ":nw;", "")
                        .sExits = .sExits & ":nw;"
                    End With
                    With Rooms(Index)
                        .sExits = Replace$(.sExits, ":se;", "")
                        .sExits = .sExits & ":se;"
                    End With
                    lblLabel.Caption = "You made an exit to the northwest."
                End If
            ElseIf lOldIn + (Rows - 1) = Index Then 'north east
                If InStr(1, Rooms(lOldIn).sExits, ":ne;") = 0 Then
                    Load lnJoin(lnJoin.UBound + 1)
                    With lnJoin(lnJoin.UBound)
                        .X1 = picRoom(lOldIn).Left + picRoom(lOldIn).Width
                        .X2 = picRoom(Index).Left
                        .Y1 = picRoom(lOldIn).Top
                        .Y2 = picRoom(Index).Top + picRoom(Index).Height
                        .Visible = True
                    End With
                    With Rooms(lOldIn)
                        .sExits = Replace$(.sExits, ":ne;", "")
                        .sExits = .sExits & ":ne;"
                    End With
                    With Rooms(Index)
                        .sExits = Replace$(.sExits, ":sw;", "")
                        .sExits = .sExits & ":sw;"
                    End With
                    lblLabel.Caption = "You made an exit to the northeast."
                End If
            ElseIf lOldIn - (Rows - 1) = Index Then 'South West
                If InStr(1, Rooms(lOldIn).sExits, ":sw;") = 0 Then
                    Load lnJoin(lnJoin.UBound + 1)
                    With lnJoin(lnJoin.UBound)
                        .X1 = picRoom(lOldIn).Left
                        .X2 = picRoom(Index).Left + picRoom(Index).Width
                        .Y1 = picRoom(lOldIn).Top + picRoom(lOldIn).Height
                        .Y2 = picRoom(Index).Top
                        .Visible = True
                    End With
                    With Rooms(lOldIn)
                        .sExits = Replace$(.sExits, ":sw;", "")
                        .sExits = .sExits & ":sw;"
                    End With
                    With Rooms(Index)
                        .sExits = Replace$(.sExits, ":ne;", "")
                        .sExits = .sExits & ":ne;"
                    End With
                    lblLabel.Caption = "You made an exit to the southwest."
                End If
            Else
                lblLabel.Caption = "Invalid selection. (Too far away)"
                Exit Sub
            End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Set picRoom(lOldIn).Picture = picPic(Rooms(lOldIn).lTag).Picture
    lOldIn = 0
    lNewIn = 0
    chkUse.Value = 0
End If
End Sub

Private Sub Form_Load()
DrawMap
With cmdMapFinished
    .Top = Me.ScaleHeight - .Height - 1
    .Left = Me.ScaleWidth - .Width - 1
End With
End Sub
Sub DrawMap()

Dim i As Long
Dim j As Long
If picRoom.UBound > 0 Then
    j = picRoom.UBound
    For i = 1 To j
        Unload picRoom(i)
    Next
End If
For i = picRoom(0).Left + 10 To frmQuickMap.ScaleWidth Step (picRoom(0).Width + 10)
    For j = picRoom(0).Top + 10 To frmQuickMap.ScaleHeight Step (picRoom(0).Height + 10)
        Load picRoom(picRoom.UBound + 1)
        With picRoom(picRoom.UBound)
            .Left = i
            .Top = j
            .Visible = True
        End With
    Next

Next
Rows = j \ (picRoom(0).Width + 10)
ReDim Rooms(1 To picRoom.UBound)
For i = 1 To UBound(Rooms)
    With Rooms(i)
        .lX = picRoom(i).Left
        .lY = picRoom(i).Top
    End With
Next
Me.Width = ScaleX(picRoom(picRoom.UBound).Left + picRoom(picRoom.UBound).Width + 20, 3, 1)
Me.Height = ScaleY(picRoom(picRoom.UBound).Top + picRoom(picRoom.UBound).Height + 40, 3, 1)
End Sub

Private Sub mnuJoin_Click()
picJoin.Visible = True
End Sub

Private Sub optDoorDef_Click()
Dim i As Long
If optDoorDef.Value = True Then
    For i = 1 To picRoom.UBound
        With picRoom(i)
            '.Appearance = 0
            '.BorderStyle = 0
            Set picRoom(i).Picture = picPic(Rooms(i).lTag)
        End With
    Next
    chkKeep.Visible = False
    chkMake.Visible = False
    lOldIn = 0
End If
End Sub

Private Sub optExitDef_Click()
Dim i As Long
If optExitDef.Value = True Then
    For i = 1 To picRoom.UBound
        With picRoom(i)
            '.Appearance = 1
            '.BorderStyle = 0
            Set picRoom(i).Picture = picPic(Rooms(i).lTag)
        End With
    Next
    chkKeep.Visible = True
    chkMake.Visible = True
    lOldIn = 0
End If
End Sub

Private Sub optRoomDef_Click()
Dim i As Long
If optRoomDef.Value = True Then
    For i = 1 To picRoom.UBound
        With picRoom(i)
            '.Appearance = 1
            '.BorderStyle = 1
            Set picRoom(i).Picture = picPic(Rooms(i).lTag)
        End With
    Next
    chkKeep.Visible = False
    chkMake.Visible = False
    lOldIn = 0
End If
End Sub

Private Sub DoDoor(Index As Integer)
If lOldIn = 0 Then
    Select Case Rooms(Index).lTag
        Case 0
            lblLabel.Caption = "Invalid selection (Not a room)."
            Exit Sub
    End Select
    lblLabel.Caption = "Choose another room next to the that has an exit."
    lOldIn = Index
    Set picRoom(Index).Picture = picPic(Rooms(Index).lTag + 9)
Else
    Select Case Rooms(Index).lTag
        Case 0
            lblLabel.Caption = "Invalid selection (Not a room)."
            Exit Sub
    End Select
    lNewIn = Index
    If chkUse.Value = 0 Then
        FillCombos
        picDoorDef.Visible = True
    Else
        cmdOK_Click
    End If
End If
End Sub

Private Sub picRoom_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    If optRoomDef.Value = True Then
        Select Case Rooms(Index).lTag
            Case 0
                Rooms(Index).lTag = 1
            Case 1
                Rooms(Index).lTag = 2
            Case 2
                Rooms(Index).lTag = 0
        End Select
        Set picRoom(Index).Picture = picPic(Rooms(Index).lTag)
    ElseIf optExitDef.Value = True Then
        If lOldIn = 0 Then
            Select Case Rooms(Index).lTag
                Case 0
                    If chkMake.Value = 1 Then
                        Rooms(Index).lTag = 1
                        Set picRoom(Index).Picture = picPic(Rooms(Index).lTag)
                    Else
                        lblLabel.Caption = "Invalid selection (Not a room)."
                        Exit Sub
                    End If
            End Select
            lblLabel.Caption = "Choose another room next to the current room."
            lOldIn = Index
            Set picRoom(Index).Picture = picPic(Rooms(Index).lTag + 3)
        Else
            Select Case Rooms(Index).lTag
                Case 0
                    If chkMake.Value = 1 Then
                        Rooms(Index).lTag = 1
                        Set picRoom(Index).Picture = picPic(Rooms(Index).lTag)
                    Else
                        lblLabel.Caption = "Invalid selection (Not a room)."
                        Exit Sub
                    End If
            End Select
            If lOldIn - Rows = Index Then 'west
                If InStr(1, Rooms(lOldIn).sExits, ":w;") = 0 Then
                    Load lnExit(lnExit.UBound + 1)
                    With lnExit(lnExit.UBound)
                        .X1 = picRoom(lOldIn).Left
                        .X2 = picRoom(Index).Left + picRoom(Index).Width
                        .Y1 = picRoom(lOldIn).Top + (picRoom(lOldIn).Height \ 2)
                        .Y2 = picRoom(Index).Top + (picRoom(Index).Height \ 2)
                        .Visible = True
                    End With
                    With Rooms(lOldIn)
                        .sExits = Replace$(.sExits, ":w;", "")
                        .sExits = .sExits & ":w;"
                    End With
                    With Rooms(Index)
                        .sExits = Replace$(.sExits, ":e;", "")
                        .sExits = .sExits & ":e;"
                    End With
                    lblLabel.Caption = "You made an exit to the west."
                End If
            ElseIf lOldIn + Rows = Index Then 'east
                If InStr(1, Rooms(lOldIn).sExits, ":e;") = 0 Then
                    Load lnExit(lnExit.UBound + 1)
                    With lnExit(lnExit.UBound)
                        .X1 = picRoom(lOldIn).Left + picRoom(lOldIn).Width
                        .X2 = picRoom(Index).Left
                        .Y1 = picRoom(lOldIn).Top + (picRoom(lOldIn).Height \ 2)
                        .Y2 = picRoom(Index).Top + (picRoom(Index).Height \ 2)
                        .Visible = True
                    End With
                    With Rooms(lOldIn)
                        .sExits = Replace$(.sExits, ":e;", "")
                        .sExits = .sExits & ":e;"
                    End With
                    With Rooms(Index)
                        .sExits = Replace$(.sExits, ":w;", "")
                        .sExits = .sExits & ":w;"
                    End With
                    lblLabel.Caption = "You made an exit to the east."
                End If
            ElseIf lOldIn - 1 = Index Then 'north
                If InStr(1, Rooms(lOldIn).sExits, ":n;") = 0 Then
                    Load lnExit(lnExit.UBound + 1)
                    With lnExit(lnExit.UBound)
                        .X1 = picRoom(lOldIn).Left + (picRoom(lOldIn).Width \ 2)
                        .X2 = .X1
                        .Y1 = picRoom(lOldIn).Top
                        .Y2 = picRoom(Index).Top + picRoom(Index).Height
                        .Visible = True
                    End With
                    With Rooms(lOldIn)
                        .sExits = Replace$(.sExits, ":n;", "")
                        .sExits = .sExits & ":n;"
                    End With
                    With Rooms(Index)
                        .sExits = Replace$(.sExits, ":s;", "")
                        .sExits = .sExits & ":s;"
                    End With
                    lblLabel.Caption = "You made an exit to the north."
                End If
            ElseIf lOldIn + 1 = Index Then 'south
                If InStr(1, Rooms(lOldIn).sExits, ":s;") = 0 Then
                    Load lnExit(lnExit.UBound + 1)
                    With lnExit(lnExit.UBound)
                        .X1 = picRoom(lOldIn).Left + (picRoom(lOldIn).Width \ 2)
                        .X2 = .X1
                        .Y1 = picRoom(lOldIn).Top + picRoom(lOldIn).Height
                        .Y2 = picRoom(Index).Top
                        .Visible = True
                    End With
                    With Rooms(lOldIn)
                        .sExits = Replace$(.sExits, ":s;", "")
                        .sExits = .sExits & ":s;"
                    End With
                    With Rooms(Index)
                        .sExits = Replace$(.sExits, ":n;", "")
                        .sExits = .sExits & ":n;"
                    End With
                    lblLabel.Caption = "You made an exit to the south."
                End If
            ElseIf lOldIn + (Rows + 1) = Index Then 'south east
                If InStr(1, Rooms(lOldIn).sExits, ":se;") = 0 Then
                    Load lnExit(lnExit.UBound + 1)
                    With lnExit(lnExit.UBound)
                        .X1 = picRoom(lOldIn).Left + picRoom(lOldIn).Width
                        .X2 = picRoom(Index).Left
                        .Y1 = picRoom(lOldIn).Top + picRoom(lOldIn).Height
                        .Y2 = picRoom(Index).Top
                        .Visible = True
                    End With
                    With Rooms(lOldIn)
                        .sExits = Replace$(.sExits, ":se;", "")
                        .sExits = .sExits & ":se;"
                    End With
                    With Rooms(Index)
                        .sExits = Replace$(.sExits, ":nw;", "")
                        .sExits = .sExits & ":nw;"
                    End With
                    lblLabel.Caption = "You made an exit to the southeast."
                End If
            ElseIf lOldIn - (Rows + 1) = Index Then 'north west
                If InStr(1, Rooms(lOldIn).sExits, ":nw;") = 0 Then
                    Load lnExit(lnExit.UBound + 1)
                    With lnExit(lnExit.UBound)
                        .X1 = picRoom(lOldIn).Left
                        .X2 = picRoom(Index).Left + picRoom(Index).Width
                        .Y1 = picRoom(lOldIn).Top
                        .Y2 = picRoom(Index).Top + picRoom(Index).Height
                        .Visible = True
                    End With
                    With Rooms(lOldIn)
                        .sExits = Replace$(.sExits, ":nw;", "")
                        .sExits = .sExits & ":nw;"
                    End With
                    With Rooms(Index)
                        .sExits = Replace$(.sExits, ":se;", "")
                        .sExits = .sExits & ":se;"
                    End With
                    lblLabel.Caption = "You made an exit to the northwest."
                End If
            ElseIf lOldIn + (Rows - 1) = Index Then 'north east
                If InStr(1, Rooms(lOldIn).sExits, ":ne;") = 0 Then
                    Load lnExit(lnExit.UBound + 1)
                    With lnExit(lnExit.UBound)
                        .X1 = picRoom(lOldIn).Left + picRoom(lOldIn).Width
                        .X2 = picRoom(Index).Left
                        .Y1 = picRoom(lOldIn).Top
                        .Y2 = picRoom(Index).Top + picRoom(Index).Height
                        .Visible = True
                    End With
                    With Rooms(lOldIn)
                        .sExits = Replace$(.sExits, ":ne;", "")
                        .sExits = .sExits & ":ne;"
                    End With
                    With Rooms(Index)
                        .sExits = Replace$(.sExits, ":sw;", "")
                        .sExits = .sExits & ":sw;"
                    End With
                    lblLabel.Caption = "You made an exit to the northeast."
                End If
            ElseIf lOldIn - (Rows - 1) = Index Then 'South West
                If InStr(1, Rooms(lOldIn).sExits, ":sw;") = 0 Then
                    Load lnExit(lnExit.UBound + 1)
                    With lnExit(lnExit.UBound)
                        .X1 = picRoom(lOldIn).Left
                        .X2 = picRoom(Index).Left + picRoom(Index).Width
                        .Y1 = picRoom(lOldIn).Top + picRoom(lOldIn).Height
                        .Y2 = picRoom(Index).Top
                        .Visible = True
                    End With
                    With Rooms(lOldIn)
                        .sExits = Replace$(.sExits, ":sw;", "")
                        .sExits = .sExits & ":sw;"
                    End With
                    With Rooms(Index)
                        .sExits = Replace$(.sExits, ":ne;", "")
                        .sExits = .sExits & ":ne;"
                    End With
                    lblLabel.Caption = "You made an exit to the southwest."
                End If
            Else
                lblLabel.Caption = "Invalid selection. (Too far away)"
                Exit Sub
            End If
            Set picRoom(lOldIn).Picture = picPic(Rooms(lOldIn).lTag)
            If chkKeep.Value = 1 Then
                lOldIn = Index
                Set picRoom(Index).Picture = picPic(Rooms(Index).lTag + 3)
            Else
                lOldIn = 0
            End If
        End If
    Else
        DoDoor Index
    End If
ElseIf Button = vbRightButton Then
    Dim Direc As Long
    If lOldIn = 0 Then
        Select Case Rooms(Index).lTag
            Case 0
                lblLabel.Caption = "Invalid selection (Not a room)."
                Exit Sub
        End Select
        lblLabel.Caption = "Choose another room next to the current room."
        lOldIn = Index
        Set picRoom(Index).Picture = picPic(Rooms(Index).lTag + 6)
    ElseIf optExitDef.Value = True Then
        Select Case Rooms(Index).lTag
            Case 0
                lblLabel.Caption = "Invalid selection (Not a room)."
                Exit Sub
        End Select
        If lOldIn - Rows = Index Then 'west
            With lnHold
                .lX1 = picRoom(lOldIn).Left
                .lX2 = picRoom(Index).Left + picRoom(Index).Width
                .lY1 = picRoom(lOldIn).Top + (picRoom(lOldIn).Height \ 2)
                .lY2 = picRoom(Index).Top + (picRoom(Index).Height \ 2)
            End With
            With Rooms(lOldIn)
                .sExits = Replace$(.sExits, ":w;", "")
            End With
            With Rooms(Index)
                .sExits = Replace$(.sExits, ":e;", "")
            End With
            lblLabel.Caption = "You deleted an exit to the west."
        ElseIf lOldIn + Rows = Index Then 'east
            With lnHold
                .lX1 = picRoom(lOldIn).Left + picRoom(lOldIn).Width
                .lX2 = picRoom(Index).Left
                .lY1 = picRoom(lOldIn).Top + (picRoom(lOldIn).Height \ 2)
                .lY2 = picRoom(Index).Top + (picRoom(Index).Height \ 2)
            End With
            With Rooms(lOldIn)
                .sExits = Replace$(.sExits, ":e;", "")
            End With
            With Rooms(Index)
                .sExits = Replace$(.sExits, ":w;", "")
            End With
            lblLabel.Caption = "You deleted an exit to the east."
        ElseIf lOldIn - 1 = Index Then 'north
            With lnHold
                .lX1 = picRoom(lOldIn).Left + (picRoom(lOldIn).Width \ 2)
                .lX2 = .lX1
                .lY1 = picRoom(lOldIn).Top
                .lY2 = picRoom(Index).Top + picRoom(Index).Height
            End With
            With Rooms(lOldIn)
                .sExits = Replace$(.sExits, ":n;", "")
            End With
            With Rooms(Index)
                .sExits = Replace$(.sExits, ":s;", "")
            End With
            lblLabel.Caption = "You deleted an exit to the north."
        ElseIf lOldIn + 1 = Index Then 'south
            With lnHold
                .lX1 = picRoom(lOldIn).Left + (picRoom(lOldIn).Width \ 2)
                .lX2 = .lX1
                .lY1 = picRoom(lOldIn).Top + picRoom(lOldIn).Height
                .lY2 = picRoom(Index).Top
            End With
            With Rooms(lOldIn)
                .sExits = Replace$(.sExits, ":s;", "")
            End With
            With Rooms(Index)
                .sExits = Replace$(.sExits, ":n;", "")
            End With
            lblLabel.Caption = "You deleted an exit to the south."
        ElseIf lOldIn + (Rows + 1) = Index Then 'south east
            Direc = 1
            With lnHold
                .lX1 = picRoom(lOldIn).Left + picRoom(lOldIn).Width
                .lX2 = picRoom(Index).Left
                .lY1 = picRoom(lOldIn).Top + picRoom(lOldIn).Height
                .lY2 = picRoom(Index).Top
            End With
            With Rooms(lOldIn)
                .sExits = Replace$(.sExits, ":se;", "")
            End With
            With Rooms(Index)
                .sExits = Replace$(.sExits, ":nw;", "")
            End With
            lblLabel.Caption = "You deleted an exit to the southeast."
        ElseIf lOldIn - (Rows + 1) = Index Then 'north west
            Direc = 1
            With lnHold
                .lX1 = picRoom(lOldIn).Left
                .lX2 = picRoom(Index).Left + picRoom(Index).Width
                .lY1 = picRoom(lOldIn).Top
                .lY2 = picRoom(Index).Top + picRoom(Index).Height
            End With
            With Rooms(lOldIn)
                .sExits = Replace$(.sExits, ":nw;", "")
            End With
            With Rooms(Index)
                .sExits = Replace$(.sExits, ":se;", "")
            End With
            lblLabel.Caption = "You deleted an exit to the northwest."
        ElseIf lOldIn + (Rows - 1) = Index Then 'north east
            Direc = 1
            With lnHold
                .lX1 = picRoom(lOldIn).Left + picRoom(lOldIn).Width
                .lX2 = picRoom(Index).Left
                .lY1 = picRoom(lOldIn).Top
                .lY2 = picRoom(Index).Top + picRoom(Index).Height
            End With
            With Rooms(lOldIn)
                .sExits = Replace$(.sExits, ":ne;", "")
            End With
            With Rooms(Index)
                .sExits = Replace$(.sExits, ":sw;", "")
            End With
            lblLabel.Caption = "You deleted an exit to the northeast."
        ElseIf lOldIn - (Rows - 1) = Index Then 'South West
            Direc = 1
            With lnHold
                .lX1 = picRoom(lOldIn).Left
                .lX2 = picRoom(Index).Left + picRoom(Index).Width
                .lY1 = picRoom(lOldIn).Top + picRoom(lOldIn).Height
                .lY2 = picRoom(Index).Top
            End With
            With Rooms(lOldIn)
                .sExits = Replace$(.sExits, ":sw;", "")
            End With
            With Rooms(Index)
                .sExits = Replace$(.sExits, ":ne;", "")
            End With
            lblLabel.Caption = "You deleted an exit to the southwest."
        Else
            lblLabel.Caption = "Invalid selection. (Too far away)"
            Exit Sub
        End If
        Set picRoom(Index).Picture = picPic(Rooms(Index).lTag)
        If chkKeep.Value = 1 Then
            lOldIn = Index
            Set picRoom(Index).Picture = picPic(Rooms(Index).lTag + 6)
        Else
            lOldIn = 0
        End If
        For i = lnExit.LBound To lnExit.UBound
            '340
            On Error Resume Next
            With lnExit(i)
                If .X1 = lnHold.lX1 Then
                    If Err.Number = 340 Then GoTo nNext
                    If .X2 = lnHold.lX2 Then
                        If .Y1 = lnHold.lY1 Then
                            If .Y2 = lnHold.lY2 Then
                                Unload lnExit(i)
                                Exit For
                            End If
                        ElseIf .Y2 = lnHold.lY1 Then
                            If .Y1 = lnHold.lY2 Then
                                Unload lnExit(i)
                                Exit For
                            End If
                        End If
                    End If
                ElseIf .X2 = lnHold.lX1 And Direc <> 1 Then
                    If Err.Number = 340 Then GoTo nNext
                    If .X2 = lnHold.lX1 Then
                        If .Y1 = lnHold.lY1 Then
                            If .Y2 = lnHold.lY2 Then
                                Unload lnExit(i)
                                Exit For
                            End If
                        ElseIf .Y2 = lnHold.lY1 Then
                            If .Y2 = lnHold.lY1 Then
                                Unload lnExit(i)
                                Exit For
                            End If
                        End If
                    End If
                End If
            End With
nNext:
Err.Number = 0
        Next
    End If
End If
End Sub

Private Sub FillCombos()
Dim i As Long
cboKeys.Clear
cboKeys.AddItem "(0) None"
For i = LBound(dbItems) To UBound(dbItems)
    With dbItems(i)
        If .sWorn = "key" Then cboKeys.AddItem "(" & .iID & ") " & .sItemName
    End With
Next
End Sub

Private Sub txtKey_Change()
modMain.SetCBOSelectByID cboKeys, txtKey.Text
End Sub

Private Sub txtRoom_Change()
modMain.SetCBOSelectByID cboRooms, txtRoom.Text
End Sub
