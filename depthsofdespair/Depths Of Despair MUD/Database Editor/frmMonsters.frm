VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMonsters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Monsters"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11535
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
   ScaleHeight     =   6495
   ScaleWidth      =   11535
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin ServerEditor.UltraBox lstMonsters 
      Height          =   5055
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   8916
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
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "(save)"
      Height          =   255
      Left            =   6480
      TabIndex        =   42
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "(new)"
      Height          =   255
      Left            =   7680
      TabIndex        =   45
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "< Previous"
      Height          =   255
      Left            =   9000
      TabIndex        =   44
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   255
      Left            =   10200
      TabIndex        =   43
      Top             =   6000
      Width           =   1095
   End
   Begin TabDlg.SSTab sTab 
      Height          =   5415
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Basic Info"
      TabPicture(0)   =   "frmMonsters.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraAttackable"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraHostile"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Picture4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Picture2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Picture1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Combat"
      TabPicture(1)   =   "frmMonsters.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture5"
      Tab(1).Control(1)=   "Picture6"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Monster Death"
      TabPicture(2)   =   "frmMonsters.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblLabel(25)"
      Tab(2).Control(1)=   "lblLabel(14)"
      Tab(2).Control(2)=   "lblLabel(10)"
      Tab(2).Control(3)=   "Label6"
      Tab(2).Control(4)=   "Line2"
      Tab(2).Control(5)=   "Label7"
      Tab(2).Control(6)=   "lblLabel(13)"
      Tab(2).Control(7)=   "txtDeathText"
      Tab(2).Control(8)=   "cboDropItem"
      Tab(2).Control(9)=   "txtGold"
      Tab(2).Control(10)=   "cboPercent"
      Tab(2).Control(11)=   "cboCorpse"
      Tab(2).Control(12)=   "lstItems"
      Tab(2).Control(13)=   "cmdAddItem"
      Tab(2).Control(14)=   "Command1"
      Tab(2).ControlCount=   15
      TabCaption(3)   =   "On Death Scripting"
      TabPicture(3)   =   "frmMonsters.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtScripting"
      Tab(3).Control(1)=   "lblMethod"
      Tab(3).Control(2)=   "Label2"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Other"
      TabPicture(4)   =   "frmMonsters.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Picture7"
      Tab(4).ControlCount=   1
      Begin VB.CommandButton Command1 
         Caption         =   "&Delete"
         Height          =   255
         Left            =   -69960
         TabIndex        =   103
         Top             =   3840
         Width           =   735
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "&Add"
         Height          =   255
         Left            =   -69960
         TabIndex        =   102
         Top             =   3480
         Width           =   735
      End
      Begin ServerEditor.UltraBox lstItems 
         Height          =   900
         Left            =   -73800
         TabIndex        =   101
         Top             =   3360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1588
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
      End
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   -74880
         ScaleHeight     =   3375
         ScaleWidth      =   8055
         TabIndex        =   90
         Top             =   480
         Width           =   8055
         Begin VB.ComboBox cboFamiliar 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1800
            Width           =   5535
         End
         Begin VB.ComboBox cboAtNight 
            Height          =   315
            Left            =   3240
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   600
            Width           =   4095
         End
         Begin VB.ComboBox cboAtDay 
            Height          =   315
            Left            =   3240
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   120
            Width           =   4095
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Change to Familiar:"
            Height          =   195
            Left            =   240
            TabIndex        =   94
            Top             =   1800
            Width           =   1395
         End
         Begin VB.Label Label10 
            Caption         =   $"frmMonsters.frx":008C
            Height          =   495
            Left            =   120
            TabIndex        =   93
            Top             =   1200
            Width           =   7455
         End
         Begin VB.Line Line4 
            X1              =   240
            X2              =   7560
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "During the night, the monster appears as:"
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   600
            Width           =   3030
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "During the day, the monster appears as:"
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   120
            Width           =   2940
         End
      End
      Begin ServerEditor.IntelliSense txtScripting 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   38
         Top             =   840
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   7011
      End
      Begin VB.ComboBox cboCorpse 
         Height          =   315
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   4320
         Width           =   2655
      End
      Begin VB.ComboBox cboPercent 
         Height          =   315
         Left            =   -70920
         TabIndex        =   36
         Text            =   "100"
         Top             =   3000
         Width           =   855
      End
      Begin ServerEditor.NumOnlyText txtGold 
         Height          =   375
         Left            =   -73800
         TabIndex        =   34
         Top             =   2520
         Width           =   2655
         _ExtentX        =   4048
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         AllowNeg        =   0   'False
         Align           =   0
         MaxLength       =   7
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   -74760
         ScaleHeight     =   3255
         ScaleWidth      =   7935
         TabIndex        =   75
         Top             =   2040
         Width           =   7935
         Begin VB.ComboBox cboSpellType 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMonsters.frx":0122
            Left            =   3960
            List            =   "frmMonsters.frx":012C
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1920
            Width           =   1215
         End
         Begin ServerEditor.NumOnlyText txtSpellID 
            Height          =   375
            Left            =   3960
            TabIndex        =   25
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   0
            Enabled         =   0   'False
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.UltraBox lstSpells 
            Height          =   2655
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   4683
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
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "&Remove"
            Height          =   375
            Left            =   4080
            TabIndex        =   32
            Top             =   2760
            Width           =   855
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   3000
            TabIndex        =   30
            Top             =   2760
            Width           =   855
         End
         Begin ServerEditor.UltraBox lstCSpells 
            Height          =   2655
            Left            =   5400
            TabIndex        =   31
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   4683
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
         End
         Begin ServerEditor.NumOnlyText txtSpellEnergy 
            Height          =   375
            Left            =   3960
            TabIndex        =   26
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   5
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtCPE 
            Height          =   375
            Left            =   3960
            TabIndex        =   27
            Top             =   1440
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   2
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtCPR 
            Height          =   375
            Left            =   3960
            TabIndex        =   29
            Top             =   2280
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   2
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cast Per Round:"
            Height          =   195
            Index           =   7
            Left            =   2700
            TabIndex        =   83
            Top             =   2280
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Spell Type:"
            Height          =   195
            Index           =   6
            Left            =   3075
            TabIndex        =   82
            Top             =   1920
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cast Per Engage:"
            Height          =   195
            Index           =   5
            Left            =   2640
            TabIndex        =   81
            Top             =   1440
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Energy Useage:"
            Height          =   195
            Index           =   4
            Left            =   2760
            TabIndex        =   80
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Spell ID:"
            Height          =   195
            Index           =   3
            Left            =   3300
            TabIndex        =   79
            Top             =   480
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Spell List:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label5 
            Caption         =   "Current Spells:"
            Height          =   255
            Left            =   5400
            TabIndex        =   77
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Spells:"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   76
            Top             =   0
            Width           =   465
         End
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   -74880
         ScaleHeight     =   1455
         ScaleWidth      =   7935
         TabIndex        =   72
         Top             =   480
         Width           =   7935
         Begin VB.ComboBox cboWeapon 
            Height          =   315
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   104
            Top             =   960
            Width           =   3735
         End
         Begin VB.ComboBox cboNoAttackItem 
            Height          =   315
            Left            =   3720
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   120
            Width           =   4095
         End
         Begin ServerEditor.NumOnlyText txtPEnergy 
            Height          =   375
            Left            =   2400
            TabIndex        =   23
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   5
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Monster will use this weapon:"
            Height          =   195
            Index           =   1
            Left            =   4200
            TabIndex        =   105
            Top             =   720
            Width           =   2115
         End
         Begin VB.Label Label4 
            Caption         =   "If the player has this item in their INVENTORY, then do not engange in combat:"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   89
            Top             =   120
            Width           =   3375
         End
         Begin VB.Line Line5 
            X1              =   0
            X2              =   7920
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   7920
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Physical attack energy usage:"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   74
            Top             =   960
            Width           =   2160
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Energy Usage:"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   73
            Top             =   720
            Width           =   1065
         End
      End
      Begin VB.ComboBox cboDropItem 
         Height          =   315
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox txtDeathText 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   720
         Width           =   8055
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   240
         ScaleHeight     =   3495
         ScaleWidth      =   4215
         TabIndex        =   61
         Top             =   480
         Width           =   4215
         Begin ServerEditor.ucCombo ucEvil 
            Height          =   375
            Left            =   1200
            TabIndex        =   107
            Top             =   1680
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin VB.TextBox txtDescription 
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   960
            Width           =   2895
         End
         Begin VB.TextBox txtMonsterName 
            Height          =   405
            Left            =   1200
            TabIndex        =   4
            Top             =   480
            Width           =   2895
         End
         Begin VB.TextBox txtID 
            Enabled         =   0   'False
            Height          =   405
            Left            =   1200
            TabIndex        =   3
            Top             =   0
            Width           =   615
         End
         Begin ServerEditor.NumOnlyText txtHP 
            Height          =   375
            Left            =   1200
            TabIndex        =   6
            Top             =   2040
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   5
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtAC 
            Height          =   375
            Left            =   1200
            TabIndex        =   8
            Top             =   2520
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   5
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtEXP 
            Height          =   375
            Left            =   1200
            TabIndex        =   10
            Top             =   3000
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   9
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtEnergy 
            Height          =   375
            Left            =   3120
            TabIndex        =   7
            Top             =   2040
            Width           =   975
            _ExtentX        =   1296
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   5
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtLevel 
            Height          =   375
            Left            =   3120
            TabIndex        =   9
            Top             =   2520
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   3
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Alignment:"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   106
            Top             =   1680
            Width           =   765
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Level:"
            Height          =   195
            Index           =   15
            Left            =   2520
            TabIndex        =   99
            Top             =   2520
            Width           =   435
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "EXP:"
            Height          =   195
            Index           =   7
            Left            =   720
            TabIndex        =   68
            Top             =   3000
            Width           =   330
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "AC:"
            Height          =   195
            Index           =   6
            Left            =   840
            TabIndex        =   67
            Top             =   2520
            Width           =   270
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Description:"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   66
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "HP:"
            Height          =   195
            Index           =   2
            Left            =   840
            TabIndex        =   65
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Monster Name:"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   64
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "ID:"
            Height          =   195
            Index           =   0
            Left            =   840
            TabIndex        =   63
            Top             =   0
            Width           =   225
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Energy:"
            Height          =   195
            Index           =   19
            Left            =   2400
            TabIndex        =   62
            Top             =   2040
            Width           =   570
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   4680
         ScaleHeight     =   4935
         ScaleWidth      =   3495
         TabIndex        =   55
         Top             =   360
         Width           =   3495
         Begin VB.CheckBox chkRoams 
            Caption         =   "Roams"
            Height          =   255
            Left            =   1800
            TabIndex        =   14
            Top             =   240
            Width           =   795
         End
         Begin VB.ComboBox cboType 
            Height          =   315
            ItemData        =   "frmMonsters.frx":0141
            Left            =   1080
            List            =   "frmMonsters.frx":014E
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   4440
            Width           =   2175
         End
         Begin VB.CheckBox chkAttackable 
            Caption         =   "Not Attackable"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkHostile 
            Caption         =   "Hostile"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtMsg1 
            Height          =   405
            Left            =   240
            MaxLength       =   30
            TabIndex        =   16
            Top             =   1320
            Width           =   3015
         End
         Begin VB.TextBox txtMsg2 
            Height          =   405
            Left            =   240
            MaxLength       =   30
            TabIndex        =   17
            Top             =   1800
            Width           =   3015
         End
         Begin VB.TextBox txtMsg3 
            Height          =   405
            Left            =   240
            MaxLength       =   30
            TabIndex        =   18
            Top             =   2280
            Width           =   3015
         End
         Begin ServerEditor.NumOnlyText txtRegenTime 
            Height          =   375
            Left            =   1080
            TabIndex        =   20
            Top             =   3960
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   6
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.Raise Raise2 
            Height          =   1575
            Left            =   120
            TabIndex        =   56
            Top             =   1200
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   2778
            Style           =   2
            Color           =   0
         End
         Begin ServerEditor.NumOnlyText txtMob 
            Height          =   375
            Left            =   1080
            TabIndex        =   19
            Top             =   3480
            Width           =   975
            _ExtentX        =   1296
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   5
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Monster Type:"
            Height          =   195
            Index           =   17
            Left            =   0
            TabIndex        =   86
            Top             =   4440
            Width           =   1050
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Messages:"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   60
            Top             =   960
            Width           =   765
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Regen Time:"
            Height          =   195
            Index           =   18
            Left            =   0
            TabIndex        =   59
            Top             =   3960
            Width           =   900
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Mob Group:"
            Height          =   195
            Index           =   8
            Left            =   0
            TabIndex        =   58
            Top             =   3480
            Width           =   840
         End
         Begin VB.Label lblEX 
            Caption         =   "EX"
            Height          =   615
            Left            =   120
            TabIndex        =   57
            Top             =   2880
            Width           =   3255
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   240
         ScaleHeight     =   975
         ScaleWidth      =   4215
         TabIndex        =   48
         Top             =   4080
         Width           =   4215
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   1080
            ScaleHeight     =   615
            ScaleWidth      =   1935
            TabIndex        =   49
            Top             =   120
            Width           =   1935
            Begin ServerEditor.NumOnlyText txtMinDamage 
               Height          =   375
               Left            =   0
               TabIndex        =   11
               Top             =   240
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   ""
               AllowNeg        =   0   'False
               Align           =   0
               MaxLength       =   4
               Enabled         =   -1  'True
               Backcolor       =   -2147483643
            End
            Begin ServerEditor.NumOnlyText txtMaxDamage 
               Height          =   375
               Left            =   1080
               TabIndex        =   12
               Top             =   240
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   ""
               AllowNeg        =   0   'False
               Align           =   0
               MaxLength       =   4
               Enabled         =   -1  'True
               Backcolor       =   -2147483643
            End
            Begin VB.Label lblLabel 
               AutoSize        =   -1  'True
               Caption         =   "Min"
               Height          =   195
               Index           =   3
               Left            =   0
               TabIndex        =   52
               Top             =   0
               Width           =   240
            End
            Begin VB.Label lblLabel 
               AutoSize        =   -1  'True
               Caption         =   "Max"
               Height          =   195
               Index           =   4
               Left            =   1080
               TabIndex        =   51
               Top             =   0
               Width           =   300
            End
            Begin VB.Label lblLabel 
               AutoSize        =   -1  'True
               Caption         =   "-"
               Height          =   195
               Index           =   42
               Left            =   840
               TabIndex        =   50
               Top             =   360
               Width           =   60
            End
         End
         Begin ServerEditor.Raise Raise1 
            Height          =   855
            Left            =   960
            TabIndex        =   53
            Top             =   0
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1508
            Style           =   2
            Color           =   0
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Damage:"
            Height          =   195
            Index           =   41
            Left            =   0
            TabIndex        =   54
            Top             =   0
            Width           =   645
         End
      End
      Begin VB.Frame fraHostile 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   47
         Top             =   2280
         Width           =   495
      End
      Begin VB.Frame fraAttackable 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   46
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblMethod 
         Height          =   375
         Left            =   -74880
         TabIndex        =   100
         Top             =   4920
         Width           =   8055
      End
      Begin VB.Label Label2 
         Caption         =   "When the monster dies, this script will be ran."
         Height          =   375
         Left            =   -74880
         TabIndex        =   88
         Top             =   360
         Width           =   7935
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Drop Corpse:"
         Height          =   195
         Index           =   13
         Left            =   -74760
         TabIndex        =   87
         Top             =   4320
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "(                     )%"
         Height          =   195
         Left            =   -71040
         TabIndex        =   85
         Top             =   3000
         Width           =   1230
      End
      Begin VB.Line Line2 
         X1              =   -74880
         X2              =   -66840
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "On death;"
         Height          =   195
         Left            =   -74880
         TabIndex        =   84
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Drop Gold:"
         Height          =   195
         Index           =   10
         Left            =   -74760
         TabIndex        =   71
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Drop Item:"
         Height          =   195
         Index           =   14
         Left            =   -74760
         TabIndex        =   70
         Top             =   3000
         Width           =   780
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Death Text:"
         Height          =   195
         Index           =   25
         Left            =   -74880
         TabIndex        =   69
         Top             =   480
         Width           =   870
      End
   End
   Begin ServerEditor.Raise Raise3 
      Height          =   495
      Left            =   6360
      TabIndex        =   95
      Top             =   5880
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise4 
      Height          =   5655
      Left            =   2880
      TabIndex        =   96
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9975
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise5 
      Height          =   5655
      Left            =   120
      TabIndex        =   97
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   9975
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise6 
      Height          =   6495
      Left            =   0
      TabIndex        =   98
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11456
      Style           =   4
      Color           =   0
   End
End
Attribute VB_Name = "frmMonsters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem*************************************************************************************
Rem*************************************************************************************
Rem***************       Code create by Chris Van Hooser          **********************
Rem***************                  (c)2002                       **********************
Rem*************** You may use this code and freely distribute it **********************
Rem***************   If you have any questions, please email me   **********************
Rem***************          at theendorbunker@attbi.com.          **********************
Rem***************       Thanks for downloading my project        **********************
Rem***************        and i hope you can use it well.         **********************
Rem***************                frmMonsters                     **********************
Rem***************                ServerEditor                    **********************
Rem***************                Editor.vbp                      **********************
Rem*************************************************************************************
Rem*************************************************************************************

Dim bIs As Boolean
Dim lcID As Long
Dim aSpells(4) As String

Private Sub chkHostile_Click()
If chkHostile.Value = 1 And ucEvil.Number < 99 Then ucEvil.Number = 99
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF4
        cmdSave_Click
    Case vbKeyF3
        cmdNext_Click
    Case vbKeyF2
        cmdPrevious_Click
End Select
End Sub
Private Sub cboDropItem_Change()
cmdAddItem.Caption = "&Add"
End Sub

Private Sub cboDropItem_Click()
cmdAddItem.Caption = "&Add"
End Sub

Private Sub cboType_Change()
cboType_Click
End Sub

Private Sub cboWeapon_Change()
If cboWeapon.list(cboWeapon.ListIndex) <> "(0) None" Then
    txtMinDamage.Text = "0"
    txtMaxDamage.Text = "0"
    txtMinDamage.Enabled = False
    txtMaxDamage.Enabled = False
    txtMsg1.Text = "attacks"
    txtMsg2.Text = "attacks"
    txtMsg3.Text = "attacks"
    txtMsg1.Enabled = False
    txtMsg2.Enabled = False
    txtMsg3.Enabled = False
Else
    txtMinDamage.Enabled = True
    txtMaxDamage.Enabled = True
    txtMsg1.Enabled = True
    txtMsg2.Enabled = True
    txtMsg3.Enabled = True
End If
End Sub

Private Sub cboWeapon_Click()
cboWeapon_Change
End Sub

Private Sub cmdAdd_Click()
If cmdAdd.Caption = "&Add" Then
    If lstCSpells.ListCount < 5 Then
        aSpells(lstCSpells.ListCount) = ":" & txtSpellID.Text & "/" & txtSpellEnergy.Text & "/E{" & txtCPE.Text & "}F{" & cboSpellType.ItemData(cboSpellType.ListIndex) & "}/" & txtCPR.Text
        lstCSpells.AddItem lstSpells.ItemText
    End If
Else
    aSpells(lstCSpells.ListIndex - 1) = ":" & txtSpellID.Text & "/" & txtSpellEnergy.Text & "/E{" & txtCPE.Text & "}F{" & cboSpellType.ItemData(cboSpellType.ListIndex) & "}/" & txtCPR.Text
    cmdAdd.Caption = "&Add"
End If
End Sub

Private Sub cmdAddItem_Click()
If cmdAddItem.Caption = "&Add" Then
    If cboDropItem.ListIndex <> -1 Then
        lstItems.AddItem cboDropItem.list(cboDropItem.ListIndex) & " @ " & CStr(Val(cboPercent.Text)) & "%"
    End If
Else
    If lstItems.ListIndex > 0 Then
        lstItems.SetItemText lstItems.ListIndex, cboDropItem.list(cboDropItem.ListIndex) & " @ " & CStr(Val(cboPercent.Text)) & "%"
    End If
    cmdAddItem.Caption = "&Add"
End If
End Sub

Private Sub cmdRemove_Click()
Dim i As Long
If lstCSpells.ListIndex < 5 Then
    For i = lstCSpells.ListIndex - 1 To 3
        aSpells(i) = aSpells(i + 1)
    Next
    aSpells(4) = ""
Else
    aSpells(4) = ""
End If
lstCSpells.RemoveItem lstCSpells.ListIndex
End Sub



Private Sub Command1_Click()
If lstItems.ListIndex > 0 Then lstItems.RemoveItem lstItems.ListIndex
cmdAddItem.Caption = "&Add"
End Sub

Private Sub lstCSpells_Click()
If lstCSpells.ListCount > 0 Then
    txtSpellID.Text = Left$(lstCSpells.list(lstCSpells.ListIndex), InStr(1, lstCSpells.list(lstCSpells.ListIndex), ")"))
    txtSpellEnergy.Text = CStr(modSpell.GetItemDurFromUnFormattedString(aSpells(lstCSpells.ListIndex - 1)))
    txtCPE.Text = CStr(modSpell.GetItemEnchantsFromUnFormattedString(aSpells(lstCSpells.ListIndex - 1)))
    txtCPR.Text = CStr(modSpell.GetItemUsesFromUnFormattedString(aSpells(lstCSpells.ListIndex - 1)))
    With dbSpells(GetSpellID(, Val(txtSpellID.Text)))
        If .iUse = 0 Then cboSpellType.ListIndex = 0 Else cboSpellType.ListIndex = 1
    End With
    cmdAdd.Caption = "&Modify"
End If
End Sub

Private Sub lstItems_Click()
Dim s As String
If lstItems.ListIndex > 0 Then
    s = lstItems.list(lstItems.ListIndex)
    s = Mid$(s, 2, InStr(1, s, ")") - 2)
    modMain.SetCBOSelectByID cboDropItem, s
    s = lstItems.list(lstItems.ListIndex)
    s = Mid$(s, InStr(1, s, "@") + 2)
    s = Left$(s, InStr(1, s, "%") - 1)
    modMain.SetListIndex cboPercent, s
    cmdAddItem.Caption = "&Modify"
End If
End Sub

Private Sub lstSpells_Click()
txtSpellID.Text = Left$(lstSpells.list(lstSpells.ListIndex), InStr(1, lstSpells.list(lstSpells.ListIndex), ")"))
txtSpellEnergy.Text = txtEnergy.Text
txtCPE.Text = "1"
txtCPR.Text = "1"
With dbSpells(GetSpellID(, Val(txtSpellID.Text)))
    If .iUse = 0 Then cboSpellType.ListIndex = 0 Else cboSpellType.ListIndex = 1
End With
cmdAdd.Caption = "&Add"
End Sub

Private Sub sTab_Click(PreviousTab As Integer)
cmdAddItem.Caption = "&Add"
cmdAdd.Caption = "&Add"
End Sub

Private Sub txtFind_Change()
lstMonsters.SetSelected lstMonsters.FindInStr(txtFind.Text), True, True
End Sub


Private Sub cboType_Click()
txtRegenTime.Enabled = False
chkHostile.Enabled = False
Select Case cboType.list(cboType.ListIndex)
    Case "(0) - Normal"
        chkHostile.Enabled = True
    Case "(2) - Boss"
        txtRegenTime.Enabled = True
        chkHostile.Enabled = True
End Select
End Sub

Private Sub cmdNew_Click()
Dim x As Long
Dim i As Long
Dim t As Boolean
MousePointer = vbHourglass
ReDim Preserve dbMonsters(1 To UBound(dbMonsters) + 1)
x = dbMonsters(UBound(dbMonsters) - 1).lID
x = x + 1
Do Until t = True
    t = True
    i = GetMonsterID(, x)
    If i <> 0 Then
        t = False
        x = x + 1
    End If
Loop
With dbMonsters(UBound(dbMonsters))
    .dEXP = 25
    .dHP = 25
    .dMoney = 3
    .iAC = 4
    .iAtDayMonster = 0
    .iAtNightMonster = 0
    .iAttackable = 1
    .iDontAttackIfItem = 0
    .iDropCorpse = 0
    .iEvil = 0
    .iHostile = 0
    .iRoams = 1
    .iTameToFam = 0
    .iType = 0
    .lEnergy = 100
    .lID = x
    .lLevel = 1
    .lMobGroup = 1
    .lPEnergy = 100
    .lRegenTime = 0
    .lRegenTimeLeft = 0
    .sAttack = "1:4"
    .sDeathText = "0"
    .sDesc = "New Monster"
    .sMessage = "0"
    .sMonsterName = "New Monster"
    .sScript = "0"
    .sSpells = "0"
End With
lcID = UBound(dbMonsters)
FillMonsters lcID, True
MousePointer = vbDefault
End Sub

Private Sub cmdNext_Click()
   On Error GoTo cmdNext_Click_Error

SaveMonsters
lcID = lcID + 1
If lcID > UBound(dbMonsters) Then lcID = 1
FillMonsters lcID

   On Error GoTo 0
   Exit Sub

cmdNext_Click_Error:
End Sub

Private Sub cmdPrevious_Click()
   On Error GoTo cmdPrevious_Click_Error

SaveMonsters
lcID = lcID - 1
If lcID < LBound(dbMonsters) Then lcID = UBound(dbMonsters)
FillMonsters lcID

   On Error GoTo 0
   Exit Sub

cmdPrevious_Click_Error:
End Sub

Private Sub cmdSave_Click()
SaveMonsters
End Sub

Private Sub Form_Load()
txtScripting.IntelliSenseAddWordsFile App.Path & "\scriptdef.aimg"
txtScripting.IntelliSenseStartSubclassing
ucEvil.ListSettings = Alignment
lcID = 1
FillCBOS
FillMonsters , True
End Sub

Private Sub lstMonsters_Click()
If bIs Then Exit Sub
Dim i As Long
MousePointer = vbHourglass
For i = LBound(dbMonsters) To UBound(dbMonsters)
    With dbMonsters(i)
        If .lID & " " & .sMonsterName = lstMonsters.ItemText Then
            FillMonsters i
            Exit For
        End If
    End With
Next
MousePointer = vbDefault
End Sub

Sub SaveMonsters()
Dim i As Long
Dim s As String
MousePointer = vbHourglass
With dbMonsters(lcID)
    .dEXP = Val(txtEXP.Text)
    .dHP = Val(txtHP.Text)
    .dMoney = Val(txtGold.Text)
    .iAC = Val(txtAC.Text)
    s = cboAtDay.list(cboAtDay.ListIndex)
    .iAtDayMonster = Val(Mid$(s, 2, InStr(1, s, ")")))
    s = cboAtNight.list(cboAtNight.ListIndex)
    .iAtNightMonster = Val(Mid$(s, 2, InStr(1, s, ")")))
    .iAttackable = chkAttackable.Value
    s = cboNoAttackItem.list(cboNoAttackItem.ListIndex)
    .iDontAttackIfItem = Val(Mid$(s, 2, InStr(1, s, ")")))
    s = cboWeapon.list(cboWeapon.ListIndex)
    .lWeapon = Val(Mid$(s, 2, InStr(1, s, ")")))
    s = cboCorpse.list(cboCorpse.ListIndex)
    .iDropCorpse = Val(Mid$(s, 2, InStr(1, s, ")")))
    .iEvil = ucEvil.Number
    .iHostile = chkHostile.Value
    .iRoams = chkRoams.Value
    s = cboFamiliar.list(cboFamiliar.ListIndex)
    .iTameToFam = Val(Mid$(s, 2, InStr(1, s, ")")))
    s = cboType.list(cboType.ListIndex)
    .iType = Val(Mid$(s, 2, InStr(1, s, ")")))
    .sDropItem = ""
    If lstItems.ListCount > 0 Then
        For i = 1 To lstItems.ListCount
            s = lstItems.list(i)
            s = Mid$(s, 2, InStr(1, s, ")") - 2)
            .sDropItem = .sDropItem & s & "/"
            s = lstItems.list(i)
            s = Mid$(s, InStr(1, s, "@") + 2)
            s = Left$(s, InStr(1, s, "%") - 1)
            .sDropItem = .sDropItem & s & ";"
        Next
    Else
        .sDropItem = "0"
    End If
    .lEnergy = Val(txtEnergy.Text)
    .lLevel = Val(txtLevel.Text)
    .lMobGroup = Val(txtMob.Text)
    .lPEnergy = Val(txtPEnergy.Text)
    .lRegenTime = Val(txtRegenTime.Text)
    .lRegenTimeLeft = .lRegenTime
    .sAttack = txtMinDamage.Text & ":" & txtMaxDamage.Text
    .sDeathText = txtDeathText.Text
    .sDesc = txtDescription.Text
    .sMessage = txtMsg1.Text & ":" & txtMsg2.Text & ":" & txtMsg3.Text
    .sMonsterName = txtMonsterName.Text
    .sScript = txtScripting.Text
    If .sScript = "" Then .sScript = "0"
    s = ""
    For i = 0 To 4
        If aSpells(i) <> "" And aSpells(i) <> "0" Then
            s = s & aSpells(i) & ";"
        End If
    Next
    .sSpells = s
    If .sSpells = "" Then .sSpells = "0"
End With
modUpdateDatabase.SaveMemoryToDatabase 2
FillMonsters lcID, True
MousePointer = vbDefault
End Sub

Sub FillMonsters(Optional Arg As Long = -1, Optional FillList As Boolean = False)
Dim i As Long, j As Long
Dim m As Long
Dim Arr() As String
Dim Arr2() As String
MousePointer = vbHourglass
bIs = True
If Arg = -1 Then Arg = LBound(dbMonsters)
If FillList Then lstMonsters.Paint = False: lstMonsters.Clear
lstCSpells.Clear
txtSpellID.Text = "0"
txtSpellEnergy.Text = "0"
txtCPE.Text = "0"
txtCPR.Text = "0"
lstItems.Clear
For i = LBound(dbMonsters) To UBound(dbMonsters)
    With dbMonsters(i)
        If FillList Then lstMonsters.AddItem .lID & " " & .sMonsterName
        If Arg = i Then
            lcID = Arg
            txtEXP.Text = .dEXP
            txtHP.Text = .dHP
            txtGold.Text = .dMoney
            txtAC.Text = .iAC
            modMain.SetCBOSelectByID cboAtDay, CStr(.iAtDayMonster)
            modMain.SetCBOSelectByID cboAtNight, CStr(.iAtNightMonster)
            chkAttackable.Value = .iAttackable
            modMain.SetCBOSelectByID cboNoAttackItem, CStr(.iDontAttackIfItem)
            modMain.SetCBOSelectByID cboCorpse, CStr(.iDropCorpse)
            modMain.SetCBOSelectByID cboWeapon, CStr(.lWeapon)
            ucEvil.Number = .iEvil
            If .sDropItem <> "0" Then
                SplitFast .sDropItem, Arr, ";"
                For j = LBound(Arr) To UBound(Arr)
                    If Arr(j) <> "" Then
                        SplitFast Arr(j), Arr2, "/"
                        lstItems.AddItem "(" & Arr2(0) & ") " & dbItems(GetItemID(, Val(Arr2(0)))).sItemName & " @ " & Arr2(1) & "%"
                    End If
                Next
            End If
            chkHostile.Value = .iHostile
            chkRoams.Value = .iRoams
            modMain.SetCBOSelectByID cboFamiliar, CStr(.iTameToFam)
            modMain.SetCBOSelectByID cboType, CStr(.iType)
            
            txtEnergy.Text = .lEnergy
            txtID.Text = .lID
            txtLevel.Text = .lLevel
            txtMob.Text = .lMobGroup
            txtPEnergy.Text = .lPEnergy
            txtRegenTime.Text = .lRegenTime
            SplitFast .sAttack, Arr, ":"
            txtMinDamage.Text = Arr(0)
            txtMaxDamage.Text = Arr(1)
            txtDeathText.Text = .sDeathText
            txtDescription.Text = .sDesc
            Erase Arr
            If .sMessage <> "0" Then
                SplitFast .sMessage, Arr, ":"
                txtMsg1.Text = Arr(0)
                txtMsg2.Text = Arr(1)
                txtMsg3.Text = Arr(2)
            Else
                txtMsg1.Text = "attacks"
                txtMsg2.Text = "attacks"
                txtMsg3.Text = "attacks"
            End If
            txtMonsterName.Text = .sMonsterName
            txtScripting.Text = .sScript
            If txtScripting.Text = "0" Then txtScripting.Text = ""
            Erase Arr
            SplitFast .sSpells, Arr, ";"
            For j = LBound(Arr) To UBound(Arr) - 1
                aSpells(j) = Arr(j)
                If aSpells(j) <> "0" And aSpells(j) <> "" Then
                    m = GetSpellID(, Val(modSpell.GetItemIDFromUnFormattedString(aSpells(j))))
                    lstCSpells.AddItem "(" & dbSpells(m).lID & ") " & dbSpells(m).sSpellName
                End If
            Next
            If Not FillList Then Exit For
        End If
    End With
Next
If FillList Then lstMonsters.Paint = True
modMain.SetLstSelected lstMonsters, txtID.Text & " " & txtMonsterName.Text
bIs = False
MousePointer = Default
End Sub

Sub FillCBOS()
Dim i As Long
cboDropItem.Clear
cboDropItem.AddItem "(0) None"
cboWeapon.Clear
cboWeapon.AddItem "(0) None"
cboCorpse.Clear
cboCorpse.AddItem "(0) None"
cboPercent.Clear
cboAtDay.Clear
cboAtNight.Clear
cboAtDay.AddItem "(0) None"
cboAtNight.AddItem "(0) None"
cboFamiliar.Clear
cboFamiliar.AddItem "(0) None"
cboNoAttackItem.Clear
cboNoAttackItem.AddItem "(0) None"
lstSpells.Clear
For i = LBound(dbItems) To UBound(dbItems)
    With dbItems(i)
        cboDropItem.AddItem "(" & .iID & ") " & .sItemName
        cboNoAttackItem.AddItem "(" & .iID & ") " & .sItemName
        If .sWorn = "corpse" Then cboCorpse.AddItem "(" & .iID & ") " & .sItemName
        If .sWorn = "weapon" Then cboWeapon.AddItem "(" & .iID & ") " & .sItemName
    End With
Next
For i = LBound(dbMonsters) To UBound(dbMonsters)
    With dbMonsters(i)
        cboAtDay.AddItem "(" & .lID & ") " & .sMonsterName
        cboAtNight.AddItem "(" & .lID & ") " & .sMonsterName
    End With
Next
For i = 1 To 100
    cboPercent.AddItem CStr(i)
Next
For i = LBound(dbFamiliars) To UBound(dbFamiliars)
    With dbFamiliars(i)
        cboFamiliar.AddItem "(" & .iID & ") " & .sFamName
    End With
Next
lstSpells.Paint = False
For i = LBound(dbSpells) To UBound(dbSpells)
    lstSpells.AddItem "(" & dbSpells(i).lID & ") " & dbSpells(i).sSpellName
Next
lstSpells.Paint = True
End Sub

Private Sub txtMsg1_Change()
lblEX.Caption = txtMonsterName.Text & " " & txtMsg1.Text & " victim!"
End Sub

Private Sub txtScripting_MethodHasChanged()
lblMethod.Caption = txtScripting.Params
End Sub
