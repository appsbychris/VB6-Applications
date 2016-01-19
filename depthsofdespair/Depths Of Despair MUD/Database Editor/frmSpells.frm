VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSpells 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Spells"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11910
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
   ScaleHeight     =   6150
   ScaleWidth      =   11910
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   240
      TabIndex        =   63
      Top             =   240
      Width           =   2655
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5055
      Left            =   3240
      TabIndex        =   5
      Top             =   240
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "General Info"
      TabPicture(0)   =   "frmSpells.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLabel(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabel(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabel(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblLabel(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblLabel(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblLabel(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblLabel(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblLabel(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblLabel(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblLabel(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblLabel(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblLabel(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblLabel(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblLabel(13)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblDamage"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblLabel(20)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtID"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtLevel"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtDifficulty"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtMana"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtTimeout"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtMinDamage"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtMaxDamage"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtLevelModify"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtLevelMax"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cboSpellType"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cboElement"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtSpellName"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtShort"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cboCast"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cboUse"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "Messages"
      TabPicture(1)   =   "frmSpells.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtWearOff"
      Tab(1).Control(1)=   "txtStatMessage"
      Tab(1).Control(2)=   "txtMessage2"
      Tab(1).Control(3)=   "txtMessageV"
      Tab(1).Control(4)=   "txtMessage"
      Tab(1).Control(5)=   "lblLabel(19)"
      Tab(1).Control(6)=   "lblLabel(18)"
      Tab(1).Control(7)=   "lblLabel(17)"
      Tab(1).Control(8)=   "Line2"
      Tab(1).Control(9)=   "lblEX"
      Tab(1).Control(10)=   "lblSpells"
      Tab(1).Control(11)=   "lblLabel(16)"
      Tab(1).Control(12)=   "lblLabel(15)"
      Tab(1).Control(13)=   "lblLabel(14)"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Flags"
      TabPicture(2)   =   "frmSpells.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "flgOpts"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cboFlags"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "tbFlags"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "picTab(0)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "picTab(1)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.ComboBox cboUse 
         Height          =   315
         ItemData        =   "frmSpells.frx":0054
         Left            =   5280
         List            =   "frmSpells.frx":006A
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   1480
         Width           =   2775
      End
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  'None
         Height          =   3255
         Index           =   1
         Left            =   -74640
         ScaleHeight     =   3255
         ScaleWidth      =   7815
         TabIndex        =   56
         Top             =   1560
         Visible         =   0   'False
         Width           =   7815
         Begin VB.CommandButton cmdRemove 
            Caption         =   "< Remove"
            Height          =   255
            Index           =   1
            Left            =   5400
            TabIndex        =   59
            Top             =   50
            Width           =   2295
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add >"
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   58
            Top             =   50
            Width           =   2295
         End
         Begin MSComctlLib.ListView lstEndCast 
            Height          =   2775
            Left            =   2160
            TabIndex        =   57
            Top             =   360
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   4895
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lblHelp 
            Caption         =   "0"
            Height          =   3015
            Index           =   1
            Left            =   50
            TabIndex        =   60
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  'None
         Height          =   3255
         Index           =   0
         Left            =   -74640
         ScaleHeight     =   3255
         ScaleWidth      =   7815
         TabIndex        =   51
         Top             =   1560
         Width           =   7815
         Begin VB.CommandButton cmdRemove 
            Caption         =   "< Remove"
            Height          =   255
            Index           =   0
            Left            =   5400
            TabIndex        =   53
            Top             =   50
            Width           =   2295
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add >"
            Height          =   255
            Index           =   0
            Left            =   2160
            TabIndex        =   52
            Top             =   50
            Width           =   2295
         End
         Begin MSComctlLib.ListView lstSpellFlags 
            Height          =   2775
            Left            =   2160
            TabIndex        =   54
            Top             =   360
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   4895
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lblHelp 
            Caption         =   "0"
            Height          =   3015
            Index           =   0
            Left            =   50
            TabIndex        =   55
            Top             =   120
            Width           =   2055
         End
      End
      Begin MSComctlLib.TabStrip tbFlags 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   50
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6800
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Begin Cast Flags"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "End Cast Flags"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cboFlags 
         Height          =   315
         ItemData        =   "frmSpells.frx":00C1
         Left            =   -74880
         List            =   "frmSpells.frx":00EC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtWearOff 
         Height          =   375
         Left            =   -71640
         TabIndex        =   47
         Top             =   4440
         Width           =   4935
      End
      Begin VB.TextBox txtStatMessage 
         Height          =   375
         Left            =   -73440
         TabIndex        =   46
         Top             =   3960
         Width           =   6735
      End
      Begin VB.TextBox txtMessage2 
         Height          =   375
         Left            =   -74760
         TabIndex        =   42
         Top             =   3120
         Width           =   8055
      End
      Begin VB.TextBox txtMessageV 
         Height          =   375
         Left            =   -74760
         TabIndex        =   41
         Top             =   2400
         Width           =   8055
      End
      Begin VB.TextBox txtMessage 
         Height          =   375
         Left            =   -74760
         TabIndex        =   40
         Top             =   1680
         Width           =   8055
      End
      Begin VB.ComboBox cboCast 
         Height          =   315
         ItemData        =   "frmSpells.frx":01E0
         Left            =   1560
         List            =   "frmSpells.frx":01F6
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txtShort 
         Height          =   375
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   32
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtSpellName 
         Height          =   375
         Left            =   1560
         TabIndex        =   31
         Top             =   1080
         Width           =   2415
      End
      Begin VB.ComboBox cboElement 
         Height          =   315
         ItemData        =   "frmSpells.frx":020C
         Left            =   5280
         List            =   "frmSpells.frx":022E
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox cboSpellType 
         Height          =   315
         ItemData        =   "frmSpells.frx":02B5
         Left            =   1560
         List            =   "frmSpells.frx":02DA
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   2520
         Width           =   1455
      End
      Begin ServerEditor.NumOnlyText txtLevelMax 
         Height          =   375
         Left            =   7320
         TabIndex        =   28
         Top             =   2400
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
      Begin ServerEditor.NumOnlyText txtLevelModify 
         Height          =   375
         Left            =   5280
         TabIndex        =   27
         Top             =   2400
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
         Text            =   "3"
         AllowNeg        =   0   'False
         Align           =   0
         MaxLength       =   5
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin ServerEditor.NumOnlyText txtMaxDamage 
         Height          =   375
         Left            =   7320
         TabIndex        =   26
         Top             =   1920
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
         Text            =   "22"
         AllowNeg        =   0   'False
         Align           =   0
         MaxLength       =   5
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin ServerEditor.NumOnlyText txtMinDamage 
         Height          =   375
         Left            =   5280
         TabIndex        =   25
         Top             =   1920
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
         Text            =   "10"
         AllowNeg        =   0   'False
         Align           =   0
         MaxLength       =   5
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin ServerEditor.NumOnlyText txtTimeout 
         Height          =   375
         Left            =   5280
         TabIndex        =   24
         Top             =   1080
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
      Begin ServerEditor.NumOnlyText txtMana 
         Height          =   375
         Left            =   1560
         TabIndex        =   23
         Top             =   3360
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
      Begin ServerEditor.NumOnlyText txtDifficulty 
         Height          =   375
         Left            =   1560
         TabIndex        =   22
         Top             =   2880
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
         AllowNeg        =   -1  'True
         Align           =   0
         MaxLength       =   4
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin ServerEditor.NumOnlyText txtLevel 
         Height          =   375
         Left            =   1560
         TabIndex        =   21
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
         MaxLength       =   3
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin ServerEditor.NumOnlyText txtID 
         Height          =   375
         Left            =   1560
         TabIndex        =   20
         Top             =   600
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
         MaxLength       =   0
         Enabled         =   0   'False
         Backcolor       =   -2147483643
      End
      Begin ServerEditor.FlagOptions flgOpts 
         Height          =   375
         Left            =   -71160
         TabIndex        =   48
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Style           =   0
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Use:"
         Height          =   195
         Index           =   20
         Left            =   4200
         TabIndex        =   62
         Top             =   1485
         Width           =   330
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Send this message when spell wears off:"
         Height          =   195
         Index           =   19
         Left            =   -74640
         TabIndex        =   45
         Top             =   4560
         Width           =   2925
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "'Stat' Message:"
         Height          =   195
         Index           =   18
         Left            =   -74640
         TabIndex        =   44
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Bless Messages:"
         Height          =   195
         Index           =   17
         Left            =   -74880
         TabIndex        =   43
         Top             =   3720
         Width           =   1170
      End
      Begin VB.Line Line2 
         X1              =   -74880
         X2              =   -67560
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label lblEX 
         Caption         =   "EX"
         Height          =   975
         Left            =   -70800
         TabIndex        =   39
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label lblSpells 
         Caption         =   "Constants"
         Height          =   975
         Left            =   -74880
         TabIndex        =   38
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Message 3 (Bystander's Perspective):"
         Height          =   195
         Index           =   16
         Left            =   -74880
         TabIndex        =   37
         Top             =   2880
         Width           =   2715
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Message 2 (Victim's Perspective):"
         Height          =   195
         Index           =   15
         Left            =   -74880
         TabIndex        =   36
         Top             =   2160
         Width           =   2385
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Message 1 (Your perspective): "
         Height          =   195
         Index           =   14
         Left            =   -74880
         TabIndex        =   35
         Top             =   1440
         Width           =   2250
      End
      Begin VB.Line Line1 
         X1              =   4080
         X2              =   4080
         Y1              =   480
         Y2              =   4920
      End
      Begin VB.Label lblDamage 
         Caption         =   "Damage"
         Height          =   2055
         Left            =   4200
         TabIndex        =   34
         Top             =   2880
         Width           =   4095
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Timeout:"
         Height          =   195
         Index           =   13
         Left            =   4200
         TabIndex        =   19
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Level Max:"
         Height          =   195
         Index           =   12
         Left            =   6240
         TabIndex        =   18
         Top             =   2280
         Width           =   780
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Level Modify:"
         Height          =   195
         Index           =   11
         Left            =   4200
         TabIndex        =   17
         Top             =   2280
         Width           =   960
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Max Damage:"
         Height          =   195
         Index           =   10
         Left            =   6240
         TabIndex        =   16
         Top             =   1920
         Width           =   990
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Min Damage:"
         Height          =   195
         Index           =   9
         Left            =   4200
         TabIndex        =   15
         Top             =   1920
         Width           =   930
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Difficulty:"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   14
         Top             =   2880
         Width           =   690
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Element:"
         Height          =   195
         Index           =   7
         Left            =   4200
         TabIndex        =   13
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Cast Per Round:"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   12
         Top             =   3840
         Width           =   1185
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Mana Cost:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   11
         Top             =   3360
         Width           =   825
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Spell Type:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   795
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Spell Level:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   810
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Spell Short:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Spell Name:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   225
      End
   End
   Begin ServerEditor.UltraBox lstSpells 
      Height          =   4695
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   8281
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
      Left            =   6840
      TabIndex        =   0
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "(new)"
      Height          =   255
      Left            =   8040
      TabIndex        =   1
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "< Previous"
      Height          =   255
      Left            =   9360
      TabIndex        =   3
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   255
      Left            =   10560
      TabIndex        =   2
      Top             =   5640
      Width           =   1095
   End
   Begin ServerEditor.Raise Raise1 
      Height          =   5295
      Left            =   3120
      TabIndex        =   64
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9340
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise2 
      Height          =   5295
      Left            =   120
      TabIndex        =   65
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   9340
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise3 
      Height          =   495
      Left            =   6720
      TabIndex        =   66
      Top             =   5520
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise4 
      Height          =   6135
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   10821
      Style           =   4
      Color           =   0
   End
End
Attribute VB_Name = "frmSpells"
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
Rem***************                frmSpells                       **********************
Rem***************                ServerEditor                    **********************
Rem***************                Editor.vbp                      **********************
Rem*************************************************************************************
Rem*************************************************************************************

Dim lcID As Long
Dim bIs As Boolean

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
Private Sub txtFind_Change()
lstSpells.SetSelected lstSpells.FindInStr(txtFind.Text), True, True
End Sub

Private Sub cboFlags_Change()
cboFlags_Click
End Sub

Private Sub cboFlags_Click()
Dim s As String
If cboFlags.ListIndex <> -1 Then
    lblHelp(0).Caption = GetHelp(cboFlags.list(cboFlags.ListIndex))
    lblHelp(1).Caption = GetHelp(cboFlags.list(cboFlags.ListIndex))
    flgOpts.ViewStyle = modMain.DeterStyle(modMain.ShortFlag(cboFlags.list(cboFlags.ListIndex)), s)
    If flgOpts.ViewStyle = ComboInputFeed Then modMain.FeedAList flgOpts, s
    cmdAdd(0).Caption = "&Add"
    cmdAdd(1).Caption = "&Add"
End If
End Sub


Private Sub cmdAdd_Click(Index As Integer)
If cmdAdd(Index).Caption = "&Add" Then
    If Index = 0 Then
        lstSpellFlags.ListItems.Add Text:=modMain.TranslateFlag(modMain.ShortFlag( _
            cboFlags.list(cboFlags.ListIndex)) & flgOpts.GetCurVal)
    Else
        lstEndCast.ListItems.Add Text:=modMain.TranslateFlag(modMain.ShortFlag( _
            cboFlags.list(cboFlags.ListIndex)) & flgOpts.GetCurVal)
    End If
Else
    If Index = 0 Then
        lstSpellFlags.SelectedItem.Text = modMain.TranslateFlag(modMain.ShortFlag( _
        cboFlags.list(cboFlags.ListIndex)) & flgOpts.GetCurVal)
    Else
        lstEndCast.SelectedItem.Text = modMain.TranslateFlag(modMain.ShortFlag( _
        cboFlags.list(cboFlags.ListIndex)) & flgOpts.GetCurVal)
    End If
    cmdAdd(Index).Caption = "&Add"
End If
End Sub

Private Sub cmdNew_Click()
MousePointer = vbHourglass
Dim x As Long
Dim i As Long
Dim t As Boolean
ReDim Preserve dbSpells(1 To UBound(dbSpells) + 1)
x = dbSpells(UBound(dbSpells) - 1).lID
x = x + 1
Do Until t = True
    t = True
    i = GetSpellID(, x)
    If i <> 0 Then
        t = False
        x = x + 1
    End If
Loop
With dbSpells(UBound(dbSpells))
    .iCast = 1
    .iDifficulty = 70
    .iLevel = 1
    .iLevelMax = 1
    .iLevelModify = 1
    .iType = 1
    .iUse = 0
    .lElement = -1
    .lID = x
    .lMana = 0
    .lMaxDam = 0
    .lMinDam = 0
    .lTimeOut = 0
    .sEndCastFlags = "0"
    .sFlags = "0"
    .sMessage = "0"
    .sMessage2 = "0"
    .sMessageV = "0"
    .sRunOutMessage = "0"
    .sShort = "nwsp"
    .sSpellName = "New Spell"
    .sStatMessage = "0"
End With
lcID = UBound(dbItems)
FillSpells lcID, True
MousePointer = vbDefault
End Sub

Private Sub cmdNext_Click()
On Error GoTo cmdNext_Click_Error
SaveSpells
lcID = lcID + 1
If lcID > UBound(dbSpells) Then lcID = 1
FillSpells lcID
On Error GoTo 0
Exit Sub
cmdNext_Click_Error:
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo cmdPrevious_Click_Error
SaveSpells
lcID = lcID - 1
If lcID < LBound(dbSpells) Then lcID = UBound(dbSpells)
FillSpells lcID
On Error GoTo 0
Exit Sub
cmdPrevious_Click_Error:
End Sub

Private Sub cmdRemove_Click(Index As Integer)
If Index = 0 Then
    If lstSpellFlags.ListItems.Count > 0 Then lstSpellFlags.ListItems.Remove lstSpellFlags.SelectedItem.Index
Else
    If lstEndCast.ListItems.Count > 0 Then lstEndCast.ListItems.Remove lstEndCast.SelectedItem.Index
End If
cmdAdd(0).Caption = "&Add"
cmdAdd(1).Caption = "&Add"
End Sub

Private Sub cmdSave_Click()
SaveSpells
End Sub

Private Sub Form_Load()
lblDamage.Caption = "Damage:" & vbCrLf & "Damage range is figured out like so:" & vbCrLf & "A random number is created from the MIN DAMAGE to the MAX DAMAGE + LEVEL MODIFY x the player's LEVEL, and that total PLUS the players INTELLECT divided by 10. Rnd(" & txtMinDamage.Text & " - " & txtMaxDamage.Text & ") + (" & txtLevelModify.Text & " x LEVEL) + (INT / 10)" & vbCrLf & "Range with these values (excluding intellect bonus, and level = LEVEL MAX:" & vbCrLf & "(0-0)"
lblSpells.Caption = "Constants:" & vbCrLf & "<%c> = Caster's Name" & vbCrLf & "<%v> = Victim's Name" & vbCrLf & "<%s> = Spell's Name" & vbCrLf & "<%d> = Damage Value" & vbCrLf
lblEX.Caption = "EX: ""<%c> cast <%s> at <%v> for <%d> damage!"" would output: ""Spike cast magic missle at Thrice for 14 damage!"", assuming those were the respective names."
lstSpellFlags.ColumnHeaders(1).Width = 10000
lstEndCast.ColumnHeaders(1).Width = 10000
FillSpells , True
End Sub

Private Sub FillSpells(Optional Arg As Long = -1, Optional FillList As Boolean = False)
Dim i As Long, j As Long
Dim m As Long
Dim Arr() As String
MousePointer = vbHourglass
bIs = True
modMain.PopulateCBOFlag cboFlags
If Arg = -1 Then Arg = LBound(dbSpells)
If FillList Then lstSpells.Paint = False: lstSpells.Clear
For i = LBound(dbSpells) To UBound(dbSpells)
    With dbSpells(i)
        If FillList Then lstSpells.AddItem CStr(.lID & " " & .sSpellName), BCOLOR:=vbWhite
        If i = Arg Then
            lcID = Arg
            txtID.Text = .lID
            txtSpellName = .sSpellName
            txtLevel.Text = .iLevel
            modMain.SetCBOlstIndex cboSpellType, .iType, [Magic Type]
            txtDifficulty.Text = .iDifficulty
            txtMana.Text = .lMana
            modMain.SetListIndex cboCast, CStr(.iCast)
            modMain.SetCBOlstIndex cboElement, .lElement, [Element Type]
            txtTimeout.Text = .lTimeOut
            txtMinDamage.Text = .lMinDam
            txtMaxDamage.Text = .lMaxDam
            txtLevelModify.Text = .iLevelModify
            txtLevelMax.Text = .iLevelMax
            txtMessage.Text = .sMessage
            txtMessageV.Text = .sMessageV
            txtShort.Text = .sShort
            txtMessage2.Text = .sMessage2
            txtStatMessage.Text = .sStatMessage
            txtWearOff.Text = .sRunOutMessage
            lstSpellFlags.ListItems.Clear
            modMain.SetCBOlstIndex cboUse, .iUse, [Spell Use]
            If .sFlags <> "0" Then
                SplitFast .sFlags, Arr, ";"
                For j = LBound(Arr) To UBound(Arr)
                    If Arr(j) <> "" And Arr(j) <> "0" Then
                        lstSpellFlags.ListItems.Add Text:=modMain.TranslateFlag(Arr(j))
                    End If
                Next
            End If
            Erase Arr
            lstEndCast.ListItems.Clear
            If .sEndCastFlags <> "0" Then
                SplitFast .sEndCastFlags, Arr, ";"
                For j = LBound(Arr) To UBound(Arr)
                    If Arr(j) <> "" And Arr(j) <> "0" Then
                        lstEndCast.ListItems.Add Text:=modMain.TranslateFlag(Arr(j))
                    End If
                Next
            End If
            If Not FillList Then Exit For
        End If
    End With
Next
lstSpells.Paint = True
modMain.SetLstSelected lstSpells, txtID.Text & " " & txtSpellName.Text
bIs = False
MousePointer = vbDefault
End Sub

Private Sub lstendcast_Click()
Dim s As String
Dim t As String
If lstEndCast.ListItems.Count < 1 Then Exit Sub
s = Left$(lstEndCast.SelectedItem.Text, 25)
s = Trim$(s)
For i = 0 To cboFlags.ListCount - 1
    If cboFlags.list(i) = s Then
        cboFlags.ListIndex = i
        Exit For
    End If
Next
t = Mid$(lstEndCast.SelectedItem.Text, 28)
t = Trim$(t)
flgOpts.SetVal Val(t)
cmdAdd(1).Caption = "&Modify"
End Sub

Private Sub lstSpellFlags_Click()
Dim s As String
Dim t As String
If lstSpellFlags.ListItems.Count < 1 Then Exit Sub
s = Left$(lstSpellFlags.SelectedItem.Text, 25)
s = Trim$(s)
For i = 0 To cboFlags.ListCount - 1
    If cboFlags.list(i) = s Then
        cboFlags.ListIndex = i
        Exit For
    End If
Next
t = Mid$(lstSpellFlags.SelectedItem.Text, 28)
t = Trim$(t)
flgOpts.SetVal Val(t)
cmdAdd(0).Caption = "&Modify"
End Sub

Private Sub SaveSpells()
Dim s As String, i As Long
MousePointer = vbHourglass
ReverseEffects lcID, Spells
With dbSpells(lcID)
    .iCast = cboCast.Text
    .iDifficulty = txtDifficulty.Text
    .iLevel = txtLevel.Text
    .iLevelMax = txtLevelMax.Text
    .iLevelModify = txtLevelModify.Text
    s = cboSpellType.list(cboSpellType.ListIndex)
    .iType = Val(Left$(s, InStr(1, s, " ")))
    s = cboUse.list(cboUse.ListIndex)
    .iUse = Val(Left$(s, InStr(1, s, " ")))
    s = txtID.Text
    txtID.Text = cboElement.Text
    .lElement = txtID.Text
    txtID.Text = s
    .lMana = txtMana.Text
    .lMaxDam = txtMaxDamage.Text
    .lMinDam = txtMinDamage.Text
    .lTimeOut = txtTimeout.Text
    .sMessage = txtMessage.Text
    .sMessage2 = txtMessage2.Text
    .sMessageV = txtMessageV.Text
    .sRunOutMessage = txtWearOff.Text
    .sShort = txtShort.Text
    .sSpellName = txtSpellName.Text
    .sStatMessage = txtStatMessage.Text
    .sFlags = ""
    If lstSpellFlags.ListItems.Count > 0 Then
        For i = 1 To lstSpellFlags.ListItems.Count
            .sFlags = .sFlags & modMain.MakeDBFlag(lstSpellFlags.ListItems(i).Text) & ";"
        Next
    Else
        .sFlags = "0"
    End If
    .sEndCastFlags = ""
    If lstEndCast.ListItems.Count > 0 Then
        For i = 1 To lstEndCast.ListItems.Count
            .sEndCastFlags = .sEndCastFlags & modMain.MakeDBFlag(lstEndCast.ListItems(i).Text) & ";"
        Next
    Else
        .sEndCastFlags = "0"
    End If
End With
SaveMemoryToDatabase Spells
DoEffects lcID, Spells
modUpdateDatabase.SaveMemoryToDatabase Players
FillSpells lcID, True
MousePointer = vbDefault
End Sub

Private Sub lstSpells_Click()
If bIs Then Exit Sub
Dim i As Long
MousePointer = vbHourglass
For i = LBound(dbSpells) To UBound(dbSpells)
    With dbSpells(i)
        If .lID & " " & .sSpellName = lstSpells.ItemText Then
            FillSpells i
            Exit For
        End If
    End With
Next
MousePointer = vbDefault
End Sub

Private Sub tbFlags_Click()
picTab(0).Visible = False
picTab(1).Visible = False
picTab(tbFlags.SelectedItem.Index - 1).Visible = True

End Sub

Private Sub txtLevelMax_Change()
txtMinDamage_Change
End Sub

Private Sub txtLevelModify_Change()
txtMinDamage_Change
End Sub

Private Sub txtMaxDamage_Change()
txtMinDamage_Change
End Sub

Private Sub txtMinDamage_Change()
Dim l&
Dim j&
l& = Val(txtMaxDamage.Text)
j& = Val(txtLevelMax.Text) * Val(txtLevelModify.Text)
lblDamage.Caption = "Damage:" & vbCrLf & "Damage range is figured out like so:" & vbCrLf & "A random number is created from the MIN DAMAGE to the MAX DAMAGE + LEVEL MODIFY x the player's LEVEL, and that total PLUS the players INTELLECT divided by 10. Rnd(" & txtMinDamage.Text & " - " & txtMaxDamage.Text & ") + (" & txtLevelModify.Text & " x LEVEL) + (INT / 10)" & vbCrLf & "Range with these values (excluding intellect bonus, and level = LEVEL MAX:" & vbCrLf & "(" & (Val(txtMinDamage.Text) + j&) & " - " & l& + j& & ")"
End Sub
