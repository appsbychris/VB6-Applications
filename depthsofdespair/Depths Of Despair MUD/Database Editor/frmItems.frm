VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Items"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   12975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   12975
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin ServerEditor.UltraBox lstItems 
      Height          =   5295
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   9340
      Style           =   3
      Color           =   0
      Fill            =   0
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
   Begin TabDlg.SSTab ssTab 
      Height          =   5655
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
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
      TabCaption(0)   =   "Basic Stats"
      TabPicture(0)   =   "frmItems.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture6"
      Tab(0).Control(1)=   "Picture4"
      Tab(0).Control(2)=   "Picture3"
      Tab(0).Control(3)=   "Picture2"
      Tab(0).Control(4)=   "Picture1"
      Tab(0).Control(5)=   "Raise1"
      Tab(0).Control(6)=   "Line4"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Flags"
      TabPicture(1)   =   "frmItems.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ssFlags"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Picture5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cboFlags"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "flgOpts"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Description"
      TabPicture(2)   =   "frmItems.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtOth3"
      Tab(2).Control(1)=   "txtOth2"
      Tab(2).Control(2)=   "txtOth1"
      Tab(2).Control(3)=   "txtVic3"
      Tab(2).Control(4)=   "txtVic2"
      Tab(2).Control(5)=   "txtVic1"
      Tab(2).Control(6)=   "txtMsg1"
      Tab(2).Control(7)=   "txtMsg2"
      Tab(2).Control(8)=   "txtMsg3"
      Tab(2).Control(9)=   "txtDescription"
      Tab(2).Control(10)=   "Line7"
      Tab(2).Control(11)=   "lblOther"
      Tab(2).Control(12)=   "lblVictim"
      Tab(2).Control(13)=   "lblYou"
      Tab(2).Control(14)=   "Label1(2)"
      Tab(2).Control(15)=   "Label1(1)"
      Tab(2).Control(16)=   "Label1(0)"
      Tab(2).Control(17)=   "lblLabel(8)"
      Tab(2).Control(18)=   "lblLabel(11)"
      Tab(2).ControlCount=   19
      TabCaption(3)   =   "Restrictions"
      TabPicture(3)   =   "frmItems.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Raise2"
      Tab(3).Control(1)=   "Picture7"
      Tab(3).Control(2)=   "Picture8"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Scripting"
      TabPicture(4)   =   "frmItems.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label4"
      Tab(4).Control(1)=   "lblMethod"
      Tab(4).Control(2)=   "txtScripting"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Projectile Definition"
      TabPicture(5)   =   "frmItems.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label2(0)"
      Tab(5).Control(1)=   "Label2(1)"
      Tab(5).Control(2)=   "lstNo"
      Tab(5).Control(3)=   "lstBows"
      Tab(5).Control(4)=   "cmdAddAmmo"
      Tab(5).Control(5)=   "cmdRemoveAmmo"
      Tab(5).Control(6)=   "txtNo"
      Tab(5).ControlCount=   7
      Begin ServerEditor.NumOnlyText txtNo 
         Height          =   615
         Left            =   -70680
         TabIndex        =   110
         Top             =   600
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
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
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin VB.CommandButton cmdRemoveAmmo 
         Height          =   1575
         Left            =   -71040
         Picture         =   "frmItems.frx":00A8
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddAmmo 
         Height          =   1575
         Left            =   -71040
         Picture         =   "frmItems.frx":04EA
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   1560
         Width           =   1695
      End
      Begin ServerEditor.UltraBox lstBows 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   104
         Top             =   960
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   8070
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
      Begin VB.TextBox txtOth3 
         Height          =   330
         Left            =   -68280
         TabIndex        =   38
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtOth2 
         Height          =   330
         Left            =   -68280
         TabIndex        =   37
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtOth1 
         Height          =   330
         Left            =   -68280
         TabIndex        =   36
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtVic3 
         Height          =   330
         Left            =   -71160
         TabIndex        =   35
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtVic2 
         Height          =   330
         Left            =   -71160
         TabIndex        =   34
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtVic1 
         Height          =   330
         Left            =   -71160
         TabIndex        =   33
         Top             =   960
         Width           =   2775
      End
      Begin ServerEditor.IntelliSense txtScripting 
         Height          =   4215
         Left            =   -74760
         TabIndex        =   48
         Top             =   840
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   7435
      End
      Begin VB.PictureBox Picture8 
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   -74760
         ScaleHeight     =   2415
         ScaleWidth      =   9135
         TabIndex        =   86
         Top             =   3000
         Width           =   9135
         Begin VB.CommandButton cmdRaceClear 
            Caption         =   "< &Delete <"
            Height          =   255
            Left            =   3480
            TabIndex        =   47
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CommandButton cmdAddRace 
            Caption         =   "> &Add >"
            Height          =   255
            Left            =   3480
            TabIndex        =   45
            Top             =   480
            Width           =   2055
         End
         Begin ServerEditor.UltraBox lstRaceRes 
            Height          =   1815
            Left            =   5640
            TabIndex        =   46
            Top             =   480
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   3201
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
         Begin ServerEditor.UltraBox lstRaces 
            Height          =   1815
            Left            =   0
            TabIndex        =   44
            Top             =   480
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   3201
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
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "All Races Available:"
            Height          =   195
            Index           =   16
            Left            =   0
            TabIndex        =   89
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Current Restrictions:"
            Height          =   195
            Index           =   15
            Left            =   5640
            TabIndex        =   88
            Top             =   240
            Width           =   1500
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Race Restrictions:"
            Height          =   195
            Index           =   14
            Left            =   0
            TabIndex        =   87
            Top             =   0
            Width           =   1305
         End
      End
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   -74760
         ScaleHeight     =   2535
         ScaleWidth      =   9135
         TabIndex        =   82
         Top             =   480
         Width           =   9135
         Begin VB.CommandButton cmdClassClear 
            Caption         =   "< &Delete <"
            Height          =   255
            Left            =   3480
            TabIndex        =   43
            Top             =   2160
            Width           =   2055
         End
         Begin VB.CommandButton cmdAddClass 
            Caption         =   "> &Add >"
            Height          =   255
            Left            =   3480
            TabIndex        =   41
            Top             =   600
            Width           =   2055
         End
         Begin ServerEditor.UltraBox lstClassRes 
            Height          =   1815
            Left            =   5640
            TabIndex        =   42
            Top             =   600
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   3201
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
         Begin ServerEditor.UltraBox lstClasses 
            Height          =   1815
            Left            =   0
            TabIndex        =   40
            Top             =   600
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   3201
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
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "All Classes Available:"
            Height          =   195
            Index           =   7
            Left            =   0
            TabIndex        =   85
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Current Restrictions:"
            Height          =   195
            Index           =   6
            Left            =   5640
            TabIndex        =   84
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Class Restrictions:"
            Height          =   195
            Index           =   21
            Left            =   0
            TabIndex        =   83
            Top             =   0
            Width           =   1320
         End
      End
      Begin ServerEditor.Raise Raise2 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   81
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9128
         Style           =   2
         Color           =   0
      End
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -74760
         ScaleHeight     =   255
         ScaleWidth      =   8775
         TabIndex        =   79
         Top             =   2880
         Width           =   8775
         Begin VB.Line Line5 
            X1              =   120
            X2              =   8520
            Y1              =   120
            Y2              =   120
         End
      End
      Begin ServerEditor.FlagOptions flgOpts 
         Height          =   375
         Left            =   5040
         TabIndex        =   24
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Style           =   0
      End
      Begin VB.ComboBox cboFlags 
         Height          =   315
         ItemData        =   "frmItems.frx":092C
         Left            =   240
         List            =   "frmItems.frx":09FF
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   600
         Width           =   4695
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   2400
         ScaleHeight     =   495
         ScaleWidth      =   6975
         TabIndex        =   77
         Top             =   4920
         Width           =   6975
         Begin VB.CommandButton cmdRemove 
            Caption         =   "&Remove"
            Height          =   375
            Left            =   5760
            TabIndex        =   28
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   4560
            TabIndex        =   27
            Top             =   0
            Width           =   1095
         End
      End
      Begin TabDlg.SSTab ssFlags 
         Height          =   4455
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   7858
         _Version        =   393216
         Style           =   1
         Tabs            =   2
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
         TabCaption(0)   =   "Flags 1"
         TabPicture(0)   =   "frmItems.frx":0F43
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstFlags1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Flags 2"
         TabPicture(1)   =   "frmItems.frx":0F5F
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstFlags2"
         Tab(1).ControlCount=   1
         Begin MSComctlLib.ListView lstFlags1 
            Height          =   3375
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   5953
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
         Begin MSComctlLib.ListView lstFlags2 
            Height          =   3375
            Left            =   -74880
            TabIndex        =   29
            Top             =   360
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   5953
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
      End
      Begin VB.PictureBox Picture4 
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
         Height          =   2295
         Left            =   -69840
         ScaleHeight     =   2295
         ScaleWidth      =   4095
         TabIndex        =   74
         Top             =   3120
         Width           =   4095
         Begin VB.CheckBox chkSD 
            Caption         =   "Check1"
            Height          =   200
            Left            =   2280
            TabIndex        =   21
            Top             =   1920
            Width           =   200
         End
         Begin VB.CheckBox chkUseFlgs2 
            Caption         =   "Check1"
            Height          =   200
            Left            =   2160
            TabIndex        =   18
            Top             =   600
            Width           =   200
         End
         Begin ServerEditor.NumOnlyText txtCost 
            Height          =   375
            Left            =   1440
            TabIndex        =   16
            Top             =   10
            Width           =   1215
            _ExtentX        =   873
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
         Begin VB.CheckBox chkMoveable 
            Caption         =   "Moveable"
            Height          =   225
            Left            =   1440
            TabIndex        =   20
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CheckBox chkLedgendary 
            Caption         =   "Ledgendary"
            Height          =   225
            Left            =   1440
            TabIndex        =   19
            Top             =   960
            Width           =   1215
         End
         Begin ServerEditor.NumOnlyText txtUses 
            Height          =   375
            Left            =   1440
            TabIndex        =   17
            Top             =   480
            Width           =   615
            _ExtentX        =   2990
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
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   4
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtSD 
            Height          =   375
            Left            =   2520
            TabIndex        =   22
            Top             =   1800
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
            Enabled         =   0   'False
            Backcolor       =   -2147483643
         End
         Begin VB.Label lblLabel 
            Caption         =   "While the Item is equiped, subtract this amount from the item's durribility each round:"
            Height          =   675
            Index           =   20
            Left            =   120
            TabIndex        =   100
            Top             =   1560
            Width           =   2325
         End
         Begin VB.Label lblUses 
            AutoSize        =   -1  'True
            Caption         =   "Amount Of Uses:"
            Height          =   195
            Left            =   120
            TabIndex        =   99
            Top             =   480
            Width           =   1230
         End
         Begin VB.Label lblLabel 
            Caption         =   "On the last use, trigger Flags 2"
            Height          =   435
            Index           =   19
            Left            =   2520
            TabIndex        =   98
            Top             =   480
            Width           =   1245
         End
         Begin VB.Line Line6 
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   2520
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "gold"
            Height          =   195
            Index           =   38
            Left            =   2760
            TabIndex        =   76
            Top             =   120
            Width           =   300
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cost:"
            Height          =   195
            Index           =   12
            Left            =   840
            TabIndex        =   75
            Top             =   0
            Width           =   390
         End
      End
      Begin VB.PictureBox Picture3 
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
         Height          =   2055
         Left            =   -74760
         ScaleHeight     =   2055
         ScaleWidth      =   4695
         TabIndex        =   68
         Top             =   3120
         Width           =   4695
         Begin ServerEditor.NumOnlyText txtDur 
            Height          =   375
            Left            =   1200
            TabIndex        =   8
            Top             =   10
            Width           =   1815
            _ExtentX        =   3201
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
         Begin ServerEditor.NumOnlyText txtLimit 
            Height          =   375
            Left            =   1200
            TabIndex        =   9
            Top             =   480
            Width           =   615
            _ExtentX        =   873
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
         Begin ServerEditor.NumOnlyText txtLevelRestriction 
            Height          =   375
            Left            =   2400
            TabIndex        =   11
            Top             =   1560
            Width           =   735
            _ExtentX        =   873
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
         Begin ServerEditor.NumOnlyText txtClassPoints 
            Height          =   375
            Left            =   2400
            TabIndex        =   10
            Top             =   1080
            Width           =   2175
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
         Begin VB.Line Line3 
            X1              =   0
            X2              =   4800
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Class Points Required To Use:"
            Height          =   195
            Index           =   18
            Left            =   240
            TabIndex        =   73
            Top             =   1200
            Width           =   2145
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "In Game Limit:"
            Height          =   195
            Index           =   25
            Left            =   0
            TabIndex        =   72
            Top             =   480
            Width           =   1020
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Level Required To Be:"
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   71
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Durability:"
            Height          =   195
            Index           =   35
            Left            =   0
            TabIndex        =   70
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Restrictions:"
            Height          =   195
            Index           =   36
            Left            =   120
            TabIndex        =   69
            Top             =   960
            Width           =   900
         End
      End
      Begin VB.PictureBox Picture2 
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
         Height          =   2535
         Left            =   -68640
         ScaleHeight     =   2535
         ScaleWidth      =   2175
         TabIndex        =   61
         Top             =   480
         Width           =   2175
         Begin ServerEditor.NumOnlyText txtSpeed 
            Height          =   375
            Left            =   1080
            TabIndex        =   15
            Top             =   2040
            Width           =   735
            _ExtentX        =   873
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
         Begin ServerEditor.NumOnlyText txtMagicLevel 
            Height          =   375
            Left            =   1080
            TabIndex        =   67
            Top             =   1560
            Width           =   735
            _ExtentX        =   873
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
         Begin ServerEditor.NumOnlyText txtAC 
            Height          =   375
            Left            =   1080
            TabIndex        =   14
            Top             =   1080
            Width           =   735
            _ExtentX        =   873
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
         Begin ServerEditor.NumOnlyText txtMinDamage 
            Height          =   375
            Left            =   1080
            TabIndex        =   12
            Top             =   0
            Width           =   975
            _ExtentX        =   873
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
         Begin ServerEditor.NumOnlyText txtMaxDamage 
            Height          =   375
            Left            =   1080
            TabIndex        =   13
            Top             =   480
            Width           =   975
            _ExtentX        =   873
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
         Begin VB.Line Line2 
            X1              =   0
            X2              =   2040
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Min Damage:"
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   66
            Top             =   0
            Width           =   930
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Max Damage:"
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   65
            Top             =   480
            Width           =   990
         End
         Begin VB.Label lblAC 
            AutoSize        =   -1  'True
            Caption         =   "Armor Class"
            Height          =   195
            Left            =   0
            TabIndex        =   64
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Magic Level:"
            Height          =   195
            Index           =   26
            Left            =   0
            TabIndex        =   63
            Top             =   1560
            Width           =   885
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Speed:"
            Height          =   195
            Index           =   9
            Left            =   360
            TabIndex        =   62
            Top             =   2040
            Width           =   510
         End
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
         Height          =   2535
         Left            =   -74760
         ScaleHeight     =   2535
         ScaleWidth      =   5895
         TabIndex        =   55
         Top             =   480
         Width           =   5895
         Begin VB.ComboBox cboWeaponType 
            Height          =   315
            ItemData        =   "frmItems.frx":0F7B
            Left            =   1080
            List            =   "frmItems.frx":0FA0
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2040
            Width           =   2415
         End
         Begin VB.ComboBox cboArmorType 
            Height          =   315
            ItemData        =   "frmItems.frx":1071
            Left            =   1080
            List            =   "frmItems.frx":1096
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1560
            Width           =   2415
         End
         Begin VB.TextBox txtID 
            Enabled         =   0   'False
            Height          =   330
            Left            =   1080
            TabIndex        =   3
            Top             =   0
            Width           =   855
         End
         Begin VB.TextBox txtItemName 
            Height          =   330
            Left            =   1080
            TabIndex        =   4
            Top             =   480
            Width           =   2415
         End
         Begin VB.ComboBox cboWorn 
            Height          =   315
            ItemData        =   "frmItems.frx":1140
            Left            =   1080
            List            =   "frmItems.frx":1142
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Line Line8 
            X1              =   5520
            X2              =   5520
            Y1              =   120
            Y2              =   2280
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   5400
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "ID:"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   60
            Top             =   0
            Width           =   225
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Item Name:"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   59
            Top             =   480
            Width           =   840
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Worn/Item:"
            Height          =   195
            Index           =   4
            Left            =   0
            TabIndex        =   58
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Weapon Type:"
            Height          =   195
            Index           =   10
            Left            =   0
            TabIndex        =   57
            Top             =   2040
            Width           =   1065
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Armor Type:"
            Height          =   195
            Index           =   17
            Left            =   0
            TabIndex        =   56
            Top             =   1560
            Width           =   900
         End
      End
      Begin VB.TextBox txtMsg1 
         Height          =   330
         Left            =   -74040
         TabIndex        =   30
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtMsg2 
         Height          =   330
         Left            =   -74040
         TabIndex        =   31
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtMsg3 
         Height          =   330
         Left            =   -74040
         TabIndex        =   32
         Top             =   1680
         Width           =   2775
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
         Height          =   2445
         Left            =   -73800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Top             =   3120
         Width           =   8295
      End
      Begin ServerEditor.Raise Raise1 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   78
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9128
         Style           =   2
         Color           =   0
      End
      Begin ServerEditor.UltraBox lstNo 
         Height          =   4575
         Left            =   -69240
         TabIndex        =   105
         Top             =   960
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   8070
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Weapons this ammo CAN NOT be used in:"
         Height          =   195
         Index           =   1
         Left            =   -69240
         TabIndex        =   107
         Top             =   720
         Width           =   3000
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "All availble projectile weapons:"
         Height          =   195
         Index           =   0
         Left            =   -74880
         TabIndex        =   106
         Top             =   720
         Width           =   2205
      End
      Begin VB.Label lblMethod 
         Height          =   375
         Left            =   -74760
         TabIndex        =   101
         Top             =   5160
         Width           =   9135
      End
      Begin VB.Line Line7 
         X1              =   -75000
         X2              =   -65520
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label lblOther 
         Caption         =   "EX"
         Height          =   975
         Left            =   -68280
         TabIndex        =   95
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label lblVictim 
         Caption         =   "EX"
         Height          =   975
         Left            =   -71160
         TabIndex        =   94
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label lblYou 
         Caption         =   "EX"
         Height          =   975
         Left            =   -74040
         TabIndex        =   93
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Other's View"
         Height          =   195
         Index           =   2
         Left            =   -68160
         TabIndex        =   92
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Victim's View"
         Height          =   195
         Index           =   1
         Left            =   -71040
         TabIndex        =   91
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Your View"
         Height          =   195
         Index           =   0
         Left            =   -73920
         TabIndex        =   90
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Scripting:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   80
         Top             =   600
         Width           =   675
      End
      Begin VB.Line Line4 
         X1              =   -70680
         X2              =   -70680
         Y1              =   480
         Y2              =   3050
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Messages:"
         Height          =   195
         Index           =   8
         Left            =   -74880
         TabIndex        =   54
         Top             =   480
         Width           =   765
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   195
         Index           =   11
         Left            =   -74880
         TabIndex        =   53
         Top             =   3120
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "(save)"
      Height          =   255
      Left            =   8040
      TabIndex        =   49
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "(new)"
      Height          =   255
      Left            =   9000
      TabIndex        =   52
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "< Previous"
      Height          =   255
      Left            =   10200
      TabIndex        =   51
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   255
      Left            =   11520
      TabIndex        =   50
      Top             =   6240
      Width           =   1215
   End
   Begin ServerEditor.Raise Raise3 
      Height          =   495
      Left            =   7920
      TabIndex        =   96
      Top             =   6120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   873
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise5 
      Height          =   5895
      Left            =   120
      TabIndex        =   102
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   10398
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise6 
      Height          =   5895
      Left            =   3000
      TabIndex        =   103
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   10398
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise4 
      Height          =   6735
      Left            =   0
      TabIndex        =   97
      Top             =   0
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   11880
      Style           =   4
      Color           =   0
   End
End
Attribute VB_Name = "frmItems"
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
Rem***************                frmItems                        **********************
Rem***************                ServerEditor                    **********************
Rem***************                Editor.vbp                      **********************
Rem*************************************************************************************
Rem*************************************************************************************
Dim lcID As Long
Dim bIs As Boolean

Private Sub cboFlags_Change()
cboFlags_Click
End Sub

Private Sub cboFlags_Click()
Dim s As String
If cboFlags.ListIndex <> -1 Then
    
    flgOpts.ViewStyle = modMain.DeterStyle(modMain.ShortFlag(cboFlags.list(cboFlags.ListIndex)), s)
    If flgOpts.ViewStyle = ComboInputFeed Then modMain.FeedAList flgOpts, s
    cmdAdd.Caption = "&Add"
End If
End Sub

Private Sub cboWeaponType_Change()
Select Case cboWeaponType.list(cboWeaponType.ListIndex)
    Case "3  = 2h Bows"
        lblUses.Caption = "Max Ammo:"
        lblUses.FontBold = True
        chkUseFlgs2.Enabled = False
        txtUses.Enabled = True
    Case Else
        lblUses.Caption = "Amount Of Uses:"
        lblUses.FontBold = False
        chkUseFlgs2.Enabled = True
        txtUses.Enabled = False
End Select
End Sub

Private Sub cboWeaponType_Click()
cboWeaponType_Change
End Sub

Private Sub cboWorn_Change()
cboWorn_Click
End Sub

Private Sub cboWorn_Click()
Dim i As Long
Dim j As Long
Select Case cboWorn.list(cboWorn.ListIndex)
    Case "weapon", "projectile"
        For i = 1 To 3
            Me.Controls("txtMsg" & i).Enabled = True
            Me.Controls("txtVic" & i).Enabled = True
            Me.Controls("txtOth" & i).Enabled = True
        Next
        cboArmorType.Enabled = False
        If cboWeaponType.list(cboWeaponType.ListIndex) = "3  = 2h Bows" Then txtUses.Enabled = False
        If cboWorn.list(cboWorn.ListIndex) = "projectile" Then
            SSTab.TabEnabled(5) = True
            SSTab.TabEnabled(1) = False
            SSTab.TabEnabled(3) = False
            SSTab.TabEnabled(4) = False
            cboWeaponType.Enabled = False
            txtSpeed.Enabled = False
            txtSD.Enabled = False
            chkSD.Enabled = False
            lblAC.Caption = "Mana Cost"
            lblAC.FontBold = True
            lblUses.Caption = "Supply Ammo:"
            lblUses.FontBold = True
        Else
            SSTab.TabEnabled(5) = False
            SSTab.TabEnabled(1) = True
            SSTab.TabEnabled(3) = True
            SSTab.TabEnabled(4) = False
            cboWeaponType.Enabled = True
            txtSpeed.Enabled = True
            txtSD.Enabled = True
            chkSD.Enabled = True
            lblAC.Caption = "Armor Class"
            lblAC.FontBold = False
            lblUses.Caption = "Amount Of Uses:"
            lblUses.FontBold = False
            For i = 1 To 3
                If Me.Controls("txtMsg" & i).Text = "" Then Me.Controls("txtMsg" & i).Text = "attack"
                If Me.Controls("txtVic" & i).Text = "" Then Me.Controls("txtVic" & i).Text = "attacks"
                If Me.Controls("txtOth" & i).Text = "" Then Me.Controls("txtOth" & i).Text = "attacks"
            Next
        End If
        txtAC.Enabled = True
        SSTab.TabEnabled(4) = False
        chkUseFlgs2.Enabled = False
        txtScripting.Enabled = False
        ssFlags.TabEnabled(1) = False
        lblLabel(2).Caption = "Min Damage:"
        lblLabel(3).Caption = "Max Damage:"
    Case "scroll", "item", "key"
        For i = 1 To 3
            If i <> 1 Then Me.Controls("txtMsg" & i).Enabled = False
            Me.Controls("txtVic" & i).Enabled = False
            Me.Controls("txtOth" & i).Enabled = False
            If i <> 1 Then Me.Controls("txtMsg" & i).Text = ""
            Me.Controls("txtVic" & i).Text = ""
            Me.Controls("txtOth" & i).Text = ""
        Next
        If cboWorn.list(cboWorn.ListIndex) = "item" Then
            SSTab.TabEnabled(4) = True
        Else
            SSTab.TabEnabled(5) = False
            SSTab.TabEnabled(4) = False
            SSTab.TabEnabled(1) = True
            SSTab.TabEnabled(3) = True
        End If
        txtMsg1.Enabled = True
        cboArmorType.Enabled = False
        cboWeaponType.Enabled = False
        txtUses.Enabled = True
        chkUseFlgs2.Enabled = True
        txtScripting.Enabled = True
        ssFlags.TabEnabled(1) = True
        txtSD.Enabled = False
        chkSD.Enabled = False
        lblLabel(2).Caption = "Min Damage:"
        lblLabel(3).Caption = "Max Damage:"
    Case "food", "ofood", "corpse"
        For i = 1 To 3
            If i <> 1 Then Me.Controls("txtMsg" & i).Enabled = False
            Me.Controls("txtVic" & i).Enabled = False
            Me.Controls("txtOth" & i).Enabled = False
            If i <> 1 Then Me.Controls("txtMsg" & i).Text = ""
            Me.Controls("txtVic" & i).Text = ""
            Me.Controls("txtOth" & i).Text = ""
        Next
        SSTab.TabEnabled(5) = False
        SSTab.TabEnabled(4) = False
        SSTab.TabEnabled(1) = True
        SSTab.TabEnabled(3) = True
        txtMsg1.Enabled = True
        cboArmorType.Enabled = False
        cboWeaponType.Enabled = False
        txtUses.Enabled = True
        chkUseFlgs2.Enabled = True
        txtScripting.Enabled = True
        ssFlags.TabEnabled(1) = True
        txtSD.Enabled = False
        chkSD.Enabled = False
        lblLabel(2).Caption = "Min Value:"
        lblLabel(3).Caption = "Max Value:"
    Case Else
        For i = 1 To 3
            Me.Controls("txtMsg" & i).Enabled = False
            Me.Controls("txtVic" & i).Enabled = False
            Me.Controls("txtOth" & i).Enabled = False
            Me.Controls("txtMsg" & i).Text = ""
            Me.Controls("txtVic" & i).Text = ""
            Me.Controls("txtOth" & i).Text = ""
        Next
        SSTab.TabEnabled(5) = False
        SSTab.TabEnabled(4) = False
        SSTab.TabEnabled(1) = True
        SSTab.TabEnabled(3) = True
        cboArmorType.Enabled = True
        cboWeaponType.Enabled = False
        txtUses.Enabled = False
        chkUseFlgs2.Enabled = False
        txtScripting.Enabled = False
        ssFlags.TabEnabled(1) = False
        txtSD.Enabled = True
        chkSD.Enabled = True
        lblLabel(2).Caption = "Min Damage:"
        lblLabel(3).Caption = "Max Damage:"
        SSTab.TabEnabled(5) = False
End Select
End Sub



Private Sub cmdAdd_Click()
If cmdAdd.Caption = "&Add" Then
    If ssFlags.Tab = 0 Then
        lstFlags1.ListItems.Add Text:=modMain.TranslateFlag(modMain.ShortFlag( _
            cboFlags.list(cboFlags.ListIndex)) & flgOpts.GetCurVal)
    Else
        lstFlags2.ListItems.Add Text:=modMain.TranslateFlag(modMain.ShortFlag( _
            cboFlags.list(cboFlags.ListIndex)) & flgOpts.GetCurVal)
    End If
Else
    If ssFlags.Tab = 0 Then
        lstFlags1.SelectedItem.Text = modMain.TranslateFlag(modMain.ShortFlag( _
        cboFlags.list(cboFlags.ListIndex)) & flgOpts.GetCurVal)
    Else
        lstFlags2.SelectedItem.Text = modMain.TranslateFlag(modMain.ShortFlag( _
        cboFlags.list(cboFlags.ListIndex)) & flgOpts.GetCurVal)
    End If
    cmdAdd.Caption = "&Add"
End If
End Sub

Private Sub cmdAddAmmo_Click()
If lstBows.ListIndex <> -1 Then lstNo.AddItem lstBows.ItemText
End Sub

Private Sub cmdAddClass_Click()
lstClassRes.AddItem lstClasses.ItemText
End Sub

Private Sub cmdAddNew_Click()
MousePointer = vbHourglass
Dim x As Long
Dim i As Long
Dim t As Boolean
ReDim Preserve dbItems(1 To UBound(dbItems) + 1)
x = dbItems(UBound(dbItems) - 1).iID
x = x + 1
Do Until t = True
    t = True
    i = GetItemID(, x)
    If i <> 0 Then
        t = False
        x = x + 1
    End If
Loop
With dbItems(UBound(dbItems))
    .dClassPoints = 0
    .dCost = 0
    .iAC = 0
    .iArmorType = 0
    .iID = x
    .iInGame = 0
    .iIsLedgenary = 0
    .iLimit = 0
    .iMagical = 0
    .iMoveable = 1
    .iOnEquipKillDur = 0
    .iSpeed = 5
    .iType = 0
    .iUses = 1
    .lDurability = 99
    .lLevel = 0
    .lOnLastUseDoFlags2 = 0
    .sClassRestriction = "0"
    .sDamage = "0:0"
    .sDesc = "New Item"
    .sFlags = "0"
    .sFlags2 = "0"
    .sItemName = "New Item"
    .sMessage2 = "0"
    .sMessageV = "0"
    .sRaceRestriction = "0"
    .sScript = "0"
    .sSwings = "none"
    .sWorn = "item"
    .sProjectile = "0"
End With
lcID = UBound(dbItems)
FillItems lcID, True
MousePointer = vbDefault
End Sub

Private Sub cmdAddRace_Click()
lstRaceRes.AddItem lstRaces.ItemText
End Sub

Private Sub cmdClassClear_Click()
lstClassRes.RemoveItem lstClassRes.ListIndex
End Sub

Private Sub cmdRaceClear_Click()
lstRaceRes.RemoveItem lstRaceRes.ListIndex
End Sub

Private Sub cmdNext_Click()
On Error GoTo cmdNext_Click_Error
SaveItems
lcID = lcID + 1
If lcID > UBound(dbItems) Then lcID = 1
FillItems lcID
On Error GoTo 0
Exit Sub
cmdNext_Click_Error:
End Sub


Private Sub cmdPrevious_Click()
On Error GoTo cmdPrevious_Click_Error
SaveItems
lcID = lcID - 1
If lcID < LBound(dbItems) Then lcID = UBound(dbItems)
FillItems lcID
On Error GoTo 0
Exit Sub
cmdPrevious_Click_Error:
End Sub


Private Sub cmdRemove_Click()
If ssFlags.Tab = 0 Then
    If lstFlags1.ListItems.Count > 0 Then lstFlags1.ListItems.Remove lstFlags1.SelectedItem.Index
Else
    If lstFlags2.ListItems.Count > 0 Then lstFlags2.ListItems.Remove lstFlags2.SelectedItem.Index
End If
cmdAdd.Caption = "&Add"
End Sub

Private Sub cmdRemoveAmmo_Click()
If lstNo.ListIndex > 0 Then lstNo.RemoveItem lstNo.ListIndex
End Sub

Private Sub cmdSave_Click()
SaveItems
End Sub


Private Sub Form_Load()
txtScripting.IntelliSenseAddWordsFile App.Path & "\scriptdef.aimg"
txtScripting.IntelliSenseStartSubclassing
lstFlags1.ColumnHeaders(1).Width = 10000
lstFlags2.ColumnHeaders(1).Width = 10000
FillCBOS
FillItems , True
End Sub

Private Sub lstClasses_DoubleClick()
cmdAddClass_Click
End Sub

Private Sub lstClassRes_DoubleClick()
cmdClassClear_Click
End Sub

Private Sub lstFlags1_Click()
Dim s As String
Dim t As String
If lstFlags1.ListItems.Count < 1 Then Exit Sub
s = Left$(lstFlags1.SelectedItem.Text, 25)
s = Trim$(s)
For i = 0 To cboFlags.ListCount - 1
    If cboFlags.list(i) = s Then
        cboFlags.ListIndex = i
        Exit For
    End If
Next
t = Mid$(lstFlags1.SelectedItem.Text, 28)
t = Trim$(t)
flgOpts.SetVal Val(t)
cmdAdd.Caption = "&Modify"
End Sub

Private Sub lstFlags2_Click()
Dim s As String
Dim t As String
If lstFlags2.ListItems.Count < 1 Then Exit Sub
s = Left$(lstFlags2.SelectedItem.Text, 25)
s = Trim$(s)
For i = 0 To cboFlags.ListCount - 1
    If cboFlags.list(i) = s Then
        cboFlags.ListIndex = i
        Exit For
    End If
Next
t = Mid$(lstFlags2.SelectedItem.Text, 28)
t = Trim$(t)
flgOpts.SetVal Val(t)
cmdAdd.Caption = "&Modify"
End Sub

Private Sub lstItems_Click()
If bIs Then Exit Sub
Dim i As Long
MousePointer = vbHourglass
For i = LBound(dbItems) To UBound(dbItems)
    With dbItems(i)
        If .iID & " " & .sItemName = lstItems.ItemText Then
            FillItems i
            Exit For
        End If
    End With
Next
MousePointer = vbDefault
End Sub

Sub FillCBOS()
cboWorn.Clear
With cboWorn
    .AddItem "item"
    .AddItem "scroll"
    .AddItem "projectile"
    .AddItem "key"
    .AddItem "food"
    .AddItem "ofood"
    .AddItem "corpse"
    .AddItem "weapon"
    .AddItem "head"
    .AddItem "neck"
    .AddItem "ears"
    .AddItem "back"
    .AddItem "face"
    .AddItem "body"
    .AddItem "hands"
    .AddItem "legs"
    .AddItem "feet"
    .AddItem "waist"
    .AddItem "arms"
    .AddItem "shield"
    .AddItem "ring"
End With
lstClasses.Clear
For i = LBound(dbClass) To UBound(dbClass)
    With dbClass(i)
        lstClasses.AddItem .sName
    End With
Next
lstRaces.Clear
For i = LBound(dbRaces) To UBound(dbRaces)
    With dbRaces(i)
        lstRaces.AddItem .sName
    End With
Next
lstBows.Clear
For i = LBound(dbItems) To UBound(dbItems)
    With dbItems(i)
        If .iType = 3 Then
            lstBows.AddItem "(" & .iID & ") " & .sItemName
        End If
    End With
Next
End Sub


Sub SaveItems()
Dim s As String, i As Long
MousePointer = vbHourglass
ReverseEffects lcID, Item
With dbItems(lcID)
    .dClassPoints = Val(txtClassPoints.Text)
    .dCost = Val(txtCost.Text)
    .iAC = Val(txtAC.Text)
    .iArmorType = Val(Trim$(Replace$(Left$(cboArmorType.list(cboArmorType.ListIndex), 2), ",", "")))
    .iIsLedgenary = chkLedgendary.Value
    .iLimit = Val(txtLimit.Text)
    .iMagical = Val(txtMagicLevel.Text)
    .iMoveable = Val(chkMoveable.Value)
    .iOnEquipKillDur = Val(txtSD.Text)
    .iSpeed = Val(txtSpeed.Text)
    .iType = Trim$(Left$(cboWeaponType.list(cboWeaponType.ListIndex), 2))
    .iUses = Val(txtUses.Text)
    .lDurability = Val(txtDur.Text)
    .lLevel = Val(Me.txtLevelRestriction.Text)
    .lOnLastUseDoFlags2 = Val(Me.chkUseFlgs2.Value)
    .sClassRestriction = ""
    If lstClassRes.ListCount > 0 Then
        For i = 1 To lstClassRes.ListCount
            .sClassRestriction = .sClassRestriction & lstClassRes.list(i)
        Next
    Else
        .sClassRestriction = "0"
    End If
    .sDamage = txtMinDamage.Text & ":" & txtMaxDamage.Text
    .sDesc = txtDescription.Text
    .sFlags = ""
    If lstFlags1.ListItems.Count > 0 Then
        For i = 1 To lstFlags1.ListItems.Count
            .sFlags = .sFlags & modMain.MakeDBFlag(lstFlags1.ListItems(i).Text) & ";"
        Next
    Else
        .sFlags = "0"
    End If
    .sFlags2 = ""
    If lstFlags2.ListItems.Count > 0 Then
        For i = 1 To lstFlags2.ListItems.Count
            .sFlags2 = .sFlags2 & modMain.MakeDBFlag(lstFlags2.ListItems(i).Text) & ";"
        Next
    Else
        .sFlags2 = "0"
    End If
    .sItemName = txtItemName.Text
    .sMessage2 = txtOth1.Text & ":" & txtOth2.Text & ":" & txtOth3.Text
    If .sMessage2 = "0::" Then .sMessage2 = "0"
    If .sMessage2 = "::" Then .sMessage2 = "0"
    .sMessageV = txtVic1.Text & ":" & txtVic2.Text & ":" & txtVic3.Text
    If .sMessageV = "0::" Then .sMessageV = "0"
    If .sMessageV = "::" Then .sMessageV = "0"
    .sRaceRestriction = ""
    If lstRaceRes.ListCount > 0 Then
        For i = 1 To lstRaceRes.ListCount
            .sRaceRestriction = .sRaceRestriction & lstRaceRes.list(i)
        Next
    Else
        .sRaceRestriction = "0"
    End If
    .sProjectile = ""
    If lstNo.ListCount > 0 Then
        For i = 1 To lstNo.ListCount
            txtNo.Text = lstNo.list(i)
            .sProjectile = .sProjectile & txtNo.Text & ";"
        Next
    Else
        .sProjectile = "0"
    End If
    .sScript = txtScripting.Text
    If .sScript = "" Then .sScript = "0"
    .sSwings = txtMsg1.Text & ":" & txtMsg2.Text & ":" & txtMsg3.Text
    If .sSwings = "0::" Then .sSwings = "0"
    If Right$(.sSwings, 2) = "::" Then .sSwings = Left$(.sSwings, Len(.sSwings) - 2)
    If .sSwings = "::" Or .sSwings = "" Then .sSwings = "0"
    .sWorn = cboWorn.list(cboWorn.ListIndex)
End With
modUpdateDatabase.SaveMemoryToDatabase Item
DoEffects lcID, Item
modUpdateDatabase.SaveMemoryToDatabase Players
FillItems lcID, True
MousePointer = vbDefault
End Sub

Sub FillItems(Optional Arg As Long = -1, Optional FillList As Boolean = False)
Dim i As Long, j As Long
Dim m As Long
Dim Arr() As String
Dim lCol As Long
Dim tCol As Long
MousePointer = vbHourglass
bIs = True
If Arg = -1 Then Arg = LBound(dbItems)
If FillList Then lstItems.Paint = False: lstItems.Clear
For i = LBound(dbItems) To UBound(dbItems)
    With dbItems(i)
        If FillList Then
            If lCol = 0 Then
                lstItems.AddItem CStr(.iID & " " & .sItemName), BCOLOR:=vbWhite
                lCol = 1
            Else
                lstItems.AddItem CStr(.iID & " " & .sItemName), BCOLOR:=&HEFEFEF
                lCol = 0
            End If
        End If
        If i = Arg Then
            lcID = Arg
            txtID.Text = .iID
            txtItemName.Text = .sItemName
            txtMinDamage.Text = Mid$(.sDamage, 1, InStr(1, .sDamage, ":") - 1)
            txtMaxDamage.Text = Mid$(.sDamage, InStr(1, .sDamage, ":") + 1, Len(.sDamage) - InStr(1, .sDamage, ":"))
            txtLimit.Text = .iLimit
            chkMoveable.Value = .iMoveable
            chkLedgendary.Value = .iIsLedgenary
            txtClassPoints.Text = .dClassPoints
            txtSD.Text = .iOnEquipKillDur
            If .iOnEquipKillDur <> 0 Then
                chkSD.Value = 1
                txtSD.Enabled = True
            Else
                chkSD.Value = 0
                txtSD.Enabled = False
            End If
            chkUseFlgs2.Value = .lOnLastUseDoFlags2
            txtAC.Text = .iAC
            If .sSwings <> "none" Then
                SplitFast .sSwings, Arr, ":"
                If UBound(Arr) < 1 Then
                    txtMsg1.Text = Arr(0)
                    For j = 2 To 3
                        Me.Controls("txtMsg" & j).Enabled = False
                    Next
                Else
                    For j = 1 To 3
                        Me.Controls("txtMsg" & j).Enabled = True
                        Me.Controls("txtMsg" & j).Text = Arr(j - 1)
                    Next
                End If
            Else
                For j = 1 To 3
                    Me.Controls("txtMsg" & j).Enabled = False
                Next
                txtMsg1.Text = "none"
            End If
            Erase Arr
            If .sMessageV <> "0" Then
                SplitFast .sMessageV, Arr, ":"
                For j = 1 To 3
                    Me.Controls("txtVic" & j).Enabled = True
                    Me.Controls("txtVic" & j).Text = Arr(j - 1)
                Next
            Else
                For j = 1 To 3
                    Me.Controls("txtVic" & j).Enabled = False
                Next
                txtVic1.Text = "0"
            End If
            Erase Arr
            If .sMessage2 <> "0" Then
                SplitFast .sMessage2, Arr, ":"
                For j = 1 To 3
                    Me.Controls("txtOth" & j).Enabled = True
                    Me.Controls("txtOth" & j).Text = Arr(j - 1)
                Next
            Else
                For j = 1 To 3
                    Me.Controls("txtOth" & j).Enabled = False
                Next
                txtOth1.Text = "0"
            End If
            txtSpeed.Text = .iSpeed
            modMain.SetCBOlstIndex cboWeaponType, .iType, [Weapon Type]
            modMain.SetListIndex cboWorn, .sWorn
            txtDescription.Text = .sDesc
            txtCost.Text = .dCost
            txtLevelRestriction.Text = .lLevel
            modMain.SetCBOlstIndex cboArmorType, .iArmorType, [Armor Type]
            If cboWorn.list(cboWorn.ListIndex) = "projectile" Then
                SSTab.TabEnabled(5) = True
                SSTab.TabEnabled(1) = False
                SSTab.TabEnabled(3) = False
                cboWeaponType.Enabled = False
                txtSpeed.Enabled = False
                txtSD.Enabled = False
                chkSD.Enabled = False
                lblAC.Caption = "Mana Cost"
                lblAC.FontBold = True
                lblUses.Caption = "Supply Ammo:"
                lblUses.FontBold = True
            ElseIf .iType = 3 Then
                lblUses.Caption = "Max Ammo:"
                lblUses.FontBold = True
                chkUseFlgs2.Enabled = False
                txtUses.Enabled = True
            ElseIf cboWorn.list(cboWorn.ListIndex) = "weapon" Then
                SSTab.TabEnabled(5) = False
                SSTab.TabEnabled(1) = True
                SSTab.TabEnabled(3) = True
                cboWeaponType.Enabled = True
                txtSpeed.Enabled = True
                txtSD.Enabled = True
                chkSD.Enabled = True
                lblAC.Caption = "Armor Class"
                lblAC.FontBold = False
                lblUses.Caption = "Amount Of Uses:"
                lblUses.FontBold = False
            End If
            Erase Arr
            lstRaceRes.Clear
            If .sRaceRestriction <> "0" Then
                SplitFast .sRaceRestriction, Arr, ";"
                For j = LBound(Arr) To UBound(Arr)
                    If Arr(j) <> "" Then
                        If tCol = 0 Then
                            lstRaceRes.AddItem Arr(j), BCOLOR:=vbWhite
                            tCol = 1
                        Else
                            lstRaceRes.AddItem Arr(j), BCOLOR:=&HEFEFEF
                            tCol = 0
                        End If
                    End If
                Next
            End If
            Erase Arr
            lstClassRes.Clear
            tCol = 0
            If .sClassRestriction <> "0" Then
                SplitFast .sClassRestriction, Arr, ";"
                For j = LBound(Arr) To UBound(Arr)
                    If Arr(j) <> "" Then
                        If tCol = 0 Then
                            lstRaceRes.AddItem Arr(j), BCOLOR:=vbWhite
                            tCol = 1
                        Else
                            lstRaceRes.AddItem Arr(j), BCOLOR:=&HEFEFEF
                            tCol = 0
                        End If
                    End If
                Next
            End If
            txtMagicLevel.Text = .iMagical
            txtScripting.Text = .sScript
            txtDur.Text = .lDurability
            txtUses.Text = .iUses
            Erase Arr
            lstFlags1.ListItems.Clear
            If .sFlags <> "0" Then
                SplitFast .sFlags, Arr, ";"
                For j = LBound(Arr) To UBound(Arr)
                    If Arr(j) <> "" And Arr(j) <> "0" Then
                        lstFlags1.ListItems.Add Text:=modMain.TranslateFlag(Arr(j))
                    End If
                Next
            End If
            Erase Arr
            lstFlags2.ListItems.Clear
            If .sFlags2 <> "0" Then
                SplitFast .sFlags2, Arr, ";"
                For j = LBound(Arr) To UBound(Arr)
                    If Arr(j) <> "" And Arr(j) <> "0" Then
                        lstFlags2.ListItems.Add Text:=modMain.TranslateFlag(Arr(j))
                    End If
                Next
            End If
            Erase Arr
            SplitFast .sProjectile, Arr, ";"
            For j = LBound(Arr) To UBound(Arr)
                If Arr(j) <> "0" And Arr(j) <> "" Then
                    lstNo.AddItem "(" & Arr(j) & ") " & dbItems(GetItemID(, Val(Arr(j)))).sItemName
                End If
            Next
            If txtScripting.Text = "0" Then txtScripting.Text = ""
            lstBows.Clear
            For j = LBound(dbItems) To UBound(dbItems)
                With dbItems(j)
                    If .iType = 3 Then
                        lstBows.AddItem "(" & .iID & ") " & .sItemName
                    End If
                End With
            Next
            If Not FillList Then Exit For
        End If
    End With
Next
lstItems.Paint = True
modMain.SetLstSelected lstItems, txtID.Text & " " & txtItemName.Text
bIs = False
MousePointer = vbDefault
End Sub


Private Sub lstRaceRes_DoubleClick()
cmdRaceClear_Click
End Sub

Private Sub lstRaces_DoubleClick()
cmdAddRace_Click
End Sub

Private Sub ssFlags_Click(PreviousTab As Integer)
cmdAdd.Caption = "&Add"
End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
cmdAdd.Caption = "&Add"
End Sub

Private Sub ssTab_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'fraAddStat.Visible = False
End Sub

Private Sub txtFind_Change()
lstItems.SetSelected lstItems.FindInStr(txtFind.Text), True, True
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

Private Sub txtMsg1_Change()
lblYou.Caption = "You " & txtMsg1.Text & " monster!"
End Sub

Private Sub txtOth1_Change()
lblOther.Caption = "Attacker " & txtOth1.Text & " victim!"
End Sub

Private Sub txtScripting_MethodHasChanged()
lblMethod.Caption = txtScripting.Params
End Sub

Private Sub txtVic1_Change()
lblVictim.Caption = "Attacker " & txtVic1.Text & " you!"
End Sub
