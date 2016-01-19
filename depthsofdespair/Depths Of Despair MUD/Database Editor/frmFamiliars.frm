VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFamiliars 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Familiars"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   12030
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
   ScaleHeight     =   6615
   ScaleWidth      =   12030
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   240
      TabIndex        =   64
      Top             =   240
      Width           =   3135
   End
   Begin ServerEditor.UltraBox lstFamiliars 
      Height          =   5175
      Left            =   240
      TabIndex        =   63
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   9128
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
   Begin TabDlg.SSTab ss1 
      Height          =   5535
      Left            =   3720
      TabIndex        =   4
      Top             =   240
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
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
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmFamiliars.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLabel(11)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabel(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabel(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtDescription"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtName"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtID"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Raise1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Picture1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtEXP"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkRide"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cboSpeed"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Combat"
      TabPicture(1)   =   "frmFamiliars.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblLabel(5)"
      Tab(1).Control(1)=   "lblLabel(6)"
      Tab(1).Control(2)=   "lblLabel(7)"
      Tab(1).Control(3)=   "lblDamage"
      Tab(1).Control(4)=   "lblLabel(20)"
      Tab(1).Control(5)=   "lblLabel(21)"
      Tab(1).Control(6)=   "txtSwings"
      Tab(1).Control(7)=   "txtLM"
      Tab(1).Control(8)=   "txtMod"
      Tab(1).Control(9)=   "txtMax"
      Tab(1).Control(10)=   "Raise4"
      Tab(1).Control(11)=   "Picture2"
      Tab(1).Control(12)=   "txtMin"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Flags"
      TabPicture(2)   =   "frmFamiliars.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "flgOpts"
      Tab(2).Control(1)=   "cboFlags"
      Tab(2).Control(2)=   "cmdAdd"
      Tab(2).Control(3)=   "cmdRemove"
      Tab(2).Control(4)=   "lstFamFlags"
      Tab(2).Control(5)=   "lblLabel(23)"
      Tab(2).Control(6)=   "lblLabel(22)"
      Tab(2).Control(7)=   "lblHelp"
      Tab(2).ControlCount=   8
      Begin ServerEditor.FlagOptions flgOpts 
         Height          =   375
         Left            =   -70800
         TabIndex        =   61
         Top             =   960
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Style           =   0
      End
      Begin VB.ComboBox cboFlags 
         Height          =   315
         ItemData        =   "frmFamiliars.frx":0054
         Left            =   -74880
         List            =   "frmFamiliars.frx":007C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   960
         Width           =   3975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   255
         Left            =   -72600
         TabIndex        =   54
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   -69360
         TabIndex        =   53
         Top             =   1440
         Width           =   2295
      End
      Begin ServerEditor.NumOnlyText txtMin 
         Height          =   255
         Left            =   -72720
         TabIndex        =   45
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   -74760
         ScaleHeight     =   3135
         ScaleWidth      =   7335
         TabIndex        =   29
         Top             =   2160
         Width           =   7335
         Begin VB.TextBox txtM2 
            Height          =   285
            Left            =   1920
            TabIndex        =   44
            Top             =   2730
            Width           =   5295
         End
         Begin VB.TextBox txtM1 
            Height          =   285
            Left            =   1920
            TabIndex        =   42
            Top             =   2400
            Width           =   5295
         End
         Begin VB.TextBox txtA2 
            Height          =   285
            Left            =   1920
            TabIndex        =   38
            Top             =   1800
            Width           =   5295
         End
         Begin VB.TextBox txtA1 
            Height          =   285
            Left            =   1920
            TabIndex        =   36
            Top             =   1470
            Width           =   5295
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "What others would see:"
            Height          =   195
            Index           =   19
            Left            =   0
            TabIndex        =   43
            Top             =   2730
            Width           =   1725
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "What you would see:"
            Height          =   195
            Index           =   18
            Left            =   0
            TabIndex        =   41
            Top             =   2400
            Width           =   1530
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Miss Messages:"
            Height          =   195
            Index           =   17
            Left            =   0
            TabIndex        =   40
            Top             =   2160
            Width           =   1110
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Hit Messages:"
            Height          =   195
            Index           =   16
            Left            =   0
            TabIndex        =   39
            Top             =   1225
            Width           =   1005
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "What others would see:"
            Height          =   195
            Index           =   15
            Left            =   0
            TabIndex        =   37
            Top             =   1800
            Width           =   1725
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   7320
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "was attacking a giant rat, and the damage the familiar hit for was 12."
            Height          =   195
            Index           =   14
            Left            =   360
            TabIndex        =   35
            Top             =   960
            Width           =   4995
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   """Your Cat attacks giant rat, and hits for 12 damage!"", assuming your familiar's name was Cat,"
            Height          =   195
            Index           =   13
            Left            =   360
            TabIndex        =   34
            Top             =   720
            Width           =   6735
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "EX: ""<%n> attacks <%m>, and hits for <%d> damage!"" would output:"
            Height          =   195
            Index           =   12
            Left            =   0
            TabIndex        =   33
            Top             =   480
            Width           =   5235
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "<%n> = Your Name && Familiars Name, <%m> = Targets Name, <%d> = Damage Amount."
            Height          =   195
            Index           =   10
            Left            =   0
            TabIndex        =   32
            Top             =   240
            Width           =   6645
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "What you would see:"
            Height          =   195
            Index           =   9
            Left            =   0
            TabIndex        =   31
            Top             =   1470
            Width           =   1530
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Messages:"
            Height          =   195
            Index           =   8
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   765
         End
      End
      Begin ServerEditor.Raise Raise4 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   28
         Top             =   2040
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   5953
         Style           =   2
         Color           =   0
      End
      Begin VB.ComboBox cboSpeed 
         Height          =   315
         ItemData        =   "frmFamiliars.frx":0119
         Left            =   1920
         List            =   "frmFamiliars.frx":0141
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CheckBox chkRide 
         Height          =   200
         Left            =   1920
         TabIndex        =   19
         Top             =   2760
         Width           =   200
      End
      Begin ServerEditor.NumOnlyText txtEXP 
         Height          =   255
         Left            =   1920
         TabIndex        =   18
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
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
      Begin VB.PictureBox Picture1 
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
         Height          =   975
         Left            =   480
         ScaleHeight     =   975
         ScaleWidth      =   4815
         TabIndex        =   12
         Top             =   3600
         Width           =   4815
         Begin ServerEditor.NumOnlyText txtMinHP 
            Height          =   255
            Left            =   1680
            TabIndex        =   15
            Top             =   120
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
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
         Begin ServerEditor.NumOnlyText txtMaxHP 
            Height          =   255
            Left            =   1680
            TabIndex        =   16
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Starting HP Max Roll:"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   14
            Top             =   480
            Width           =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Starting HP Min Roll:"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   120
            Width           =   1455
         End
      End
      Begin ServerEditor.Raise Raise1 
         Height          =   1215
         Left            =   360
         TabIndex        =   11
         Top             =   3480
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2143
         Style           =   2
         Color           =   0
      End
      Begin VB.TextBox txtID 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtDescription 
         Height          =   885
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1440
         Width           =   3495
      End
      Begin ServerEditor.NumOnlyText txtMax 
         Height          =   255
         Left            =   -72720
         TabIndex        =   46
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      Begin ServerEditor.NumOnlyText txtMod 
         Height          =   255
         Left            =   -72720
         TabIndex        =   47
         Top             =   1200
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      Begin ServerEditor.NumOnlyText txtLM 
         Height          =   255
         Left            =   -69960
         TabIndex        =   49
         Top             =   1200
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      Begin ServerEditor.NumOnlyText txtSwings 
         Height          =   255
         Left            =   -71880
         TabIndex        =   52
         Top             =   1680
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      Begin MSComctlLib.ListView lstFamFlags 
         Height          =   3615
         Left            =   -72600
         TabIndex        =   62
         Top             =   1800
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6376
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
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   195
         Index           =   23
         Left            =   -70800
         TabIndex        =   58
         Top             =   720
         Width           =   450
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Flag:"
         Height          =   195
         Index           =   22
         Left            =   -74880
         TabIndex        =   57
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblHelp 
         Caption         =   "0"
         Height          =   3495
         Left            =   -74760
         TabIndex        =   55
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Amount of Swings Per Combat Round:"
         Height          =   195
         Index           =   21
         Left            =   -74760
         TabIndex        =   51
         Top             =   1680
         Width           =   2745
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Level Damge Mod. Max:"
         Height          =   195
         Index           =   20
         Left            =   -71880
         TabIndex        =   50
         Top             =   1200
         Width           =   1725
      End
      Begin VB.Label lblDamage 
         Caption         =   "Assuming the familiar is level 15, the familiars damage would range from <%min> to <%max>."
         Height          =   975
         Left            =   -71880
         TabIndex        =   48
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Level Damage Modification:"
         Height          =   195
         Index           =   7
         Left            =   -74760
         TabIndex        =   27
         Top             =   1200
         Width           =   1965
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Base Max Damage:"
         Height          =   195
         Index           =   6
         Left            =   -74760
         TabIndex        =   26
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Base Min Damage:"
         Height          =   195
         Index           =   5
         Left            =   -74760
         TabIndex        =   25
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Speed:"
         Height          =   195
         Index           =   4
         Left            =   1320
         TabIndex        =   22
         Top             =   3000
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ridable:"
         Height          =   195
         Index           =   3
         Left            =   1275
         TabIndex        =   20
         Top             =   2760
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "EXP Needed Per Level:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Width           =   1635
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   225
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   9
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   195
         Index           =   11
         Left            =   960
         TabIndex        =   8
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "(save)"
      Height          =   255
      Left            =   6960
      TabIndex        =   3
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "(new)"
      Height          =   255
      Left            =   7920
      TabIndex        =   2
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "< Previous"
      Height          =   255
      Left            =   9240
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   255
      Left            =   10440
      TabIndex        =   0
      Top             =   6120
      Width           =   1095
   End
   Begin ServerEditor.Raise Raise2 
      Height          =   5775
      Left            =   3600
      TabIndex        =   23
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   10186
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise3 
      Height          =   5775
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   10186
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise5 
      Height          =   495
      Left            =   6840
      TabIndex        =   59
      Top             =   6000
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   873
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise6 
      Height          =   6615
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   11668
      Style           =   4
      Color           =   0
   End
End
Attribute VB_Name = "frmFamiliars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
lstFamiliars.SetSelected lstFamiliars.FindInStr(txtFind.Text), True, True
End Sub

Sub SaveFams()
Dim s As String
MousePointer = vbHourglass
With dbFamiliars(lcID)
    .dEXPPerLevel = Val(txtEXP.Text)
    .lLevelMax = Val(txtLM.Text)
    .lLevelMod = Val(txtMod.Text)
    .lMaxDam = Val(txtMax.Text)
    .lMinDam = Val(txtMin.Text)
    .lRidable = Val(chkRide.Value)
    .lSpeed = cboSpeed.ListIndex
    .lStartHPMax = Val(txtMaxHP.Text)
    .lStartHPMin = Val(txtMinHP.Text)
    .lSwings = Val(txtSwings.Text)
    .sAttackMessage = txtA1.Text
    .sDescription = txtDescription.Text
    .sFamName = txtName.Text
    .sMessage2 = txtA2.Text
    .sMissMessage = txtM1.Text
    .sMissMessage2 = txtM2.Text
    For i = 1 To lstFamFlags.ListItems.Count
        s = s & modMain.MakeDBFlag(lstFamFlags.ListItems(i).Text) & ";"
    Next
    .sFlags = s
End With
modUpdateDatabase.SaveMemoryToDatabase Familiars
FillFams lcID, True
MousePointer = vbDefault
End Sub

Sub FillFams(Optional Arg As Long = -1, Optional FillList As Boolean = False)
Dim i As Long, j As Long
Dim m As Long
Dim Arr() As String
MousePointer = vbHourglass
bIs = True
If Arg = -1 Then Arg = LBound(dbFamiliars)
If FillList Then lstFamiliars.Clear
For i = LBound(dbFamiliars) To UBound(dbFamiliars)
    With dbFamiliars(i)
        If FillList Then lstFamiliars.AddItem CStr(.iID & " " & .sFamName)
        If i = Arg Then
            lstFamFlags.ListItems.Clear
            lcID = i
            txtID.Text = .iID
            txtName.Text = .sFamName
            txtDescription.Text = .sDescription
            txtEXP.Text = .dEXPPerLevel
            txtSwings.Text = .lSwings
            chkRide.Value = .lRidable
            Select Case chkRide.Value
                Case 0
                    cboSpeed.Enabled = False
                Case Else
                    cboSpeed.Enabled = True
            End Select
            modMain.SetCBOSelectByID cboSpeed, CStr(.lSpeed)
            Arr = Split(.sFlags, ";")
            For j = LBound(Arr) To UBound(Arr)
                If Arr(j) <> "" And Arr(j) <> "0" Then
                    lstFamFlags.ListItems.Add Text:=modMain.TranslateFlag(Arr(j))
                End If
            Next
            txtMinHP.Text = .lStartHPMin
            txtMaxHP.Text = .lStartHPMax
            txtMin.Text = .lMinDam
            txtMax.Text = .lMaxDam
            txtMod.Text = .lLevelMod
            txtLM.Text = .lLevelMax
            txtA1.Text = .sAttackMessage
            txtA2.Text = .sMessage2
            txtM1.Text = .sMissMessage
            txtM2.Text = .sMissMessage2
            If Not FillList Then Exit For
        End If
    End With
Next
modMain.SetLstSelected lstFamiliars, txtID.Text & " " & txtName.Text
bIs = False
MousePointer = vbDefault
End Sub

Private Sub cboFlags_Change()
If cboFlags.ListIndex <> -1 Then
    lblHelp.Caption = GetHelp(cboFlags.list(cboFlags.ListIndex))
    cmdAdd.Caption = "Add"
End If
End Sub

Private Sub cboFlags_Click()
cboFlags_Change
End Sub

Private Sub chkRide_Click()
Select Case chkRide.Value
    Case 0
        cboSpeed.Enabled = False
    Case Else
        cboSpeed.Enabled = True
End Select
End Sub

Private Sub cmdAdd_Click()
If cmdAdd.Caption = "Modify" Then
    If cboFlags.ListIndex <> -1 Then _
        lstFamFlags.SelectedItem.Text = modMain.TranslateFlag(modMain.ShortFlag( _
            cboFlags.list(cboFlags.ListIndex)) & flgOpts.GetCurVal)
Else
    If cboFlags.ListIndex <> -1 Then _
        lstFamFlags.ListItems.Add Text:=modMain.TranslateFlag(modMain.ShortFlag( _
            cboFlags.list(cboFlags.ListIndex)) & flgOpts.GetCurVal)
End If
cmdAdd.Caption = "Add"
End Sub

Private Sub cmdAddNew_Click()
Dim x As Long
Dim i As Long
Dim t As Boolean
MousePointer = vbHourglass
ReDim Preserve dbFamiliars(1 To UBound(dbFamiliars) + 1)
x = dbFamiliars(UBound(dbFamiliars) - 1).iID
x = x + 1
Do Until t = True
    t = True
    i = GetFamID(x)
    If i <> 0 Then
        t = False
        x = x + 1
    End If
Loop
With dbFamiliars(UBound(dbFamiliars))
    .iID = x
    .sFamName = "New Familiar"
    .dEXPPerLevel = 300
    .lLevelMax = 25
    .lLevelMod = 1
    .lMaxDam = 5
    .lMinDam = 1
    .lRidable = 0
    .lSpeed = 0
    .lStartHPMax = 25
    .lStartHPMin = 15
    .lSwings = 1
    .sAttackMessage = "<%n> attacks <%m> for <%d> damage!"
    .sMessage2 = "<%n> does an attack on <%m>!"
    .sMissMessage = "<%n> misses their attack on <%m>!"
    .sMissMessage2 = "<%n> misses when they attack <%m>!"
    .sFlags = "0"
    .sDescription = "A New Familiar"
End With
lcID = UBound(dbFamiliars)
FillFams lcID, True
MousePointer = vbDefault
End Sub

Private Sub cmdNext_Click()
On Error GoTo cmdNext_Click_Error
SaveFams
lcID = lcID + 1
If lcID > UBound(dbFamiliars) Then lcID = 1
FillFams lcID
On Error GoTo 0
Exit Sub
cmdNext_Click_Error:
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo cmdPrevious_Click_Error
SaveFams
lcID = lcID - 1
If lcID < LBound(dbFamiliars) Then lcID = UBound(dbFamiliars)
FillFams lcID
On Error GoTo 0
Exit Sub
cmdPrevious_Click_Error:
End Sub

Private Sub cmdRemove_Click()
If lstFamFlags.SelectedItem.Index <> -1 Then lstFamFlags.ListItems.Remove (lstFamFlags.SelectedItem.Index)
cmdAdd.Caption = "Add"
End Sub

Private Sub cmdSave_Click()
SaveFams
End Sub

Private Sub Form_Load()
FillFams FillList:=True
lstFamFlags.ColumnHeaders(1).Width = 10000
End Sub

Private Sub lstFamFlags_Click()
Dim s As String
Dim t As String
If lstFamFlags.ListItems.Count < 1 Then Exit Sub
s = Left$(lstFamFlags.SelectedItem.Text, 25)
s = Trim$(s)
For i = 0 To cboFlags.ListCount - 1
    If cboFlags.list(i) = s Then
        cboFlags.ListIndex = i
        Exit For
    End If
Next
t = Mid$(lstFamFlags.SelectedItem.Text, 28)
t = Trim$(t)
flgOpts.SetVal Val(t)
cmdAdd.Caption = "Modify"
End Sub

Private Sub lstFamiliars_Click()
If bIs Then Exit Sub
Dim i As Long
MousePointer = vbHourglass
For i = LBound(dbFamiliars) To UBound(dbFamiliars)
    With dbFamiliars(i)
        If .iID & " " & .sFamName = lstFamiliars.ItemText Then
            FillFams i
            Exit For
        End If
    End With
Next
MousePointer = vbDefault
End Sub

Sub DoFamDam()
Dim lMin As String
Dim lMax As String
Dim lLevel As Long
Dim lMod As Long
lblDamage.Caption = "Assuming the familiar is level 15, the familiars MAX damage would range from <%min> to <%max>. (note: The mod bonus is given at random)"
If Val(txtLM.Text) < 15 Then lLevel = Val(txtLM.Text) Else lLevel = 15
lMod = Val(txtMod.Text) * lLevel
lMin = Val(txtMin.Text) + lMod
lMax = Val(txtMax.Text) + lMod
lblDamage.Caption = Replace$(lblDamage.Caption, "<%min>", txtMin.Text)
lblDamage.Caption = Replace$(lblDamage.Caption, "<%max>", lMax)

End Sub

Private Sub ss1_Click(PreviousTab As Integer)
cmdAdd.Caption = "Add"
End Sub

Private Sub txtLM_Change()
DoFamDam
End Sub

Private Sub txtMax_Change()
DoFamDam
End Sub

Private Sub txtMin_Change()
DoFamDam
End Sub

Private Sub txtMod_Change()
DoFamDam
End Sub
