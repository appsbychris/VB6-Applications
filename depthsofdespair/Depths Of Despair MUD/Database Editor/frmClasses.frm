VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmClasses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Classes"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   9
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   10935
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   240
      TabIndex        =   67
      Top             =   240
      Width           =   2175
   End
   Begin ServerEditor.UltraBox lstClasses 
      Height          =   4215
      Left            =   240
      TabIndex        =   65
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   7435
      Style           =   3
      Color           =   16777215
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
   Begin TabDlg.SSTab ssClass 
      Height          =   4575
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8070
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
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
      TabPicture(0)   =   "frmClasses.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblLabel(0)"
      Tab(0).Control(1)=   "lblLabel(1)"
      Tab(0).Control(2)=   "lblLabel(4)"
      Tab(0).Control(3)=   "lblLabel(5)"
      Tab(0).Control(4)=   "lblLabel(6)"
      Tab(0).Control(5)=   "lblLabel(7)"
      Tab(0).Control(6)=   "lblLabel(8)"
      Tab(0).Control(7)=   "lblLabel(10)"
      Tab(0).Control(8)=   "lblLabel(12)"
      Tab(0).Control(9)=   "lblLabel(26)"
      Tab(0).Control(10)=   "txtMaxMana"
      Tab(0).Control(11)=   "txtMinMana"
      Tab(0).Control(12)=   "txtMagicLevel"
      Tab(0).Control(13)=   "txtEXP"
      Tab(0).Control(14)=   "txtID"
      Tab(0).Control(15)=   "txtName"
      Tab(0).Control(16)=   "cboWeaponType"
      Tab(0).Control(17)=   "cboArmorType"
      Tab(0).Control(18)=   "cboType"
      Tab(0).Control(19)=   "cboMagicLevel"
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Class Bonuses"
      TabPicture(1)   =   "frmClasses.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ssClassB"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cboFlags"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "flgFlags"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Flags"
      TabPicture(2)   =   "frmClasses.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblHelp(5)"
      Tab(2).Control(1)=   "cboCFlags"
      Tab(2).Control(2)=   "cmdAdd(5)"
      Tab(2).Control(3)=   "cmdRemove(5)"
      Tab(2).Control(4)=   "flgOpts"
      Tab(2).Control(5)=   "lstClassFlags"
      Tab(2).ControlCount=   6
      Begin MSComctlLib.ListView lstClassFlags 
         Height          =   3135
         Left            =   -72720
         TabIndex        =   59
         Top             =   1320
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   5530
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
      Begin ServerEditor.FlagOptions flgFlags 
         Height          =   375
         Left            =   3360
         TabIndex        =   58
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Style           =   0
      End
      Begin ServerEditor.FlagOptions flgOpts 
         Height          =   375
         Left            =   -70920
         TabIndex        =   57
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Style           =   0
      End
      Begin VB.ComboBox cboFlags 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmClasses.frx":0054
         Left            =   120
         List            =   "frmClasses.frx":008B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   3135
      End
      Begin TabDlg.SSTab ssClassB 
         Height          =   3495
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   6165
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
         TabCaption(0)   =   "Base Bonus"
         TabPicture(0)   =   "frmClasses.frx":0170
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblHelp(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lstBase"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdRemove(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmdAdd(0)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Beginner Bonus"
         TabPicture(1)   =   "frmClasses.frx":018C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdRemove(1)"
         Tab(1).Control(1)=   "cmdAdd(1)"
         Tab(1).Control(2)=   "txtClsPts(1)"
         Tab(1).Control(3)=   "lstBegin"
         Tab(1).Control(4)=   "lblLabel(3)"
         Tab(1).Control(5)=   "lblHelp(1)"
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "Intermediate Bonus"
         TabPicture(2)   =   "frmClasses.frx":01A8
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdRemove(2)"
         Tab(2).Control(1)=   "cmdAdd(2)"
         Tab(2).Control(2)=   "txtClsPts(2)"
         Tab(2).Control(3)=   "lstInter"
         Tab(2).Control(4)=   "lblLabel(13)"
         Tab(2).Control(5)=   "lblHelp(2)"
         Tab(2).ControlCount=   6
         TabCaption(3)   =   "Master Bonus"
         TabPicture(3)   =   "frmClasses.frx":01C4
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "cmdRemove(3)"
         Tab(3).Control(1)=   "cmdAdd(3)"
         Tab(3).Control(2)=   "txtClsPts(3)"
         Tab(3).Control(3)=   "lstMaster"
         Tab(3).Control(4)=   "lblLabel(14)"
         Tab(3).Control(5)=   "lblHelp(3)"
         Tab(3).ControlCount=   6
         TabCaption(4)   =   "Guru Bonus"
         TabPicture(4)   =   "frmClasses.frx":01E0
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "cmdRemove(4)"
         Tab(4).Control(1)=   "cmdAdd(4)"
         Tab(4).Control(2)=   "txtClsPts(4)"
         Tab(4).Control(3)=   "lstGuru"
         Tab(4).Control(4)=   "lblLabel(15)"
         Tab(4).Control(5)=   "lblHelp(4)"
         Tab(4).ControlCount=   6
         Begin VB.CommandButton cmdRemove 
            Caption         =   "< Remove"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   -72360
            TabIndex        =   26
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add >"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   -72360
            TabIndex        =   25
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "< Remove"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   -72360
            TabIndex        =   23
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add >"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   -72360
            TabIndex        =   22
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "< Remove"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   -72360
            TabIndex        =   20
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add >"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   -72360
            TabIndex        =   19
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "< Remove"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   -72360
            TabIndex        =   17
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add >"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   -72360
            TabIndex        =   16
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add >"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   13
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "< Remove"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   14
            Top             =   840
            Width           =   1095
         End
         Begin ServerEditor.NumOnlyText txtClsPts 
            Height          =   255
            Index           =   1
            Left            =   -73560
            TabIndex        =   15
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
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
            Text            =   "10"
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   6
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtClsPts 
            Height          =   255
            Index           =   2
            Left            =   -73560
            TabIndex        =   18
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
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
            Text            =   "50"
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   6
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtClsPts 
            Height          =   255
            Index           =   3
            Left            =   -73560
            TabIndex        =   21
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
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
            Text            =   "100"
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   6
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtClsPts 
            Height          =   255
            Index           =   4
            Left            =   -73560
            TabIndex        =   24
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
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
            Text            =   "200"
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   6
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin MSComctlLib.ListView lstBase 
            Height          =   3015
            Left            =   3840
            TabIndex        =   60
            Top             =   360
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   5318
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
         Begin MSComctlLib.ListView lstBegin 
            Height          =   3015
            Left            =   -71160
            TabIndex        =   61
            Top             =   360
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   5318
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
         Begin MSComctlLib.ListView lstInter 
            Height          =   3015
            Left            =   -71160
            TabIndex        =   62
            Top             =   360
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   5318
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
         Begin MSComctlLib.ListView lstMaster 
            Height          =   3015
            Left            =   -71160
            TabIndex        =   63
            Top             =   360
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   5318
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
         Begin MSComctlLib.ListView lstGuru 
            Height          =   3015
            Left            =   -71160
            TabIndex        =   64
            Top             =   360
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   5318
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
            Caption         =   "Cls. Pts. Req.:"
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
            Index           =   15
            Left            =   -74880
            TabIndex        =   55
            Top             =   480
            Width           =   1050
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cls. Pts. Req.:"
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
            Index           =   14
            Left            =   -74880
            TabIndex        =   54
            Top             =   480
            Width           =   1050
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cls. Pts. Req.:"
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
            Index           =   13
            Left            =   -74880
            TabIndex        =   53
            Top             =   480
            Width           =   1050
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cls. Pts. Req.:"
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
            Index           =   3
            Left            =   -74880
            TabIndex        =   52
            Top             =   480
            Width           =   1050
         End
         Begin VB.Label lblHelp 
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
            Height          =   2055
            Index           =   4
            Left            =   -74880
            TabIndex        =   51
            Top             =   1200
            Width           =   3615
         End
         Begin VB.Label lblHelp 
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
            Height          =   2055
            Index           =   3
            Left            =   -74880
            TabIndex        =   50
            Top             =   1200
            Width           =   3615
         End
         Begin VB.Label lblHelp 
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
            Height          =   2055
            Index           =   2
            Left            =   -74880
            TabIndex        =   49
            Top             =   1200
            Width           =   3615
         End
         Begin VB.Label lblHelp 
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
            Height          =   2055
            Index           =   1
            Left            =   -74880
            TabIndex        =   48
            Top             =   1200
            Width           =   3615
         End
         Begin VB.Label lblHelp 
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
            Height          =   2055
            Index           =   0
            Left            =   120
            TabIndex        =   47
            Top             =   1200
            Width           =   3615
         End
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "< Remove"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -69480
         TabIndex        =   29
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add >"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -72720
         TabIndex        =   28
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cboCFlags 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmClasses.frx":01FC
         Left            =   -74760
         List            =   "frmClasses.frx":0227
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   480
         Width           =   3735
      End
      Begin VB.ComboBox cboMagicLevel 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmClasses.frx":0316
         Left            =   -73440
         List            =   "frmClasses.frx":032C
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3720
         Width           =   2775
      End
      Begin VB.ComboBox cboType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmClasses.frx":037D
         Left            =   -73440
         List            =   "frmClasses.frx":03A2
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3360
         Width           =   2775
      End
      Begin VB.ComboBox cboArmorType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmClasses.frx":042D
         Left            =   -73440
         List            =   "frmClasses.frx":0479
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3000
         Width           =   2775
      End
      Begin VB.ComboBox cboWeaponType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmClasses.frx":066D
         Left            =   -73440
         List            =   "frmClasses.frx":06A7
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73440
         MaxLength       =   16
         TabIndex        =   3
         Top             =   840
         Width           =   2775
      End
      Begin ServerEditor.NumOnlyText txtID 
         Height          =   255
         Left            =   -73440
         TabIndex        =   1
         Top             =   480
         Width           =   495
         _ExtentX        =   873
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
         Enabled         =   0   'False
         Backcolor       =   -2147483643
      End
      Begin ServerEditor.NumOnlyText txtEXP 
         Height          =   255
         Left            =   -72360
         TabIndex        =   2
         Top             =   480
         Width           =   1695
         _ExtentX        =   873
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
      Begin ServerEditor.NumOnlyText txtMagicLevel 
         Height          =   255
         Left            =   -73440
         TabIndex        =   4
         Top             =   1200
         Width           =   495
         _ExtentX        =   873
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
      Begin ServerEditor.NumOnlyText txtMinMana 
         Height          =   255
         Left            =   -73440
         TabIndex        =   5
         Top             =   1920
         Width           =   735
         _ExtentX        =   873
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
      Begin ServerEditor.NumOnlyText txtMaxMana 
         Height          =   255
         Left            =   -71400
         TabIndex        =   6
         Top             =   1920
         Width           =   735
         _ExtentX        =   1296
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
      Begin VB.Label lblHelp 
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
         Height          =   3015
         Index           =   5
         Left            =   -74880
         TabIndex        =   56
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Magic Level"
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
         Index           =   26
         Left            =   -74760
         TabIndex        =   46
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "EXP"
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
         Index           =   12
         Left            =   -72720
         TabIndex        =   45
         Top             =   480
         Width           =   270
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Max Mana"
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
         Index           =   10
         Left            =   -72600
         TabIndex        =   44
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Min Mana"
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
         Index           =   8
         Left            =   -74760
         TabIndex        =   43
         Top             =   1920
         Width           =   675
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Magic Level"
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
         Index           =   7
         Left            =   -74760
         TabIndex        =   42
         Top             =   3720
         Width           =   825
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Magic Type"
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
         Index           =   6
         Left            =   -74760
         TabIndex        =   41
         Top             =   3360
         Width           =   810
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Armor Type"
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
         Index           =   5
         Left            =   -74760
         TabIndex        =   40
         Top             =   3000
         Width           =   840
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Weapon Type"
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
         Index           =   4
         Left            =   -74760
         TabIndex        =   39
         Top             =   2640
         Width           =   1005
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Name"
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
         Index           =   1
         Left            =   -74760
         TabIndex        =   38
         Top             =   840
         Width           =   405
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "ID"
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
         Index           =   0
         Left            =   -74760
         TabIndex        =   37
         Top             =   480
         Width           =   165
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   32
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "< Previous"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   31
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "(new)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   33
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "(save)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   30
      Top             =   5160
      Width           =   855
   End
   Begin ServerEditor.Raise Raise2 
      Height          =   4815
      Left            =   120
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   8493
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise3 
      Height          =   495
      Left            =   6120
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5040
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   873
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise1 
      Height          =   4815
      Left            =   2640
      TabIndex        =   66
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8493
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise4 
      Height          =   5655
      Left            =   0
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   9975
      Style           =   4
      Color           =   0
   End
End
Attribute VB_Name = "frmClasses"
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
Rem***************                frmClasses                      **********************
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
lstClasses.SetSelected lstClasses.FindInStr(txtFind.Text), True, True
End Sub

Private Sub cboCFlags_Change()
Dim s As String
If cboCFlags.ListIndex <> -1 Then
    lblHelp(5).Caption = GetHelp(cboCFlags.list(cboCFlags.ListIndex))
    flgOpts.ViewStyle = modMain.DeterStyle(modMain.ShortFlag(cboCFlags.list(cboCFlags.ListIndex)), s)
    If flgOpts.ViewStyle = ComboInputFeed Then modMain.FeedAList flgOpts, s
    cmdAdd(5).Caption = "Add >"
End If
End Sub

Private Sub cboCFlags_Click()
cboCFlags_Change
End Sub

Private Sub cboFlags_Change()
Dim s As String
If cboFlags.ListIndex <> -1 Then
    lblHelp(ssClassB.Tab).Caption = GetHelp(cboFlags.list(cboFlags.ListIndex))
    flgFlags.ViewStyle = modMain.DeterStyle(modMain.ShortFlag(cboFlags.list(cboFlags.ListIndex)), s)
    If flgFlags.ViewStyle = ComboInputFeed Then modMain.FeedAList flgFlags, s
    For i = 0 To 4
        cmdAdd(i).Caption = "Add >"
    Next
End If
End Sub

Private Sub cboFlags_Click()
cboFlags_Change
End Sub



Private Sub cmdAdd_Click(Index As Integer)
If cmdAdd(Index).Caption = "Modify >" Then
    Select Case Index
        Case 0
            If cboFlags.ListIndex <> -1 Then _
                lstBase.SelectedItem.Text = modMain.TranslateFlag(modMain.ShortFlag( _
                    cboFlags.list(cboFlags.ListIndex)) & flgFlags.GetCurVal)
        Case 1
            If cboFlags.ListIndex <> -1 Then _
                lstBegin.SelectedItem.Text = modMain.TranslateFlag(modMain.ShortFlag( _
                    cboFlags.list(cboFlags.ListIndex)) & flgFlags.GetCurVal)
        Case 2
            If cboFlags.ListIndex <> -1 Then _
                lstInter.SelectedItem.Text = modMain.TranslateFlag(modMain.ShortFlag( _
                    cboFlags.list(cboFlags.ListIndex)) & flgFlags.GetCurVal)
        Case 3
            If cboFlags.ListIndex <> -1 Then _
                lstMaster.SelectedItem.Text = modMain.TranslateFlag(modMain.ShortFlag( _
                    cboFlags.list(cboFlags.ListIndex)) & flgFlags.GetCurVal)
        Case 4
            If cboFlags.ListIndex <> -1 Then _
                lstGuru.SelectedItem.Text = modMain.TranslateFlag(modMain.ShortFlag( _
                    cboFlags.list(cboFlags.ListIndex)) & flgFlags.GetCurVal)
        Case 5
            If cboCFlags.ListIndex <> -1 Then _
                lstClassFlags.SelectedItem.Text = modMain.TranslateFlag(modMain.ShortFlag( _
                    cboCFlags.list(cboCFlags.ListIndex)) & flgOpts.GetCurVal)
    End Select
Else
    Select Case Index
        Case 0
            If cboFlags.ListIndex <> -1 Then _
                lstBase.ListItems.Add Text:=modMain.TranslateFlag(modMain.ShortFlag( _
                    cboFlags.list(cboFlags.ListIndex)) & flgFlags.GetCurVal)
        Case 1
            If cboFlags.ListIndex <> -1 Then _
                lstBegin.ListItems.Add Text:=modMain.TranslateFlag(modMain.ShortFlag( _
                    cboFlags.list(cboFlags.ListIndex)) & flgFlags.GetCurVal)
        Case 2
            If cboFlags.ListIndex <> -1 Then _
                lstInter.ListItems.Add Text:=modMain.TranslateFlag(modMain.ShortFlag( _
                    cboFlags.list(cboFlags.ListIndex)) & flgFlags.GetCurVal)
        Case 3
            If cboFlags.ListIndex <> -1 Then _
                lstMaster.ListItems.Add Text:=modMain.TranslateFlag(modMain.ShortFlag( _
                    cboFlags.list(cboFlags.ListIndex)) & flgFlags.GetCurVal)
        Case 4
            If cboFlags.ListIndex <> -1 Then _
                lstGuru.ListItems.Add Text:=modMain.TranslateFlag(modMain.ShortFlag( _
                    cboFlags.list(cboFlags.ListIndex)) & flgFlags.GetCurVal)
        Case 5
            If cboCFlags.ListIndex <> -1 Then _
                lstClassFlags.ListItems.Add Text:=modMain.TranslateFlag(modMain.ShortFlag( _
                    cboCFlags.list(cboCFlags.ListIndex)) & flgOpts.GetCurVal)
    End Select
End If
For i = 0 To 5
    cmdAdd(i).Caption = "Add >"
Next
End Sub

Private Sub cmdAddNew_Click()
Dim x As Long
Dim i As Long
Dim t As Boolean
MousePointer = vbHourglass
ReDim Preserve dbClass(1 To UBound(dbClass) + 1)
x = dbClass(UBound(dbClass) - 1).iID
x = x + 1
Do Until t = True
    t = True
    i = GetClassID(, x)
    If i <> 0 Then
        t = False
        x = x + 1
    End If
Loop
With dbClass(UBound(dbClass))
    .iID = x
    .sName = "New Class"
    .dEXP = 0
    .dBeginnerMax = 10
    .dGuru = 200
    .dIntermediateMax = 50
    .dMasterMax = 100
    .iArmorType = 0
    .iMaxMana = 0
    .iMinMana = 0
    .iSpellLevel = 0
    .iSpellType = 0
    .iUseMagical = 1
    .sFlags = "0"
    .sBaseBonus = "0"
    .sBBonus = "0"
    .sGBonus = "0"
    .sIBonus = "0"
    .sMBonus = "0"
End With
lcID = UBound(dbClass)
FillClass lcID, True
MousePointer = vbDefault
End Sub

Private Sub cmdNext_Click()
On Error GoTo cmdNext_Click_Error
SaveClass
lcID = lcID + 1
If lcID > UBound(dbClass) Then lcID = 1
FillClass lcID
On Error GoTo 0
Exit Sub
cmdNext_Click_Error:

End Sub

Private Sub cmdPrevious_Click()
On Error GoTo cmdPrevious_Click_Error
SaveClass
lcID = lcID - 1
If lcID < LBound(dbClass) Then lcID = UBound(dbClass)
FillClass lcID
On Error GoTo 0
Exit Sub
cmdPrevious_Click_Error:
End Sub

Private Sub cmdRemove_Click(Index As Integer)
Select Case Index
    Case 0
        If lstBase.SelectedItem.Index <> -1 Then lstBase.ListItems.Remove (lstBase.SelectedItem.Index)
    Case 1
        If lstBegin.SelectedItem.Index <> -1 Then lstBegin.ListItems.Remove (lstBegin.SelectedItem.Index)
    Case 2
        If lstInter.SelectedItem.Index <> -1 Then lstInter.ListItems.Remove (lstInter.SelectedItem.Index)
    Case 3
        If lstMaster.SelectedItem.Index <> -1 Then lstMaster.ListItems.Remove (lstMaster.SelectedItem.Index)
    Case 4
        If lstGuru.SelectedItem.Index <> -1 Then lstGuru.ListItems.Remove (lstGuru.SelectedItem.Index)
    Case 5
        If lstClassFlags.SelectedItem.Index <> -1 Then lstClassFlags.ListItems.Remove (lstClassFlags.SelectedItem.Index)
End Select
For i = 0 To 5
    cmdAdd(i).Caption = "Add >"
Next
End Sub

Private Sub cmdSave_Click()
SaveClass
End Sub

Private Sub Form_Load()
lcID = 1
FillClass FillList:=True
lstClassFlags.ColumnHeaders(1).Width = 10000
lstBase.ColumnHeaders(1).Width = 10000
lstBegin.ColumnHeaders(1).Width = 10000
lstInter.ColumnHeaders(1).Width = 10000
lstMaster.ColumnHeaders(1).Width = 10000
lstGuru.ColumnHeaders(1).Width = 10000
modMain.PopulateCBOFlag cboCFlags
End Sub

Private Sub lstBase_Click()
Dim s As String
Dim t As String
If lstBase.ListItems.Count < 1 Then Exit Sub
s = Left$(lstBase.SelectedItem.Text, 25)
s = Trim$(s)
For i = 0 To cboFlags.ListCount - 1
    If cboFlags.list(i) = s Then
        cboFlags.ListIndex = i
        Exit For
    End If
Next
t = Mid$(lstBase.SelectedItem.Text, 28)
t = Trim$(t)
flgFlags.SetVal Val(t)
cmdAdd(0).Caption = "Modify >"
End Sub

Private Sub lstBegin_Click()
Dim s As String
Dim t As String
If lstBegin.ListItems.Count < 1 Then Exit Sub
s = Left$(lstBegin.SelectedItem.Text, 25)
s = Trim$(s)
For i = 0 To cboFlags.ListCount - 1
    If cboFlags.list(i) = s Then
        cboFlags.ListIndex = i
        Exit For
    End If
Next
t = Mid$(lstBegin.SelectedItem.Text, 28)
t = Trim$(t)
flgFlags.SetVal Val(t)
cmdAdd(1).Caption = "Modify >"
End Sub

Private Sub lstClasses_Click()
If bIs Then Exit Sub
Dim i As Long
MousePointer = vbHourglass
For i = LBound(dbClass) To UBound(dbClass)
    With dbClass(i)
        If .iID & " " & .sName = lstClasses.ItemText Then
            FillClass i
            Exit For
        End If
    End With
Next
MousePointer = vbDefault
End Sub

Sub SaveClass()
Dim s As String
MousePointer = vbHourglass
ReverseEffects lcID, Class
With dbClass(lcID)
    .dEXP = CDbl(txtEXP.Text)
    .dBeginnerMax = CDbl(txtClsPts(1).Text)
    .dGuru = CDbl(txtClsPts(4).Text)
    .dIntermediateMax = CDbl(txtClsPts(2).Text)
    .dMasterMax = CDbl(txtClsPts(3).Text)
    .iArmorType = CLng(Trim$(Replace$(Left$(cboArmorType.list(cboArmorType.ListIndex), 2), ",", "")))
    .iMaxMana = CLng(txtMaxMana.Text)
    .iMinMana = CLng(txtMinMana.Text)
    .iWeapon = CLng(Trim$(Left$(cboWeaponType.list(cboWeaponType.ListIndex), 2)))
    .iSpellLevel = CLng(Left$(cboMagicLevel.list(cboMagicLevel.ListIndex), 1))
    .iSpellType = CLng(Trim$(Left$(cboType.list(cboType.ListIndex), 2)))
    For i = 1 To lstBase.ListItems.Count
        s = s & modMain.MakeDBFlag(lstBase.ListItems(i).Text) & ":"
    Next
    If s = "" Then s = "0"
    .sBaseBonus = s
    s = ""
    For i = 1 To lstBegin.ListItems.Count
        s = s & modMain.MakeDBFlag(lstBegin.ListItems(i).Text) & ":"
    Next
    If s = "" Then s = "0"
    .sBBonus = s
    s = ""
    For i = 1 To lstInter.ListItems.Count
        s = s & modMain.MakeDBFlag(lstInter.ListItems(i).Text) & ":"
    Next
    If s = "" Then s = "0"
    .sIBonus = s
    s = ""
    For i = 1 To lstMaster.ListItems.Count
        s = s & modMain.MakeDBFlag(lstMaster.ListItems(i).Text) & ":"
    Next
    If s = "" Then s = "0"
    .sMBonus = s
    s = ""
    For i = 1 To lstGuru.ListItems.Count
        s = s & modMain.MakeDBFlag(lstGuru.ListItems(i).Text) & ":"
    Next
    If s = "" Then s = "0"
    .sGBonus = s
    s = ""
    For i = 1 To lstClassFlags.ListItems.Count
        s = s & modMain.MakeDBFlag(lstClassFlags.ListItems(i).Text) & ";"
    Next
    s = ""
    .sName = txtName.Text
    .iUseMagical = CLng(txtMagicLevel.Text)
End With
modUpdateDatabase.SaveMemoryToDatabase Class
DoEffects lcID, Class
modUpdateDatabase.SaveMemoryToDatabase Players
FillClass lcID, True
MousePointer = vbDefault
End Sub

Sub FillClass(Optional Arg As Long = -1, Optional FillList As Boolean = False)
Dim i As Long, j As Long
Dim m As Long
Dim Arr() As String
MousePointer = vbHourglass
bIs = True
If Arg = -1 Then Arg = LBound(dbClass)
'm = lstClasses.ListIndex
If FillList Then lstClasses.Clear
For i = LBound(dbClass) To UBound(dbClass)
    With dbClass(i)
        If FillList Then lstClasses.AddItem CStr(.iID & " " & .sName)
        If i = Arg Then
            lcID = i
            txtID.Text = .iID
            txtEXP.Text = .dEXP
            txtName.Text = .sName
            txtMagicLevel.Text = .iUseMagical
            txtMinMana.Text = .iMinMana
            txtMaxMana.Text = .iMaxMana
            
            lstBase.ListItems.Clear
            Arr = Split(.sBaseBonus, ":")
            For j = LBound(Arr) To UBound(Arr)
                If Arr(j) <> "" And Arr(j) <> "0" Then
                    lstBase.ListItems.Add Text:=modMain.TranslateFlag(Arr(j))
                End If
            Next
            Erase Arr
            
            lstBegin.ListItems.Clear
            Arr = Split(.sBBonus, ":")
            For j = LBound(Arr) To UBound(Arr)
                If Arr(j) <> "" And Arr(j) <> "0" Then
                    lstBegin.ListItems.Add Text:=modMain.TranslateFlag(Arr(j))
                End If
            Next
            Erase Arr
            
            lstInter.ListItems.Clear
            Arr = Split(.sIBonus, ":")
            For j = LBound(Arr) To UBound(Arr)
                If Arr(j) <> "" And Arr(j) <> "0" Then
                    lstInter.ListItems.Add Text:=modMain.TranslateFlag(Arr(j))
                End If
            Next
            Erase Arr
            
            lstMaster.ListItems.Clear
            Arr = Split(.sMBonus, ":")
            For j = LBound(Arr) To UBound(Arr)
                If Arr(j) <> "" And Arr(j) <> "0" Then
                    lstMaster.ListItems.Add Text:=modMain.TranslateFlag(Arr(j))
                End If
            Next
            Erase Arr
            
            lstGuru.ListItems.Clear
            Arr = Split(.sGBonus, ":")
            For j = LBound(Arr) To UBound(Arr)
                If Arr(j) <> "" And Arr(j) <> "0" Then
                    lstGuru.ListItems.Add Text:=modMain.TranslateFlag(Arr(j))
                End If
            Next
            Erase Arr
            
            'lstClassFlags.ListItems.Clear
            'Arr = Split(.sFlags, ";")
            'For j = LBound(Arr) To UBound(Arr)
            '    If Arr(j) <> "" And Arr(j) <> "0" Then
            '        lstClassFlags.ListItems.Add Text:=modMain.TranslateFlag(Arr(j))
            '    End If
           ' Next
           ' Erase Arr
            
            txtClsPts(1).Text = .dBeginnerMax
            txtClsPts(2).Text = .dIntermediateMax
            txtClsPts(3).Text = .dMasterMax
            txtClsPts(4).Text = .dGuru
            
            modMain.SetCBOlstIndex cboWeaponType, .iWeapon, [Weapon Type]
            modMain.SetCBOlstIndex cboArmorType, .iArmorType, [Armor Type]
            modMain.SetCBOlstIndex cboType, .iSpellType, [Magic Type]
            modMain.SetCBOlstIndex cboMagicLevel, .iSpellLevel, [Magic Level]
            If Not FillList Then Exit For
        End If
    End With
Next
modMain.SetLstSelected lstClasses, txtID.Text & " " & txtName.Text
bIs = False
MousePointer = vbDefault
End Sub



Private Sub lstClassFlags_Click()
Dim s As String
Dim t As String
If lstClassFlags.ListItems.Count < 1 Then Exit Sub
s = Left$(lstClassFlags.SelectedItem.Text, 25)
s = Trim$(s)
For i = 0 To cboCFlags.ListCount - 1
    If cboCFlags.list(i) = s Then
        cboCFlags.ListIndex = i
        Exit For
    End If
Next
t = Mid$(lstClassFlags.SelectedItem.Text, 28)
t = Trim$(t)
t = Replace$(t, "TRUE", "0")
t = Replace$(t, "FALSE", "1")
flgOpts.SetVal Val(t)
cmdAdd(5).Caption = "Modify >"
End Sub

Private Sub lstGuru_Click()
Dim s As String
Dim t As String
If lstGuru.ListItems.Count < 1 Then Exit Sub
s = Left$(lstGuru.SelectedItem.Text, 25)
s = Trim$(s)
For i = 0 To cboFlags.ListCount - 1
    If cboFlags.list(i) = s Then
        cboFlags.ListIndex = i
        Exit For
    End If
Next
t = Mid$(lstGuru.SelectedItem.Text, 28)
t = Trim$(t)
flgFlags.SetVal Val(t)
cmdAdd(0).Caption = "Modify >"
End Sub

Private Sub lstInter_Click()
Dim s As String
Dim t As String
If lstInter.ListItems.Count < 1 Then Exit Sub
s = Left$(lstInter.SelectedItem.Text, 25)
s = Trim$(s)
For i = 0 To cboFlags.ListCount - 1
    If cboFlags.list(i) = s Then
        cboFlags.ListIndex = i
        Exit For
    End If
Next
t = Mid$(lstInter.SelectedItem.Text, 28)
t = Trim$(t)
flgFlags.SetVal Val(t)
cmdAdd(2).Caption = "Modify >"
End Sub

Private Sub lstMaster_Click()
Dim s As String
Dim t As String
If lstMaster.ListItems.Count < 1 Then Exit Sub
s = Left$(lstMaster.SelectedItem.Text, 25)
s = Trim$(s)
For i = 0 To cboFlags.ListCount - 1
    If cboFlags.list(i) = s Then
        cboFlags.ListIndex = i
        Exit For
    End If
Next
t = Mid$(lstMaster.SelectedItem.Text, 28)
t = Trim$(t)
flgFlags.SetVal Val(t)
cmdAdd(3).Caption = "Modify >"
End Sub

Private Sub ssClass_DblClick()
Dim i As Long
For i = 0 To 5
    cmdAdd(i).Caption = "Add >"
Next
End Sub

Private Sub ssClassB_Click(PreviousTab As Integer)
Dim i As Long
For i = 0 To 4
    cmdAdd(i).Caption = "Add >"
Next
End Sub

