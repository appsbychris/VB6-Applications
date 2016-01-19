VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMapp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Editor"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12255
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
   ScaleHeight     =   7110
   ScaleWidth      =   12255
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   240
      TabIndex        =   109
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   255
      Left            =   10800
      TabIndex        =   108
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "< Previous"
      Height          =   255
      Left            =   9480
      TabIndex        =   107
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "(new)"
      Height          =   255
      Left            =   8280
      TabIndex        =   106
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "(save)"
      Height          =   255
      Left            =   7320
      TabIndex        =   105
      Top             =   6600
      Width           =   855
   End
   Begin TabDlg.SSTab ssTabMain 
      Height          =   6015
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10610
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmMapp.frx":0000
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
      Tab(0).Control(7)=   "txtGold"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtMaxRegen"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtRoomTitle"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtDescription"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtMobGroup"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "sldLight"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkSafeRoom"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdQuick"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Special"
      TabPicture(1)   =   "frmMapp.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblLabel(7)"
      Tab(1).Control(1)=   "lblLabel(8)"
      Tab(1).Control(2)=   "lblLabel(9)"
      Tab(1).Control(3)=   "lblLabel(10)"
      Tab(1).Control(4)=   "lblLabel(11)"
      Tab(1).Control(5)=   "lblLabel(12)"
      Tab(1).Control(6)=   "lblLabel(13)"
      Tab(1).Control(7)=   "ucShop"
      Tab(1).Control(8)=   "ucTrainClass"
      Tab(1).Control(9)=   "ucEnvironment"
      Tab(1).Control(10)=   "ucDeathRoom"
      Tab(1).Control(11)=   "ucRoomType"
      Tab(1).Control(12)=   "ucSpecialMon"
      Tab(1).Control(13)=   "ucSpecialItem"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Exits"
      TabPicture(2)   =   "frmMapp.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Scripting"
      TabPicture(3)   =   "frmMapp.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblMethod"
      Tab(3).Control(1)=   "txtScripting"
      Tab(3).ControlCount=   2
      Begin VB.CommandButton cmdQuick 
         Caption         =   "Quick Draw Map Editor"
         Height          =   495
         Left            =   6960
         TabIndex        =   114
         Top             =   5400
         Width           =   1455
      End
      Begin ServerEditor.IntelliSense txtScripting 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   104
         Top             =   480
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8705
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   30
         Top             =   360
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   9763
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Exits"
         TabPicture(0)   =   "frmMapp.frx":0070
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lblLabel(14)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblLabel(15)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblLabel(16)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblLabel(17)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblLabel(18)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblLabel(19)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lblLabel(20)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lblLabel(21)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "lblLabel(22)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "lblLabel(23)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "ucDown"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "ucUp"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "ucSW"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "ucNW"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "ucSE"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "ucNE"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "ucEast"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "ucWest"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "ucSouth"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "ucNorth"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).ControlCount=   20
         TabCaption(1)   =   "Doors"
         TabPicture(1)   =   "frmMapp.frx":008C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cboGate(9)"
         Tab(1).Control(1)=   "cboGate(8)"
         Tab(1).Control(2)=   "cboGate(7)"
         Tab(1).Control(3)=   "cboGate(6)"
         Tab(1).Control(4)=   "cboGate(5)"
         Tab(1).Control(5)=   "cboGate(4)"
         Tab(1).Control(6)=   "cboGate(3)"
         Tab(1).Control(7)=   "cboGate(2)"
         Tab(1).Control(8)=   "cboGate(1)"
         Tab(1).Control(9)=   "cboGate(0)"
         Tab(1).Control(10)=   "txtB(0)"
         Tab(1).Control(11)=   "txtP(0)"
         Tab(1).Control(12)=   "chKDoor(0)"
         Tab(1).Control(13)=   "chKDoor(1)"
         Tab(1).Control(14)=   "chKDoor(2)"
         Tab(1).Control(15)=   "chKDoor(3)"
         Tab(1).Control(16)=   "chKDoor(4)"
         Tab(1).Control(17)=   "chKDoor(5)"
         Tab(1).Control(18)=   "chKDoor(6)"
         Tab(1).Control(19)=   "chKDoor(7)"
         Tab(1).Control(20)=   "chKDoor(8)"
         Tab(1).Control(21)=   "chKDoor(9)"
         Tab(1).Control(22)=   "ucD(0)"
         Tab(1).Control(23)=   "ucD(1)"
         Tab(1).Control(24)=   "ucD(2)"
         Tab(1).Control(25)=   "ucD(3)"
         Tab(1).Control(26)=   "ucD(4)"
         Tab(1).Control(27)=   "ucD(5)"
         Tab(1).Control(28)=   "ucD(6)"
         Tab(1).Control(29)=   "ucD(7)"
         Tab(1).Control(30)=   "ucD(8)"
         Tab(1).Control(31)=   "ucD(9)"
         Tab(1).Control(32)=   "txtP(1)"
         Tab(1).Control(33)=   "txtP(2)"
         Tab(1).Control(34)=   "txtP(3)"
         Tab(1).Control(35)=   "txtP(4)"
         Tab(1).Control(36)=   "txtP(5)"
         Tab(1).Control(37)=   "txtP(6)"
         Tab(1).Control(38)=   "txtP(7)"
         Tab(1).Control(39)=   "txtP(8)"
         Tab(1).Control(40)=   "txtP(9)"
         Tab(1).Control(41)=   "txtB(1)"
         Tab(1).Control(42)=   "txtB(2)"
         Tab(1).Control(43)=   "txtB(3)"
         Tab(1).Control(44)=   "txtB(4)"
         Tab(1).Control(45)=   "txtB(5)"
         Tab(1).Control(46)=   "txtB(6)"
         Tab(1).Control(47)=   "txtB(7)"
         Tab(1).Control(48)=   "txtB(8)"
         Tab(1).Control(49)=   "txtB(9)"
         Tab(1).Control(50)=   "lblLabel(36)"
         Tab(1).Control(51)=   "lblLabel(35)"
         Tab(1).Control(52)=   "lblLabel(34)"
         Tab(1).Control(53)=   "lblLabel(33)"
         Tab(1).Control(54)=   "lblLabel(32)"
         Tab(1).Control(55)=   "lblLabel(31)"
         Tab(1).Control(56)=   "lblLabel(30)"
         Tab(1).Control(57)=   "lblLabel(29)"
         Tab(1).Control(58)=   "lblLabel(28)"
         Tab(1).Control(59)=   "lblLabel(27)"
         Tab(1).Control(60)=   "lblLabel(26)"
         Tab(1).Control(61)=   "lblLabel(25)"
         Tab(1).Control(62)=   "lblLabel(24)"
         Tab(1).ControlCount=   63
         TabCaption(2)   =   "Outdoor Food"
         TabPicture(2)   =   "frmMapp.frx":00A8
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdAdd"
         Tab(2).Control(1)=   "cmdRemove"
         Tab(2).Control(2)=   "lstODFood"
         Tab(2).Control(3)=   "lstFood"
         Tab(2).Control(4)=   "Label1(1)"
         Tab(2).Control(5)=   "Label1(0)"
         Tab(2).ControlCount=   6
         TabCaption(3)   =   "Mini Map"
         TabPicture(3)   =   "frmMapp.frx":00C4
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "lN(0)"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "lN(1)"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "lN(2)"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "lN(3)"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).Control(4)=   "lN(4)"
         Tab(3).Control(4).Enabled=   0   'False
         Tab(3).Control(5)=   "lN(5)"
         Tab(3).Control(5).Enabled=   0   'False
         Tab(3).Control(6)=   "lN(6)"
         Tab(3).Control(6).Enabled=   0   'False
         Tab(3).Control(7)=   "lN(7)"
         Tab(3).Control(7).Enabled=   0   'False
         Tab(3).Control(8)=   "lblUp"
         Tab(3).Control(8).Enabled=   0   'False
         Tab(3).Control(9)=   "lblDown"
         Tab(3).Control(9).Enabled=   0   'False
         Tab(3).Control(10)=   "picMap(0)"
         Tab(3).Control(10).Enabled=   0   'False
         Tab(3).Control(11)=   "picMap(1)"
         Tab(3).Control(11).Enabled=   0   'False
         Tab(3).Control(12)=   "picMap(2)"
         Tab(3).Control(12).Enabled=   0   'False
         Tab(3).Control(13)=   "picMap(3)"
         Tab(3).Control(13).Enabled=   0   'False
         Tab(3).Control(14)=   "picMap(4)"
         Tab(3).Control(14).Enabled=   0   'False
         Tab(3).Control(15)=   "picMap(5)"
         Tab(3).Control(15).Enabled=   0   'False
         Tab(3).Control(16)=   "picMap(6)"
         Tab(3).Control(16).Enabled=   0   'False
         Tab(3).Control(17)=   "picMap(7)"
         Tab(3).Control(17).Enabled=   0   'False
         Tab(3).Control(18)=   "picMap(8)"
         Tab(3).Control(18).Enabled=   0   'False
         Tab(3).Control(19)=   "picMap(9)"
         Tab(3).Control(19).Enabled=   0   'False
         Tab(3).Control(20)=   "picMap(10)"
         Tab(3).Control(20).Enabled=   0   'False
         Tab(3).ControlCount=   21
         Begin VB.PictureBox picMap 
            Height          =   1215
            Index           =   10
            Left            =   5160
            ScaleHeight     =   1155
            ScaleWidth      =   1155
            TabIndex        =   152
            Top             =   3000
            Width           =   1215
            Begin VB.Label lblMap 
               Caption         =   "0"
               Height          =   1095
               Index           =   10
               Left            =   0
               TabIndex        =   153
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.PictureBox picMap 
            Height          =   1215
            Index           =   9
            Left            =   5160
            ScaleHeight     =   1155
            ScaleWidth      =   1155
            TabIndex        =   150
            Top             =   1440
            Width           =   1215
            Begin VB.Label lblMap 
               Caption         =   "0"
               Height          =   1095
               Index           =   9
               Left            =   0
               TabIndex        =   151
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.PictureBox picMap 
            Height          =   1215
            Index           =   8
            Left            =   840
            ScaleHeight     =   1155
            ScaleWidth      =   1155
            TabIndex        =   148
            Top             =   3600
            Width           =   1215
            Begin VB.Label lblMap 
               Caption         =   "0"
               Height          =   1095
               Index           =   8
               Left            =   0
               TabIndex        =   149
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.PictureBox picMap 
            Height          =   1215
            Index           =   7
            Left            =   3720
            ScaleHeight     =   1155
            ScaleWidth      =   1155
            TabIndex        =   146
            Top             =   3600
            Width           =   1215
            Begin VB.Label lblMap 
               Caption         =   "0"
               Height          =   1095
               Index           =   7
               Left            =   0
               TabIndex        =   147
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.PictureBox picMap 
            Height          =   1215
            Index           =   6
            Left            =   840
            ScaleHeight     =   1155
            ScaleWidth      =   1155
            TabIndex        =   144
            Top             =   720
            Width           =   1215
            Begin VB.Label lblMap 
               Caption         =   "0"
               Height          =   1095
               Index           =   6
               Left            =   0
               TabIndex        =   145
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.PictureBox picMap 
            Height          =   1215
            Index           =   5
            Left            =   3720
            ScaleHeight     =   1155
            ScaleWidth      =   1155
            TabIndex        =   142
            Top             =   720
            Width           =   1215
            Begin VB.Label lblMap 
               Caption         =   "0"
               Height          =   1095
               Index           =   5
               Left            =   0
               TabIndex        =   143
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.PictureBox picMap 
            Height          =   1215
            Index           =   4
            Left            =   840
            ScaleHeight     =   1155
            ScaleWidth      =   1155
            TabIndex        =   140
            Top             =   2160
            Width           =   1215
            Begin VB.Label lblMap 
               Caption         =   "0"
               Height          =   1095
               Index           =   4
               Left            =   0
               TabIndex        =   141
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.PictureBox picMap 
            Height          =   1215
            Index           =   3
            Left            =   3720
            ScaleHeight     =   1155
            ScaleWidth      =   1155
            TabIndex        =   138
            Top             =   2160
            Width           =   1215
            Begin VB.Label lblMap 
               Caption         =   "0"
               Height          =   1095
               Index           =   3
               Left            =   0
               TabIndex        =   139
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.PictureBox picMap 
            Height          =   1215
            Index           =   2
            Left            =   2280
            ScaleHeight     =   1155
            ScaleWidth      =   1155
            TabIndex        =   136
            Top             =   3600
            Width           =   1215
            Begin VB.Label lblMap 
               Caption         =   "0"
               Height          =   1095
               Index           =   2
               Left            =   0
               TabIndex        =   137
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.PictureBox picMap 
            Height          =   1215
            Index           =   1
            Left            =   2280
            ScaleHeight     =   1155
            ScaleWidth      =   1155
            TabIndex        =   134
            Top             =   720
            Width           =   1215
            Begin VB.Label lblMap 
               Caption         =   "0"
               Height          =   1095
               Index           =   1
               Left            =   0
               TabIndex        =   135
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.PictureBox picMap 
            Height          =   1215
            Index           =   0
            Left            =   2280
            ScaleHeight     =   1155
            ScaleWidth      =   1155
            TabIndex        =   132
            Top             =   2160
            Width           =   1215
            Begin VB.Label lblMap 
               Caption         =   "0"
               Height          =   1095
               Index           =   0
               Left            =   0
               TabIndex        =   133
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "< &Add"
            Height          =   375
            Left            =   -71760
            TabIndex        =   128
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "&Remove >"
            Height          =   375
            Left            =   -71760
            TabIndex        =   127
            Top             =   1320
            Width           =   1095
         End
         Begin VB.ComboBox cboGate 
            Height          =   315
            Index           =   9
            ItemData        =   "frmMapp.frx":00E0
            Left            =   -74040
            List            =   "frmMapp.frx":00EA
            Style           =   2  'Dropdown List
            TabIndex        =   125
            Top             =   4980
            Width           =   975
         End
         Begin VB.ComboBox cboGate 
            Height          =   315
            Index           =   8
            ItemData        =   "frmMapp.frx":00FA
            Left            =   -74040
            List            =   "frmMapp.frx":0104
            Style           =   2  'Dropdown List
            TabIndex        =   124
            Top             =   4500
            Width           =   975
         End
         Begin VB.ComboBox cboGate 
            Height          =   315
            Index           =   7
            ItemData        =   "frmMapp.frx":0114
            Left            =   -74040
            List            =   "frmMapp.frx":011E
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Top             =   4020
            Width           =   975
         End
         Begin VB.ComboBox cboGate 
            Height          =   315
            Index           =   6
            ItemData        =   "frmMapp.frx":012E
            Left            =   -74040
            List            =   "frmMapp.frx":0138
            Style           =   2  'Dropdown List
            TabIndex        =   122
            Top             =   3540
            Width           =   975
         End
         Begin VB.ComboBox cboGate 
            Height          =   315
            Index           =   5
            ItemData        =   "frmMapp.frx":0148
            Left            =   -74040
            List            =   "frmMapp.frx":0152
            Style           =   2  'Dropdown List
            TabIndex        =   121
            Top             =   3060
            Width           =   975
         End
         Begin VB.ComboBox cboGate 
            Height          =   315
            Index           =   4
            ItemData        =   "frmMapp.frx":0162
            Left            =   -74040
            List            =   "frmMapp.frx":016C
            Style           =   2  'Dropdown List
            TabIndex        =   120
            Top             =   2580
            Width           =   975
         End
         Begin VB.ComboBox cboGate 
            Height          =   315
            Index           =   3
            ItemData        =   "frmMapp.frx":017C
            Left            =   -74040
            List            =   "frmMapp.frx":0186
            Style           =   2  'Dropdown List
            TabIndex        =   119
            Top             =   2100
            Width           =   975
         End
         Begin VB.ComboBox cboGate 
            Height          =   315
            Index           =   2
            ItemData        =   "frmMapp.frx":0196
            Left            =   -74040
            List            =   "frmMapp.frx":01A0
            Style           =   2  'Dropdown List
            TabIndex        =   118
            Top             =   1620
            Width           =   975
         End
         Begin VB.ComboBox cboGate 
            Height          =   315
            Index           =   1
            ItemData        =   "frmMapp.frx":01B0
            Left            =   -74040
            List            =   "frmMapp.frx":01BA
            Style           =   2  'Dropdown List
            TabIndex        =   117
            Top             =   1140
            Width           =   975
         End
         Begin VB.ComboBox cboGate 
            Height          =   315
            Index           =   0
            ItemData        =   "frmMapp.frx":01CA
            Left            =   -74040
            List            =   "frmMapp.frx":01D4
            Style           =   2  'Dropdown List
            TabIndex        =   116
            Top             =   660
            Width           =   975
         End
         Begin ServerEditor.NumOnlyText txtB 
            Height          =   375
            Index           =   0
            Left            =   -67560
            TabIndex        =   91
            Top             =   660
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":01E4
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtP 
            Height          =   375
            Index           =   0
            Left            =   -68280
            TabIndex        =   81
            Top             =   660
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":020C
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin VB.CheckBox chKDoor 
            Caption         =   "Door"
            Height          =   255
            Index           =   0
            Left            =   -74280
            TabIndex        =   80
            Top             =   660
            Width           =   735
         End
         Begin VB.CheckBox chKDoor 
            Caption         =   "Door"
            Height          =   255
            Index           =   1
            Left            =   -74280
            TabIndex        =   79
            Top             =   1140
            Width           =   735
         End
         Begin VB.CheckBox chKDoor 
            Caption         =   "Door"
            Height          =   255
            Index           =   2
            Left            =   -74280
            TabIndex        =   78
            Top             =   1620
            Width           =   735
         End
         Begin VB.CheckBox chKDoor 
            Caption         =   "Door"
            Height          =   255
            Index           =   3
            Left            =   -74280
            TabIndex        =   77
            Top             =   2100
            Width           =   735
         End
         Begin VB.CheckBox chKDoor 
            Caption         =   "Door"
            Height          =   255
            Index           =   4
            Left            =   -74280
            TabIndex        =   76
            Top             =   2580
            Width           =   735
         End
         Begin VB.CheckBox chKDoor 
            Caption         =   "Door"
            Height          =   255
            Index           =   5
            Left            =   -74280
            TabIndex        =   75
            Top             =   3060
            Width           =   735
         End
         Begin VB.CheckBox chKDoor 
            Caption         =   "Door"
            Height          =   255
            Index           =   6
            Left            =   -74280
            TabIndex        =   74
            Top             =   3540
            Width           =   735
         End
         Begin VB.CheckBox chKDoor 
            Caption         =   "Door"
            Height          =   255
            Index           =   7
            Left            =   -74280
            TabIndex        =   73
            Top             =   4020
            Width           =   735
         End
         Begin VB.CheckBox chKDoor 
            Caption         =   "Door"
            Height          =   255
            Index           =   8
            Left            =   -74280
            TabIndex        =   72
            Top             =   4500
            Width           =   735
         End
         Begin VB.CheckBox chKDoor 
            Caption         =   "Door"
            Height          =   255
            Index           =   9
            Left            =   -74280
            TabIndex        =   71
            Top             =   4980
            Width           =   735
         End
         Begin ServerEditor.ucCombo ucNorth 
            Height          =   375
            Left            =   -73680
            TabIndex        =   31
            Top             =   540
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucSouth 
            Height          =   375
            Left            =   -73680
            TabIndex        =   32
            Top             =   1020
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucWest 
            Height          =   375
            Left            =   -73680
            TabIndex        =   33
            Top             =   1500
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucEast 
            Height          =   375
            Left            =   -73680
            TabIndex        =   34
            Top             =   1980
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucNE 
            Height          =   375
            Left            =   -73680
            TabIndex        =   35
            Top             =   2460
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucSE 
            Height          =   375
            Left            =   -73680
            TabIndex        =   36
            Top             =   2940
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucNW 
            Height          =   375
            Left            =   -73680
            TabIndex        =   37
            Top             =   3420
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucSW 
            Height          =   375
            Left            =   -73680
            TabIndex        =   38
            Top             =   3900
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucUp 
            Height          =   375
            Left            =   -73680
            TabIndex        =   39
            Top             =   4380
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucDown 
            Height          =   375
            Left            =   -73680
            TabIndex        =   40
            Top             =   4860
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucD 
            Height          =   375
            Index           =   0
            Left            =   -72960
            TabIndex        =   51
            Top             =   660
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucD 
            Height          =   375
            Index           =   1
            Left            =   -72960
            TabIndex        =   52
            Top             =   1140
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucD 
            Height          =   375
            Index           =   2
            Left            =   -72960
            TabIndex        =   53
            Top             =   1620
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucD 
            Height          =   375
            Index           =   3
            Left            =   -72960
            TabIndex        =   54
            Top             =   2100
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucD 
            Height          =   375
            Index           =   4
            Left            =   -72960
            TabIndex        =   55
            Top             =   2580
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucD 
            Height          =   375
            Index           =   5
            Left            =   -72960
            TabIndex        =   56
            Top             =   3060
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucD 
            Height          =   375
            Index           =   6
            Left            =   -72960
            TabIndex        =   57
            Top             =   3540
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucD 
            Height          =   375
            Index           =   7
            Left            =   -72960
            TabIndex        =   58
            Top             =   4020
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucD 
            Height          =   375
            Index           =   8
            Left            =   -72960
            TabIndex        =   59
            Top             =   4500
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.ucCombo ucD 
            Height          =   375
            Index           =   9
            Left            =   -72960
            TabIndex        =   60
            Top             =   4980
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   661
         End
         Begin ServerEditor.NumOnlyText txtP 
            Height          =   375
            Index           =   1
            Left            =   -68280
            TabIndex        =   82
            Top             =   1140
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":0234
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtP 
            Height          =   375
            Index           =   2
            Left            =   -68280
            TabIndex        =   83
            Top             =   1620
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":025C
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtP 
            Height          =   375
            Index           =   3
            Left            =   -68280
            TabIndex        =   84
            Top             =   2100
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":0284
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtP 
            Height          =   375
            Index           =   4
            Left            =   -68280
            TabIndex        =   85
            Top             =   2580
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":02AC
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtP 
            Height          =   375
            Index           =   5
            Left            =   -68280
            TabIndex        =   86
            Top             =   3060
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":02D4
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtP 
            Height          =   375
            Index           =   6
            Left            =   -68280
            TabIndex        =   87
            Top             =   3540
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":02FC
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtP 
            Height          =   375
            Index           =   7
            Left            =   -68280
            TabIndex        =   88
            Top             =   4020
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":0324
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtP 
            Height          =   375
            Index           =   8
            Left            =   -68280
            TabIndex        =   89
            Top             =   4500
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":034C
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtP 
            Height          =   375
            Index           =   9
            Left            =   -68280
            TabIndex        =   90
            Top             =   4980
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":0374
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtB 
            Height          =   375
            Index           =   1
            Left            =   -67560
            TabIndex        =   92
            Top             =   1140
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":039C
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtB 
            Height          =   375
            Index           =   2
            Left            =   -67560
            TabIndex        =   93
            Top             =   1620
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":03C4
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtB 
            Height          =   375
            Index           =   3
            Left            =   -67560
            TabIndex        =   94
            Top             =   2100
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":03EC
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtB 
            Height          =   375
            Index           =   4
            Left            =   -67560
            TabIndex        =   95
            Top             =   2580
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":0414
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtB 
            Height          =   375
            Index           =   5
            Left            =   -67560
            TabIndex        =   96
            Top             =   3060
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":043C
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtB 
            Height          =   375
            Index           =   6
            Left            =   -67560
            TabIndex        =   97
            Top             =   3540
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":0464
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtB 
            Height          =   375
            Index           =   7
            Left            =   -67560
            TabIndex        =   98
            Top             =   4020
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":048C
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtB 
            Height          =   375
            Index           =   8
            Left            =   -67560
            TabIndex        =   99
            Top             =   4500
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":04B4
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtB 
            Height          =   375
            Index           =   9
            Left            =   -67560
            TabIndex        =   100
            Top             =   4980
            Width           =   615
            _extentx        =   1085
            _extenty        =   661
            font            =   "frmMapp.frx":04DC
            text            =   "0"
            allowneg        =   0
            align           =   0
            maxlength       =   0
            enabled         =   0
            backcolor       =   -2147483643
         End
         Begin ServerEditor.UltraBox lstODFood 
            Height          =   4455
            Left            =   -74760
            TabIndex        =   126
            Top             =   840
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   7858
            Style           =   0
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
         Begin ServerEditor.UltraBox lstFood 
            Height          =   4455
            Left            =   -70560
            TabIndex        =   129
            Top             =   840
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   7858
            Style           =   0
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
         Begin VB.Label lblDown 
            Caption         =   "Down:"
            Height          =   255
            Left            =   5160
            TabIndex        =   155
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label lblUp 
            Caption         =   "Up:"
            Height          =   255
            Left            =   5160
            TabIndex        =   154
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Line lN 
            BorderWidth     =   3
            Index           =   7
            X1              =   3480
            X2              =   3720
            Y1              =   3360
            Y2              =   3600
         End
         Begin VB.Line lN 
            BorderWidth     =   3
            Index           =   6
            X1              =   2040
            X2              =   2280
            Y1              =   3600
            Y2              =   3360
         End
         Begin VB.Line lN 
            BorderWidth     =   3
            Index           =   5
            X1              =   2040
            X2              =   2280
            Y1              =   1920
            Y2              =   2160
         End
         Begin VB.Line lN 
            BorderWidth     =   3
            Index           =   4
            X1              =   3480
            X2              =   3720
            Y1              =   2160
            Y2              =   1920
         End
         Begin VB.Line lN 
            BorderWidth     =   3
            Index           =   3
            X1              =   2040
            X2              =   2280
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Line lN 
            BorderWidth     =   3
            Index           =   2
            X1              =   3480
            X2              =   3720
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Line lN 
            BorderWidth     =   3
            Index           =   1
            X1              =   2880
            X2              =   2880
            Y1              =   3600
            Y2              =   3360
         End
         Begin VB.Line lN 
            BorderWidth     =   3
            Index           =   0
            X1              =   2880
            X2              =   2880
            Y1              =   2160
            Y2              =   1920
         End
         Begin VB.Label Label1 
            Caption         =   "All Food Available In Database:"
            Height          =   255
            Index           =   1
            Left            =   -70560
            TabIndex        =   131
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Food that can drop here:"
            Height          =   255
            Index           =   0
            Left            =   -74760
            TabIndex        =   130
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Bash:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   36
            Left            =   -67560
            TabIndex        =   103
            Top             =   420
            Width           =   450
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Pick:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   35
            Left            =   -68280
            TabIndex        =   102
            Top             =   420
            Width           =   390
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Key:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   34
            Left            =   -72960
            TabIndex        =   101
            Top             =   420
            Width           =   360
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Down:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   33
            Left            =   -74880
            TabIndex        =   70
            Top             =   4980
            Width           =   510
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Up:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   32
            Left            =   -74880
            TabIndex        =   69
            Top             =   4500
            Width           =   270
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "SW:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   31
            Left            =   -74880
            TabIndex        =   68
            Top             =   4020
            Width           =   315
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "NW:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   30
            Left            =   -74880
            TabIndex        =   67
            Top             =   3540
            Width           =   315
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "SE:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   29
            Left            =   -74880
            TabIndex        =   66
            Top             =   3060
            Width           =   240
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "NE:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   28
            Left            =   -74880
            TabIndex        =   65
            Top             =   2580
            Width           =   240
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "East:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   27
            Left            =   -74880
            TabIndex        =   64
            Top             =   2100
            Width           =   405
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "West:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   26
            Left            =   -74880
            TabIndex        =   63
            Top             =   1620
            Width           =   480
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "South:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   25
            Left            =   -74880
            TabIndex        =   62
            Top             =   1140
            Width           =   540
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "North:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   24
            Left            =   -74880
            TabIndex        =   61
            Top             =   660
            Width           =   510
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Down:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   23
            Left            =   -74760
            TabIndex        =   50
            Top             =   4860
            Width           =   510
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Up:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   22
            Left            =   -74760
            TabIndex        =   49
            Top             =   4380
            Width           =   270
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "SW:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   -74760
            TabIndex        =   48
            Top             =   3900
            Width           =   315
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "NW:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   20
            Left            =   -74760
            TabIndex        =   47
            Top             =   3420
            Width           =   315
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "SE:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   19
            Left            =   -74760
            TabIndex        =   46
            Top             =   2940
            Width           =   240
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "NE:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   18
            Left            =   -74760
            TabIndex        =   45
            Top             =   2460
            Width           =   240
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "East:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   -74760
            TabIndex        =   44
            Top             =   1980
            Width           =   405
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "West:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   -74760
            TabIndex        =   43
            Top             =   1500
            Width           =   480
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "South:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   -74760
            TabIndex        =   42
            Top             =   1020
            Width           =   540
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "North:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   -74760
            TabIndex        =   41
            Top             =   540
            Width           =   510
         End
      End
      Begin ServerEditor.ucCombo ucSpecialItem 
         Height          =   375
         Left            =   -73680
         TabIndex        =   16
         Top             =   720
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   661
      End
      Begin VB.CheckBox chkSafeRoom 
         Caption         =   "Safe Room"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   4440
         Width           =   255
      End
      Begin MSComctlLib.Slider sldLight 
         Height          =   510
         Left            =   1800
         TabIndex        =   7
         Top             =   3840
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   900
         _Version        =   393216
         LargeChange     =   10
         Min             =   -200
         Max             =   200
         TickFrequency   =   10
      End
      Begin ServerEditor.NumOnlyText txtMobGroup 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   2520
         Width           =   735
         _extentx        =   1296
         _extenty        =   661
         font            =   "frmMapp.frx":0504
         text            =   "1"
         allowneg        =   0
         align           =   0
         maxlength       =   0
         enabled         =   -1
         backcolor       =   -2147483643
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
         Height          =   1215
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Text            =   "frmMapp.frx":052C
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox txtRoomTitle 
         Height          =   375
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   2
         Text            =   "Room Title"
         Top             =   720
         Width           =   4095
      End
      Begin ServerEditor.NumOnlyText txtMaxRegen 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   3000
         Width           =   735
         _extentx        =   1296
         _extenty        =   661
         font            =   "frmMapp.frx":0538
         text            =   "1"
         allowneg        =   0
         align           =   0
         maxlength       =   0
         enabled         =   -1
         backcolor       =   -2147483643
      End
      Begin ServerEditor.NumOnlyText txtGold 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   3480
         Width           =   735
         _extentx        =   1296
         _extenty        =   661
         font            =   "frmMapp.frx":0560
         text            =   "1"
         allowneg        =   0
         align           =   0
         maxlength       =   0
         enabled         =   -1
         backcolor       =   -2147483643
      End
      Begin ServerEditor.ucCombo ucSpecialMon 
         Height          =   375
         Left            =   -73680
         TabIndex        =   17
         Top             =   1320
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   661
      End
      Begin ServerEditor.ucCombo ucRoomType 
         Height          =   375
         Left            =   -73680
         TabIndex        =   18
         Top             =   1920
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   661
      End
      Begin ServerEditor.ucCombo ucDeathRoom 
         Height          =   375
         Left            =   -73680
         TabIndex        =   19
         Top             =   2520
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   661
      End
      Begin ServerEditor.ucCombo ucEnvironment 
         Height          =   375
         Left            =   -73680
         TabIndex        =   20
         Top             =   3120
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   661
      End
      Begin ServerEditor.ucCombo ucTrainClass 
         Height          =   375
         Left            =   -73680
         TabIndex        =   27
         Top             =   3720
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   661
      End
      Begin ServerEditor.ucCombo ucShop 
         Height          =   375
         Left            =   -73680
         TabIndex        =   28
         Top             =   4320
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   661
      End
      Begin VB.Label lblMethod 
         Height          =   375
         Left            =   -74880
         TabIndex        =   115
         Top             =   5520
         Width           =   8295
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Shop:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   -74880
         TabIndex        =   29
         Top             =   4320
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Train Class:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   -74880
         TabIndex        =   26
         Top             =   3720
         Width           =   960
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Environment:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   -74880
         TabIndex        =   25
         Top             =   3120
         Width           =   1125
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Death Room:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   -74880
         TabIndex        =   24
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Room Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   -74880
         TabIndex        =   23
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Special Mon:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   -74880
         TabIndex        =   22
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Special Item:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   -74880
         TabIndex        =   21
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Safe Room:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Top             =   4440
         Width           =   960
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Light:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   3960
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Gold:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Max Regen:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   990
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Mob Group:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   960
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Room Description:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1545
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Room Title:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   960
      End
   End
   Begin ServerEditor.UltraBox lstRooms 
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   9975
      Style           =   0
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
   Begin ServerEditor.Raise Raise5 
      Height          =   6255
      Index           =   0
      Left            =   120
      TabIndex        =   110
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   11033
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise5 
      Height          =   6255
      Index           =   1
      Left            =   3360
      TabIndex        =   111
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11033
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise5 
      Height          =   495
      Index           =   2
      Left            =   7200
      TabIndex        =   112
      Top             =   6480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   873
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise5 
      Height          =   7095
      Index           =   3
      Left            =   0
      TabIndex        =   113
      Top             =   0
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   12515
      Style           =   2
      Color           =   0
   End
End
Attribute VB_Name = "frmMapp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cID As Long
Private bU As Boolean


Private Sub cmdAdd_Click()
If lstFood.ListIndex < 1 Then Exit Sub
If lstODFood.Find(lstFood.list(lstFood.ListIndex)) <> 0 Then Exit Sub
lstODFood.AddItem lstFood.list(lstFood.ListIndex)
End Sub

Private Sub cmdAddNew_Click()
MousePointer = vbHourglass
Dim x As Long
Dim i As Long
Dim t As Boolean
ReDim Preserve dbMap(1 To UBound(dbMap) + 1)
x = dbMap(UBound(dbMap) - 1).lRoomID
x = x + 1
Do Until t = True
    t = True
    i = GetMapIndex(x)
    If i <> 0 Then
        t = False
        x = x + 1
    End If
Loop
With dbMap(UBound(dbMap))
    .iMaxRegen = 2
    .lDeathRoom = 232
    .lRoomID = x
    .sExits = "0"
    .sHidden = "0"
    .sItems = "0"
    .sMonsters = "0"
    .sRoomDesc = "New Room"
    .sRoomTitle = "New Room"
    .sScript = "0"
    .sShopItems = "0"
End With
cID = UBound(dbMap)
FillList
lstRooms.SetSelected lstRooms.Find("(" & dbMap(cID).lRoomID & ") " & dbMap(cID).sRoomTitle), True, True
LoadRoom cID
MousePointer = vbDefault
End Sub

Private Sub cmdNext_Click()
lcd = lcd + 1
If GetMapIndex(lcd) = 0 Then lcd = 1
FillList
LoadRoom lcd
End Sub

Private Sub cmdQuick_Click()
Load frmMapEdit
frmMapEdit.Show
End Sub

Private Sub cmdRemove_Click()
If lstODFood.ListCount < 1 Then Exit Sub
If lstODFood.ListIndex < 1 Then Exit Sub
lstODFood.RemoveItem lstODFood.ListIndex
End Sub

Private Sub lblMap_Click(Index As Integer)
With dbMap(cID)
    Select Case Index
        Case 1
            LoadRoom .lNorth
        Case 2
            LoadRoom .lSouth
        Case 3
            LoadRoom .lEast
        Case 4
            LoadRoom .lWest
        Case 5
            LoadRoom .lNorthEast
        Case 6
            LoadRoom .lNorthWest
        Case 7
            LoadRoom .lSouthEast
        Case 8
            LoadRoom .lSouthWest
        Case 9
            LoadRoom .lUp
        Case 10
            LoadRoom .lDown
    End Select
End With
End Sub

Private Sub txtFind_Change()
lstRooms.SetSelected lstRooms.FindInStr(txtFind.Text), True, True
End Sub

Private Sub chKDoor_Click(Index As Integer)
ucD(Index).Enabled = chKDoor(Index).Value
txtP(Index).Enabled = chKDoor(Index).Value
txtB(Index).Enabled = chKDoor(Index).Value
End Sub

Private Sub cmdSave_Click()
MousePointer = vbHourglass
Me.Enabled = False
SaveMap
Me.Enabled = True
MousePointer = vbNormal
End Sub

Public Sub SaveMap()
Dim s As String
Dim i As Long
With dbMap(cID)
    .dGold = Val(txtGold.Text)
    .iInDoor = ucEnvironment.Number
    .iMaxRegen = Val(txtMaxRegen.Text)
    .iMobGroup = Val(txtMobGroup.Text)
    .iSafeRoom = chkSafeRoom.Value
    .iTrainClass = ucTrainClass.Number
    .iType = ucRoomType.Number
    .lBD = Val(txtB(9).Text)
    .lBE = Val(txtB(3).Text)
    .lBN = Val(txtB(0).Text)
    .lBNE = Val(txtB(4).Text)
    .lBNW = Val(txtB(6).Text)
    .lBS = Val(txtB(1).Text)
    .lBSE = Val(txtB(5).Text)
    .lBSW = Val(txtB(7).Text)
    .lBU = Val(txtB(8).Text)
    .lBW = Val(txtB(2).Text)
    .lPD = Val(txtP(9).Text)
    .lPE = Val(txtP(3).Text)
    .lPN = Val(txtP(0).Text)
    .lPNE = Val(txtP(4).Text)
    .lPNW = Val(txtP(6).Text)
    .lPS = Val(txtP(1).Text)
    .lPSE = Val(txtP(5).Text)
    .lPSW = Val(txtP(7).Text)
    .lPU = Val(txtP(8).Text)
    .lPW = Val(txtP(2).Text)
    modMapFlags.SetMapFlag cID, mapGate, cboGate(9).ListIndex, Down
    modMapFlags.SetMapFlag cID, mapGate, cboGate(3).ListIndex, East
    modMapFlags.SetMapFlag cID, mapGate, cboGate(0).ListIndex, North
    modMapFlags.SetMapFlag cID, mapGate, cboGate(4).ListIndex, NorthEast
    modMapFlags.SetMapFlag cID, mapGate, cboGate(6).ListIndex, NorthWest
    modMapFlags.SetMapFlag cID, mapGate, cboGate(1).ListIndex, South
    modMapFlags.SetMapFlag cID, mapGate, cboGate(5).ListIndex, SouthEast
    modMapFlags.SetMapFlag cID, mapGate, cboGate(7).ListIndex, SouthWest
    modMapFlags.SetMapFlag cID, mapGate, cboGate(8).ListIndex, Up
    modMapFlags.SetMapFlag cID, mapGate, cboGate(2).ListIndex, West
    .lDD = Val(chKDoor(9).Value)
    .lDE = Val(chKDoor(3).Value)
    .lDN = Val(chKDoor(0).Value)
    .lDNE = Val(chKDoor(4).Value)
    .lDNW = Val(chKDoor(6).Value)
    .lDS = Val(chKDoor(1).Value)
    .lDSE = Val(chKDoor(5).Value)
    .lDSW = Val(chKDoor(7).Value)
    .lDU = Val(chKDoor(8).Value)
    .lDW = Val(chKDoor(2).Value)
    .lKD = Val(ucD(9).Number)
    .lKE = Val(ucD(3).Number)
    .lKN = Val(ucD(0).Number)
    .lKNE = Val(ucD(4).Number)
    .lKNW = Val(ucD(6).Number)
    .lKS = Val(ucD(1).Number)
    .lKSE = Val(ucD(5).Number)
    .lKSW = Val(ucD(7).Number)
    .lKU = Val(ucD(8).Number)
    .lKW = Val(ucD(2).Number)
    .lLight = sldLight.Value
    .lDeathRoom = ucDeathRoom.Number
    .lDown = ucDown.Number
    .lEast = ucEast.Number
    .lNorth = ucNorth.Number
    .lNorthEast = ucNE.Number
    .lNorthWest = ucNW.Number
    .lSouth = ucSouth.Number
    .lSouthEast = ucSE.Number
    .lSouthWest = ucSW.Number
    .lSpecialItem = ucSpecialItem.Number
    .lSpecialMon = ucSpecialMon.Number
    .lUp = ucUp.Number
    .lWest = ucWest.Number
    .sRoomDesc = txtDescription.Text
    If .sRoomDesc = "" Then .sRoomDesc = "None"
    .sRoomTitle = txtRoomTitle.Text
    If .sRoomTitle = "" Then .sRoomTitle = "None"
    .sScript = txtScripting.Text
    If .sScript = "" Then .sScript = "0"
    .sShopItems = ucShop.Number
    If lstODFood.ListCount = 0 Then
        s = "0;"
    Else
        s = ""
        For i = 1 To lstODFood.ListCount
            s = s & Mid$(lstODFood.list(i), 2, InStr(1, lstODFood.list(i), ")") - 2) & ";"
        Next
    End If
    .sOutDoorFood = s
End With
bU = True
modUpdateDatabase.SaveMemoryToDatabase Map
modUpdateDatabase.LoadDatabaseIntoMemory False
FillList
If cID = 0 Then Exit Sub
lstRooms.SetSelected lstRooms.Find("(" & dbMap(cID).lRoomID & ") " & dbMap(cID).sRoomTitle), True, True
bU = False
LoadRoom dbMap(cID).lRoomID
End Sub

Private Sub FillList()
Dim i As Long
lstRooms.Paint = False
lstRooms.Clear
For i = 1 To UBound(dbMap)
    With dbMap(i)
        lstRooms.AddItem "(" & .lRoomID & ") " & .sRoomTitle
    End With
    DoEvents
Next
lstRooms.Paint = True
lstFood.Paint = False
lstFood.Clear
For i = 1 To UBound(dbItems)
    With dbItems(i)
        If .sWorn = "ofood" Then
            lstFood.AddItem "(" & .iID & ") " & .sItemName
        End If
    End With
    DoEvents
Next
lstFood.Paint = True
ucDeathRoom.ListSettings = Rooms
ucNorth.ListSettings = Rooms
ucSouth.ListSettings = Rooms
ucEast.ListSettings = Rooms
ucWest.ListSettings = Rooms
ucNW.ListSettings = Rooms
ucSW.ListSettings = Rooms
ucNE.ListSettings = Rooms
ucSE.ListSettings = Rooms
ucUp.ListSettings = Rooms
ucDown.ListSettings = Rooms
End Sub

Private Sub Form_Load()
Dim i As Long
txtScripting.IntelliSenseAddWordsFile App.Path & "\scriptdef.aimg"
txtScripting.IntelliSenseStartSubclassing
ucSpecialItem.ListSettings = Items
ucSpecialMon.ListSettings = Monsters
ucRoomType.ListSettings = RoomType
ucEnvironment.ListSettings = Environment
ucTrainClass.ListSettings = Classes
ucShop.ListSettings = Shops
For i = 0 To 9
    ucD(i).ListSettings = Keys
    ucD(i).Enabled = False
Next
FillList
LoadRoom 1
End Sub

Private Sub lstRooms_Click()
Dim s As String
Dim j As Long
If bU Then Exit Sub
s = lstRooms.list(lstRooms.ListIndex)
s = Mid$(s, 2, InStr(1, s, ")"))
j = CLng(Val(s))
LoadRoom j
End Sub

Private Sub txtScripting_MethodHasChanged()
lblMethod.Caption = txtScripting.Params
End Sub

Private Sub ucRoomType_Change()
ucTrainClass.Enabled = False
ucSpecialMon.Enabled = False
ucShop.Enabled = False
Select Case ucRoomType.Number
    Case 1
        ucShop.Enabled = True
    Case 4
        ucSpecialMon.Enabled = True
    Case 6
        ucTrainClass.Enabled = True
End Select
End Sub

Public Sub LoadRoom(lRID As Long)
Dim Arr() As String
Dim i As Long
Dim dd As Long
cID = GetMapIndex(lRID)
If cID = 0 Then Exit Sub
bU = True
With dbMap(cID)
    txtRoomTitle.Text = .sRoomTitle
    txtDescription.Text = .sRoomDesc
    txtMobGroup.Text = .iMobGroup
    txtMaxRegen.Text = .iMaxRegen
    txtGold.Text = .dGold
    sldLight.Value = .lLight
    chkSafeRoom.Value = .iSafeRoom
    ucSpecialItem.Number = .lSpecialItem
    ucSpecialMon.Number = .lSpecialMon
    ucRoomType.Number = .iType
    ucDeathRoom.Number = .lDeathRoom
    ucEnvironment.Number = .iInDoor
    ucTrainClass.Number = .iTrainClass
    ucShop.Number = .sShopItems
    ucNorth.Number = .lNorth
    ucSouth.Number = .lSouth
    ucWest.Number = .lWest
    ucEast.Number = .lEast
    ucNE.Number = .lNorthEast
    ucNW.Number = .lNorthWest
    ucSE.Number = .lSouthEast
    ucSW.Number = .lSouthWest
    ucUp.Number = .lUp
    ucDown.Number = .lDown
    chKDoor(0).Value = IIf(.lDN > 0, 1, 0)
    chKDoor(1).Value = IIf(.lDS > 0, 1, 0)
    chKDoor(2).Value = IIf(.lDW > 0, 1, 0)
    chKDoor(3).Value = IIf(.lDE > 0, 1, 0)
    chKDoor(4).Value = IIf(.lDNE > 0, 1, 0)
    chKDoor(5).Value = IIf(.lDSE > 0, 1, 0)
    chKDoor(6).Value = IIf(.lDNW > 0, 1, 0)
    chKDoor(7).Value = IIf(.lDSW > 0, 1, 0)
    chKDoor(8).Value = IIf(.lDU > 0, 1, 0)
    chKDoor(9).Value = IIf(.lDD > 0, 1, 0)
    ucD(0).Number = .lKN
    ucD(1).Number = .lKS
    ucD(2).Number = .lKW
    ucD(3).Number = .lKE
    ucD(4).Number = .lKNE
    ucD(5).Number = .lKSE
    ucD(6).Number = .lKNW
    ucD(7).Number = .lKSW
    ucD(8).Number = .lKU
    ucD(9).Number = .lKD
    txtP(0).Text = .lPN
    txtP(1).Text = .lPS
    txtP(2).Text = .lPW
    txtP(3).Text = .lPE
    txtP(4).Text = .lPNE
    txtP(5).Text = .lPSE
    txtP(6).Text = .lPNW
    txtP(7).Text = .lPSW
    txtP(8).Text = .lPU
    txtP(9).Text = .lPD
    txtB(0).Text = .lBN
    txtB(1).Text = .lBS
    txtB(2).Text = .lBW
    txtB(3).Text = .lBE
    txtB(4).Text = .lBNE
    txtB(5).Text = .lBSE
    txtB(6).Text = .lBNW
    txtB(7).Text = .lBSW
    txtB(8).Text = .lBU
    txtB(9).Text = .lBD
    cboGate(0).ListIndex = modMapFlags.GetMapFlag(cID, mapGate, North)
    cboGate(1).ListIndex = modMapFlags.GetMapFlag(cID, mapGate, South)
    cboGate(2).ListIndex = modMapFlags.GetMapFlag(cID, mapGate, West)
    cboGate(3).ListIndex = modMapFlags.GetMapFlag(cID, mapGate, East)
    cboGate(4).ListIndex = modMapFlags.GetMapFlag(cID, mapGate, NorthEast)
    cboGate(5).ListIndex = modMapFlags.GetMapFlag(cID, mapGate, SouthEast)
    cboGate(6).ListIndex = modMapFlags.GetMapFlag(cID, mapGate, NorthWest)
    cboGate(7).ListIndex = modMapFlags.GetMapFlag(cID, mapGate, SouthWest)
    cboGate(8).ListIndex = modMapFlags.GetMapFlag(cID, mapGate, Up)
    cboGate(9).ListIndex = modMapFlags.GetMapFlag(cID, mapGate, Down)
    lstODFood.Clear
    SplitFast .sOutDoorFood, Arr, ";"
    If UBound(Arr) < 1 And Arr(i) = "0" Then
    
    Else
        For i = LBound(Arr) To UBound(Arr)
            If Arr(i) <> "0" Then
                dd = GetItemID(, CLng(Val(Arr(i))))
                If dd <> 0 Then lstODFood.AddItem "(" & dbItems(dd).iID & ") " & dbItems(dd).sItemName
            End If
        Next
    End If
    txtScripting.Text = .sScript
    lblMap(0).Caption = "(" & .lRoomID & ") " & .sRoomTitle
    If .lNorth <> 0 Then
        lN(0).Visible = True
        picMap(1).Visible = True
        With dbMap(GetMapIndex(.lNorth))
            lblMap(1).Caption = "(" & .lRoomID & ") " & .sRoomTitle
        End With
    Else
        picMap(1).Visible = False
        lN(0).Visible = False
    End If
    If .lSouth <> 0 Then
        lN(1).Visible = True
        picMap(2).Visible = True
        With dbMap(GetMapIndex(.lSouth))
            lblMap(2).Caption = "(" & .lRoomID & ") " & .sRoomTitle
        End With
    Else
        picMap(2).Visible = False
        lN(1).Visible = False
    End If
    If .lEast <> 0 Then
        lN(2).Visible = True
        picMap(3).Visible = True
        With dbMap(GetMapIndex(.lEast))
            lblMap(3).Caption = "(" & .lRoomID & ") " & .sRoomTitle
        End With
    Else
        picMap(3).Visible = False
        lN(2).Visible = False
    End If
    If .lWest <> 0 Then
        lN(3).Visible = True
        picMap(4).Visible = True
        With dbMap(GetMapIndex(.lWest))
            lblMap(4).Caption = "(" & .lRoomID & ") " & .sRoomTitle
        End With
    Else
        picMap(4).Visible = False
        lN(3).Visible = False
    End If
    If .lNorthEast <> 0 Then
        lN(4).Visible = True
        picMap(5).Visible = True
        With dbMap(GetMapIndex(.lNorthEast))
            lblMap(5).Caption = "(" & .lRoomID & ") " & .sRoomTitle
        End With
    Else
        picMap(5).Visible = False
        lN(4).Visible = False
    End If
    If .lNorthWest <> 0 Then
        lN(5).Visible = True
        picMap(6).Visible = True
        With dbMap(GetMapIndex(.lNorthWest))
            lblMap(6).Caption = "(" & .lRoomID & ") " & .sRoomTitle
        End With
    Else
        picMap(6).Visible = False
        lN(5).Visible = False
    End If
    If .lSouthEast <> 0 Then
        lN(7).Visible = True
        picMap(7).Visible = True
        With dbMap(GetMapIndex(.lSouthEast))
            lblMap(7).Caption = "(" & .lRoomID & ") " & .sRoomTitle
        End With
    Else
        picMap(7).Visible = False
        lN(7).Visible = False
    End If
    If .lSouthWest <> 0 Then
        lN(6).Visible = True
        picMap(8).Visible = True
        With dbMap(GetMapIndex(.lSouthWest))
            lblMap(8).Caption = "(" & .lRoomID & ") " & .sRoomTitle
        End With
    Else
        picMap(8).Visible = False
        lN(6).Visible = False
    End If
    If .lUp <> 0 Then
        lblUp.Visible = True
        picMap(9).Visible = True
        With dbMap(GetMapIndex(.lUp))
            lblMap(9).Caption = "(" & .lRoomID & ") " & .sRoomTitle
        End With
    Else
        lblUp.Visible = False
        picMap(9).Visible = False
    End If
    If .lDown <> 0 Then
        lblDown.Visible = True
        picMap(10).Visible = True
        With dbMap(GetMapIndex(.lDown))
            lblMap(10).Caption = "(" & .lRoomID & ") " & .sRoomTitle
        End With
    Else
        lblDown.Visible = False
        picMap(10).Visible = False
    End If
    lstRooms.SetSelected lstRooms.Find("(" & dbMap(cID).lRoomID & ") " & dbMap(cID).sRoomTitle), True, True
    bU = False
End With
End Sub

