VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMap 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   12705
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   6
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   12705
   Begin VB.CommandButton cmdSave 
      Caption         =   "(save)"
      Height          =   255
      Left            =   10560
      TabIndex        =   0
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "(new)"
      Height          =   255
      Left            =   11640
      TabIndex        =   1
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox txtDoor 
      Height          =   255
      Index           =   9
      Left            =   13080
      TabIndex        =   161
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   4200
      Width           =   255
   End
   Begin VB.TextBox txtDoor 
      Height          =   255
      Index           =   8
      Left            =   13080
      TabIndex        =   160
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txtDoor 
      Height          =   255
      Index           =   7
      Left            =   13080
      TabIndex        =   159
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox txtDoor 
      Height          =   255
      Index           =   6
      Left            =   13080
      TabIndex        =   158
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3120
      Width           =   255
   End
   Begin VB.TextBox txtDoor 
      Height          =   255
      Index           =   5
      Left            =   13080
      TabIndex        =   157
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtDoor 
      Height          =   255
      Index           =   4
      Left            =   13080
      TabIndex        =   156
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txtDoor 
      Height          =   255
      Index           =   3
      Left            =   13080
      TabIndex        =   155
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txtDoor 
      Height          =   255
      Index           =   2
      Left            =   13080
      TabIndex        =   154
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtDoor 
      Height          =   255
      Index           =   1
      Left            =   13080
      TabIndex        =   153
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtDoor 
      Height          =   255
      Index           =   0
      Left            =   13080
      TabIndex        =   152
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   960
      Width           =   255
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   6495
      Left            =   3120
      TabIndex        =   6
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Basic Info"
      TabPicture(0)   =   "frmMap.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblLabel(23)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabel(21)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabel(20)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblLabel(19)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblLabel(17)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblLabel(15)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblLabel(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblLabel(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblLabel(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblLabel(18)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblLabel(13)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblLabel(14)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblLabel(16)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblLabel(24)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtTrainClass"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtSpecialItem"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cboType"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtDescription"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtSpecialMonster"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtMobGroup"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtMaxRegen"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtRoomTitle"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtID"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cboMonsters"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cboItems"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtShop"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cboShops"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtGold"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cboEnv"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cboRooms(11)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtDeathRoom"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "chkSafeRoom"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cboClass"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).ControlCount=   33
      TabCaption(1)   =   "Exits"
      TabPicture(1)   =   "frmMap.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblLabel(12)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblLabel(11)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblLabel(10)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblLabel(9)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblLabel(8)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblLabel(7)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblLabel(6)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblLabel(5)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblLabel(4)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblLabel(3)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtDown"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtUp"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtSE"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtSW"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtNE"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtNW"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtWest"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtEast"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtSouth"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtNorth"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cboRooms(0)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "chKDoor(0)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "picDoors(0)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "cboRooms(1)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "chKDoor(1)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "picDoors(1)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "cboRooms(2)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "chKDoor(2)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "picDoors(2)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "cboRooms(3)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "chKDoor(3)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "picDoors(3)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "cboRooms(4)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "chKDoor(4)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "picDoors(4)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "cboRooms(5)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "chKDoor(5)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "picDoors(5)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "cboRooms(6)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "chKDoor(6)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "picDoors(6)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "cboRooms(7)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "chKDoor(7)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "picDoors(7)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "cboRooms(8)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "chKDoor(8)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "picDoors(8)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "cboRooms(9)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "chKDoor(9)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "picDoors(9)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).ControlCount=   50
      TabCaption(2)   =   """Go"" Commands && Scripting"
      TabPicture(2)   =   "frmMap.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2(0)"
      Tab(2).Control(1)=   "Label2(1)"
      Tab(2).Control(2)=   "Label2(2)"
      Tab(2).Control(3)=   "Label2(3)"
      Tab(2).Control(4)=   "Label2(4)"
      Tab(2).Control(5)=   "Label3"
      Tab(2).Control(6)=   "Label4"
      Tab(2).Control(7)=   "txtGoRoom"
      Tab(2).Control(8)=   "txtGoCom"
      Tab(2).Control(9)=   "cboRooms(10)"
      Tab(2).Control(10)=   "txtPlayersDesc"
      Tab(2).Control(11)=   "txtCurRoomMessage"
      Tab(2).Control(12)=   "txtOtherRoomMessage"
      Tab(2).Control(13)=   "txtScripting"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Map"
      TabPicture(3)   =   "frmMap.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "shpRoom(0)"
      Tab(3).Control(1)=   "shpRoom(3)"
      Tab(3).Control(2)=   "shpRoom(9)"
      Tab(3).Control(3)=   "shpRoom(10)"
      Tab(3).Control(4)=   "shpRoom(1)"
      Tab(3).Control(5)=   "shpRoom(5)"
      Tab(3).Control(6)=   "linDir(5)"
      Tab(3).Control(7)=   "shpRoom(7)"
      Tab(3).Control(8)=   "shpRoom(4)"
      Tab(3).Control(9)=   "shpRoom(8)"
      Tab(3).Control(10)=   "shpRoom(2)"
      Tab(3).Control(11)=   "shpRoom(6)"
      Tab(3).Control(12)=   "lblRoom(7)"
      Tab(3).Control(13)=   "lblRoom(6)"
      Tab(3).Control(14)=   "lblRoom(5)"
      Tab(3).Control(15)=   "lblRoom(4)"
      Tab(3).Control(16)=   "lblRoom(3)"
      Tab(3).Control(17)=   "lblRoom(2)"
      Tab(3).Control(18)=   "lblRoom(1)"
      Tab(3).Control(19)=   "lblRoom(0)"
      Tab(3).Control(20)=   "linDir(9)"
      Tab(3).Control(21)=   "linDir(8)"
      Tab(3).Control(22)=   "linDir(6)"
      Tab(3).Control(23)=   "linDir(1)"
      Tab(3).Control(24)=   "linDir(4)"
      Tab(3).Control(25)=   "linDir(3)"
      Tab(3).Control(26)=   "linDir(7)"
      Tab(3).Control(27)=   "linDir(0)"
      Tab(3).Control(28)=   "linDir(2)"
      Tab(3).Control(29)=   "linDir(11)"
      Tab(3).Control(30)=   "linDir(10)"
      Tab(3).Control(31)=   "lblRoom(9)"
      Tab(3).Control(32)=   "lblRoom(8)"
      Tab(3).Control(33)=   "Label1"
      Tab(3).Control(34)=   "fraDirs"
      Tab(3).ControlCount=   35
      Begin VB.ComboBox cboClass 
         Height          =   240
         ItemData        =   "frmMap.frx":0070
         Left            =   -67680
         List            =   "frmMap.frx":0072
         Style           =   2  'Dropdown List
         TabIndex        =   184
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CheckBox chkSafeRoom 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -67680
         TabIndex        =   182
         Top             =   3600
         Width           =   255
      End
      Begin ServerEditor.NumOnlyText txtDeathRoom 
         Height          =   255
         Left            =   -68280
         TabIndex        =   181
         Top             =   3285
         Width           =   495
         _extentx        =   873
         _extenty        =   450
         font            =   "frmMap.frx":0074
         text            =   ""
         allowneg        =   0
         align           =   0
         maxlength       =   4
         enabled         =   -1
         backcolor       =   -2147483643
      End
      Begin VB.ComboBox cboRooms 
         Height          =   240
         Index           =   11
         ItemData        =   "frmMap.frx":009C
         Left            =   -67680
         List            =   "frmMap.frx":009E
         Style           =   2  'Dropdown List
         TabIndex        =   179
         Top             =   3285
         Width           =   2055
      End
      Begin VB.ComboBox cboEnv 
         Height          =   240
         ItemData        =   "frmMap.frx":00A0
         Left            =   -67680
         List            =   "frmMap.frx":00AD
         Style           =   2  'Dropdown List
         TabIndex        =   177
         Top             =   2910
         Width           =   2055
      End
      Begin VB.Frame fraDirs 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
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
         Height          =   1335
         Left            =   -69600
         TabIndex        =   175
         Top             =   1560
         Width           =   1215
         Begin VB.TextBox txtDirections 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   0
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            MultiLine       =   -1  'True
            TabIndex        =   176
            Text            =   "frmMap.frx":00DB
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.TextBox txtGold 
         Height          =   285
         Left            =   -67680
         MaxLength       =   5
         TabIndex        =   162
         Top             =   2520
         Width           =   2055
      End
      Begin VB.ComboBox cboShops 
         Height          =   240
         Left            =   -67680
         Style           =   2  'Dropdown List
         TabIndex        =   150
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtShop 
         Height          =   180
         Left            =   -67680
         TabIndex        =   149
         Text            =   "Text1"
         Top             =   2160
         Width           =   255
      End
      Begin VB.ComboBox cboItems 
         Height          =   240
         Left            =   -67680
         Style           =   2  'Dropdown List
         TabIndex        =   148
         Top             =   1800
         Width           =   2055
      End
      Begin VB.ComboBox cboMonsters 
         Height          =   240
         Left            =   -67680
         Style           =   2  'Dropdown List
         TabIndex        =   147
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtScripting 
         Height          =   2655
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   146
         Top             =   3720
         Width           =   9015
      End
      Begin VB.TextBox txtOtherRoomMessage 
         Height          =   615
         Left            =   -72960
         TabIndex        =   141
         Text            =   "0"
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtCurRoomMessage 
         Height          =   615
         Left            =   -72960
         TabIndex        =   139
         Text            =   "0"
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txtPlayersDesc 
         Height          =   615
         Left            =   -72960
         TabIndex        =   137
         Text            =   "0"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.ComboBox cboRooms 
         Height          =   240
         Index           =   10
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   135
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtGoCom 
         Height          =   255
         Left            =   -72960
         TabIndex        =   134
         Text            =   "0"
         Top             =   480
         Width           =   2775
      End
      Begin VB.PictureBox picDoors 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   4680
         ScaleHeight     =   435
         ScaleWidth      =   4635
         TabIndex        =   126
         Top             =   5880
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox txtPick 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   4090
            TabIndex        =   132
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkPick 
            Caption         =   "Pick Chance"
            Height          =   375
            Index           =   9
            Left            =   3240
            TabIndex        =   131
            Top             =   40
            Width           =   855
         End
         Begin VB.TextBox txtBash 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   2650
            TabIndex        =   130
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkBack 
            Caption         =   "Bash Str:"
            Height          =   375
            Index           =   9
            Left            =   1920
            TabIndex        =   129
            Top             =   40
            Width           =   735
         End
         Begin VB.ComboBox cboKeys 
            Enabled         =   0   'False
            Height          =   240
            Index           =   9
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   128
            Top             =   120
            Width           =   1290
         End
         Begin VB.CheckBox chkKey 
            Caption         =   "Key"
            Height          =   255
            Index           =   9
            Left            =   10
            TabIndex        =   127
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.CheckBox chKDoor 
         Caption         =   "Door"
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
         Index           =   9
         Left            =   3960
         TabIndex        =   125
         Top             =   6000
         Width           =   735
      End
      Begin VB.ComboBox cboRooms 
         Height          =   240
         Index           =   9
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   124
         Top             =   6000
         Width           =   2415
      End
      Begin VB.PictureBox picDoors 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   4680
         ScaleHeight     =   435
         ScaleWidth      =   4635
         TabIndex        =   117
         Top             =   5280
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox txtPick 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   4090
            TabIndex        =   123
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkPick 
            Caption         =   "Pick Chance"
            Height          =   375
            Index           =   8
            Left            =   3240
            TabIndex        =   122
            Top             =   40
            Width           =   855
         End
         Begin VB.TextBox txtBash 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   2650
            TabIndex        =   121
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkBack 
            Caption         =   "Bash Str:"
            Height          =   375
            Index           =   8
            Left            =   1920
            TabIndex        =   120
            Top             =   40
            Width           =   735
         End
         Begin VB.ComboBox cboKeys 
            Enabled         =   0   'False
            Height          =   240
            Index           =   8
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   119
            Top             =   120
            Width           =   1290
         End
         Begin VB.CheckBox chkKey 
            Caption         =   "Key"
            Height          =   255
            Index           =   8
            Left            =   10
            TabIndex        =   118
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.CheckBox chKDoor 
         Caption         =   "Door"
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
         Index           =   8
         Left            =   3960
         TabIndex        =   116
         Top             =   5400
         Width           =   735
      End
      Begin VB.ComboBox cboRooms 
         Height          =   240
         Index           =   8
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   5400
         Width           =   2415
      End
      Begin VB.PictureBox picDoors 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   4680
         ScaleHeight     =   435
         ScaleWidth      =   4635
         TabIndex        =   108
         Top             =   4680
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox txtPick 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   4090
            TabIndex        =   114
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkPick 
            Caption         =   "Pick Chance"
            Height          =   375
            Index           =   7
            Left            =   3240
            TabIndex        =   113
            Top             =   40
            Width           =   855
         End
         Begin VB.TextBox txtBash 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   2650
            TabIndex        =   112
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkBack 
            Caption         =   "Bash Str:"
            Height          =   375
            Index           =   7
            Left            =   1920
            TabIndex        =   111
            Top             =   40
            Width           =   735
         End
         Begin VB.ComboBox cboKeys 
            Enabled         =   0   'False
            Height          =   240
            Index           =   7
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   110
            Top             =   120
            Width           =   1290
         End
         Begin VB.CheckBox chkKey 
            Caption         =   "Key"
            Height          =   255
            Index           =   7
            Left            =   10
            TabIndex        =   109
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.CheckBox chKDoor 
         Caption         =   "Door"
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
         Index           =   7
         Left            =   3960
         TabIndex        =   107
         Top             =   4800
         Width           =   735
      End
      Begin VB.ComboBox cboRooms 
         Height          =   240
         Index           =   7
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   4800
         Width           =   2415
      End
      Begin VB.PictureBox picDoors 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   4680
         ScaleHeight     =   435
         ScaleWidth      =   4635
         TabIndex        =   99
         Top             =   4080
         Visible         =   0   'False
         Width           =   4695
         Begin VB.ComboBox cboKeys 
            Enabled         =   0   'False
            Height          =   240
            Index           =   6
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   104
            Top             =   120
            Width           =   1290
         End
         Begin VB.CheckBox chkBack 
            Caption         =   "Bash Str:"
            Height          =   375
            Index           =   6
            Left            =   1920
            TabIndex        =   103
            Top             =   40
            Width           =   735
         End
         Begin VB.TextBox txtBash 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   2650
            TabIndex        =   102
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkPick 
            Caption         =   "Pick Chance"
            Height          =   375
            Index           =   6
            Left            =   3240
            TabIndex        =   101
            Top             =   40
            Width           =   855
         End
         Begin VB.TextBox txtPick 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   4090
            TabIndex        =   100
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkKey 
            Caption         =   "Key"
            Height          =   255
            Index           =   6
            Left            =   10
            TabIndex        =   105
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.CheckBox chKDoor 
         Caption         =   "Door"
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
         Index           =   6
         Left            =   3960
         TabIndex        =   98
         Top             =   4200
         Width           =   735
      End
      Begin VB.ComboBox cboRooms 
         Height          =   240
         Index           =   6
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   4200
         Width           =   2415
      End
      Begin VB.PictureBox picDoors 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   4680
         ScaleHeight     =   435
         ScaleWidth      =   4635
         TabIndex        =   90
         Top             =   3480
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox txtPick 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   4090
            TabIndex        =   96
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkPick 
            Caption         =   "Pick Chance"
            Height          =   375
            Index           =   5
            Left            =   3240
            TabIndex        =   95
            Top             =   40
            Width           =   855
         End
         Begin VB.TextBox txtBash 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   2650
            TabIndex        =   94
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkBack 
            Caption         =   "Bash Str:"
            Height          =   375
            Index           =   5
            Left            =   1920
            TabIndex        =   93
            Top             =   40
            Width           =   735
         End
         Begin VB.ComboBox cboKeys 
            Enabled         =   0   'False
            Height          =   240
            Index           =   5
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   92
            Top             =   120
            Width           =   1290
         End
         Begin VB.CheckBox chkKey 
            Caption         =   "Key"
            Height          =   255
            Index           =   5
            Left            =   10
            TabIndex        =   91
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.CheckBox chKDoor 
         Caption         =   "Door"
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
         Index           =   5
         Left            =   3960
         TabIndex        =   89
         Top             =   3600
         Width           =   735
      End
      Begin VB.ComboBox cboRooms 
         Height          =   240
         Index           =   5
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   3600
         Width           =   2415
      End
      Begin VB.PictureBox picDoors 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   4680
         ScaleHeight     =   435
         ScaleWidth      =   4635
         TabIndex        =   81
         Top             =   2880
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox txtPick 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   4090
            TabIndex        =   87
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkPick 
            Caption         =   "Pick Chance"
            Height          =   375
            Index           =   4
            Left            =   3240
            TabIndex        =   86
            Top             =   40
            Width           =   855
         End
         Begin VB.TextBox txtBash 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   2650
            TabIndex        =   85
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkBack 
            Caption         =   "Bash Str:"
            Height          =   375
            Index           =   4
            Left            =   1920
            TabIndex        =   84
            Top             =   40
            Width           =   735
         End
         Begin VB.ComboBox cboKeys 
            Enabled         =   0   'False
            Height          =   240
            Index           =   4
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   83
            Top             =   120
            Width           =   1290
         End
         Begin VB.CheckBox chkKey 
            Caption         =   "Key"
            Height          =   255
            Index           =   4
            Left            =   10
            TabIndex        =   82
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.CheckBox chKDoor 
         Caption         =   "Door"
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
         Index           =   4
         Left            =   3960
         TabIndex        =   80
         Top             =   3000
         Width           =   735
      End
      Begin VB.ComboBox cboRooms 
         Height          =   240
         Index           =   4
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   3000
         Width           =   2415
      End
      Begin VB.PictureBox picDoors 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   4680
         ScaleHeight     =   435
         ScaleWidth      =   4635
         TabIndex        =   72
         Top             =   2280
         Visible         =   0   'False
         Width           =   4695
         Begin VB.ComboBox cboKeys 
            Enabled         =   0   'False
            Height          =   240
            Index           =   3
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   120
            Width           =   1290
         End
         Begin VB.CheckBox chkBack 
            Caption         =   "Bash Str:"
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   76
            Top             =   40
            Width           =   735
         End
         Begin VB.TextBox txtBash 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   2650
            TabIndex        =   75
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkPick 
            Caption         =   "Pick Chance"
            Height          =   375
            Index           =   3
            Left            =   3240
            TabIndex        =   74
            Top             =   40
            Width           =   855
         End
         Begin VB.TextBox txtPick 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   4090
            TabIndex        =   73
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkKey 
            Caption         =   "Key"
            Height          =   255
            Index           =   3
            Left            =   10
            TabIndex        =   78
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.CheckBox chKDoor 
         Caption         =   "Door"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   71
         Top             =   2400
         Width           =   735
      End
      Begin VB.ComboBox cboRooms 
         Height          =   240
         Index           =   3
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   2400
         Width           =   2415
      End
      Begin VB.PictureBox picDoors 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   4680
         ScaleHeight     =   435
         ScaleWidth      =   4635
         TabIndex        =   63
         Top             =   1680
         Visible         =   0   'False
         Width           =   4695
         Begin VB.ComboBox cboKeys 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   120
            Width           =   1290
         End
         Begin VB.CheckBox chkBack 
            Caption         =   "Bash Str:"
            Height          =   375
            Index           =   2
            Left            =   1920
            TabIndex        =   67
            Top             =   40
            Width           =   735
         End
         Begin VB.TextBox txtBash 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2650
            TabIndex        =   66
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkPick 
            Caption         =   "Pick Chance"
            Height          =   375
            Index           =   2
            Left            =   3240
            TabIndex        =   65
            Top             =   40
            Width           =   855
         End
         Begin VB.TextBox txtPick 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   4090
            TabIndex        =   64
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkKey 
            Caption         =   "Key"
            Height          =   255
            Index           =   2
            Left            =   10
            TabIndex        =   69
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.CheckBox chKDoor 
         Caption         =   "Door"
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
         Index           =   2
         Left            =   3960
         TabIndex        =   62
         Top             =   1800
         Width           =   735
      End
      Begin VB.ComboBox cboRooms 
         Height          =   240
         Index           =   2
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   1800
         Width           =   2415
      End
      Begin VB.PictureBox picDoors 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   4680
         ScaleHeight     =   435
         ScaleWidth      =   4635
         TabIndex        =   54
         Top             =   1080
         Visible         =   0   'False
         Width           =   4695
         Begin VB.ComboBox cboKeys 
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   120
            Width           =   1290
         End
         Begin VB.CheckBox chkBack 
            Caption         =   "Bash Str:"
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   58
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox txtBash 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2650
            TabIndex        =   57
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkPick 
            Caption         =   "Pick Chance"
            Height          =   375
            Index           =   1
            Left            =   3240
            TabIndex        =   56
            Top             =   40
            Width           =   855
         End
         Begin VB.TextBox txtPick 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   4090
            TabIndex        =   55
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkKey 
            Caption         =   "Key"
            Height          =   255
            Index           =   1
            Left            =   10
            TabIndex        =   60
            Top             =   120
            Width           =   600
         End
      End
      Begin VB.CheckBox chKDoor 
         Caption         =   "Door"
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
         Index           =   1
         Left            =   3960
         TabIndex        =   53
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox cboRooms 
         Height          =   240
         Index           =   1
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   1200
         Width           =   2415
      End
      Begin VB.PictureBox picDoors 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   4680
         ScaleHeight     =   435
         ScaleWidth      =   4635
         TabIndex        =   45
         Top             =   480
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox txtPick 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   4090
            TabIndex        =   51
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkPick 
            Caption         =   "Pick Chance"
            Height          =   375
            Index           =   0
            Left            =   3240
            TabIndex        =   50
            Top             =   40
            Width           =   855
         End
         Begin VB.TextBox txtBash 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2650
            TabIndex        =   49
            Text            =   "0"
            Top             =   120
            Width           =   495
         End
         Begin VB.CheckBox chkBack 
            Caption         =   "Bash Str:"
            Height          =   375
            Index           =   0
            Left            =   1920
            TabIndex        =   48
            Top             =   40
            Width           =   735
         End
         Begin VB.ComboBox cboKeys 
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   120
            Width           =   1290
         End
         Begin VB.CheckBox chkKey 
            Caption         =   "Key"
            Height          =   255
            Index           =   0
            Left            =   10
            TabIndex        =   46
            Top             =   120
            Width           =   600
         End
      End
      Begin VB.CheckBox chKDoor 
         Caption         =   "Door"
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
         Index           =   0
         Left            =   3960
         TabIndex        =   44
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox cboRooms 
         Height          =   240
         Index           =   0
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtNorth 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   32
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtSouth 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   31
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtEast 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   30
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtWest 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   29
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtNW 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   28
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox txtNE 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   27
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox txtSW 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   26
         Top             =   4200
         Width           =   495
      End
      Begin VB.TextBox txtSE 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   25
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox txtUp 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   24
         Top             =   5400
         Width           =   495
      End
      Begin VB.TextBox txtDown 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   23
         Top             =   6000
         Width           =   495
      End
      Begin VB.TextBox txtID 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73800
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtRoomTitle 
         Height          =   285
         Left            =   -73800
         TabIndex        =   13
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox txtMaxRegen 
         Height          =   285
         Left            =   -67680
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtMobGroup 
         Height          =   285
         Left            =   -67680
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtSpecialMonster 
         Height          =   180
         Left            =   -67680
         MaxLength       =   5
         TabIndex        =   10
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtDescription 
         Height          =   885
         Left            =   -73800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   2160
         Width           =   3975
      End
      Begin VB.ComboBox cboType 
         Height          =   240
         ItemData        =   "frmMap.frx":010F
         Left            =   -73800
         List            =   "frmMap.frx":0128
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtSpecialItem 
         Height          =   180
         Left            =   -67680
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtGoRoom 
         Height          =   180
         Left            =   -72960
         TabIndex        =   143
         Text            =   "Text1"
         Top             =   840
         Width           =   180
      End
      Begin ServerEditor.NumOnlyText txtTrainClass 
         Height          =   255
         Left            =   -68280
         TabIndex        =   186
         Top             =   3960
         Width           =   495
         _extentx        =   873
         _extenty        =   450
         font            =   "frmMap.frx":0181
         text            =   ""
         allowneg        =   0
         align           =   0
         maxlength       =   4
         enabled         =   -1
         backcolor       =   -2147483643
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Train Class:"
         Height          =   120
         Index           =   24
         Left            =   -69720
         TabIndex        =   185
         Top             =   3960
         Width           =   1080
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Safe Room:"
         Height          =   120
         Index           =   16
         Left            =   -69720
         TabIndex        =   183
         Top             =   3600
         Width           =   900
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Death Room:"
         Height          =   120
         Index           =   14
         Left            =   -69735
         TabIndex        =   180
         Top             =   3300
         Width           =   990
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Enviroment:"
         Height          =   120
         Index           =   13
         Left            =   -69720
         TabIndex        =   178
         Top             =   2925
         Width           =   990
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "YOU"
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
         Left            =   -72960
         TabIndex        =   164
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label lblRoom 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   8
         Left            =   -71040
         TabIndex        =   174
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblRoom 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   9
         Left            =   -71040
         TabIndex        =   173
         Top             =   2520
         Width           =   855
      End
      Begin VB.Line linDir 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   6
         Index           =   10
         X1              =   -70800
         X2              =   -70800
         Y1              =   1920
         Y2              =   2280
      End
      Begin VB.Line linDir 
         BorderColor     =   &H00FF00FF&
         BorderWidth     =   6
         Index           =   11
         X1              =   -70440
         X2              =   -70440
         Y1              =   2400
         Y2              =   2760
      End
      Begin VB.Line linDir 
         BorderColor     =   &H00000000&
         BorderWidth     =   6
         Index           =   2
         X1              =   -73680
         X2              =   -73080
         Y1              =   1560
         Y2              =   2160
      End
      Begin VB.Line linDir 
         BorderColor     =   &H00000000&
         BorderWidth     =   6
         Index           =   0
         X1              =   -73680
         X2              =   -73080
         Y1              =   3240
         Y2              =   2640
      End
      Begin VB.Line linDir 
         BorderColor     =   &H00000000&
         BorderWidth     =   6
         Index           =   7
         X1              =   -72765
         X2              =   -72765
         Y1              =   2760
         Y2              =   3120
      End
      Begin VB.Line linDir 
         BorderColor     =   &H00000000&
         BorderWidth     =   6
         Index           =   3
         X1              =   -72480
         X2              =   -71880
         Y1              =   2640
         Y2              =   3240
      End
      Begin VB.Line linDir 
         BorderColor     =   &H00000000&
         BorderWidth     =   6
         Index           =   4
         X1              =   -72480
         X2              =   -71880
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line linDir 
         BorderColor     =   &H00000000&
         BorderWidth     =   6
         Index           =   1
         X1              =   -72480
         X2              =   -71880
         Y1              =   2160
         Y2              =   1560
      End
      Begin VB.Line linDir 
         BorderColor     =   &H00000000&
         BorderWidth     =   6
         Index           =   6
         X1              =   -72765
         X2              =   -72765
         Y1              =   1680
         Y2              =   2040
      End
      Begin VB.Line linDir 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   6
         Index           =   8
         X1              =   -72960
         X2              =   -72960
         Y1              =   1920
         Y2              =   2280
      End
      Begin VB.Line linDir 
         BorderColor     =   &H00FF00FF&
         BorderWidth     =   6
         Index           =   9
         X1              =   -72600
         X2              =   -72600
         Y1              =   2520
         Y2              =   2880
      End
      Begin VB.Label lblRoom 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   0
         Left            =   -74400
         TabIndex        =   172
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblRoom 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   1
         Left            =   -73200
         TabIndex        =   171
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblRoom 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   2
         Left            =   -72000
         TabIndex        =   170
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblRoom 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   3
         Left            =   -74400
         TabIndex        =   169
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblRoom 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   4
         Left            =   -72000
         TabIndex        =   168
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblRoom 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   5
         Left            =   -74400
         TabIndex        =   167
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblRoom 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   6
         Left            =   -73200
         TabIndex        =   166
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblRoom 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   7
         Left            =   -72000
         TabIndex        =   165
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Gold:"
         Height          =   120
         Index           =   18
         Left            =   -69720
         TabIndex        =   163
         Top             =   2520
         Width           =   450
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Shop ID:"
         Height          =   120
         Index           =   2
         Left            =   -69720
         TabIndex        =   151
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Scripting:"
         Height          =   120
         Left            =   -74880
         TabIndex        =   145
         Top             =   3480
         Width           =   900
      End
      Begin VB.Label Label3 
         Caption         =   $"frmMap.frx":01A9
         Height          =   855
         Left            =   -70080
         TabIndex        =   144
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Other Room's Message:"
         Height          =   120
         Index           =   4
         Left            =   -74880
         TabIndex        =   142
         Top             =   2640
         Width           =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Current Room Message:"
         Height          =   120
         Index           =   3
         Left            =   -74880
         TabIndex        =   140
         Top             =   1920
         Width           =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Players Message:"
         Height          =   120
         Index           =   2
         Left            =   -74880
         TabIndex        =   138
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Go To Room:"
         Height          =   120
         Index           =   1
         Left            =   -74880
         TabIndex        =   136
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Go Command:"
         Height          =   120
         Index           =   0
         Left            =   -74880
         TabIndex        =   133
         Top             =   480
         Width           =   990
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "North:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   42
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "South"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   240
         TabIndex        =   41
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "East"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   240
         TabIndex        =   40
         Top             =   1800
         Width           =   330
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "West"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   240
         TabIndex        =   39
         Top             =   2400
         Width           =   405
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "NW"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   240
         TabIndex        =   38
         Top             =   3000
         Width           =   315
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "NE"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   240
         TabIndex        =   37
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "SW"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   240
         TabIndex        =   36
         Top             =   4200
         Width           =   300
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "SE"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   240
         TabIndex        =   35
         Top             =   4800
         Width           =   225
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Up"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   240
         TabIndex        =   34
         Top             =   5400
         Width           =   210
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Down"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   240
         TabIndex        =   33
         Top             =   6000
         Width           =   420
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "RoomID:"
         Height          =   120
         Index           =   0
         Left            =   -74880
         TabIndex        =   22
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Room Title:"
         Height          =   120
         Index           =   1
         Left            =   -74880
         TabIndex        =   21
         Top             =   1680
         Width           =   990
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Max Regen:"
         Height          =   120
         Index           =   15
         Left            =   -69720
         TabIndex        =   20
         Top             =   720
         Width           =   900
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Mob Group:"
         Height          =   120
         Index           =   17
         Left            =   -69720
         TabIndex        =   19
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Special Regen Monster:"
         Height          =   120
         Index           =   19
         Left            =   -69720
         TabIndex        =   18
         Top             =   1440
         Width           =   1980
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   120
         Index           =   20
         Left            =   -74880
         TabIndex        =   17
         Top             =   2160
         Width           =   1080
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Room Type:"
         Height          =   120
         Index           =   21
         Left            =   -74880
         TabIndex        =   16
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Special Item:"
         Height          =   120
         Index           =   23
         Left            =   -69720
         TabIndex        =   15
         Top             =   1800
         Width           =   1170
      End
      Begin VB.Shape shpRoom 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   6
         Left            =   -72000
         Shape           =   3  'Circle
         Top             =   960
         Width           =   855
      End
      Begin VB.Shape shpRoom 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   2
         Left            =   -72000
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   855
      End
      Begin VB.Shape shpRoom 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   8
         Left            =   -72000
         Shape           =   3  'Circle
         Top             =   3120
         Width           =   855
      End
      Begin VB.Shape shpRoom 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   4
         Left            =   -73200
         Shape           =   3  'Circle
         Top             =   3120
         Width           =   855
      End
      Begin VB.Shape shpRoom 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   7
         Left            =   -74400
         Shape           =   3  'Circle
         Top             =   3120
         Width           =   855
      End
      Begin VB.Line linDir 
         BorderColor     =   &H00000000&
         BorderWidth     =   6
         Index           =   5
         X1              =   -73680
         X2              =   -73080
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Shape shpRoom 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   5
         Left            =   -74400
         Shape           =   3  'Circle
         Top             =   960
         Width           =   855
      End
      Begin VB.Shape shpRoom 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   1
         Left            =   -73200
         Shape           =   3  'Circle
         Top             =   960
         Width           =   855
      End
      Begin VB.Shape shpRoom 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   10
         Left            =   -71040
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   855
      End
      Begin VB.Shape shpRoom 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   9
         Left            =   -71040
         Shape           =   3  'Circle
         Top             =   1440
         Width           =   855
      End
      Begin VB.Shape shpRoom 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   3
         Left            =   -74400
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   855
      End
      Begin VB.Shape shpRoom 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   0
         Left            =   -73200
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdGoto 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   6360
      Width           =   255
   End
   Begin VB.TextBox txtGoTo 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   6360
      Width           =   1695
   End
   Begin VB.ListBox lstRooms 
      Height          =   6180
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "GOTO:"
      Height          =   120
      Index           =   22
      Left            =   120
      TabIndex        =   4
      Top             =   6360
      Width           =   450
   End
End
Attribute VB_Name = "frmMap"
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
Rem***************                frmMap                          **********************
Rem***************                ServerEditor                    **********************
Rem***************                Editor.vbp                      **********************
Rem*************************************************************************************
Rem*************************************************************************************
Dim RSmID As Integer
Dim FillingBox As Boolean
Dim bFillMap As Boolean

Sub FillCBOS()
Dim i As Long
With RSMap
    For i = cboRooms.LBound To cboRooms.UBound
        cboRooms(i).Clear
        If i <> 11 Then
            cboRooms(i).AddItem "(0) None"
            cboRooms(i).AddItem "(???) Undefined"
        End If
        .MoveFirst
        Do
            cboRooms(i).AddItem "(" & !RoomID & ") " & !RoomTitle
            .MoveNext
        Loop Until .EOF
    Next
End With
With RSItem
    For i = cboKeys.LBound To cboKeys.UBound
        cboKeys(i).Clear
        If i = 0 Then
            cboItems.Clear
            cboItems.AddItem "(0) None"
        End If
        cboKeys(i).AddItem "(0) None"
        .MoveFirst
        Do
            If i = 0 Then cboItems.AddItem "(" & !ID & ") " & !ItemName
            If !Worn = "key" Then
                cboKeys(i).AddItem "(" & !ID & ") " & !ItemName
                .MoveNext
            ElseIf Not .EOF Then
                .MoveNext
            End If
        Loop Until .EOF
    Next
End With
With RSMonster
    .MoveFirst
    cboMonsters.Clear
    cboMonsters.AddItem "(0) None"
    Do
        cboMonsters.AddItem "(" & !ID & ") " & !MonsterName
        .MoveNext
    Loop Until .EOF
End With
With RSShops
    .MoveFirst
    cboShops.Clear
    cboShops.AddItem "(0) None"
    Do
        cboShops.AddItem "(" & !ID & ") " & !ShopName
        .MoveNext
    Loop Until .EOF
End With
With RSClass
    .MoveFirst
    cboClass.Clear
    cboClass.AddItem "(0) None"
    Do
        cboClass.AddItem "(" & !ID & ") " & !Name
        .MoveNext
    Loop Until .EOF
End With
End Sub

Private Sub cboClass_Change()
txtTrainClass.Text = Mid$(cboClass.list(cboClass.ListIndex), 2, InStr(1, cboClass.list(cboClass.ListIndex), ")") - 2)
End Sub

Private Sub cboClass_Click()
cboClass_Change
End Sub

Private Sub cboItems_Change()
txtSpecialItem.Text = Mid$(cboItems.list(cboItems.ListIndex), 2, InStr(1, cboItems.list(cboItems.ListIndex), ")") - 2)
End Sub

Private Sub cboItems_Click()
txtSpecialItem.Text = Mid$(cboItems.list(cboItems.ListIndex), 2, InStr(1, cboItems.list(cboItems.ListIndex), ")") - 2)
End Sub

Private Sub cboKeys_Change(Index As Integer)
If cboKeys(Index).list(cboKeys(Index).ListIndex) = "(0) None" And chKDoor(Index).Value = 1 Then
    txtDoor(Index).Text = "1"
ElseIf chKDoor(Index).Value = 1 Then
    txtDoor(Index).Text = "2"
End If
End Sub

Private Sub cboKeys_Click(Index As Integer)
If cboKeys(Index).list(cboKeys(Index).ListIndex) = "(0) None" And chKDoor(Index).Value = 1 Then
    txtDoor(Index).Text = "1"
ElseIf chKDoor(Index).Value = 1 Then
    txtDoor(Index).Text = "2"
End If
End Sub

Private Sub cboMonsters_Change()
txtSpecialMonster.Text = Mid$(cboMonsters.list(cboMonsters.ListIndex), 2, InStr(1, cboMonsters.list(cboMonsters.ListIndex), ")") - 2)
End Sub

Private Sub cboMonsters_Click()
txtSpecialMonster.Text = Mid$(cboMonsters.list(cboMonsters.ListIndex), 2, InStr(1, cboMonsters.list(cboMonsters.ListIndex), ")") - 2)
End Sub

Private Sub cboRooms_Change(Index As Integer)
Select Case Index
    Case 0
        txtNorth.Text = Mid$(cboRooms(Index).list(cboRooms(Index).ListIndex), 2, InStr(1, cboRooms(Index).list(cboRooms(Index).ListIndex), ")") - 2)
    Case 1
        txtSouth.Text = Mid$(cboRooms(Index).list(cboRooms(Index).ListIndex), 2, InStr(1, cboRooms(Index).list(cboRooms(Index).ListIndex), ")") - 2)
    Case 2
        txtEast.Text = Mid$(cboRooms(Index).list(cboRooms(Index).ListIndex), 2, InStr(1, cboRooms(Index).list(cboRooms(Index).ListIndex), ")") - 2)
    Case 3
        txtWest.Text = Mid$(cboRooms(Index).list(cboRooms(Index).ListIndex), 2, InStr(1, cboRooms(Index).list(cboRooms(Index).ListIndex), ")") - 2)
    Case 4
        txtNW.Text = Mid$(cboRooms(Index).list(cboRooms(Index).ListIndex), 2, InStr(1, cboRooms(Index).list(cboRooms(Index).ListIndex), ")") - 2)
    Case 5
        txtNE.Text = Mid$(cboRooms(Index).list(cboRooms(Index).ListIndex), 2, InStr(1, cboRooms(Index).list(cboRooms(Index).ListIndex), ")") - 2)
    Case 6
        txtSW.Text = Mid$(cboRooms(Index).list(cboRooms(Index).ListIndex), 2, InStr(1, cboRooms(Index).list(cboRooms(Index).ListIndex), ")") - 2)
    Case 7
        txtSE.Text = Mid$(cboRooms(Index).list(cboRooms(Index).ListIndex), 2, InStr(1, cboRooms(Index).list(cboRooms(Index).ListIndex), ")") - 2)
    Case 8
        txtUp.Text = Mid$(cboRooms(Index).list(cboRooms(Index).ListIndex), 2, InStr(1, cboRooms(Index).list(cboRooms(Index).ListIndex), ")") - 2)
    Case 9
        txtDown.Text = Mid$(cboRooms(Index).list(cboRooms(Index).ListIndex), 2, InStr(1, cboRooms(Index).list(cboRooms(Index).ListIndex), ")") - 2)
    Case 10
        txtGoRoom.Text = Mid$(cboRooms(Index).list(cboRooms(Index).ListIndex), 2, InStr(1, cboRooms(Index).list(cboRooms(Index).ListIndex), ")") - 2)
    Case 11
        txtDeathRoom.Text = Mid$(cboRooms(Index).list(cboRooms(Index).ListIndex), 2, InStr(1, cboRooms(Index).list(cboRooms(Index).ListIndex), ")") - 2)
End Select
End Sub

Private Sub cboRooms_Click(Index As Integer)
cboRooms_Change Index
End Sub

Private Sub cboShops_Change()
txtShop.Text = Mid$(cboShops.list(cboShops.ListIndex), 2, InStr(1, cboShops.list(cboShops.ListIndex), ")") - 2)
End Sub

Private Sub cboShops_Click()
txtShop.Text = Mid$(cboShops.list(cboShops.ListIndex), 2, InStr(1, cboShops.list(cboShops.ListIndex), ")") - 2)
End Sub

Private Sub cboType_Click()
'cmdEditShop.Enabled = False
cboShops.Enabled = False
cboMonsters.Enabled = False
Select Case cboType.Text
    Case "1 - Shop":
        cboShops.Enabled = True
        cboMonsters.Enabled = True
    Case "2 - Trainer":
        cboMonsters.Enabled = True
    Case "4 - Boss":
        cboMonsters.Enabled = True
    Case "5 - Bank":
        cboMonsters.Enabled = True
End Select
End Sub

Private Sub chkAutoExits_Click()
txtRoomExits.Enabled = False
If chkAutoExits.Value = 0 Then txtRoomExits.Enabled = True
If chkAutoExits.Value = 1 Then GetExits
End Sub

Private Sub chkBack_Click(Index As Integer)
txtBash(Index).Enabled = chkBack(Index).Value
End Sub

Private Sub chKDoor_Click(Index As Integer)
picDoors(Index).Visible = chKDoor(Index).Value
Select Case chKDoor(Index).Value
    Case 0
        txtDoor(Index).Text = "0"
        chkKey(Index).Value = 0
        chkPick(Index).Value = 0
        txtBash(Index).Text = "0"
        txtPick(Index).Text = "0"
    Case 1
        If chkKey(Index).Value = 0 And chkPick(Index).Value = 0 Then
            txtDoor(Index).Text = "1"
        Else
            txtDoor(Index).Text = "2"
        End If
End Select
End Sub

Private Sub chkKey_Click(Index As Integer)
cboKeys(Index).Enabled = chkKey(Index).Value
If chkKey(Index).Value = 0 Then
    txtDoor(Index).Text = "1"
Else
    txtDoor(Index).Text = "2"
End If
End Sub

Private Sub chkPick_Click(Index As Integer)
txtPick(Index).Enabled = chkPick(Index).Value
End Sub



Private Sub cmdGoto_Click()
SaveMap
With RSMap
    .MoveFirst
    Do
        If CInt(!RoomID) = CInt(txtGoTo.Text) Then
            RSmID = !RoomID
            DrawMapOut
            FillMap
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
End Sub

Private Sub cmdNew_Click()
SaveMap
With RSMap
    .MoveLast
    Dim x As Integer
    x = !RoomID
    x = x + 1
    .AddNew
    !RoomID = x
    !North = 0
    !South = 0
    !East = 0
    !West = 0
    !NorthWest = 0
    !NorthEast = 0
    !SouthWest = 0
    !SouthEast = 0
    !SpecialItem = "0"
    !UP = 0
    !Down = 0
    !Items = "0"
    !RoomTitle = "New Room"
    !RoomDesc = "None"
    !Monsters = "0"
    !MaxRegen = "1"
    !Type = "0"
    !ShopItems = "0"
    !MobGroup = "0"
    !Gold = "0"
    !SpecialMon = "0"
    !DN = "0"
    !DS = "0"
    !DE = "0"
    !DW = "0"
    !DNW = "0"
    !DNE = "0"
    !DSW = "0"
    !DSE = "0"
    !DU = "0"
    !DD = "0"
    !KN = "0"
    !KS = "0"
    !KE = "0"
    !KW = "0"
    !KNW = "0"
    !KNE = "0"
    !KSW = "0"
    !KSE = "0"
    !KU = "0"
    !KD = "0"
    !BN = "-1"
    !BS = "-1"
    !BE = "-1"
    !BW = "-1"
    !BNW = "-1"
    !BNE = "-1"
    !BSW = "-1"
    !BSE = "-1"
    !BU = "-1"
    !BD = "-1"
    !PN = "-1"
    !PS = "-1"
    !PE = "-1"
    !PW = "-1"
    !PNW = "-1"
    !PNE = "-1"
    !PSW = "-1"
    !PSE = "-1"
    !PU = "-1"
    !PD = "-1"
    !Scripting = "0"
    !GoCom = "0"
    !GoDesc = "0"
    !Hidden = "0"
    !GoOthersDescAway = "0"
    !GoOthersDescTo = "0"
    !InDoor = "0"
    !TrainClass = "0"
    !DeathRoom = "232"
    !SafeRoom = "0"
    .Update
    RSmID = x
    DrawMapOut
    FillMap
End With
FillListBox
End Sub

Private Sub cmdSave_Click()
SaveMap
FillListBox
SetLstSelected lstRooms, "(" & txtID.Text & ") " & txtRoomTitle.Text
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
RSmID = 1
FillListBox
DrawMapOut
FillMap
End Sub

Private Sub lblRoom_Click(Index As Integer)
'SaveMap
Select Case Index
    Case 0:
        GetNewID "nw"
    Case 1:
        GetNewID "n"
    Case 2:
        GetNewID "ne"
    Case 3:
        GetNewID "w"
    Case 4:
        GetNewID "e"
    Case 5:
        GetNewID "sw"
    Case 6:
        GetNewID "s"
    Case 7:
        GetNewID "se"
    Case 8:
        GetNewID "u"
    Case 9:
        GetNewID "d"
End Select
DrawMapOut
FillMap
End Sub

Private Sub lstRooms_Click()
If FillingBox = False Then
    With RSMap
        .MoveFirst
        Do
            If "(" & !RoomID & ") " & !RoomTitle = lstRooms.Text Then
                RSmID = !RoomID
                DrawMapOut
                FillMap
                Exit Do
            ElseIf Not .EOF Then
                .MoveNext
            End If
        Loop Until .EOF
    End With
End If
End Sub

Private Sub txtDeathRoom_Change()
If bFillMap Then Exit Sub
On Error Resume Next
SetListIndex cboRooms(11), "(" & txtDeathRoom.Text & ") " & GetMapName(CInt(txtDeathRoom.Text))
End Sub

Private Sub txtDoor_Change(Index As Integer)
Select Case CInt(txtDoor(Index).Text)
    Case 0
        If chKDoor(Index).Value <> 0 Then chKDoor(Index).Value = 0
    Case 1 To 3
        If chKDoor(Index).Value <> 1 Then chKDoor(Index).Value = 1
End Select
End Sub

Private Sub txtDown_Change()
If bFillMap Then Exit Sub
On Error Resume Next
If GetMapName(CInt(txtDown.Text)) = "" And txtDown.Text <> "0" Then
    SetListIndex cboRooms(9), "(???) Undefined"
ElseIf txtDown.Text <> "0" Then
    SetListIndex cboRooms(9), "(" & txtDown.Text & ") " & GetMapName(CInt(txtDown.Text))
ElseIf txtDown.Text <> "" Then
    SetListIndex cboRooms(9), "(0) None"
End If
End Sub

Private Sub txtEast_Change()
If bFillMap Then Exit Sub
On Error Resume Next
If GetMapName(CInt(txtEast.Text)) = "" And txtEast.Text <> "0" Then
    SetListIndex cboRooms(2), "(???) Undefined"
ElseIf txtEast.Text <> "0" Then
    SetListIndex cboRooms(2), "(" & txtEast.Text & ") " & GetMapName(CInt(txtEast.Text))
ElseIf txtEast.Text <> "" Then
    SetListIndex cboRooms(2), "(0) None"
End If
End Sub

Private Sub txtGoRoom_Change()
If bFillMap Then Exit Sub
On Error Resume Next
If GetMapName(CInt(txtGoRoom.Text)) = "" And txtGoRoom.Text <> "0" Then
    SetListIndex cboRooms(10), "(???) Undefined"
ElseIf txtGoRoom.Text <> "0" Then
    SetListIndex cboRooms(10), "(" & txtGoRoom.Text & ") " & GetMapName(CInt(txtGoRoom.Text))
ElseIf txtGoRoom.Text <> "" Then
    SetListIndex cboRooms(10), "(0) None"
End If
End Sub

Private Sub txtGoTo_Change()
If Not IsNumeric(txtGoTo.Text) Then txtGoTo.Text = ""
End Sub

Sub FillListBox()
MousePointer = vbHourglass
FillingBox = True
lstRooms.Clear
Dim tID%
If lstRooms.ListIndex <> -1 Then tID% = lstRooms.ListIndex
With RSMap
    .MoveFirst
    Do
        lstRooms.AddItem "(" & !RoomID & ") " & !RoomTitle
        .MoveNext
    Loop Until .EOF
End With
lstRooms.Selected(tID%) = True
FillingBox = False
MousePointer = vbDefault
End Sub

Sub SaveMap()
MousePointer = vbHourglass
If IsSaveAble = False Then
    MsgBox "Not everything is filled out correctly, or not everything is filled in." & vbCrLf & "Please check to make sure everything is correct." & vbCrLf & "Save will not continue.", vbCritical, "Error"
    MousePointer = vbDefault
    Exit Sub
End If
With RSMap
    .MoveFirst
    Do
        If CInt(!RoomID) = RSmID Then
            .Edit
            !RoomTitle = txtRoomTitle.Text
            !North = txtNorth.Text
            !South = txtSouth.Text
            !East = txtEast.Text
            !SpecialItem = txtSpecialItem.Text
            !West = txtWest.Text
            !NorthWest = txtNW.Text
            !NorthEast = txtNE.Text
            !SouthWest = txtSW.Text
            !SouthEast = txtSE.Text
            !UP = txtUp.Text
            !Down = txtDown.Text
            !MaxRegen = txtMaxRegen.Text
            !InDoor = Left$(cboEnv.list(cboEnv.ListIndex), 1)
            !ShopItems = txtShop.Text
            !MobGroup = txtMobGroup.Text
            !Gold = txtGold.Text
            !SpecialMon = txtSpecialMonster.Text
            !RoomDesc = txtDescription.Text
            !Type = Left$(cboType.list(cboType.ListIndex), 1)
            !TrainClass = txtTrainClass.Text
            !DN = txtDoor(0).Text
            !DS = txtDoor(1).Text
            !DE = txtDoor(2).Text
            !DW = txtDoor(3).Text
            !DNW = txtDoor(4).Text
            !DNE = txtDoor(5).Text
            !DSW = txtDoor(6).Text
            !DSE = txtDoor(7).Text
            !DU = txtDoor(8).Text
            !DD = txtDoor(9).Text
            !KN = Mid$(cboKeys(0).list(cboKeys(0).ListIndex), 2, InStr(1, cboKeys(0).list(cboKeys(0).ListIndex), ")") - 2)
            !KS = Mid$(cboKeys(1).list(cboKeys(1).ListIndex), 2, InStr(1, cboKeys(1).list(cboKeys(1).ListIndex), ")") - 2)
            !KE = Mid$(cboKeys(2).list(cboKeys(2).ListIndex), 2, InStr(1, cboKeys(2).list(cboKeys(2).ListIndex), ")") - 2)
            !KW = Mid$(cboKeys(3).list(cboKeys(3).ListIndex), 2, InStr(1, cboKeys(3).list(cboKeys(3).ListIndex), ")") - 2)
            !KNW = Mid$(cboKeys(4).list(cboKeys(4).ListIndex), 2, InStr(1, cboKeys(4).list(cboKeys(4).ListIndex), ")") - 2)
            !KNE = Mid$(cboKeys(5).list(cboKeys(5).ListIndex), 2, InStr(1, cboKeys(5).list(cboKeys(5).ListIndex), ")") - 2)
            !KSW = Mid$(cboKeys(6).list(cboKeys(6).ListIndex), 2, InStr(1, cboKeys(6).list(cboKeys(6).ListIndex), ")") - 2)
            !KSE = Mid$(cboKeys(7).list(cboKeys(7).ListIndex), 2, InStr(1, cboKeys(7).list(cboKeys(7).ListIndex), ")") - 2)
            !KU = Mid$(cboKeys(8).list(cboKeys(8).ListIndex), 2, InStr(1, cboKeys(8).list(cboKeys(8).ListIndex), ")") - 2)
            !KD = Mid$(cboKeys(9).list(cboKeys(9).ListIndex), 2, InStr(1, cboKeys(9).list(cboKeys(9).ListIndex), ")") - 2)
            !BN = txtBash(0).Text
            !BS = txtBash(1).Text
            !BE = txtBash(2).Text
            !BW = txtBash(3).Text
            !BNW = txtBash(4).Text
            !BNE = txtBash(5).Text
            !BSW = txtBash(6).Text
            !BSE = txtBash(7).Text
            !BU = txtBash(8).Text
            !BD = txtBash(9).Text
            !PN = txtPick(0).Text
            !PS = txtPick(1).Text
            !PE = txtPick(2).Text
            !PW = txtPick(3).Text
            !PNW = txtPick(4).Text
            !PNE = txtPick(5).Text
            !PSW = txtPick(6).Text
            !PSE = txtPick(7).Text
            !PU = txtPick(8).Text
            !PD = txtPick(9).Text
            !Scripting = txtScripting.Text
            !GoCom = txtGoCom.Text
            !GoRoom = txtGoRoom.Text
            !GoDesc = txtPlayersDesc.Text
            !GoOthersDescAway = txtCurRoomMessage.Text
            !GoOthersDescTo = txtOtherRoomMessage.Text
            !DeathRoom = Mid$(cboRooms(11).list(cboRooms(11).ListIndex), 2, InStr(1, cboRooms(11).list(cboRooms(11).ListIndex), ")") - 2)
            !SafeRoom = chkSafeRoom.Value
            .Update
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
DrawMapOut
MousePointer = vbDefault
End Sub

Sub FillMap()
MousePointer = vbHourglass
bFillMap = True
FillCBOS
With RSMap
    .MoveFirst
    Do
        If CInt(!RoomID) = RSmID Then
            txtID.Text = !RoomID
            txtRoomTitle.Text = !RoomTitle
            txtNorth.Text = !North
            txtSouth.Text = !South
            txtEast.Text = !East
            txtSpecialItem.Text = !SpecialItem
            txtWest.Text = !West
            txtNW.Text = !NorthWest
            txtNE.Text = !NorthEast
            txtSW.Text = !SouthWest
            txtSE.Text = !SouthEast
            txtUp.Text = !UP
            txtDown.Text = !Down
            txtMaxRegen.Text = !MaxRegen
            txtMobGroup.Text = !MobGroup
            txtGold.Text = !Gold
            txtSpecialMonster.Text = !SpecialMon
            txtDescription.Text = !RoomDesc
            txtGoCom.Text = !GoCom
            txtGoRoom.Text = !GoRoom
            txtPlayersDesc.Text = !GoDesc
            txtCurRoomMessage.Text = !GoOthersDescAway
            txtOtherRoomMessage.Text = !GoOthersDescTo
            txtScripting.Text = !Scripting
            txtShop.Text = !ShopItems
            txtTrainClass.Text = !TrainClass
            If txtTrainClass.Text <> "0" Then
                SetListIndex cboClass, "(" & txtTrainClass.Text & ") " & GetClassName(CInt(txtTrainClass.Text))
            Else
                SetListIndex cboClass, "(0) None"
            End If
            chkSafeRoom.Value = !SafeRoom
            For i = chKDoor.LBound To chKDoor.UBound
                txtDoor(i).Text = "0"
                chKDoor(i).Value = 0
                chkKey(i).Value = 0
                chkBack(i).Value = 0
                chkPick(i).Value = 0
            Next
            '----
            txtDoor(0).Text = !DN
            If !KN <> "0" Then
                chkKey(0).Value = 1
                SetListIndex cboKeys(0), "(" & !KN & ") " & GetItemName(CInt(!KN))
            Else
                SetListIndex cboKeys(0), "(0) None"
            End If
            txtBash(0).Text = !BN
            If !BN <> "-1" Then chkBack(0).Value = 1
            txtPick(0).Text = !PN
            If !PN <> "-1" Then chkPick(0).Value = 1
            '----
            
            '----
            txtDoor(1).Text = !DS
            If !KS <> "0" Then
                chkKey(1).Value = 1
                SetListIndex cboKeys(1), "(" & !KS & ") " & GetItemName(CInt(!KS))
            Else
                SetListIndex cboKeys(1), "(0) None"
            End If
            txtBash(1).Text = !BS
            If !BS <> "-1" Then chkBack(1).Value = 1
            txtPick(1).Text = !PS
            If !PS <> "-1" Then chkPick(1).Value = 1
            '----
            
            '----
            txtDoor(2).Text = !DE
            If !KE <> "0" Then
                chkKey(2).Value = 1
                SetListIndex cboKeys(2), "(" & !KE & ") " & GetItemName(CInt(!KE))
            Else
                SetListIndex cboKeys(2), "(0) None"
            End If
            txtBash(2).Text = !BE
            If !BE <> "-1" Then chkBack(2).Value = 1
            txtPick(2).Text = !PE
            If !PE <> "-1" Then chkPick(2).Value = 1
            '----
            Select Case !InDoor
'            0 - Outdoor
'1 - Indoor
'2 - Underground

                Case "0"
                    SetListIndex cboEnv, "0 - Outdoor"
                Case "1"
                    SetListIndex cboEnv, "1 - Indoor"
                Case "2"
                    SetListIndex cboEnv, "2 - Underground"
            End Select
            '----
            txtDoor(3).Text = !DW
            If !KW <> "0" Then
                chkKey(3).Value = 1
                SetListIndex cboKeys(3), "(" & !KW & ") " & GetItemName(CInt(!KW))
            Else
                SetListIndex cboKeys(3), "(0) None"
            End If
            txtBash(3).Text = !BW
            If !BW <> "-1" Then chkBack(3).Value = 1
            txtPick(3).Text = !PW
            If !PW <> "-1" Then chkPick(3).Value = 1
            '----
            
            '----
            txtDoor(4).Text = !DNW
            If !KNW <> "0" Then
                chkKey(4).Value = 1
                SetListIndex cboKeys(4), "(" & !KNW & ") " & GetItemName(CInt(!KNW))
            Else
                SetListIndex cboKeys(4), "(0) None"
            End If
            txtBash(4).Text = !BNW
            If !BNW <> "-1" Then chkBack(4).Value = 1
            txtPick(4).Text = !PNW
            If !PNW <> "-1" Then chkPick(4).Value = 1
            '----
            
            '----
            txtDoor(5).Text = !DNE
            If !KNE <> "0" Then
                chkKey(5).Value = 1
                SetListIndex cboKeys(5), "(" & !KNE & ") " & GetItemName(CInt(!KNE))
            
            Else
                SetListIndex cboKeys(5), "(0) None"
            End If
            txtBash(5).Text = !BNE
            If !BNE <> "-1" Then chkBack(5).Value = 1
            txtPick(5).Text = !PNE
            If !PNE <> "-1" Then chkPick(5).Value = 1
            '----
            
            '----
            txtDoor(6).Text = !DSW
            If !KSW <> "0" Then
                chkKey(6).Value = 1
                SetListIndex cboKeys(6), "(" & !KSW & ") " & GetItemName(CInt(!KSW))
            Else
                SetListIndex cboKeys(6), "(0) None"
            End If
            txtBash(6).Text = !BSW
            If !BSW <> "-1" Then chkBack(6).Value = 1
            txtPick(6).Text = !PSW
            If !PSW <> "-1" Then chkPick(6).Value = 1
            '----
            
            '----
            txtDoor(7).Text = !DSE
            If !KSE <> "0" Then
                chkKey(7).Value = 1
                SetListIndex cboKeys(7), "(" & !KSE & ") " & GetItemName(CInt(!KSE))
            Else
                SetListIndex cboKeys(7), "(0) None"
            End If
            txtBash(7).Text = !BSE
            If !BSE <> "-1" Then chkBack(7).Value = 1
            txtPick(7).Text = !PSE
            If !PSE <> "-1" Then chkPick(7).Value = 1
            '----
             '----
            txtDoor(8).Text = !DU
            If !KU <> "0" Then
                chkKey(8).Value = 1
                SetListIndex cboKeys(8), "(" & !KU & ") " & GetItemName(CInt(!KU))
            Else
                SetListIndex cboKeys(8), "(0) None"
            End If
            txtBash(8).Text = !BU
            If !BU <> "-1" Then chkBack(8).Value = 1
            txtPick(8).Text = !PU
            If !PU <> "-1" Then chkPick(8).Value = 1
            '----
            
             '----
            txtDoor(9).Text = !DD
            If !KD <> "0" Then
                chkKey(9).Value = 1
                SetListIndex cboKeys(9), "(" & !KD & ") " & GetItemName(CInt(!KD))
            Else
                SetListIndex cboKeys(9), "(0) None"
            End If
            txtBash(9).Text = !BD
            If !BD <> "-1" Then chkBack(9).Value = 1
            txtPick(9).Text = !PD
            If !PD <> "-1" Then chkPick(9).Value = 1
            '----
            Select Case !Type
                Case "0":
                    SetListIndex cboType, "0 - Normal"
                Case "1":
                    SetListIndex cboType, "1 - Shop"
                Case "2":
                    SetListIndex cboType, "2 - Trainer"
                Case "3":
                    SetListIndex cboType, "3 - Arena"
                Case "4":
                    SetListIndex cboType, "4 - Boss"
                Case "5":
                    SetListIndex cboType, "5 - Bank"
            End Select
            SetLstSelected lstRooms, "(" & !RoomID & ") " & !RoomTitle
            If txtNorth.Text <> "0" Then
                SetListIndex cboRooms(0), "(" & txtNorth.Text & ") " & GetMapName(CInt(txtNorth.Text))
            Else
                SetListIndex cboRooms(0), "(0) None"
            End If
            If txtSouth.Text <> "0" Then
                SetListIndex cboRooms(1), "(" & txtSouth.Text & ") " & GetMapName(CInt(txtSouth.Text))
            Else
                SetListIndex cboRooms(1), "(0) None"
            End If
            If txtEast.Text <> "0" Then
                SetListIndex cboRooms(2), "(" & txtEast.Text & ") " & GetMapName(CInt(txtEast.Text))
            Else
                SetListIndex cboRooms(2), "(0) None"
            End If
            If txtWest.Text <> "0" Then
                SetListIndex cboRooms(3), "(" & txtWest.Text & ") " & GetMapName(CInt(txtWest.Text))
            Else
                SetListIndex cboRooms(3), "(0) None"
            End If
            If txtNW.Text <> "0" Then
                SetListIndex cboRooms(4), "(" & txtNW.Text & ") " & GetMapName(CInt(txtNW.Text))
            Else
                SetListIndex cboRooms(4), "(0) None"
            End If
            If txtSW.Text <> "0" Then
                SetListIndex cboRooms(5), "(" & txtSW.Text & ") " & GetMapName(CInt(txtSW.Text))
            Else
                SetListIndex cboRooms(5), "(0) None"
            End If
            If txtNE.Text <> "0" Then
                SetListIndex cboRooms(6), "(" & txtNE.Text & ") " & GetMapName(CInt(txtNE.Text))
            Else
                SetListIndex cboRooms(6), "(0) None"
            End If
            If txtSE.Text <> "0" Then
                SetListIndex cboRooms(7), "(" & txtSE.Text & ") " & GetMapName(CInt(txtSE.Text))
            Else
                SetListIndex cboRooms(7), "(0) None"
            End If
            If txtUp.Text <> "0" Then
                SetListIndex cboRooms(8), "(" & txtUp.Text & ") " & GetMapName(CInt(txtUp.Text))
            Else
                SetListIndex cboRooms(8), "(0) None"
            End If
            If txtDown.Text <> "0" Then
                SetListIndex cboRooms(9), "(" & txtDown.Text & ") " & GetMapName(CInt(txtDown.Text))
            Else
                SetListIndex cboRooms(9), "(0) None"
            End If
            If txtGoRoom.Text <> "0" Then
                SetListIndex cboRooms(10), "(" & txtGoRoom.Text & ") " & GetMapName(CInt(txtGoRoom.Text))
            Else
                SetListIndex cboRooms(10), "(0) None"
            End If
            If txtSpecialMonster.Text <> "0" Then
                SetListIndex cboMonsters, "(" & txtSpecialMonster.Text & ") " & GetMonsterName(CInt(txtSpecialMonster.Text))
            Else
                SetListIndex cboMonsters, "(0) None"
            End If
            If txtSpecialItem.Text <> "0" Then
                SetListIndex cboItems, "(" & txtSpecialItem.Text & ") " & GetItemName(CInt(txtSpecialItem.Text))
            Else
                SetListIndex cboItems, "(0) None"
            End If
            If txtShop.Text <> "0" Then
                SetListIndex cboShops, "(" & txtShop.Text & ") " & GetShopName(CInt(txtShop.Text))
            Else
                SetListIndex cboShops, "(0) None"
            End If
            SetListIndex cboRooms(11), "(" & !DeathRoom & ") " & GetMapName(CInt(!DeathRoom))
            bFillMap = False
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
MousePointer = vbDefault
End Sub

Function GetShopName(iShopID As Integer) As String
With RSShops
    .MoveFirst
    Do
        If CInt(!ID) = iShopID Then
            GetShopName = !ShopName
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
End Function

Function GetItemName(iItemID As Integer) As String
With RSItem
    .MoveFirst
    Do
        If CInt(!ID) = iItemID Then
            GetItemName = !ItemName
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
End Function

Function GetMapName(ID As Integer) As String
With RSMap
    .MoveFirst
    Do
        If CInt(!RoomID) = ID Then
            GetMapName = !RoomTitle
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
End Function

Function GetMonsterName(ID As Integer) As String
With RSMonster
    .MoveFirst
    Do
        If CInt(!ID) = ID Then
            GetMonsterName = !MonsterName
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
End Function

Function GetClassName(ID As Integer) As String
With RSClass
    .MoveFirst
    Do
        If CInt(!ID) = ID Then
            GetClassName = !Name
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
End Function

Sub DrawMapOut()
MousePointer = vbHourglass
For i = 0 To shpRoom.UBound
    shpRoom(i).FillColor = &H404040
Next
For i = 0 To linDir.UBound
    linDir(i).BorderColor = &H404040
Next
For i = 0 To lblRoom.UBound
    lblRoom(i).Caption = ""
Next
With RSMap
    .MoveFirst
    Do
        If CInt(!RoomID) = RSmID Then
            shpRoom(0).Visible = True
            shpRoom(0).FillColor = vbBlack
            If CInt(!North) <> 0 Then
                linDir(6).BorderColor = vbBlack
                shpRoom(1).FillColor = vbBlack
                lblRoom(1).Caption = vbCrLf & !North
            End If
            If CInt(!South) <> 0 Then
                linDir(7).BorderColor = vbBlack
                shpRoom(4).FillColor = vbBlack
                lblRoom(6).Caption = vbCrLf & !South
            End If
            If CInt(!East) <> 0 Then
                linDir(4).BorderColor = vbBlack
                shpRoom(2).FillColor = vbBlack
                lblRoom(4).Caption = vbCrLf & !East
            End If
            If CInt(!West) <> 0 Then
                linDir(5).BorderColor = vbBlack
                shpRoom(3).FillColor = vbBlack
                lblRoom(3).Caption = vbCrLf & !West
            End If
            If CInt(!NorthWest) <> 0 Then
                linDir(2).BorderColor = vbBlack
                shpRoom(5).FillColor = vbBlack
                lblRoom(0).Caption = vbCrLf & !NorthWest
            End If
            If CInt(!NorthEast) <> 0 Then
                linDir(1).BorderColor = vbBlack
                shpRoom(6).FillColor = vbBlack
                lblRoom(2).Caption = vbCrLf & !NorthEast
            End If
            If CInt(!SouthEast) <> 0 Then
                linDir(3).BorderColor = vbBlack
                shpRoom(8).FillColor = vbBlack
                lblRoom(7).Caption = vbCrLf & !SouthEast
            End If
            If CInt(!SouthWest) <> 0 Then
                linDir(0).BorderColor = vbBlack
                shpRoom(7).FillColor = vbBlack
                lblRoom(5).Caption = vbCrLf & !SouthWest
            End If
            If CInt(!UP) <> 0 Then
                linDir(10).BorderColor = &HFF0000
                linDir(8).BorderColor = &HFF0000
                shpRoom(9).FillColor = vbBlack
                lblRoom(8).Caption = vbCrLf & !UP
            End If
            If CInt(!Down) <> 0 Then
                linDir(9).BorderColor = &HFF00FF
                linDir(11).BorderColor = &HFF00FF
                shpRoom(10).FillColor = vbBlack
                lblRoom(9).Caption = vbCrLf & !Down
            End If
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
MousePointer = vbDefault
End Sub

Sub GetNewID(WhichDir As String)
Dim BackUp As Integer
BackUp = RSmID
With RSMap
    .MoveFirst
    Do
        If CInt(!RoomID) = RSmID Then
            Select Case WhichDir
                Case "nw"
                    RSmID = !NorthWest
                Case "n"
                    RSmID = !North
                Case "ne"
                    RSmID = !NorthEast
                Case "w"
                    RSmID = !West
                Case "e"
                    RSmID = !East
                Case "sw"
                    RSmID = !SouthWest
                Case "s"
                    RSmID = !South
                Case "se"
                    RSmID = !SouthEast
                Case "u"
                    RSmID = !UP
                Case "d"
                    RSmID = !Down
            End Select
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
If RSmID = 0 Then RSmID = BackUp
End Sub

Sub GetExits()
Dim tVar As String
If txtNorth.Text <> "0" Then tVar = tVar & "north,"
If txtSouth.Text <> "0" Then tVar = tVar & "south,"
If txtEast.Text <> "0" Then tVar = tVar & "east,"
If txtWest.Text <> "0" Then tVar = tVar & "west,"
If txtNW.Text <> "0" Then tVar = tVar & "northwest,"
If txtNE.Text <> "0" Then tVar = tVar & "northeast,"
If txtSW.Text <> "0" Then tVar = tVar & "southwest,"
If txtSE.Text <> "0" Then tVar = tVar & "southeast,"
If txtUp.Text <> "0" Then tVar = tVar & "up,"
If txtDown.Text <> "0" Then tVar = tVar & "down,"
If tVar = "" Then
    tVar = "none"
    txtRoomExits.Text = tVar
    Exit Sub
End If
tVar = Left$(tVar, Len(tVar) - 1)
txtRoomExits.Text = tVar
End Sub

Function IsSaveAble() As Boolean
'If txtRoomExits.Text = "" Then txtRoomExits.Text = "None"
If Not IsNumeric(txtNorth.Text) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtSouth.Text) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtEast.Text) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtWest.Text) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtNW.Text) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtNE.Text) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtSW.Text) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtSE.Text) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtUp.Text) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtDown.Text) Then IsSaveAble = False: Exit Function

If txtRoomTitle.Text = "" Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtMaxRegen.Text) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtMobGroup.Text) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtGold.Text) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtSpecialMonster.Text) Then IsSaveAble = False: Exit Function
IsSaveAble = True
End Function

Private Sub txtNE_Change()
If bFillMap Then Exit Sub
On Error Resume Next
If GetMapName(CInt(txtNE.Text)) = "" And txtNE.Text <> "0" Then
    SetListIndex cboRooms(5), "(???) Undefined"
ElseIf txtNE.Text <> "0" Then
    SetListIndex cboRooms(5), "(" & txtNE.Text & ") " & GetMapName(CInt(txtNE.Text))
ElseIf txtNE.Text <> "" Then
    SetListIndex cboRooms(5), "(0) None"
End If
End Sub

Private Sub txtNorth_Change()
If bFillMap Then Exit Sub
On Error Resume Next
If GetMapName(CInt(txtNorth.Text)) = "" And txtNorth.Text <> "0" Then
    SetListIndex cboRooms(0), "(???) Undefined"
ElseIf txtNorth.Text <> "0" Then
    SetListIndex cboRooms(0), "(" & txtNorth.Text & ") " & GetMapName(CInt(txtNorth.Text))
ElseIf txtNorth.Text <> "" Then
    SetListIndex cboRooms(0), "(0) None"
End If
End Sub

Private Sub txtNW_Change()
If bFillMap Then Exit Sub
On Error Resume Next
If GetMapName(CInt(txtNW.Text)) = "" And txtNW.Text <> "0" Then
    SetListIndex cboRooms(4), "(???) Undefined"
ElseIf txtNW.Text <> "0" Then
    SetListIndex cboRooms(4), "(" & txtNW.Text & ") " & GetMapName(CInt(txtNW.Text))
ElseIf txtNW.Text <> "" Then
    SetListIndex cboRooms(4), "(0) None"
End If
End Sub

Private Sub txtPick_Change(Index As Integer)
If txtPick(Index).Text <> "-1" Then txtDoor(Index).Text = "2"
End Sub

Private Sub txtSE_Change()
If bFillMap Then Exit Sub
On Error Resume Next
If GetMapName(CInt(txtSE.Text)) = "" And txtSE.Text <> "0" Then
    SetListIndex cboRooms(7), "(???) Undefined"
ElseIf txtSE.Text <> "0" Then
    SetListIndex cboRooms(7), "(" & txtSE.Text & ") " & GetMapName(CInt(txtSE.Text))
ElseIf txtSE.Text <> "" Then
    SetListIndex cboRooms(7), "(0) None"
End If
End Sub

Private Sub txtSouth_Change()
If bFillMap Then Exit Sub
On Error Resume Next
If GetMapName(CInt(txtSouth.Text)) = "" And txtSouth.Text <> "0" Then
    SetListIndex cboRooms(1), "(???) Undefined"
ElseIf txtNorth.Text <> "0" Then
    SetListIndex cboRooms(1), "(" & txtSouth.Text & ") " & GetMapName(CInt(txtSouth.Text))
ElseIf txtNorth.Text <> "" Then
    SetListIndex cboRooms(1), "(0) None"
End If
End Sub

Private Sub txtSW_Change()
If bFillMap Then Exit Sub
On Error Resume Next
If GetMapName(CInt(txtSW.Text)) = "" And txtSW.Text <> "0" Then
    SetListIndex cboRooms(6), "(???) Undefined"
ElseIf txtSW.Text <> "0" Then
    SetListIndex cboRooms(6), "(" & txtSW.Text & ") " & GetMapName(CInt(txtSW.Text))
ElseIf txtSW.Text <> "" Then
    SetListIndex cboRooms(6), "(0) None"
End If
End Sub

Private Sub txtTrainClass_Change()
If bFillMap Then Exit Sub
On Error Resume Next
If txtTrainClass.Text = "0" Then
    SetListIndex cboClass, "(0) None"
    SetListIndex cboType, "0 - Normal"
Else
    SetListIndex cboType, "6 - Class Trainer"
    SetListIndex cboClass, "(" & txtTrainClass.Text & ") " & GetMapName(CInt(txtTrainClass.Text))
End If

End Sub

Private Sub txtUp_Change()
If bFillMap Then Exit Sub
On Error Resume Next
If GetMapName(CInt(txtUp.Text)) = "" And txtUp.Text <> "0" Then
    SetListIndex cboRooms(8), "(???) Undefined"
ElseIf txtUp.Text <> "0" Then
    SetListIndex cboRooms(8), "(" & txtUp.Text & ") " & GetMapName(CInt(txtUp.Text))
ElseIf txtUp.Text <> "" Then
    SetListIndex cboRooms(8), "(0) None"
End If
End Sub

Private Sub txtWest_Change()
If bFillMap Then Exit Sub
On Error Resume Next
If GetMapName(CInt(txtWest.Text)) = "" And txtWest.Text <> "0" Then
    SetListIndex cboRooms(3), "(???) Undefined"
ElseIf txtWest.Text <> "0" Then
    SetListIndex cboRooms(3), "(" & txtWest.Text & ") " & GetMapName(CInt(txtWest.Text))
ElseIf txtWest.Text <> "" Then
    SetListIndex cboRooms(3), "(0) None"
End If
End Sub
