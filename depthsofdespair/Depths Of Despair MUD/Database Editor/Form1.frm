VERSION 5.00
Begin VB.Form frmPlayers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Players:"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   12525
   Begin VB.TextBox txtDamage 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   90
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox txtBank 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5520
      MaxLength       =   15
      TabIndex        =   88
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txtDodge 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   86
      Top             =   4200
      Width           =   615
   End
   Begin VB.ComboBox cboArmorType 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form1.frx":0000
      Left            =   7440
      List            =   "Form1.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   85
      Top             =   1560
      Width           =   1575
   End
   Begin VB.ComboBox cboWeaponType 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form1.frx":009A
      Left            =   7440
      List            =   "Form1.frx":00BC
      Style           =   2  'Dropdown List
      TabIndex        =   84
      Top             =   1200
      Width           =   1575
   End
   Begin VB.ComboBox cboFamiliars 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7440
      Style           =   2  'Dropdown List
      TabIndex        =   83
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox txtFamiliar 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7440
      MaxLength       =   4
      TabIndex        =   81
      Top             =   4560
      Width           =   615
   End
   Begin VB.ComboBox cboType 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form1.frx":01C0
      Left            =   7440
      List            =   "Form1.frx":01E5
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1920
      Width           =   1575
   End
   Begin VB.ComboBox cboMagicLevel 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form1.frx":0270
      Left            =   7440
      List            =   "Form1.frx":0286
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "(save)"
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
      Left            =   9120
      TabIndex        =   38
      Top             =   5040
      Width           =   1095
   End
   Begin VB.ListBox lstPlayers 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5100
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Width           =   2415
   End
   Begin VB.CheckBox chkSysop 
      Caption         =   "Check1"
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
      Left            =   3840
      TabIndex        =   11
      Top             =   2280
      Width           =   255
   End
   Begin VB.ComboBox cboSpells 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7440
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   23
      Left            =   11880
      TabIndex        =   26
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   22
      Left            =   10200
      MaxLength       =   6
      TabIndex        =   25
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   3840
      TabIndex        =   72
      Top             =   2280
      Width           =   150
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   19
      Left            =   3840
      TabIndex        =   36
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   18
      Left            =   3840
      TabIndex        =   34
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3840
      MaxLength       =   16
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   3840
      MaxLength       =   16
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   7440
      MaxLength       =   10
      TabIndex        =   15
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   7440
      TabIndex        =   16
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   4800
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   4800
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   4800
      MaxLength       =   4
      TabIndex        =   10
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   5760
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   5760
      MaxLength       =   4
      TabIndex        =   9
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   3840
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   10200
      MaxLength       =   6
      TabIndex        =   23
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   11880
      TabIndex        =   24
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   5520
      TabIndex        =   35
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   7440
      MaxLength       =   4
      TabIndex        =   14
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtPField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   5520
      MaxLength       =   4
      TabIndex        =   37
      Top             =   3840
      Width           =   615
   End
   Begin VB.ComboBox cboWeapon 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   10200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   960
      Width           =   2295
   End
   Begin VB.ComboBox cboHead 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   10200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ComboBox cboBody 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   10200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   1680
      Width           =   2295
   End
   Begin VB.ComboBox cboArms 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   10200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   2040
      Width           =   2295
   End
   Begin VB.ComboBox cboLegs 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   10200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   2760
      Width           =   2295
   End
   Begin VB.ComboBox cboWaist 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   10200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   2400
      Width           =   2295
   End
   Begin VB.ListBox lstItems 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      IntegralHeight  =   0   'False
      Left            =   7440
      TabIndex        =   20
      Top             =   3000
      Width           =   1575
   End
   Begin VB.ComboBox cboR 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   3840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2640
      Width           =   2295
   End
   Begin VB.ComboBox cboC 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   3840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton cmdPNext 
      Caption         =   "Next >"
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
      Left            =   11520
      TabIndex        =   39
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdPPrevious 
      Caption         =   "< Previous"
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
      Left            =   10440
      TabIndex        =   40
      Top             =   5040
      Width           =   975
   End
   Begin VB.ComboBox cboFeet 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   10200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   3120
      Width           =   2295
   End
   Begin VB.ComboBox cboItems 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   7440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   4125
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddInv 
      Caption         =   "Add to INV"
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
      Left            =   6240
      TabIndex        =   22
      Top             =   4125
      Width           =   1095
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Damage:"
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
      Index           =   41
      Left            =   2640
      TabIndex        =   91
      Top             =   4560
      Width           =   645
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Bank:"
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
      Index           =   40
      Left            =   5040
      TabIndex        =   89
      Top             =   4200
      Width           =   420
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Dodge:"
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
      Index           =   39
      Left            =   2640
      TabIndex        =   87
      Top             =   4200
      Width           =   525
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Familiar:"
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
      Index           =   38
      Left            =   6360
      TabIndex        =   82
      Top             =   4560
      Width           =   675
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Weapon Type:"
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
      Index           =   37
      Left            =   6240
      TabIndex        =   80
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Spells:"
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
      Index           =   36
      Left            =   6240
      TabIndex        =   79
      Top             =   2640
      Width           =   510
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Spell Level:"
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
      Index           =   35
      Left            =   6240
      TabIndex        =   78
      Top             =   2280
      Width           =   885
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Spell Type:"
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
      Index           =   34
      Left            =   6240
      TabIndex        =   77
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Max Mana"
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
      Index           =   33
      Left            =   10920
      TabIndex        =   76
      Top             =   480
      Width           =   810
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Mana"
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
      Index           =   32
      Left            =   9360
      TabIndex        =   75
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Armor Type:"
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
      Index           =   31
      Left            =   6240
      TabIndex        =   74
      Top             =   1560
      Width           =   945
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Is Sysop:"
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
      Index           =   30
      Left            =   2640
      TabIndex        =   73
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Crits:"
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
      Index           =   29
      Left            =   2640
      TabIndex        =   71
      Top             =   3840
      Width           =   420
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Accuracy:"
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
      Index           =   28
      Left            =   2640
      TabIndex        =   70
      Top             =   3480
      Width           =   780
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Back Up Loc:"
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
      Index           =   27
      Left            =   2640
      TabIndex        =   69
      Top             =   1920
      Width           =   990
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "ID"
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
      Index           =   0
      Left            =   2640
      TabIndex        =   68
      Top             =   120
      Width           =   210
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
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
      Index           =   1
      Left            =   2640
      TabIndex        =   67
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
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
      Index           =   2
      Left            =   2640
      TabIndex        =   66
      Top             =   840
      Width           =   750
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "EXP:"
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
      Left            =   6240
      TabIndex        =   65
      Top             =   480
      Width           =   360
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "EXP Needed:"
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
      Left            =   6240
      TabIndex        =   64
      Top             =   840
      Width           =   990
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Str:"
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
      Left            =   4320
      TabIndex        =   63
      Top             =   1200
      Width           =   300
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Agil:"
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
      Left            =   4320
      TabIndex        =   62
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Int:"
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
      Left            =   4320
      TabIndex        =   61
      Top             =   1920
      Width           =   285
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Cha:"
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
      Left            =   5280
      TabIndex        =   60
      Top             =   1200
      Width           =   330
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Dex:"
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
      Left            =   5280
      TabIndex        =   59
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Location:"
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
      Left            =   2640
      TabIndex        =   58
      Top             =   1560
      Width           =   705
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Index:"
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
      Left            =   2640
      TabIndex        =   57
      Top             =   1200
      Width           =   510
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "HP:"
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
      Left            =   9360
      TabIndex        =   56
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "MaxHP:"
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
      Index           =   13
      Left            =   10920
      TabIndex        =   55
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "AC:"
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
      Index           =   14
      Left            =   5040
      TabIndex        =   54
      Top             =   3480
      Width           =   270
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Level:"
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
      Index           =   15
      Left            =   6240
      TabIndex        =   53
      Top             =   120
      Width           =   450
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Gold:"
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
      Index           =   16
      Left            =   5040
      TabIndex        =   52
      Top             =   3840
      Width           =   390
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Weapon:"
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
      Index           =   17
      Left            =   9360
      TabIndex        =   51
      Top             =   960
      Width           =   675
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Head:"
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
      Index           =   18
      Left            =   9360
      TabIndex        =   50
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Body:"
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
      Index           =   19
      Left            =   9360
      TabIndex        =   49
      Top             =   1680
      Width           =   420
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Arms:"
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
      Index           =   20
      Left            =   9360
      TabIndex        =   48
      Top             =   2040
      Width           =   435
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Legs:"
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
      Index           =   21
      Left            =   9360
      TabIndex        =   47
      Top             =   2760
      Width           =   390
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Waist:"
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
      Index           =   22
      Left            =   9360
      TabIndex        =   46
      Top             =   2400
      Width           =   510
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Inventory:"
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
      Index           =   23
      Left            =   6240
      TabIndex        =   45
      Top             =   3000
      Width           =   810
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Race:"
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
      Index           =   24
      Left            =   2640
      TabIndex        =   44
      Top             =   2640
      Width           =   420
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Class:"
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
      Index           =   25
      Left            =   2640
      TabIndex        =   43
      Top             =   3000
      Width           =   435
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Feet:"
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
      Index           =   26
      Left            =   9360
      TabIndex        =   42
      Top             =   3120
      Width           =   375
   End
End
Attribute VB_Name = "frmPlayers"
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
Rem***************                frmPlayers                      **********************
Rem***************                ServerEditor                    **********************
Rem***************                Editor.vbp                      **********************
Rem*************************************************************************************
Rem*************************************************************************************

Dim RSpID As Integer
Dim Filling As Boolean

Private Sub cboArms_Click(Index As Integer)
If Filling = False Then
    RemoveHPMANABonus "arms"
    If IsWearable(cboArms(Index)) = False Then ReFillCantUse "arms"
    AddAC
End If
End Sub

Private Sub cboBody_Click(Index As Integer)
If Filling = False Then
    RemoveHPMANABonus "body"
    If IsWearable(cboBody(Index)) = False Then ReFillCantUse "body"
    AddAC
End If
End Sub

Private Sub cboC_Change(Index As Integer)
For i = 0 To cboC(0).ListCount - 1
    If cboC(0).List(cboC(0).ListIndex) = cboC(0).List(i) Then
        MousePointer = vbHourglass
        UpdateClass
        MousePointer = vbNormal
        Exit For
    End If
Next
End Sub

Private Sub cboC_Click(Index As Integer)
cboC_Change Index
End Sub

Private Sub cboFamiliars_Click()
txtFamiliar.Text = Mid$(cboFamiliars.List(cboFamiliars.ListIndex), 2, InStr(1, cboFamiliars.List(cboFamiliars.ListIndex), ")") - 2)
End Sub

Private Sub cboFeet_Click(Index As Integer)
If Filling = False Then
    RemoveHPMANABonus "feet"
    If IsWearable(cboFeet(Index)) = False Then ReFillCantUse "feet"
    AddAC
End If
End Sub

Private Sub cboHead_Click(Index As Integer)
If Filling = False Then
    RemoveHPMANABonus "head"
    If IsWearable(cboHead(Index)) = False Then ReFillCantUse "head"
    AddAC
End If
End Sub

Private Sub cboLegs_Click(Index As Integer)
If Filling = False Then
    RemoveHPMANABonus "legs"
    If IsWearable(cboLegs(Index)) = False Then ReFillCantUse "legs"
    AddAC
End If
End Sub

Private Sub cboWaist_Click(Index As Integer)
If Filling = False Then
    RemoveHPMANABonus "waist"
    If IsWearable(cboWaist(Index)) = False Then ReFillCantUse "waist"
    AddAC
End If
End Sub

Private Sub cboWeapon_Click(Index As Integer)
If Filling = False Then
    RemoveHPMANABonus "weapon"
    If IsWearable(cboWeapon(Index)) = False Then
        ReFillCantUse "weapon"
    End If
    AddAC
End If
End Sub

Private Sub chkSysop_Click()
If txtPField(20).Text <> chkSysop.Value Then
    txtPField(20).Text = chkSysop.Value
End If
End Sub

Private Sub cmdAddInv_Click()
lstItems.AddItem cboItems(0).List(cboItems(0).ListIndex)
SavePlayer
End Sub

Private Sub cmdPNext_Click()
On Error GoTo cmdPNext_Click_Error
SavePlayer
With RSPlayers
    .MoveFirst
    Do
        If CInt(!PlayerId) = RSpID Then
            .MoveNext
            RSpID = !PlayerId
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
FillPlayer
AddAC
On Error GoTo 0
Exit Sub
cmdPNext_Click_Error:
End Sub

Private Sub cmdPPrevious_Click()
On Error GoTo cmdPPrevious_Click_Error
SavePlayer
With RSPlayers
    .MoveFirst
    Do
        If CInt(!PlayerId) = RSpID Then
            .MovePrevious
            RSpID = !PlayerId
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
FillPlayer
AddAC
On Error GoTo 0
Exit Sub
cmdPPrevious_Click_Error:
End Sub

Private Sub cmdSave_Click()
SavePlayer
SetLstSelected lstPlayers, "(" & txtPField(0).Text & ") " & txtPField(1).Text
End Sub

Private Sub Form_Load()
RSpID = 1
FillPlayer
FillLstPlayers
End Sub


Private Sub lstItems_DblClick()
If lstItems.ListIndex > -1 Then lstItems.RemoveItem lstItems.ListIndex
End Sub

Private Sub lstPlayers_Click()
If Filling = False Then
    With RSPlayers
        .MoveFirst
        Do
            If "(" & !PlayerId & ") " & !PlayerName = lstPlayers.Text Then
                RSpID = CInt(!PlayerId)
                FillPlayer
                Exit Do
            ElseIf Not .EOF Then
                .MoveNext
            End If
        Loop Until .EOF
    End With
End If
End Sub

Private Sub txtPField_Change(Index As Integer)
If Index = 20 Then chkSysop.Value = Val(txtPField(Index).Text)
End Sub

Function IsSaveAble() As Boolean
If txtPField(1).Text = "" Then IsSaveAble = False: Exit Function
If txtPField(2).Text = "" Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtPField(10)) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtPField(17)) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtPField(3)) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtPField(5)) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtPField(6)) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtPField(7)) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtPField(8)) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtPField(9)) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtPField(12)) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtPField(22)) Then IsSaveAble = False: Exit Function
If Not IsNumeric(txtPField(16)) Then IsSaveAble = False: Exit Function
IsSaveAble = True
End Function

Function IsWearable(CBO As ComboBox) As Boolean
With RSItem
    .MoveFirst
    Do
        If !ItemName = CBO.List(CBO.ListIndex) Then
            If ClassCanuseMagical(cboC(0).List(cboC(0).ListIndex), !Magical) = False Then
                IsWearable = False
                Exit Do
            End If
            If Val(!Type) <= CInt(Left$(cboWeaponType.List(cboWeaponType.ListIndex), 1)) Then
                If Val(!ArmorType) <= CInt(Left$(cboArmorType.List(cboArmorType.ListIndex), 1)) Then
                    If Val(!Level) <= Val(txtPField(15).Text) Then
                        IsWearable = True
                        Exit Do
                    Else
                        IsWearable = False
                        Exit Do
                    End If
                Else
                    IsWearable = False
                    Exit Do
                End If
            Else
                IsWearable = False
                Exit Do
            End If
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
End Function

Function ClassCanuseMagical(PlayersClass As String, Magical As Integer) As Boolean
With RSClass
    .MoveFirst
    Do
        If LCase(PlayersClass) = LCase(!Name) Then
            If Magical <= CInt(!UseMagical) Then
                ClassCanuseMagical = True
            Else
                ClassCanuseMagical = False
            End If
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
End Function

Sub AddAC()
Dim tVar As Integer, tCrits As Integer, tAcc As Integer, tDodge As Integer
Dim tDam As Integer, tHP As Integer, tMana As Integer
MousePointer = vbHourglass
With RSItem
    .MoveFirst
    Do
        If !ItemName = cboWeapon(0).List(cboWeapon(0).ListIndex) Then
            tVar = tVar + Val(!AC)
            tCrits = tCrits + Val(!Crits)
            tAcc = tAcc + Val(!Acc)
            tDodge = tDodge + Val(!DodgeBonus)
            tDam = tDam + Val(!DamageBonus)
            tHP = tHP + Val(!HPBonus)
            tMana = tMana + Val(!ManaBonus)
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
    .MoveFirst
    Do
        If !ItemName = cboHead(0).List(cboHead(0).ListIndex) Then
            tVar = tVar + Val(!AC)
            tCrits = tCrits + Val(!Crits)
            tAcc = tAcc + Val(!Acc)
            tDodge = tDodge + Val(!DodgeBonus)
            tDam = tDam + Val(!DamageBonus)
            tHP = tHP + Val(!HPBonus)
            tMana = tMana + Val(!ManaBonus)
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
    .MoveFirst
    Do
        If !ItemName = cboBody(0).List(cboBody(0).ListIndex) Then
            tVar = tVar + Val(!AC)
            tCrits = tCrits + Val(!Crits)
            tAcc = tAcc + Val(!Acc)
            tDodge = tDodge + Val(!DodgeBonus)
            tDam = tDam + Val(!DamageBonus)
            tHP = tHP + Val(!HPBonus)
            tMana = tMana + Val(!ManaBonus)
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
    .MoveFirst
    Do
        If !ItemName = cboArms(0).List(cboArms(0).ListIndex) Then
            tVar = tVar + Val(!AC)
            tCrits = tCrits + Val(!Crits)
            tAcc = tAcc + Val(!Acc)
            tDodge = tDodge + Val(!DodgeBonus)
            tDam = tDam + Val(!DamageBonus)
            tHP = tHP + Val(!HPBonus)
            tMana = tMana + Val(!ManaBonus)
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
    .MoveFirst
    Do
        If !ItemName = cboLegs(0).List(cboLegs(0).ListIndex) Then
            tVar = tVar + Val(!AC)
            tCrits = tCrits + Val(!Crits)
            tAcc = tAcc + Val(!Acc)
            tDodge = tDodge + Val(!DodgeBonus)
            tDam = tDam + Val(!DamageBonus)
            tHP = tHP + Val(!HPBonus)
            tMana = tMana + Val(!ManaBonus)
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
    .MoveFirst
    Do
        If !ItemName = cboWaist(0).List(cboWaist(0).ListIndex) Then
            tVar = tVar + Val(!AC)
            tCrits = tCrits + Val(!Crits)
            tAcc = tAcc + Val(!Acc)
            tDodge = tDodge + Val(!DodgeBonus)
            tDam = tDam + Val(!DamageBonus)
            tHP = tHP + Val(!HPBonus)
            tMana = tMana + Val(!ManaBonus)
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
    .MoveFirst
    Do
        If !ItemName = cboFeet(0).List(cboFeet(0).ListIndex) Then
            tVar = tVar + Val(!AC)
            tCrits = tCrits + Val(!Crits)
            tAcc = tAcc + Val(!Acc)
            tDodge = tDodge + Val(!DodgeBonus)
            tDam = tDam + Val(!DamageBonus)
            tHP = tHP + Val(!HPBonus)
            tMana = tMana + Val(!ManaBonus)
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
With RSClass
    .MoveFirst
    Do
        If !Name = cboC(0).List(cboC(0).ListIndex) Then
            tAcc = tAcc + Val(!Acc)
            tCrits = tCrits + Val(!Crits)
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
Dim FamID As Integer
Dim aStats(2) As String
Dim SkipFam As Boolean
SkipFam = False
FamID = txtFamiliar.Text
If FamID = 0 Then SkipFam = True
If SkipFam = False Then
    With RSFamiliars
        .MoveFirst
        Do
            If CInt(!ID) = FamID Then
                aStats(0) = !FamAbility1
                aStats(1) = !FamAbility2
                aStats(2) = !FamAbility3
                Exit Do
            ElseIf Not .EOF Then
                .MoveNext
            End If
            DoEvents
        Loop Until .EOF
    End With
    For i = 0 To 2
        Select Case Left$(aStats(i), 3)
            Case "pac"
                tVar = tVar + CLng(Replace(aStats(i), "pac:", ""))
            Case "acc"
                tAcc = tAcc + CLng(Replace(aStats(i), "acc:", ""))
            Case "cri"
                tCrits = tCrits + CLng(Replace(aStats(i), "cri:", ""))
            Case "dam"
                tDam = tDam + CLng(Replace(aStats(i), "dam:", ""))
            Case "dod"
                tDodge = tDodge + CLng(Replace(aStats(i), "dod:", ""))
            Case "mhp"
                tHP = tHP + CLng(Replace(aStats(i), "mhp:", ""))
            Case "mma"
                tMana = tMana + CLng(Replace(aStats(i), "mma:", ""))
        End Select
    Next
End If
txtPField(14).Text = tVar
txtPField(18).Text = tAcc
txtPField(19).Text = tCrits
txtPField(13).Text = CInt(txtPField(13).Text) + tHP
txtPField(23).Text = CInt(txtPField(23).Text) + tMana
txtDodge.Text = tDodge
txtDamage.Text = tDam
MousePointer = vbDefault
End Sub

Sub ClearCBOS()
cboSpells.Clear
cboFamiliars.Clear
For i = 0 To cboC.UBound
    cboC(i).Clear
Next
For i = 0 To cboR.UBound
    cboR(i).Clear
Next
For i = 0 To cboWeapon.UBound
    cboWeapon(i).Clear
Next
For i = 0 To cboBody.UBound
    cboBody(i).Clear
Next
For i = 0 To cboArms.UBound
    cboArms(i).Clear
Next
For i = 0 To cboLegs.UBound
    cboLegs(i).Clear
Next
For i = 0 To cboWaist.UBound
    cboWaist(i).Clear
Next
For i = 0 To cboFeet.UBound
    cboFeet(i).Clear
Next
For i = 0 To cboHead.UBound
    cboHead(i).Clear
Next
For i = 0 To cboItems.UBound
    cboItems(i).Clear
Next
End Sub

Sub FillPlayer()
Dim Spells$
Filling = True
MousePointer = vbHourglass
ClearCBOS
FillCBOS
With RSPlayers
    .MoveFirst
    Do
        If CInt(!PlayerId) = RSpID Then
            txtPField(23).Text = !MaxMana
            txtPField(13).Text = !MaxHP
            txtPField(0).Text = !PlayerId
            txtPField(1).Text = !PlayerName
            txtPField(2).Text = !PlayerPW
            txtPField(3).Text = !Exp
            txtPField(4).Text = !EXPNeeded
            txtPField(5).Text = !Str
            txtPField(6).Text = !AGIL
            txtPField(7).Text = !Int
            txtPField(8).Text = !CHA
            txtPField(9).Text = !DEX
            txtBank.Text = !Bank
            txtDodge.Text = !Dodge
            txtFamiliar.Text = !FamID
            txtPField(10).Text = ""
            RSMap.MoveFirst
            Do
                If CInt(RSMap!RoomID) = CInt(!Location) Then
                    txtPField(10).Text = !Location
                    Exit Do
                ElseIf Not RSMap.EOF Then
                    RSMap.MoveNext
                End If
            Loop Until RSMap.EOF
            If txtPField(10) = "" Then txtPField(10).Text = "1"
            .Edit
            !Location = "1"
            .Update
            txtPField(11).Text = !Index
            txtPField(12).Text = !HP
            txtPField(14).Text = !AC
            txtPField(15).Text = !Level
            txtPField(16).Text = !Gold
            txtPField(17).Text = ""
            RSMap.MoveFirst
            Do
                If CInt(RSMap!RoomID) = CInt(!BackUpLoc) Then
                    txtPField(17).Text = !BackUpLoc
                    Exit Do
                ElseIf Not RSMap.EOF Then
                    RSMap.MoveNext
                End If
            Loop Until RSMap.EOF
            If txtPField(17) = "" Then txtPField(17).Text = "1"
            .Edit
            !BackUpLoc = "1"
            .Update
            txtPField(18).Text = !Acc
            txtPField(19).Text = !Crits
            txtPField(20).Text = !IsSysop
            Select Case !ArmorType
                Case "0"
                    SetListIndex cboArmorType, "0 - None"
                Case "1"
                    SetListIndex cboArmorType, "1 - Basic"
                Case "2"
                    SetListIndex cboArmorType, "2 - Leather"
                Case "3"
                    SetListIndex cboArmorType, "3 - Metals"
            End Select
            txtPField(22).Text = !Mana
            Select Case !SpellType
                Case "0"
                    SetListIndex cboType, "0 - None"
                Case "1"
                    SetListIndex cboType, "1 - Magery"
                Case "2"
                    SetListIndex cboType, "2 - Druish"
                Case "3"
                    SetListIndex cboType, "3 - Priestly"
                Case "4"
                    SetListIndex cboType, "4 - Kai"
                Case "5"
                    SetListIndex cboType, "5 - General"
                Case "6"
                    SetListIndex cboType, "6 - Unholy"
                Case "7"
                    SetListIndex cboType, "7 - Psychic"
                Case "8"
                    SetListIndex cboType, "8 - Bardic"
                Case "9"
                    SetListIndex cboType, "9 - Witch"
                Case "10"
                    SetListIndex cboType, "10 - Teleporter"
            End Select
            Select Case !SpellLevel
                Case "0"
                    SetListIndex cboMagicLevel, "0 - None"
                Case "1"
                    SetListIndex cboMagicLevel, "1 - Basic"
                Case "2"
                    SetListIndex cboMagicLevel, "2 - Intermediate"
                Case "3"
                    SetListIndex cboMagicLevel, "3 - Advanced"
                Case "4"
                    SetListIndex cboMagicLevel, "4 - Expert"
                Case "5"
                    SetListIndex cboMagicLevel, "5 - Master"
            End Select
            Select Case !Weapon
                Case "0"
                    SetListIndex cboWeaponType, "0 - No Weapons"
                Case "1"
                    SetListIndex cboWeaponType, "1 - 1H - Staves or less"
                Case "2"
                    SetListIndex cboWeaponType, "2 - 1H - Short blades or less"
                Case "3"
                    SetListIndex cboWeaponType, "3 - 1H - Small Blunt or less"
                Case "4"
                    SetListIndex cboWeaponType, "4 - 1H - Long blades or less"
                Case "5"
                    SetListIndex cboWeaponType, "5 - 1H - Large Blunt or less"
                Case "6"
                    SetListIndex cboWeaponType, "6 - 2H - Staves or less"
                Case "7"
                    SetListIndex cboWeaponType, "7 - 2H - Long Blades or less"
                Case "8"
                    SetListIndex cboWeaponType, "8 - 2H - Large Blunt or less"
                Case "9"
                    SetListIndex cboWeaponType, "9 - Any"
            End Select
            SetListIndex cboWeapon(0), !Weapon
            SetListIndex cboHead(0), !Head
            SetListIndex cboBody(0), !Body
            SetListIndex cboArms(0), !Arms
            SetListIndex cboLegs(0), !Legs
            SetListIndex cboWaist(0), !Waist
            SetListIndex cboFeet(0), !Feet
            If txtFamiliar.Text <> "0" Then
                SetListIndex cboFamiliars, "(" & txtFamiliar.Text & ") " & !FamName
            Else
                SetListIndex cboFamiliars, "(0) None"
            End If
            If !Spells <> "0" Then
                Spells$ = Replace(!Spells, ":", "")
            End If
            Dim tArr() As String
            Dim tVar As String
            tVar = !Inv
            lstItems.Clear
            tVar = Left$(tVar, Len(tVar) - 1)
            If tVar = "Nothing" Then GoTo SkipItems
            tArr = Split(tVar, ",")
            For i = 0 To UBound(tArr)
                lstItems.AddItem tArr(i)
            Next
SkipItems:
            SetListIndex cboR(0), !Race
            SetListIndex cboC(0), !Class
            If Spells$ <> "" Then
                Erase tArr
                tArr() = Split(Left$(Spells$, Len(Spells$) - 1), ";")
            End If
            SetLstSelected lstPlayers, "(" & !PlayerId & ") " & !PlayerName
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
If Spells$ <> "" Then
    With RSSpells
        For i = 0 To UBound(tArr)
            .MoveFirst
            Do
                If CInt(!ID) = CInt(tArr(i)) Then
                    cboSpells.AddItem !SpellName
                    Exit Do
                ElseIf Not .EOF Then
                    .MoveNext
                End If
            Loop Until .EOF
        Next
    End With
    SetListIndex cboSpells, cboSpells.List(0)
Else
    cboSpells.AddItem "None"
    SetListIndex cboSpells, cboSpells.List(0)
End If
Filling = False
MousePointer = vbDefault
End Sub

Sub FillCBOS()
With RSItem
    .MoveFirst
    Do
        cboItems(0).AddItem !ItemName
        .MoveNext
    Loop Until .EOF
    cboItems(0).ListIndex = 0
    .MoveFirst
    Do
        If !Worn = "weapon" Then
            cboWeapon(0).AddItem !ItemName
            .MoveNext
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
    .MoveFirst
    cboHead(0).AddItem "Nothing"
    Do
        If !Worn = "head" Then
            cboHead(0).AddItem !ItemName
            .MoveNext
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
    .MoveFirst
    cboBody(0).AddItem "Nothing"
    Do
        If !Worn = "body" Then
            cboBody(0).AddItem !ItemName
            .MoveNext
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
    .MoveFirst
    cboArms(0).AddItem "Nothing"
    Do
        If !Worn = "arms" Then
            cboArms(0).AddItem !ItemName
            .MoveNext
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
    .MoveFirst
    cboLegs(0).AddItem "Nothing"
    Do
        If !Worn = "legs" Then
            cboLegs(0).AddItem !ItemName
            .MoveNext
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
    .MoveFirst
    cboWaist(0).AddItem "Nothing"
    Do
        If !Worn = "waist" Then
            cboWaist(0).AddItem !ItemName
            .MoveNext
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
    .MoveFirst
    cboFeet(0).AddItem "Nothing"
    Do
        If !Worn = "feet" Then
            cboFeet(0).AddItem !ItemName
            .MoveNext
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
With RSClass
    .MoveFirst
    Do
        cboC(0).AddItem !Name
        .MoveNext
    Loop Until .EOF
End With
With RSRace
    .MoveFirst
    Do
        cboR(0).AddItem !Name
        .MoveNext
    Loop Until .EOF
End With
cboFamiliars.AddItem "(0) None"
With RSFamiliars
    .MoveFirst
    Do
        cboFamiliars.AddItem "(" & !ID & ") " & !FamName
        .MoveNext
    Loop Until .EOF
End With
End Sub

Sub SavePlayer()
MousePointer = vbHourglass
If IsSaveAble = False Then
    MsgBox "Not everything is filled out correctly, or not everything is filled in." & vbCrLf & "Please check to make sure everything is correct." & vbCrLf & "Save will not continue.", vbCritical, "Error"
    Exit Sub
End If
With RSPlayers
    .MoveFirst
    Do
        If CInt(!PlayerId) = RSpID Then
            .Edit
            !PlayerName = txtPField(1).Text
            !PlayerPW = txtPField(2).Text
            !Exp = txtPField(3).Text
            !Str = txtPField(5).Text
            !Dodge = txtDodge.Text
            !Bank = txtBank.Text
            !AGIL = txtPField(6).Text
            !Int = txtPField(7).Text
            !CHA = txtPField(8).Text
            !FamID = txtFamiliar.Text
            !FamName = Mid$(cboFamiliars.List(cboFamiliars.ListIndex), InStr(1, cboFamiliars.List(cboFamiliars.ListIndex), ")") + 2, Len(cboFamiliars.List(cboFamiliars.ListIndex)) - InStr(1, cboFamiliars.List(cboFamiliars.ListIndex), ")") + 2)
            !DEX = txtPField(9).Text
            !Location = txtPField(10).Text
            !HP = txtPField(12).Text
            !MaxHP = txtPField(13).Text
            !AC = txtPField(14).Text
            !Gold = txtPField(16).Text
            !BackUpLoc = txtPField(17).Text
            !Acc = txtPField(18).Text
            !Crits = txtPField(19).Text
            !IsSysop = txtPField(20).Text
            !ArmorType = Left$(cboArmorType.List(cboArmorType.ListIndex), 1)
            !Mana = txtPField(22).Text
            !MaxMana = txtPField(23).Text
            !SpellType = Left$(cboType.List(cboType.ListIndex), 1)
            !SpellLevel = Left$(cboMagicLevel.List(cboMagicLevel.ListIndex), 1)
            !Weapons = Left$(cboWeaponType.List(cboWeaponType.ListIndex), 1)
            !Weapon = cboWeapon(0).List(cboWeapon(0).ListIndex)
            !Head = cboHead(0).List(cboHead(0).ListIndex)
            !Body = cboBody(0).List(cboBody(0).ListIndex)
            !Arms = cboArms(0).List(cboArms(0).ListIndex)
            !Legs = cboLegs(0).List(cboLegs(0).ListIndex)
            !Waist = cboWaist(0).List(cboWaist(0).ListIndex)
            !Feet = cboFeet(0).List(cboFeet(0).ListIndex)
            Dim tVar As String
            For i = 0 To lstItems.ListCount - 1
                tVar = tVar & lstItems.List(i) & ","
            Next
            If tVar = "" Or tVar = "," Then tVar = "Nothing,"
            !Inv = tVar
            For i = 0 To cboR(0).ListCount - 1
                If cboR(0).Text = cboR(0).List(i) Then
                    !Race = cboR(0).Text
                    Exit For
                End If
            Next
            !Class = cboC(0).List(cboC(0).ListIndex)
            Dim IDs$, Shorts$
            If cboSpells.ListCount > 0 And cboSpells.List(0) <> "None" Then
                With RSSpells
                    For i = 0 To cboSpells.ListCount - 1
                        .MoveFirst
                        Do
                            If !SpellName = cboSpells.List(i) Then
                                IDs$ = IDs$ & ":" & !ID & ";"
                                Shorts$ = Shorts$ & !Short & ";"
                                Exit Do
                            ElseIf Not .EOF Then
                                .MoveNext
                            End If
                        Loop Until .EOF
                    Next
                End With
            End If
            .Update
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
FillLstPlayers
MousePointer = vbDefault
End Sub

Sub UpdateClass()
With RSClass
    .MoveFirst
    Do
        If !Name = cboC(0).Text Then
            Select Case !Weapon
                Case "0"
                    SetListIndex cboWeaponType, "0 - No Weapons"
                Case "1"
                    SetListIndex cboWeaponType, "1 - 1H - Staves or less"
                Case "2"
                    SetListIndex cboWeaponType, "2 - 1H - Short blades or less"
                Case "3"
                    SetListIndex cboWeaponType, "3 - 1H - Small Blunt or less"
                Case "4"
                    SetListIndex cboWeaponType, "4 - 1H - Long blades or less"
                Case "5"
                    SetListIndex cboWeaponType, "5 - 1H - Large Blunt or less"
                Case "6"
                    SetListIndex cboWeaponType, "6 - 2H - Staves or less"
                Case "7"
                    SetListIndex cboWeaponType, "7 - 2H - Long Blades or less"
                Case "8"
                    SetListIndex cboWeaponType, "8 - 2H - Large Blunt or less"
                Case "9"
                    SetListIndex cboWeaponType, "9 - Any"
            End Select
            Select Case !ArmorType
                Case "0"
                    SetListIndex cboArmorType, "0 - No Armor"
                Case "1"
                    SetListIndex cboArmorType, "1 - Silk"
                Case "2"
                    SetListIndex cboArmorType, "2 - Robes"
                Case "3"
                    SetListIndex cboArmorType, "3 - Soft Leather"
                Case "4"
                    SetListIndex cboArmorType, "4 - Hard Leather"
                Case "4"
                    SetListIndex cboArmorType, "5 - Chainmail"
                Case "6"
                    SetListIndex cboArmorType, "6 - Light Platemail"
                Case "7"
                    SetListIndex cboArmorType, "7 - Platemail"
            End Select
            Select Case !SpellType
                Case "1"
                    SetListIndex cboType, "1 - Magery"
                Case "2"
                    SetListIndex cboType, "2 - Druish"
                Case "3"
                    SetListIndex cboType, "3 - Priestly"
                Case "4"
                    SetListIndex cboType, "4 - Kai"
                Case "5"
                    SetListIndex cboType, "5 - General"
                Case "6"
                    SetListIndex cboType, "6 - Unholy"
                Case "7"
                    SetListIndex cboType, "7 - Psychic"
                Case "8"
                    SetListIndex cboType, "8 - Bardic"
                Case "9"
                    SetListIndex cboType, "9 - Witch"
                Case "10"
                    SetListIndex cboType, "10 - Teleporter"
            End Select
            Select Case !SpellLevel
                Case "1"
                    SetListIndex cboMagicLevel, "1 - Basic"
                Case "2"
                    SetListIndex cboMagicLevel, "2 - Intermediate"
                Case "3"
                    SetListIndex cboMagicLevel, "3 - Advanced"
                Case "4"
                    SetListIndex cboMagicLevel, "4 - Expert"
                Case "5"
                    SetListIndex cboMagicLevel, "5 - Master"
            End Select
            AddAC
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
End Sub

Sub ReFillCantUse(WhichCBO As String)
Filling = True
With RSPlayers
    .MoveFirst
    Select Case WhichCBO
        Case "weapon":
            Do
                If CInt(!PlayerId) = RSpID Then
                    SetListIndex cboWeapon(0), !Weapon
                    Exit Do
                ElseIf Not .EOF Then
                    .MoveNext
                End If
            Loop Until .EOF
        Case "head":
            Do
                If CInt(!PlayerId) = RSpID Then
                    SetListIndex cboHead(0), !Head
                    Exit Do
                ElseIf Not .EOF Then
                    .MoveNext
                End If
            Loop Until .EOF
        Case "body":
            Do
                If CInt(!PlayerId) = RSpID Then
                    SetListIndex cboBody(0), !Body
                    Exit Do
                ElseIf Not .EOF Then
                    .MoveNext
                End If
            Loop Until .EOF
        Case "arms":
            Do
                If CInt(!PlayerId) = RSpID Then
                    SetListIndex cboArms(0), !Arms
                    Exit Do
                ElseIf Not .EOF Then
                    .MoveNext
                End If
            Loop Until .EOF
        Case "waist":
            Do
                If CInt(!PlayerId) = RSpID Then
                    SetListIndex cboWaist(0), !Waist
                    Exit Do
                ElseIf Not .EOF Then
                    .MoveNext
                End If
            Loop Until .EOF
        Case "legs":
            Do
                If CInt(!PlayerId) = RSpID Then
                    SetListIndex cboLegs(0), !Legs
                    Exit Do
                ElseIf Not .EOF Then
                    .MoveNext
                End If
            Loop Until .EOF
        Case "feet":
            Do
                If CInt(!PlayerId) = RSpID Then
                    SetListIndex cboFeet(0), !Feet
                    Exit Do
                ElseIf Not .EOF Then
                    .MoveNext
                End If
            Loop Until .EOF
    End Select
End With
Filling = False
End Sub

Sub FillLstPlayers()
lstPlayers.Clear
With RSPlayers
    .MoveFirst
    Do
        lstPlayers.AddItem "(" & !PlayerId & ") " & !PlayerName
        .MoveNext
    Loop Until .EOF
End With
End Sub

Sub RemoveHPMANABonus(Slot As String)
Dim sItem As String
With RSPlayers
    .MoveFirst
    Do
        If CInt(!PlayerId) = CInt(txtPField(0).Text) Then
            Select Case Slot
                Case "weapon"
                    sItem = !Weapon
                Case "head"
                    sItem = !Head
                Case "body"
                    sItem = !Body
                Case "arms"
                    sItem = !Arms
                Case "waist"
                    sItem = !Waist
                Case "legs"
                    sItem = !Legs
                Case "feet"
                    sItem = !Feet
            End Select
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
With RSItem
    .MoveFirst
    Do
        If LCase(sItem) = LCase(!ItemName) Then
            txtPField(13).Text = Val(txtPField(13).Text) - Val(!HPBonus)
            txtPField(23).Text = Val(txtPField(23).Text) - Val(!ManaBonus)
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
End Sub
