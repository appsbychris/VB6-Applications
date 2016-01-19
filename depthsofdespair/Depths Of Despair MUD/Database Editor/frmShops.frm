VERSION 5.00
Begin VB.Form frmShops 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shops"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   12855
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   240
      TabIndex        =   83
      Top             =   240
      Width           =   2895
   End
   Begin ServerEditor.UltraBox lstShops 
      Height          =   7215
      Left            =   240
      TabIndex        =   82
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   12726
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
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   3600
      ScaleHeight     =   1215
      ScaleWidth      =   5295
      TabIndex        =   69
      Top             =   360
      Width           =   5295
      Begin ServerEditor.NumOnlyText txtMarkUp 
         Height          =   375
         Left            =   960
         TabIndex        =   81
         Top             =   840
         Width           =   4095
         _ExtentX        =   7223
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
         Text            =   "0"
         AllowNeg        =   -1  'True
         Align           =   0
         MaxLength       =   5
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin VB.TextBox txtShopName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   960
         TabIndex        =   71
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox txtID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   70
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Shop Name:"
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
         Left            =   0
         TabIndex        =   74
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mark up:"
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
         Left            =   165
         TabIndex        =   73
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
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
         Left            =   600
         TabIndex        =   72
         Top             =   0
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   3600
      ScaleHeight     =   5775
      ScaleWidth      =   8895
      TabIndex        =   4
      Top             =   1920
      Width           =   8895
      Begin VB.ComboBox cboItems 
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
         Index           =   0
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtQ 
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
         Left            =   615
         TabIndex        =   34
         Text            =   "0"
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cboItems 
         BackColor       =   &H00E0E0E0&
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
         Index           =   1
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtQ 
         BackColor       =   &H00E0E0E0&
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
         Left            =   615
         TabIndex        =   32
         Text            =   "0"
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cboItems 
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
         Index           =   2
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtQ 
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
         Left            =   615
         TabIndex        =   30
         Text            =   "0"
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox cboItems 
         BackColor       =   &H00E0E0E0&
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
         Index           =   3
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtQ 
         BackColor       =   &H00E0E0E0&
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
         Left            =   615
         TabIndex        =   28
         Text            =   "0"
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox cboItems 
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
         Index           =   4
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtQ 
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
         Left            =   615
         TabIndex        =   26
         Text            =   "0"
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox cboItems 
         BackColor       =   &H00E0E0E0&
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
         Index           =   5
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox txtQ 
         BackColor       =   &H00E0E0E0&
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
         Left            =   615
         TabIndex        =   24
         Text            =   "0"
         Top             =   2040
         Width           =   975
      End
      Begin VB.ComboBox cboItems 
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
         Index           =   6
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox txtQ 
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
         Index           =   6
         Left            =   615
         TabIndex        =   22
         Text            =   "0"
         Top             =   2400
         Width           =   975
      End
      Begin VB.ComboBox cboItems 
         BackColor       =   &H00E0E0E0&
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
         Index           =   7
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox txtQ 
         BackColor       =   &H00E0E0E0&
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
         Index           =   7
         Left            =   615
         TabIndex        =   20
         Text            =   "0"
         Top             =   2760
         Width           =   975
      End
      Begin VB.ComboBox cboItems 
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
         Index           =   8
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox txtQ 
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
         Index           =   8
         Left            =   615
         TabIndex        =   18
         Text            =   "0"
         Top             =   3120
         Width           =   975
      End
      Begin VB.ComboBox cboItems 
         BackColor       =   &H00E0E0E0&
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
         Index           =   9
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3480
         Width           =   2415
      End
      Begin VB.TextBox txtQ 
         BackColor       =   &H00E0E0E0&
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
         Index           =   9
         Left            =   615
         TabIndex        =   16
         Text            =   "0"
         Top             =   3480
         Width           =   975
      End
      Begin VB.ComboBox cboItems 
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
         Index           =   10
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3840
         Width           =   2415
      End
      Begin VB.TextBox txtQ 
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
         Index           =   10
         Left            =   615
         TabIndex        =   14
         Text            =   "0"
         Top             =   3840
         Width           =   975
      End
      Begin VB.ComboBox cboItems 
         BackColor       =   &H00E0E0E0&
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
         Index           =   11
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   4200
         Width           =   2415
      End
      Begin VB.TextBox txtQ 
         BackColor       =   &H00E0E0E0&
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
         Index           =   11
         Left            =   615
         TabIndex        =   12
         Text            =   "0"
         Top             =   4200
         Width           =   975
      End
      Begin VB.ComboBox cboItems 
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
         Index           =   12
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   4560
         Width           =   2415
      End
      Begin VB.TextBox txtQ 
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
         Index           =   12
         Left            =   615
         TabIndex        =   10
         Text            =   "0"
         Top             =   4560
         Width           =   975
      End
      Begin VB.ComboBox cboItems 
         BackColor       =   &H00E0E0E0&
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
         Index           =   13
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   4920
         Width           =   2415
      End
      Begin VB.TextBox txtQ 
         BackColor       =   &H00E0E0E0&
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
         Index           =   13
         Left            =   615
         TabIndex        =   8
         Text            =   "0"
         Top             =   4920
         Width           =   975
      End
      Begin VB.ComboBox cboItems 
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
         Index           =   14
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   5280
         Width           =   2415
      End
      Begin VB.TextBox txtQ 
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
         Index           =   14
         Left            =   615
         TabIndex        =   6
         Text            =   "0"
         Top             =   5280
         Width           =   975
      End
      Begin ServerEditor.NumOnlyText txtItems 
         Height          =   350
         Index           =   0
         Left            =   4695
         TabIndex        =   5
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MaxLength       =   4
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin ServerEditor.NumOnlyText txtItems 
         Height          =   350
         Index           =   1
         Left            =   4695
         TabIndex        =   36
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MaxLength       =   4
         Enabled         =   -1  'True
         Backcolor       =   14737632
      End
      Begin ServerEditor.NumOnlyText txtItems 
         Height          =   350
         Index           =   2
         Left            =   4695
         TabIndex        =   37
         Top             =   960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MaxLength       =   4
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin ServerEditor.NumOnlyText txtItems 
         Height          =   350
         Index           =   3
         Left            =   4695
         TabIndex        =   38
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MaxLength       =   4
         Enabled         =   -1  'True
         Backcolor       =   14737632
      End
      Begin ServerEditor.NumOnlyText txtItems 
         Height          =   350
         Index           =   4
         Left            =   4695
         TabIndex        =   39
         Top             =   1680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MaxLength       =   4
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin ServerEditor.NumOnlyText txtItems 
         Height          =   350
         Index           =   5
         Left            =   4695
         TabIndex        =   40
         Top             =   2040
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MaxLength       =   4
         Enabled         =   -1  'True
         Backcolor       =   14737632
      End
      Begin ServerEditor.NumOnlyText txtItems 
         Height          =   350
         Index           =   6
         Left            =   4695
         TabIndex        =   41
         Top             =   2400
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MaxLength       =   4
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin ServerEditor.NumOnlyText txtItems 
         Height          =   350
         Index           =   7
         Left            =   4695
         TabIndex        =   42
         Top             =   2760
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MaxLength       =   4
         Enabled         =   -1  'True
         Backcolor       =   14737632
      End
      Begin ServerEditor.NumOnlyText txtItems 
         Height          =   350
         Index           =   8
         Left            =   4695
         TabIndex        =   43
         Top             =   3120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MaxLength       =   4
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin ServerEditor.NumOnlyText txtItems 
         Height          =   350
         Index           =   9
         Left            =   4695
         TabIndex        =   44
         Top             =   3480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MaxLength       =   4
         Enabled         =   -1  'True
         Backcolor       =   14737632
      End
      Begin ServerEditor.NumOnlyText txtItems 
         Height          =   350
         Index           =   10
         Left            =   4695
         TabIndex        =   45
         Top             =   3840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MaxLength       =   4
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin ServerEditor.NumOnlyText txtItems 
         Height          =   350
         Index           =   11
         Left            =   4695
         TabIndex        =   46
         Top             =   4200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MaxLength       =   4
         Enabled         =   -1  'True
         Backcolor       =   14737632
      End
      Begin ServerEditor.NumOnlyText txtItems 
         Height          =   350
         Index           =   12
         Left            =   4695
         TabIndex        =   47
         Top             =   4560
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MaxLength       =   4
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin ServerEditor.NumOnlyText txtItems 
         Height          =   350
         Index           =   13
         Left            =   4695
         TabIndex        =   48
         Top             =   4920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MaxLength       =   4
         Enabled         =   -1  'True
         Backcolor       =   14737632
      End
      Begin ServerEditor.NumOnlyText txtItems 
         Height          =   350
         Index           =   14
         Left            =   4695
         TabIndex        =   49
         Top             =   5280
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MaxLength       =   4
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Items:"
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
         Index           =   2
         Left            =   0
         TabIndex        =   68
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Quantity In Stock"
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
         Left            =   495
         TabIndex        =   67
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item Name                                     /   ID"
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
         Left            =   2175
         TabIndex        =   66
         Top             =   0
         Width           =   2805
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   5415
         TabIndex        =   65
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "End Price (With no Charm Deductions/Increases)"
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
         Left            =   5295
         TabIndex        =   64
         Top             =   0
         Width           =   3495
      End
      Begin VB.Line Line1 
         X1              =   1935
         X2              =   1935
         Y1              =   0
         Y2              =   5520
      End
      Begin VB.Line Line2 
         X1              =   5175
         X2              =   5175
         Y1              =   0
         Y2              =   5520
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Index           =   1
         Left            =   5415
         TabIndex        =   63
         Top             =   600
         Width           =   90
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   2
         Left            =   5415
         TabIndex        =   62
         Top             =   960
         Width           =   90
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Index           =   3
         Left            =   5415
         TabIndex        =   61
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   4
         Left            =   5415
         TabIndex        =   60
         Top             =   1680
         Width           =   90
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Index           =   5
         Left            =   5415
         TabIndex        =   59
         Top             =   2040
         Width           =   90
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   6
         Left            =   5415
         TabIndex        =   58
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Index           =   7
         Left            =   5415
         TabIndex        =   57
         Top             =   2760
         Width           =   90
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   8
         Left            =   5415
         TabIndex        =   56
         Top             =   3120
         Width           =   90
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Index           =   9
         Left            =   5415
         TabIndex        =   55
         Top             =   3480
         Width           =   90
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   10
         Left            =   5415
         TabIndex        =   54
         Top             =   3840
         Width           =   90
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Index           =   11
         Left            =   5415
         TabIndex        =   53
         Top             =   4200
         Width           =   90
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   12
         Left            =   5415
         TabIndex        =   52
         Top             =   4560
         Width           =   90
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Index           =   13
         Left            =   5415
         TabIndex        =   51
         Top             =   4920
         Width           =   90
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   14
         Left            =   5415
         TabIndex        =   50
         Top             =   5280
         Width           =   90
      End
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
      Left            =   7680
      TabIndex        =   3
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
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
      Left            =   8880
      TabIndex        =   2
      Top             =   8160
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
      Left            =   10320
      TabIndex        =   1
      Top             =   8160
      Width           =   1095
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
      Left            =   11520
      TabIndex        =   0
      Top             =   8160
      Width           =   1095
   End
   Begin ServerEditor.Raise Raise1 
      Height          =   1455
      Left            =   3480
      TabIndex        =   75
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2566
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise2 
      Height          =   6015
      Left            =   3480
      TabIndex        =   76
      Top             =   1800
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10610
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise3 
      Height          =   7815
      Left            =   3360
      TabIndex        =   77
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   13785
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise4 
      Height          =   7815
      Left            =   120
      TabIndex        =   78
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   13785
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise5 
      Height          =   495
      Left            =   7560
      TabIndex        =   79
      Top             =   8040
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   873
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise6 
      Height          =   8655
      Left            =   0
      TabIndex        =   80
      Top             =   0
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   15266
      Style           =   4
      Color           =   0
   End
End
Attribute VB_Name = "frmShops"
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
lstShops.SetSelected lstShops.FindInStr(txtFind.Text), True, True
End Sub

Private Sub cboItems_Change(Index As Integer)
txtItems(Index).Text = Mid$(cboItems(Index).list(cboItems(Index).ListIndex), 2, InStr(1, cboItems(Index).list(cboItems(Index).ListIndex), ")") - 2)
End Sub

Private Sub cboItems_Click(Index As Integer)
txtItems(Index).Text = Mid$(cboItems(Index).list(cboItems(Index).ListIndex), 2, InStr(1, cboItems(Index).list(cboItems(Index).ListIndex), ")") - 2)
End Sub

Private Sub cmdNew_Click()
Dim x As Long
Dim i As Long
Dim t As Boolean
MousePointer = vbHourglass
ReDim Preserve dbShops(1 To UBound(dbShops) + 1)
x = dbShops(UBound(dbShops) - 1).iID
x = x + 1
Do Until t = True
    t = True
    i = GetShopID(x)
    If i <> 0 Then
        t = False
        x = x + 1
    End If
Loop
With dbShops(UBound(dbShops))
    .iID = x
    .sShopName = "New Shop"
    .iMarkUp = 0
    For i = LBound(.iItems) To UBound(.iItems)
        .iItems(i) = 0
        .iQ(i) = 0
    Next
End With
FillShops UBound(dbShops), True
MousePointer = vbDefault
End Sub

Private Sub cmdNext_Click()
On Error GoTo cmdNext_Click_Error
SaveShops
lcID = lcID + 1
If lcID > UBound(dbShops) Then lcID = 1
FillShops lcID
On Error GoTo 0
Exit Sub
cmdNext_Click_Error:
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo cmdPrevious_Click_Error
SaveShops
lcID = lcID - 1
If lcID < LBound(dbShops) Then lcID = UBound(dbShops)
FillShops lcID
On Error GoTo 0
Exit Sub
cmdPrevious_Click_Error:
End Sub

Private Sub cmdSave_Click()
SaveShops
End Sub

Private Sub Form_Load()
FillCBOS
FillShops FillList:=True
End Sub

Sub FillShops(Optional Arg As Long = -1, Optional FillList As Boolean = False)
Dim i As Long, j As Long
Dim m As Long
Dim Arr() As String
MousePointer = vbHourglass
bIs = True
If Arg = -1 Then Arg = LBound(dbShops)
If FillList Then lstShops.Clear
For i = LBound(dbShops) To UBound(dbShops)
    With dbShops(i)
        If FillList Then lstShops.AddItem CStr(.iID & " " & .sShopName)
        If Arg = i Then
            lcID = i
            txtShopName.Text = .sShopName
            txtID.Text = .iID
            txtMarkUp.Text = .iMarkUp
            For j = cboItems.LBound To cboItems.UBound
                txtQ(j).Text = .iQ(j)
                txtItems(j).Text = .iItems(j)
                If .iItems(j) <> 0 Then
                    m = GetItemID(, .iItems(j))
                    If m <> 0 Then
                        lblP(j).Caption = CStr(Round(dbItems(m).dCost + (dbItems(m).dCost * (.iMarkUp / 100)), 0)) & " gold"
                        If lblP(j).Caption = "0 gold" Then lblP(j).Caption = "Free"
                    End If
                Else
                    lblP(j).Caption = "N/A"
                End If
            Next
            If Not FillList Then Exit For
        End If
    End With
Next
modMain.SetLstSelected lstShops, txtID.Text & " " & txtShopName.Text
bIs = False
MousePointer = vbDefault
End Sub

Sub FillCBOS()
Dim i As Long
Dim j As Long
For i = cboItems.LBound To cboItems.UBound
    cboItems(i).Clear
    cboItems(i).AddItem "(0) None"
    For j = LBound(dbItems) To UBound(dbItems)
        With dbItems(j)
            cboItems(i).AddItem "(" & .iID & ") " & .sItemName
        End With
    Next
Next
End Sub

Sub SaveShops()
Dim i As Long
MousePointer = vbHourglass
With dbShops(lcID)
    .iID = Val(txtID.Text)
    For i = cboItems.LBound To cboItems.UBound
        .iItems(i) = Val(txtItems(i).Text)
        .iQ(i) = Val(txtQ(i).Text)
    Next
    .iMarkUp = Val(txtMarkUp.Text)
    .sShopName = txtShopName.Text
End With
SaveMemoryToDatabase Shop
MousePointer = vbDefault
FillShops lcID, True
End Sub

Private Sub lstShops_Click()
If bIs Then Exit Sub
Dim i As Long
MousePointer = vbHourglass
For i = LBound(dbShops) To UBound(dbShops)
    With dbShops(i)
        If .iID & " " & .sShopName = lstShops.ItemText Then
            FillShops i
            Exit For
        End If
    End With
Next
MousePointer = vbDefault
End Sub

Private Sub txtItems_Change(Index As Integer)
Dim m As Long
If txtItems(Index).Text <> "0" Then
    modMain.SetCBOSelectByID cboItems(Index), txtItems(Index).Text
    m = GetItemID(, Val(txtItems(Index).Text))
    If m <> 0 Then
        lblP(Index).Caption = CStr(Round(dbItems(m).dCost + (dbItems(m).dCost * (Val(txtMarkUp.Text) / 100)), 0)) & " gold"
        If lblP(Index).Caption = "0 gold" Then lblP(Index).Caption = "Free"
    End If
Else
    SetListIndex cboItems(Index), "(0) None"
    lblP(Index).Caption = "N/A"
End If
End Sub

Private Sub txtMarkUp_Change()
Dim i As Long
Dim m As Long
For i = txtQ.LBound To txtQ.UBound
    If txtItems(i).Text <> "0" Then
        m = GetItemID(, Val(txtItems(i).Text))
        If m <> 0 Then
            lblP(i).Caption = CStr(Round(dbItems(m).dCost + (dbItems(m).dCost * (Val(txtMarkUp.Text) / 100)), 0)) & " gold"
            If lblP(i).Caption = "0 gold" Then lblP(i).Caption = "Free"
        End If
    Else
        lblP(i).Caption = "N/A"
    End If
Next
End Sub
