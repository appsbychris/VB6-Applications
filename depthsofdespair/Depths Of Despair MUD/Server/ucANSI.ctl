VERSION 5.00
Begin VB.UserControl ucANSI 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8955
   ScaleHeight     =   2730
   ScaleWidth      =   8955
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   115
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   116
      Top             =   2280
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ê"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   93
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Â"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   94
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "‰"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   88
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "‚"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   92
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Á"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   90
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "·"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   86
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Î"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   101
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ó"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   103
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "‡"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   80
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   2280
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "˝"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   113
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":0000
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   65
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":0005
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   61
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "È"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   89
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "í"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   62
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ù"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   73
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":000A
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   45
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   1200
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ï"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   100
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   1200
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "˜"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   109
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   2280
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "˚"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   105
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   2280
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Û"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   36
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2280
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":000F
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   66
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "™"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   69
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "©"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   72
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ﬁ"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   81
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   79
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   1080
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "˛"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   114
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   115
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "¯"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   112
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "˘"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   111
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "˙"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   108
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "¸"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   107
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ô"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   102
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   2280
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   ""
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   99
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ú"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   98
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   2280
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ì"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   97
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   2280
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ò"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   96
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   2280
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Í"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   95
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   2280
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "„"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   91
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   2280
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ë"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   87
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   2280
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "‹"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   84
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "≤"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   82
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   1080
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ﬂ"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   78
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ü"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   76
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "®"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   75
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ú"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   74
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":0014
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   67
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":0019
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   64
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   ""
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   60
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":001E
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   59
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":0023
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   57
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":0028
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   58
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":002D
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   43
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":0032
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   39
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":0037
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   53
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   600
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":003C
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   42
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   600
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":0041
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   48
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":0046
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   52
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":004B
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   49
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":0050
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   54
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":0055
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   51
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":005A
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   50
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ø"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   47
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   600
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":005F
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   46
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   1080
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":0064
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   44
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":0069
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   41
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "º"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "π"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   30
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Œ"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   33
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ω"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   35
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ÿ"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   32
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ª"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "æ"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "µ"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "∏"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   29
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "–"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "œ"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "À"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "≥"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1320
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "”"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "¡"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ÿ"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ƒ"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Õ"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "¿"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "∫"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "‘"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "∆"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "»"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ã"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "…"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "¥"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   34
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ø"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   40
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "∂"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   38
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "∑"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "—"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "’"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "≈"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   31
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "√"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "¬"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "⁄"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "◊"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   37
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "«"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   116
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "“"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   28
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":006E
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   63
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":0073
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   55
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"ucANSI.ctx":0078
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   56
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "±"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   83
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   1080
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "›"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   85
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "∞"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   77
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   1080
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ˆ"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   110
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   2280
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Æ"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   71
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   600
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ı"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   106
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ù"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   104
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   1080
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "û"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   70
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   1800
      Width           =   375
   End
   Begin VB.OptionButton optChar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "õ"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   68
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   1440
      Width           =   375
   End
End
Attribute VB_Name = "ucANSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : ucANSI
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Private bInit As Boolean

Public StringChar As String

Private Sub optChar_Click(Index As Integer)
StringChar = optChar(Index).Caption
If Index = 115 Then StringChar = Chr$(0)
End Sub

Private Sub UserControl_Initialize()
optChar(115).Value = True
StringChar = Chr$(0)
End Sub

Private Sub UserControl_Resize()
With UserControl
    .Height = 2730
    .Width = 8955
End With
End Sub


