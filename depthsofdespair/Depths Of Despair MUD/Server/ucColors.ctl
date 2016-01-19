VERSION 5.00
Begin VB.UserControl ucColors 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
   ScaleHeight     =   1080
   ScaleWidth      =   6450
   Begin VB.OptionButton optCol 
      BackColor       =   &H00808000&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   19
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   360
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H00800080&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   18
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   360
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H00800000&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   17
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   360
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H00008080&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   16
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   360
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H00008000&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   15
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   360
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H00000080&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   14
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   360
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   600
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H00FFFF00&
      Height          =   375
      Index           =   12
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   600
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H00FF00FF&
      Height          =   375
      Index           =   11
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   600
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H00FF0000&
      Height          =   375
      Index           =   10
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   9
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H0000FF00&
      Height          =   375
      Index           =   8
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Index           =   7
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   6
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H00808000&
      Height          =   375
      Index           =   5
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H00800080&
      Height          =   375
      Index           =   4
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H00800000&
      Height          =   375
      Index           =   3
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H0000C0C0&
      Height          =   375
      Index           =   2
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H00008000&
      Height          =   375
      Index           =   1
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optCol 
      BackColor       =   &H00000080&
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "ucColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : ucColors
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public lColor As Long
Public IsBC As Boolean

Private Sub optCol_Click(Index As Integer)
If Index < 14 Then
    lColor = optCol(Index).BackColor
    IsBC = False
Else
    lColor = optCol(Index).BackColor
    IsBC = True
End If
End Sub

Private Sub UserControl_Initialize()
lColor = &HC0C0C0
IsBC = False
End Sub

Private Sub UserControl_Resize()
With UserControl
    .Height = 1080
    .Width = 6450
End With
End Sub
