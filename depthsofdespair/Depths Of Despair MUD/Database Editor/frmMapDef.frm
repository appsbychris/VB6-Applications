VERSION 5.00
Begin VB.Form frmMapDef 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Room Definition"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3765
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
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optIndoor 
      Caption         =   "&Outdoor"
      Height          =   375
      Index           =   0
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3720
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton optIndoor 
      Caption         =   "&Indoor"
      Height          =   375
      Index           =   1
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4200
      Width           =   1095
   End
   Begin VB.OptionButton optIndoor 
      Caption         =   "&Underground"
      Height          =   375
      Index           =   2
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "Join To Existing Room"
      Height          =   375
      Left            =   1200
      TabIndex        =   13
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox txtDeath 
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox txtLight 
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox txtMax 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtMob 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txtTitle 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3495
   End
   Begin VB.TextBox txtDesc 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Do not Auto Enchance"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Death Room:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   4680
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Light:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Mac Regen:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Mob Group:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Room Title:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "frmMapDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
With udtMapArea(frmMapEdit.prevSel)
    .lAuto = Me.Check1.Value
End With
End Sub


Private Sub cmdJoin_Click()
Load frmRoomJoin
frmRoomJoin.Show
Me.Enabled = False
frmMapEdit.Enabled = False
End Sub

Private Sub optIndoor_Click(Index As Integer)
udtMapArea(frmMapEdit.prevSel).lIndoor = Index
End Sub

Private Sub txtDeath_Change()
With udtMapArea(frmMapEdit.prevSel)
    .lDeath = Val(txtDeath.Text)
End With
End Sub

Private Sub txtDesc_Change()
With udtMapArea(frmMapEdit.prevSel)
    .sDesc = Me.txtDesc.Text
End With
End Sub

Private Sub txtLight_Change()
With udtMapArea(frmMapEdit.prevSel)
    .lLight = Val(txtLight.Text)
End With
End Sub

Private Sub txtMax_Change()
With udtMapArea(frmMapEdit.prevSel)
    .lMaxRegen = Val(txtMax.Text)
End With
End Sub

Private Sub txtMob_Change()
With udtMapArea(frmMapEdit.prevSel)
    .lMob = Val(txtMob.Text)
End With
End Sub

Private Sub txtTitle_Change()
With udtMapArea(frmMapEdit.prevSel)
    .sTitle = Me.txtTitle.Text
End With
End Sub
