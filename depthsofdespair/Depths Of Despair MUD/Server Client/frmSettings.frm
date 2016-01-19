VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Connection Settings"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
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
   MDIChild        =   -1  'True
   ScaleHeight     =   2250
   ScaleWidth      =   6795
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "&Save As..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtAddress 
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
      Left            =   3960
      TabIndex        =   4
      Text            =   "127.0.0.1"
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CheckBox chkAutoLog 
      Caption         =   "Auto Log On When Connected?"
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
      Left            =   3960
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   720
      Width           =   2655
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
      Left            =   3960
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.FileListBox filCharacter 
      Height          =   1455
      Left            =   120
      Pattern         =   "*.char"
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telnet Address:"
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
      Left            =   2760
      TabIndex        =   11
      Top             =   1440
      Width           =   1140
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Left            =   3120
      TabIndex        =   10
      Top             =   720
      Width           =   750
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Character Name:"
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
      Left            =   2640
      TabIndex        =   9
      Top             =   360
      Width           =   1230
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a character:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1470
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub filCharacter_Click()
LoadCharacter filCharacter.Path & "\" & filCharacter.List(filCharacter.ListIndex)
End Sub

Private Sub Form_Load()
filCharacter = App.Path

End Sub

Private Sub LoadCharacter(sPath As String)
Dim s As String
Dim i As Long
Dim j As Long
Open sPath For Input As #1
    s = Input$(LOF(1), 1)
Close #1
'FILE HEAD:
'*.CHAR:name.*
'*.PASS:pass.*
'*.AUTO:0.*
'*.ADDR:127.*
i = InStr(1, s, "*.CHAR:") + 7
j = InStr(i, s, ".*")
txtName.Text = Mid$(s, i, j - i)
i = InStr(1, s, "*.PASS:") + 7
j = InStr(i, s, ".*")
txtPassword.Text = Mid$(s, i, j - i)
i = InStr(1, s, "*.AUTO:") + 7
j = InStr(i, s, ".*")
chkAutoLog.Value = Val(Mid$(s, i, j - i))
i = InStr(1, s, "*.ADDR:") + 7
j = InStr(i, s, ".*")
txtAddress.Text = Mid$(s, i, j - i)
End Sub
