VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2715
   ClientLeft      =   3210
   ClientTop       =   2280
   ClientWidth     =   4440
   LinkTopic       =   "Form2"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   2715
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblPer 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " Loading () [0%] ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1665
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " Please wait... "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub LoadServer()
Screen.MousePointer = vbHourglass
lblPer.Caption = " Loading (Loading Database) [15%] ..."
LoadDatabaseIntoMemory
lblPer.Caption = " Loading (Setting Defaults) [100%] ..."
Screen.MousePointer = vbDefault
AlwaysOnTop Me, False
Load mdiMain
mdiMain.Show
Unload Me
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width \ 2) - (Me.Width \ 2), (Screen.Height \ 2) - (Me.Height \ 2)
AlwaysOnTop Me, True

End Sub
