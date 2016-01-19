VERSION 5.00
Begin VB.Form frmAdvanced 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Advanced"
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   30
   ClientWidth     =   6720
   Icon            =   "frmAdvanced.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picStore 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
      Begin VB.TextBox ServerMessage 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   1080
         Width           =   2895
      End
      Begin DoDMudServer.UltraBox List1 
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1085
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
         Mult            =   0   'False
         Sort            =   0   'False
      End
      Begin DoDMudServer.eButton cmdSend 
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Style           =   2
         Cap             =   "&Send"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         hCol            =   12632256
         bCol            =   12632256
         CA              =   2
      End
      Begin DoDMudServer.Raise Raise 
         Height          =   615
         Index           =   1
         Left            =   0
         TabIndex        =   6
         Top             =   960
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   1085
         Style           =   2
         Color           =   14737632
      End
      Begin DoDMudServer.eButton cmdBoot 
         Height          =   375
         Left            =   4630
         TabIndex        =   7
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Style           =   2
         Cap             =   "&Boot selected user"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         hCol            =   12632256
         bCol            =   12632256
         CA              =   2
      End
      Begin DoDMudServer.Raise Raise 
         Height          =   855
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   1508
         Style           =   2
         Color           =   14737632
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Send a server message-"
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
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1785
      End
   End
   Begin DoDMudServer.Raise cmdBack 
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Style           =   2
      Color           =   14737632
   End
End
Attribute VB_Name = "frmAdvanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBoot_Click()
On Error Resume Next
If List1.ListIndex > 0 Then
    'Sub for booting players
    'Closes the winsok
    If Form1.ws(List1.ListIndex).State = sckConnected Then Form1.ws(List1.ListIndex).Close
    List1.SetItemText List1.ListIndex, "[Line " & CStr(List1.ListIndex) & " - Open]"   'resets list1
    'adjust the online value
    If Val(Form1.Online.Caption) > 0 Then Form1.Online.Caption = Val(Form1.Online.Caption) - 1
    'clear out defaults
    dbPlayers(GetPlayerIndexNumber(List1.ListIndex)).iIndex = 0
    X(List1.ListIndex) = ""
    PNAME(List1.ListIndex) = ""
    pPoint(List1.ListIndex) = 0
    UpdateList "}bLine " & (List1.ListIndex) & " has been booted from the server. }b(}n}i" & Time & "}n}b)"
End If
End Sub

Private Sub cmdSend_Click()
On Error Resume Next
'Sends a server message to all online
SendToAll GREEN & "[SERVER MESSAGE]: " & ServerMessage.Text & vbCrLf & WHITE
ServerMessage.Text = ""
End Sub

Private Sub Form_Load()
With cmdBack
    .Left = 0
    .Top = 0
    .Width = Me.ScaleWidth '- 50
    .Height = Me.ScaleHeight '- 50
End With
End Sub

Private Sub ServerMessage_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdSend_Click 'if Enter, then call the sub
End Sub


