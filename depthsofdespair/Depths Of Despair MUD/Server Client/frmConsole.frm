VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConsole 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Console - XXX"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   9615
   FillColor       =   &H00FFFFFF&
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   9615
   Begin ServerClient.ctlTelnet ctlTelnet1 
      Height          =   4500
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9600
      _ExtentX        =   25400
      _ExtentY        =   10583
   End
   Begin VB.Frame frmFrame 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   5880
      Width           =   10575
      Begin MSComctlLib.Toolbar CTB 
         Height          =   585
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   1032
         ButtonWidth     =   2302
         ButtonHeight    =   979
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Auto-Attack"
               Style           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Auto-Get Cash"
               Style           =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Auto-Rest"
               Style           =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Auto-Heal"
               Style           =   1
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Line lneSep 
         BorderColor     =   &H00000080&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   10515
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line lneSep 
         BorderColor     =   &H00000040&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Index           =   1
         X1              =   45
         X2              =   10560
         Y1              =   150
         Y2              =   150
      End
      Begin VB.Label lblStatline 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "[HP=0/0, MA=0,0]"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1920
      End
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   8640
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   23
   End
   Begin VB.Timer timFlash 
      Interval        =   400
      Left            =   10080
      Top             =   120
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuHangUp 
         Caption         =   "&Hang Up"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuDashkdafk21l23 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSize 
         Caption         =   "&Size"
         Begin VB.Menu mnuSmall 
            Caption         =   "&Small Version"
         End
         Begin VB.Menu mnuMedium 
            Caption         =   "&Medium Version"
         End
         Begin VB.Menu mnuLarge 
            Caption         =   "&Large Version"
         End
      End
      Begin VB.Menu mnuDashdfsa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim sTyped As String
Private Const UP_ARROW        As String = "[A"
Private Const DOWN_ARROW      As String = "[B"
Private Const RIGHT_ARROW     As String = "[C"
Private Const LEFT_ARROW      As String = "[D"

Private Sub CTB_ButtonClick(ByVal Button As MSComctlLib.Button)
With uSettings
    Select Case Button.Index
        Case 1
            .iAutoAttack = IIf(.iAutoAttack = 1, 0, 1)
        Case 2
            .iAutoGetCash = IIf(.iAutoGetCash = 1, 0, 1)
        Case 3
            .lRestIfBelow = CLng(InputBox("below what?", "rest below", "15"))
            .iAutoRest = IIf(.iAutoHeal = 1, 0, 1)
        Case 4
            .sHealSpell = InputBox("what spell?", "what spell", "mihe")
            .iAutoHeal = IIf(.iAutoHeal = 1, 0, 1)
    End Select
End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp
        If ws.State = 7 Then Send_Commands UP_ARROW
    Case vbKeyDown
        If ws.State = 7 Then Send_Commands DOWN_ARROW
    Case vbKeyLeft
        If ws.State = 7 Then Send_Commands LEFT_ARROW
    Case vbKeyRight
        If ws.State = 7 Then Send_Commands RIGHT_ARROW
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Static i As Long
Select Case KeyAscii
    Case 32, 33 To 126, 128, 145, 146, 161 To 255
        'ctlTelnet1.TypedText Chr$(KeyAscii)
        If ws.State = 7 Then Send_Commands Chr$(KeyAscii): i = i + 1
    Case 13
        'ctlTelnet1.TypedEnter
        If ws.State = 7 Then Send_Commands vbCrLf: i = 0
    Case vbKeyBack
        If i > 0 Then ctlTelnet1.TypedBackspace: i = i - 1
        If ws.State = 7 Then Send_Commands Chr$(vbKeyBack)
End Select
End Sub

Private Sub Form_Load()
ctlTelnet1.SetCursorEnabled True
End Sub

Private Sub Form_Unload(Cancel As Integer)
ws.Close
End Sub

Private Sub mnuConnect_Click()
With ws
    mnuHangUp_Click
    .RemoteHost = InputBox("Host:")
    .Connect
End With
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuHangUp_Click()
On Error Resume Next
ws.Close
ctlTelnet1.FeedMe "[0m[44m[Connection Closed][0m[37m"
End Sub

Private Sub mnuLarge_Click()
ctlTelnet1.SetFontSize 14
Me.Width = ctlTelnet1.Width + 100
Me.Height = ctlTelnet1.Height + 375
End Sub

Private Sub mnuMedium_Click()
ctlTelnet1.SetFontSize 12
Me.Width = ctlTelnet1.Width + 100
Me.Height = ctlTelnet1.Height + 375
End Sub

Private Sub mnuSmall_Click()
ctlTelnet1.SetFontSize 9
Me.Width = ctlTelnet1.Width + 100
Me.Height = ctlTelnet1.Height + 375
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
Dim sRec As String
ws.GetData sRec, vbString
ctlTelnet1.FeedMe sRec
End Sub

Sub Send_Commands(SendWhat As String)
With ws
    .SendData SendWhat
End With
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim stext As String
Select Case Number
    Case sckAddressNotAvailable, sckHostNotFound, sckHostNotFoundTryAgain
        stext = stext & vbCrLf & BGBLUE & "[Host Not Found]" & WHITE & vbCrLf
    Case sckAlreadyConnected
        ws.Close
    Case sckConnectionRefused
        stext = stext & vbCrLf & BGBLUE & "[Connection Refused]" & WHITE & vbCrLf
    Case sckTimedout
        stext = stext & vbCrLf & BGBLUE & "[Connection Timed Out]" & WHITE & vbCrLf
    Case Else
        stext = stext & vbCrLf & BGBLUE & "[Unknown Error Occured]" & WHITE & vbCrLf
End Select
ctlTelnet1.FeedMe stext
End Sub


