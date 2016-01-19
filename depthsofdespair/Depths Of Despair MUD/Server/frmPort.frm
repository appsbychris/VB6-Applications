VERSION 5.00
Begin VB.Form frmPort 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose another port"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPort.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Quit"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&OK"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdCustom 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Custom Port"
      Height          =   975
      Left            =   2160
      Picture         =   "frmPort.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdTest 
      BackColor       =   &H00C0C0C0&
      Caption         =   "C&heck if port is in use"
      Height          =   975
      Left            =   2160
      Picture         =   "frmPort.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.ListBox lstPorts 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin DoDMudServer.Raise Raise1 
      Height          =   2775
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4895
      Style           =   2
      Color           =   14737632
   End
End
Attribute VB_Name = "frmPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : frmPort
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Private lPort As Long

Private Sub cmdCustom_Click()
s = InputBox$("Input a custom port number.", "Port", "24")
If IsNumeric(s) Then
    With frmMain
        .ws(0).LocalPort = Val(s)
        On Error Resume Next
        .ws(0).Listen
        If Err.Number = 10048 Then
            MsgBox "That port is in use. Please choose another.", vbExclamation + vbOKOnly, "Already In Use"
            Err.Number = 0
        Else
            MsgBox "That port is open and ready for use.", vbInformation + vbOKOnly, "Port is ready for use"
            lPort = CLng(s)
            .ws(0).Close
        End If
    End With
End If
End Sub

Private Sub cmdOK_Click()
With frmMain
    .ws(0).LocalPort = lPort
    .ws(0).Listen
End With
AlwaysOnTop frmSplash, True
Unload Me
End Sub

Private Sub cmdQuit_Click()
frmMain.ShutDownServer
modDatabase.CloseRecordsets
modDatabase.CloseDatabase
Unload frmMain
Unload frmSplash
Unload frmUserDefined
Unload Me

End Sub

Private Sub cmdTest_Click()
With frmMain
    .ws(0).LocalPort = lstPorts.List(lstPorts.ListIndex)
    On Error Resume Next
    .ws(0).Listen
    If Err.Number = 10048 Then
        MsgBox "That port is in use. Please choose another.", vbExclamation + vbOKOnly, "Already In Use"
        Err.Number = 0
    Else
        MsgBox "That port is open and ready for use.", vbInformation + vbOKOnly, "Port is ready for use"
        lPort = CLng(lstPorts.List(lstPorts.ListIndex))
        .ws(0).Close
    End If
End With
End Sub

Private Sub Form_Load()
Dim i As Long
For i = 24 To 49
    lstPorts.AddItem i
Next
Screen.MousePointer = vbDefault
End Sub

