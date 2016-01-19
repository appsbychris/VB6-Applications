VERSION 5.00
Begin VB.Form frmIPS 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Blocked IPs"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIPS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstIPs 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin DoDMudServer.eButton cmdDel 
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Style           =   2
      Cap             =   "&Delete Selected"
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
   Begin DoDMudServer.eButton cmdClose 
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Style           =   2
      Cap             =   "&Close"
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
End
Attribute VB_Name = "frmIPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : frmIPS
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDel_Click()
If lstIPs.ListIndex <> -1 Then lstIPs.RemoveItem lstIPs.ListIndex
End Sub

Private Sub Form_Load()
Dim FileNumber&
Dim s As String
Dim tArr() As String
Dim i As Long
FileNumber = FreeFile
Open App.Path & "\ipb.list" For Binary As #FileNumber
    If DE Then DoEvents
    s = Input$(LOF(1), 1)
Close #FileNumber
SplitFast s, tArr, vbCrLf
For i = LBound(tArr) To UBound(tArr)
    If tArr(i) <> "" Then lstIPs.AddItem tArr(i)
    If DE Then DoEvents
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim s As String
Dim i As Long
Dim FileNumber&
For i = 0 To lstIPs.ListCount - 1
    s = s & lstIPs.List(i) & vbcrkf
    If DE Then DoEvents
Next
FileNumber = FreeFile
Open App.Path & "\ipb.list" For Output As #FileNumber
    If DE Then DoEvents
    Print #FileNumber, s
Close #FileNumber
End Sub
