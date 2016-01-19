VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About MyApp"
   ClientHeight    =   6600
   ClientLeft      =   2295
   ClientTop       =   1500
   ClientWidth     =   8190
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":08CA
   ScaleHeight     =   4555.438
   ScaleMode       =   0  'User
   ScaleWidth      =   7690.833
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1920
      Picture         =   "frmAbout.frx":14B24
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   0
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label cmdCredits 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   5400
      MouseIcon       =   "frmAbout.frx":153EE
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label cmdOK 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      MouseIcon       =   "frmAbout.frx":156F8
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Image imgHCredits 
      Height          =   465
      Left            =   5367
      Picture         =   "frmAbout.frx":15A02
      Top             =   5028
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image imgHOK 
      Height          =   450
      Left            =   3589
      Picture         =   "frmAbout.frx":16303
      Top             =   5028
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   675
      Left            =   6360
      MouseIcon       =   "frmAbout.frx":16B89
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Chris Van Hooser."
      Height          =   225
      Left            =   3840
      TabIndex        =   3
      Top             =   3840
      Width           =   2235
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      UseMnemonic     =   0   'False
      Width           =   1290
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   225
      Left            =   3840
      TabIndex        =   2
      Top             =   2880
      Width           =   585
   End
   Begin VB.Image imgHX 
      Height          =   675
      Left            =   6360
      Picture         =   "frmAbout.frx":16E93
      Top             =   1080
      Visible         =   0   'False
      Width           =   675
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************
'*************************************************************************************
'***************       Code create by Chris Van Hooser          **********************
'***************                  (c)2001                       **********************
'*************** You may use this code and freely distribute it **********************
'***************   If you have any questions, please email me   **********************
'***************          at theendorbunker@attbi.com.          **********************
'***************       Thanks for downloading my project        **********************
'***************        and i hope you can use it well.         **********************
'***************                TicBoard                        **********************
'***************                TicBoard.vbp                    **********************
'*************************************************************************************
'*************************************************************************************


Private Sub cmdCredits_Click()
On Error GoTo cmdCredits_Click_Error
Load Form6 'load credits form
Form6.Show 1
On Error GoTo 0
Exit Sub
cmdCredits_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdCredits_Click in Form, frmAbout"
End Sub

Private Sub cmdCredits_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo cmdCredits_MouseMove_Error
Call UnHall 'unhighlight all buttons
imgHCredits.Visible = True 'make this on visible
On Error GoTo 0
Exit Sub
cmdCredits_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdCredits_MouseMove in Form, frmAbout"
End Sub

Private Sub cmdOK_Click()
On Error GoTo cmdOK_Click_Error
  Unload Me 'unload the form
On Error GoTo 0
Exit Sub
cmdOK_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdOK_Click in Form, frmAbout"
End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo cmdOK_MouseMove_Error
Call UnHall 'unhighlight all the buttons
imgHOK.Visible = True 'make this one true
On Error GoTo 0
Exit Sub
cmdOK_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdOK_MouseMove in Form, frmAbout"
End Sub

Private Sub Form_Load()
On Error GoTo Form_Load_Error
Me.Caption = "About " & App.Title 'get the app title
lblVersion.Caption = "Version " & App.Major & "." & App.Minor 'get the app version
lblTitle.Caption = App.Title 'get the title
On Error GoTo 0
Exit Sub
Form_Load_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_Load in Form, frmAbout"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Form_MouseDown_Error
MousePointer = 15 '4 arrow pointer
Call ReleaseCapture 'move the form with the cursor
Call SendMessage(hwnd, &HA1, 2, 0&)
MousePointer = 1 'default cursor
On Error GoTo 0
Exit Sub
Form_MouseDown_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_MouseDown in Form, frmAbout"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Form_MouseMove_Error
Call UnHall 'unhighlight all the buttons
On Error GoTo 0
Exit Sub
Form_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_MouseMove in Form, frmAbout"
End Sub

Private Sub Label2_Click()
On Error GoTo Label2_Click_Error
Unload Me 'unload the form
On Error GoTo 0
Exit Sub
Label2_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Label2_Click in Form, frmAbout"
End Sub

Sub UnHall()
imgHCredits.Visible = False
imgHOK.Visible = False 'unhighlight all the buttons
imgHX.Visible = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Label2_MouseMove_Error
Call UnHall 'unhighlight the buttons
imgHX.Visible = True 'make this on visible
On Error GoTo 0
Exit Sub
Label2_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Label2_MouseMove in Form, frmAbout"
End Sub
