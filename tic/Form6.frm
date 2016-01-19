VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   Caption         =   "Credits"
   ClientHeight    =   7260
   ClientLeft      =   1800
   ClientTop       =   2310
   ClientWidth     =   3345
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   7260
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   Begin VB.Label cmdOK 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1800
      MouseIcon       =   "Form6.frx":4FEA2
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Image imgHOK 
      Height          =   540
      Left            =   1850
      Picture         =   "Form6.frx":501AC
      Top             =   5660
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label lblCredits 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
On Error GoTo cmdOK_Click_Error
Unload Me 'unload the form
On Error GoTo 0
Exit Sub
cmdOK_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdOK_Click in Form, Form6"
End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo cmdOK_MouseMove_Error
imgHOK.Visible = True 'make the hightlight visible
On Error GoTo 0
Exit Sub
cmdOK_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdOK_MouseMove in Form, Form6"
End Sub

Private Sub Form_Load()
Dim Credits$
On Error GoTo Form_Load_Error
'make the credits
Credits$ = "Thanks go out to-" & vbCrLf & vbCrLf & _
        "Beta Testers:" & vbCrLf & _
        "  Dave Morrison" & vbCrLf & _
        "  Justin Wilson" & vbCrLf & _
        "  Zach Anderson" & vbCrLf & _
        "  Mike Magnision" & vbCrLf & vbCrLf & _
        "Help with code:" & vbCrLf & _
        "  ""The Hand""" & vbCrLf & vbCrLf & _
        "Places that helped with code:" & vbCrLf & _
        "  www.visualbasicforum.com" & vbCrLf & _
        "  www.planetsourcecode.com" & vbCrLf & _
        "  www.allapi.net" & vbCrLf & _
        "  www.elitevb.com/home.asp" & vbCrLf & vbCrLf & _
        "Thanks to everyone listed." & vbCrLf & _
        "If I forgot to list you, sorry."
lblCredits.Caption = Credits$ 'set the credits
On Error GoTo 0
Exit Sub
Form_Load_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_Load in Form, Form6"
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
imgHOK.Visible = False 'make the highlight invisible
On Error GoTo 0
Exit Sub
Form_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_MouseMove in Form, Form6"
End Sub

Private Sub lblCredits_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo lblCredits_MouseMove_Error
imgHOK.Visible = False 'make the highlight disapear
On Error GoTo 0
Exit Sub
lblCredits_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: lblCredits_MouseMove in Form, Form6"
End Sub
