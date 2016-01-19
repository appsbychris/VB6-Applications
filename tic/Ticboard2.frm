VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Connect to server"
   ClientHeight    =   6195
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   8850
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Ticboard2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Ticboard2.frx":08CA
   ScaleHeight     =   6195
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   3720
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2400
      Width           =   4695
   End
   Begin VB.PictureBox picTexts 
      AutoRedraw      =   -1  'True
      Height          =   1335
      Left            =   3600
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   405
      TabIndex        =   3
      Top             =   6000
      Width           =   6135
      Begin VB.Image imgText2 
         Height          =   375
         Left            =   120
         Top             =   720
         Width           =   5775
      End
      Begin VB.Image imgText1 
         Height          =   375
         Left            =   120
         Top             =   120
         Width           =   5775
      End
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Connect"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3960
      MouseIcon       =   "Ticboard2.frx":C2DD
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Image imgHConnect 
      Height          =   570
      Left            =   4080
      Picture         =   "Ticboard2.frx":C5E7
      Top             =   4800
      Width           =   1665
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   6840
      MouseIcon       =   "Ticboard2.frx":CE52
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C&ancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5640
      MouseIcon       =   "Ticboard2.frx":D15C
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Image imgHCancel 
      Height          =   540
      Left            =   5760
      Picture         =   "Ticboard2.frx":D466
      Top             =   4800
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Image imghX 
      Height          =   555
      Left            =   6800
      Picture         =   "Ticboard2.frx":DCA0
      Top             =   880
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "Form2"
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

Private Sub Command1_Click()
On Error GoTo Command1_Click_Error
'make sure theres stuff in the text box
If Text1.Text <> "" And Text2.Text <> "" Then
    Form1.ws.Close 'close any connections
    Form1.ws.RemoteHost = Text1.Text 'set the IP
    YName$ = Text2.Text 'save the nickname
    Form1.ws.Connect 'connect to the server
    Unload Me 'unload this form
End If
On Error GoTo 0
Exit Sub
Command1_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Command1_Click in Form, Form2"
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Command1_MouseMove_Error
Call RemoveH 'remove all the highlighted buttons
imgHConnect.Visible = True 'and make the highlighted connect button visible
On Error GoTo 0
Exit Sub
Command1_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Command1_MouseMove in Form, Form2"
End Sub

Private Sub Form_Load()
On Error GoTo Form_Load_Error
Form2.ScaleMode = 3 'pixels
picTexts.Cls 'clear the picbox
'get the pic behind the first box
BitBlt picTexts.hdc, 0, 0, Text1.Width, Text1.Height, Form2.hdc, Text1.Left, Text1.Top, vbSrcCopy
'set the pic to an image box
imgText1.Picture = picTexts.Image
picTexts.Cls 'clear the picbox
'get the pic behind the second box
BitBlt picTexts.hdc, 0, 0, Text2.Width, Text2.Height, Form2.hdc, Text2.Left, Text2.Top, vbSrcCopy
'set that pic to an imagebox
imgText2.Picture = picTexts.Image
'make that pic box invisible
picTexts.Visible = False
'set back to twips
Form2.ScaleMode = 1
'set our pattern brushes so we can create transparent textboxes
txtConBox1 = CreatePatternBrush(imgText1.Picture.Handle)
txtConBox2 = CreatePatternBrush(imgText2.Picture.Handle)
'begin subclassing
oldWindowProc3 = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf NewWindowProc3)
'unhighlight all buttons
Call RemoveH
On Error GoTo 0
Exit Sub
Form_Load_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_Load in Form, Form2"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Form_MouseDown_Error
'4 arraw cursor
MousePointer = 15
'move the form according to the currsor
Call ReleaseCapture
Call SendMessage(hwnd, &HA1, 2, 0&)
MousePointer = 1 'default the mouse currsor
On Error GoTo 0
Exit Sub
Form_MouseDown_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_MouseDown in Form, Form2"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Form_MouseMove_Error
Call RemoveH 'unhighlight the buttons
On Error GoTo 0
Exit Sub
Form_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_MouseMove in Form, Form2"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'unsubclass the window
SetWindowLong Me.hwnd, GWL_WNDPROC, oldWindowProc3
'delete the pattern brushes
DeleteObject txtConBox1
DeleteObject txtConBox2
End Sub

Private Sub Label1_Click()
On Error GoTo Label1_Click_Error
Unload Me 'unload the form
On Error GoTo 0
Exit Sub
Label1_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Label1_Click in Form, Form2"
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Label1_MouseMove_Error
Call RemoveH 'unhighlight the buttons
imgHCancel.Visible = True 'make the cancel highlight visible
On Error GoTo 0
Exit Sub
Label1_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Label1_MouseMove in Form, Form2"
End Sub

Private Sub Label2_Click()
On Error GoTo Label2_Click_Error
Unload Me 'unload the form
On Error GoTo 0
Exit Sub
Label2_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Label2_Click in Form, Form2"
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Label2_MouseMove_Error
Call RemoveH 'unhighlight all the buttons
imgHX.Visible = True 'highlight the X button
On Error GoTo 0
Exit Sub
Label2_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Label2_MouseMove in Form, Form2"
End Sub

Sub RemoveH()
On Error GoTo RemoveH_Error
'unihighlight all the buttons
imgHConnect.Visible = False
imgHCancel.Visible = False
imgHX.Visible = False
On Error GoTo 0
Exit Sub
RemoveH_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: RemoveH in Form, Form2"
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then 'enter key
    Call Command1_Click 'submit the form
End If
End Sub
