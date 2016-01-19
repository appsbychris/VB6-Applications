VERSION 5.00
Begin VB.Form Form7 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sort Hand"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3780
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   133
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   252
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHoldTrans 
      AutoRedraw      =   -1  'True
      Height          =   2535
      Left            =   240
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
      Begin VB.Image imgTransSort 
         Height          =   2415
         Left            =   240
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.PictureBox picCard 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   2400
      ScaleHeight     =   1500
      ScaleWidth      =   1125
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1125
   End
   Begin TicBoard.Deck Deck1 
      Left            =   1920
      Top             =   0
      _ExtentX        =   1429
      _ExtentY        =   2090
      Picture         =   "Form7.frx":1910A
   End
   Begin VB.ListBox lstSort 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1815
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label cmdMove 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   1
      Left            =   1800
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.Label cmdMove 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.Image imgDown 
      Height          =   510
      Left            =   1575
      Picture         =   "Form7.frx":19126
      Top             =   1170
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgUp 
      Height          =   540
      Left            =   1575
      Picture         =   "Form7.frx":19AF8
      Top             =   165
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FU As New Functions 'the class module
Dim pHand() As Integer 'tempary array

Private Sub cmdMove_Click(Index As Integer)
Dim tMoveStr%, tAboveStr% 'holds values
Dim tMoveInt%, tAboveInt%
On Error GoTo cmdMove_Click_Error
Select Case Index
    Case 0: 'if moveing up 1
        If lstSort.ListIndex > 0 Then 'make sure it can go up 1
            tMoveStr% = pHand(lstSort.ListIndex) 'store the number
            tMoveInt% = lstSort.ListIndex 'store the index
            tAboveStr% = pHand(lstSort.ListIndex - 1) 'get the new number
            tAboveInt% = lstSort.ListIndex - 1 'get the new index
            pHand(tAboveInt%) = tMoveStr% 'switch em around
            pHand(tMoveInt%) = tAboveStr%
        End If
    Case 1:
        If lstSort.ListIndex < lstSort.ListCount - 1 Then 'if can move down 1
            tMoveStr% = pHand(lstSort.ListIndex) 'store the number
            tMoveInt% = lstSort.ListIndex 'store the index
            tAboveStr% = pHand(lstSort.ListIndex + 1) 'store the new number
            tAboveInt% = lstSort.ListIndex + 1 'store the new index
            pHand(tAboveInt%) = tMoveStr% 'switch em around
            pHand(tMoveInt%) = tAboveStr%
        End If
End Select
lstSort.Clear 'clear the box
For i = 0 To UBound(pHand) 're-insert everything
    If pHand(i) <> -1 Then lstSort.AddItem Deck1.GetCardName(pHand(i)) 'get the card name
Next
lstSort.Selected(tAboveInt%) = True 'reselect the current item
On Error GoTo 0
Exit Sub
cmdMove_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdMove_Click in Form, Form7"
End Sub

Private Sub cmdMove_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UHAll 'unhighlight all the buttons
Select Case Index
    Case 0:
        imgUp.Visible = True 'highlight hte up arrow
    Case 1:
        imgDown.Visible = True 'highlight the down arrow
End Select
End Sub

Private Sub Form_Load()
Dim ret As Long, rRect As RECT 'stuff to trim off border
On Error GoTo Form_Load_Error
pHand() = CardsInHand() 'get the array
For i = 0 To UBound(pHand) 'add it the the lstbox
    If pHand(i) <> -1 Then lstSort.AddItem Deck1.GetCardName(pHand(i))
Next
'get the new values for lstsorts height and width
rRect.lTop = 1
rRect.lLeft = 1
rRect.lRight = lstSort.Width - 1
rRect.lBottom = lstSort.Height - 1
'get the new rect for lstsort
ret = CreateRectRgnIndirect(rRect)
SetWindowRgn lstSort.hwnd, ret, True 'set the new rect for lstsort
picHoldTrans.Cls 'clear  the pic box
'get the picture that is under lstsort
BitBlt picHoldTrans.hdc, 0, 0, lstSort.Width + 1, lstSort.Height + 1, Me.hdc, lstSort.Left + 1, lstSort.Top + 1, vbSrcCopy
'set that picture to an image box
imgTransSort.Picture = picHoldTrans.Image
Me.ScaleMode = 1 'twips
SortBrush = CreatePatternBrush(imgTransSort.Picture.Handle) 'create a new pattern brush
'begin subclassing
oldWindowSortProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf NewWindowSortProc)
oldSortListProc = SetWindowLong(lstSort.hwnd, GWL_WNDPROC, AddressOf NewSortListProc)
'lstSort.Refresh 'refresh the listbox
cmdMove(0).Top = imgUp.Top 'set the top/left/width/height for the labels
    'that act as our buttons
cmdMove(0).Left = imgUp.Left
cmdMove(0).Height = imgUp.Height
cmdMove(0).Width = imgUp.Width
cmdMove(1).Top = imgDown.Top
cmdMove(1).Left = imgDown.Left
cmdMove(1).Height = imgDown.Height
cmdMove(1).Width = imgDown.Width
Me.Show
Dim Temp() As String
ReDim Temp(lstSort.ListCount - 1) As String
For i = 0 To lstSort.ListCount - 1
    Temp(i) = lstSort.List(i)
Next
lstSort.Clear
lstSort.Refresh
For i = 0 To UBound(Temp)
    lstSort.AddItem Temp(i)
Next
lstSort.Refresh
On Error GoTo 0
Exit Sub
Form_Load_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_Load in Form, Form7"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Form_MouseMove_Error
Call UHAll 'unhighlight all the buttons
On Error GoTo 0
Exit Sub
Form_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_MouseMove in Form, Form7"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Form_Unload_Error
For i = 0 To UBound(CardsInHand) 'reset the array to the cards inhand
    CardsInHand(i) = pHand(i)
Next
Form1.RedoHand 'redraw the hand on form1
'unsubclass
SetWindowLong Me.hwnd, GWL_WNDPROC, oldWindowSortProc
SetWindowLong lstSort.hwnd, GWL_WNDPROC, oldSortListProc
'delete our brush
DeleteObject SortBrush
'take off top
FU.PutOnTop Me, False
On Error GoTo 0
Exit Sub
Form_Unload_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_Unload in Form, Form7"
End Sub

Private Sub lstSort_Click()
'get the picture ofthe card
On Error GoTo lstSort_Click_Error
Deck1.GetAnotherCard = pHand(lstSort.ListIndex)
picCard.Picture = Deck1.Picture 'set the pic to the pic box
picCard.Visible = True 'make sure thepic box is visible true
On Error GoTo 0
Exit Sub
lstSort_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: lstSort_Click in Form, Form7"
End Sub

Private Sub lstSort_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo lstSort_MouseMove_Error
Call UHAll 'unhighlight the buttons
On Error GoTo 0
Exit Sub
lstSort_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: lstSort_MouseMove in Form, Form7"
End Sub

Private Sub picCard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo picCard_MouseMove_Error
Call UHAll 'unhighlight the buttons
On Error GoTo 0
Exit Sub
picCard_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: picCard_MouseMove in Form, Form7"
End Sub

Sub UHAll()
'make the highlihghs visible false
On Error GoTo UHAll_Error
imgUp.Visible = False
imgDown.Visible = False
On Error GoTo 0
Exit Sub
UHAll_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: UHAll in Form, Form7"
End Sub
