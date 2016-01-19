VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TIC Server"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   6645
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   307
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   443
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   3360
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   381
      TabIndex        =   11
      Top             =   4200
      Width           =   5775
      Begin VB.Image Image1 
         Height          =   855
         Left            =   120
         Top             =   0
         Width           =   5655
      End
   End
   Begin VB.ListBox List3 
      Height          =   5325
      Left            =   8760
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Height          =   5325
      Left            =   7560
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   975
      IntegralHeight  =   0   'False
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   0
      Left            =   6120
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3220
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start Game"
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
      Height          =   615
      Left            =   3600
      MouseIcon       =   "Form1.frx":8847
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Image imgHCommand1 
      Height          =   810
      Left            =   3600
      Picture         =   "Form1.frx":8B51
      Top             =   1440
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your IP address is:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3840
      TabIndex        =   13
      Top             =   2520
      Width           =   1545
   End
   Begin VB.Label lblIP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   3480
      TabIndex        =   12
      Top             =   2880
      Width           =   120
   End
   Begin VB.Label Albl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Turn"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   2280
      TabIndex        =   9
      Top             =   1560
      Width           =   1050
   End
   Begin VB.Label Albl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player's Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   720
      TabIndex        =   8
      Top             =   1560
      Width           =   1155
   End
   Begin VB.Label Albl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Points"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   465
   End
   Begin VB.Label Albl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-Line Status-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   2640
      TabIndex        =   6
      Top             =   0
      Width           =   990
   End
   Begin VB.Label lblPoints 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblPointer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<-------"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2400
      TabIndex        =   4
      Top             =   1800
      Width           =   840
   End
   Begin VB.Label lblTurn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuStartGame 
         Caption         =   "&Start Game"
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAmount 
         Caption         =   "&Amount of Decks to Use"
         Begin VB.Menu mnuAmountDeck 
            Caption         =   "2 (&Default)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuAmountDeck 
            Caption         =   "3"
            Index           =   1
         End
         Begin VB.Menu mnuAmountDeck 
            Caption         =   "4"
            Index           =   2
         End
         Begin VB.Menu mnuAmountDeck 
            Caption         =   "5"
            Index           =   3
         End
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHMen 
      Caption         =   "&Help"
      Begin VB.Menu mnuIndex 
         Caption         =   "&Index"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form1"
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
'***************                TicServer                       **********************
'***************                tic.vbp                         **********************
'*************************************************************************************
'*************************************************************************************
Dim aryDeck() As Integer 'The deck of cards
Dim Online% 'amount of players online
Dim PlayerTic% 'player who Ticed out
Dim Players(5) As String 'amount of players
Dim Round% 'the current round
Dim CardOnTop% 'the card on the top of the deck
Dim NumberofDecks% 'how many decks are being used
Dim LastTurn As Boolean 'if its the last turn
Dim Refreshing As Boolean 'if the server is refreshing
Dim tPN(5) As String 'array for refreshing
Dim AmountToR As Integer 'array for determining if all players have connected

Private Sub Albl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Albl_MouseMove_Error
imgHCommand1.Visible = False 'make the highlighted button visible true
On Error GoTo 0
Exit Sub
Albl_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Albl_MouseMove in Form, Form1"
End Sub

Private Sub Command1_Click()
Dim a As Integer
On Error Resume Next
'*******************
'make sure there are users connected
Dim Counter% 'counter
Counter% = 0 'defaul to 0
'loop through the list box
For i = 0 To List1.ListCount - 1
    If List1.List(i) <> "<no player>" Then
        Counter% = Counter% + 1 'if there is somone there, then add 1
        Exit For 'exit the for since we got someone
    End If
Next
If Counter% = 0 Then 'if there is noone
    'tell the user
    MsgBox "There are no users connected to the server.", vbCritical, "Error"
    'exit the sub
    Exit Sub
End If
'********************
List3.Clear 'clear the listboxes
List2.Clear
For a = 0 To UBound(aryDeck) 'add the cards
    List2.AddItem aryDeck(a)
Next
Call Shuffle 'shuffle the deck
Round% = 3 'set to the first round
Call MakeTurn 'setup the turns
WaitFor 1 'wait for 300 Miliseconds
Call DealHand 'deal out the hands
WaitFor 1 'wait for 300 MS
Call FlipFirst 'flip over the first card
WaitFor 1 'wait for 300 MS
Call InformWhosTurn 'tell the players who's turn it is
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Command1_MouseMove_Error
imgHCommand1.Visible = True 'make the highlighted button visible
Command1.ForeColor = vbWhite
On Error GoTo 0
Exit Sub
Command1_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Command1_MouseMove in Form, Form1"
End Sub

Private Sub Form_Load()
Dim b As Integer, a As Integer
On Error GoTo Form_Load_Error
If App.PrevInstance = True Then Unload Me 'only 1 instance can be running
Me.Caption = "Tic Server [ver " & App.Major & "." & App.Minor & "]" 'set the caption
Randomize 'initiate the random number generator
NumberofDecks% = 2 'set the default number of decks
Round% = 3 'set the defaul round
lblIP.Caption = GetIPAddress 'set the IP address
For a = 0 To 5 'add users to the list box
    List1.AddItem "<no player>"
Next
ws(0).Close 'close the winsock
ws(0).Listen 'make it listen
For a = 1 To 6 'load winsocks
    Load ws(a)
    ws(a).LocalPort = 3220 'set port
    ws(a).Protocol = sckTCPProtocol 'set to TCP/IP
Next a
b = 0 'default
ReDim aryDeck(NumberofDecks% * 52) As Integer 'get amount of cards
For a = 0 To UBound(aryDeck) 'and add to the deck array
    Select Case b
        'this is used to make sure the cards are
            'numbered from 0-51
        Case 0 To 51
            aryDeck(a) = b
        Case 52 To 103
            aryDeck(a) = b - 52
        Case 104 To 155
            aryDeck(a) = b - 104
        Case 156 To 207
            aryDeck(a) = b - 156
        Case 208 To 259
            aryDeck(a) = b - 208
    End Select
    b = b + 1
Next
Dim ret As Long, rRect As RECT
'stuff to trim off border of the listbox
rRect.lTop = 1
rRect.lLeft = 1
rRect.lRight = List1.Width - 1
rRect.lBottom = List1.Height - 1
ret = CreateRectRgnIndirect(rRect)
'trim it off
SetWindowRgn List1.hWnd, ret, True
'get the picture behind list1
BitBlt Picture1.hdc, 0, 0, List1.Width + 1, List1.Height + 1, Me.hdc, List1.Left + 1, List1.Top + 1, vbSrcCopy
'make image1's picture the bit'd image
Image1.Picture = Picture1.Image
Me.ScaleMode = 1 'pixels
gBGBrush = CreatePatternBrush(Image1.Picture.Handle) 'set the pattern brush
Picture1.Visible = False
'begin subclassing
oldWindowProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf NewWindowProc)
oldLbx1Proc = SetWindowLong(List1.hWnd, GWL_WNDPROC, AddressOf NewLbxProc)
On Error GoTo 0
Exit Sub
Form_Load_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_Load in Form, Form1"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Form_MouseMove_Error
imgHCommand1.Visible = False 'make the highlight go away
Command1.ForeColor = vbBlack 'change back to black
On Error GoTo 0
Exit Sub
Form_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_MouseMove in Form, Form1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Form_Unload_Error
For i = 0 To ws.UBound 'close all winsocks
    ws(i).Close
Next
'unsubclass
SetWindowLong Me.hWnd, GWL_WNDPROC, oldWindowProc
SetWindowLong List1.hWnd, GWL_WNDPROC, oldLbx1Proc
'delete the pattern brush
DeleteObject gBGBrush
On Error GoTo 0
Exit Sub
Form_Unload_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_Unload in Form, Form1"
End Sub

Private Sub lblPointer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo lblPointer_MouseMove_Error
imgHCommand1.Visible = False 'make the highligh invisible
Command1.ForeColor = vbBlack 'change back to black
On Error GoTo 0
Exit Sub
lblPointer_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: lblPointer_MouseMove in Form, Form1"
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo List1_MouseMove_Error
imgHCommand1.Visible = False 'make the highligh invisible
Command1.ForeColor = vbBlack 'change back to black
On Error GoTo 0
Exit Sub
List1_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: List1_MouseMove in Form, Form1"
End Sub

Private Sub mnuAmountDeck_Click(Index As Integer)
Dim i As Integer
On Error GoTo mnuAmountDeck_Click_Error
For i = 0 To mnuAmountDeck.UBound 'checking/unchecking of the menu system
    If i <> Index Then
        mnuAmountDeck(i).Checked = False
    Else
        mnuAmountDeck(i).Checked = True
    End If
Next
'set the number of decks to their choice
NumberofDecks% = Index + 2
Dim b As Integer
Dim a As Integer
b = 0
'refresh the amount of cards
ReDim aryDeck(NumberofDecks% * 52) As Integer
'and readd em all
For a = 0 To UBound(aryDeck)
    Select Case b
        Case 0 To 51
            aryDeck(a) = b
        Case 52 To 103
            aryDeck(a) = b - 52
        Case 104 To 155
            aryDeck(a) = b - 104
        Case 156 To 207
            aryDeck(a) = b - 156
        Case 208 To 259
            aryDeck(a) = b - 208
    End Select
    b = b + 1
Next
On Error GoTo 0
Exit Sub
mnuAmountDeck_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: mnuAmountDeck_Click in Form, Form1"
End Sub

Private Sub mnuExit_Click()
On Error GoTo mnuExit_Click_Error
Unload Me 'unload the form
On Error GoTo 0
Exit Sub
mnuExit_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: mnuExit_Click in Form, Form1"
End Sub

Private Sub mnuIndex_Click()
On Error GoTo mnuIndex_Click_Error
Load Form2 'load the help form
Form2.Show
On Error GoTo 0
Exit Sub
mnuIndex_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: mnuIndex_Click in Form, Form1"
End Sub

Private Sub mnuRefresh_Click()
'get conformation
On Error GoTo mnuRefresh_Click_Error
If MsgBox( _
    "Refreshing should only be used if a user lost the connection, " & vbCrLf & _
    "and you want to resume the game.  This will redistribute the points according" _
    & vbCrLf & _
    "to the players, and will remove player's points of whos no longer connected." _
    , vbOKCancel + vbQuestion, "Refresh") = vbOK Then
        Dim i As Integer, b As Integer 'counters
        AmountToR = 0 'set to 0
        Erase tPN 'clear out the array
        For i = 0 To UBound(tPN) 'get all the players names and points
            If List1.List(i) <> "<no player>" Then
                tPN(i) = List1.List(i) & "[p]" & lblPoints(i).Caption
                AmountToR = AmountToR + 1 'add 1 to it
            End If
        Next
        b = 0 'set to 0
        For i = 0 To UBound(tPN)
        'get rid of any empty gaps in the array
            If tPN(i) <> "" Then
                If i <> b Then
                    tPN(b) = tPN(i)
                    tPN(i) = ""
                End If
                b = b + 1
            End If
        Next
        'unload all the point and turn labels
        For i = 1 To lblTurn.UBound
            Unload lblTurn(i)
            Unload lblPoints(i)
        Next
        'set the captions of the turn label and point label to nothing
        lblTurn(0).Caption = ""
        lblPoints(0).Caption = ""
        lblPointer.Top = lblTurn(0).Top 'put the pointer at the top
        For b = 0 To List1.ListCount - 1 'send out the refresh code
            If List1.List(b) <> "<no player>" Then
                ws(b + 1).SendData "rÉFRË§µ¸" 'send it"
                List1.List(b) = "<no player>" 'clear them out of the listbox
                DoEvents
            End If
        Next
        'close all the connections
        For i = 1 To ws.UBound
            ws(i).Close
        Next
        'set this to true
        Refreshing = True
End If
On Error GoTo 0
Exit Sub
mnuRefresh_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: mnuRefresh_Click in Form, Form1"

End Sub

Private Sub mnuStartGame_Click()
On Error GoTo mnuStartGame_Click_Error
Call Command1_Click 'call command1
On Error GoTo 0
Exit Sub
mnuStartGame_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: mnuStartGame_Click in Form, Form1"
End Sub

Private Sub ws_Close(Index As Integer)
On Error GoTo ws_Close_Error
List1.List(Index - 1) = "<no player>" 'get rid of any instance of the player
ws(Index).Close 'close the winsock connection
Unload ws(Index)
Load ws(Index)
ws(Index).LocalPort = 3220
ws(Index).Protocol = sckTCPProtocol
On Error GoTo 0
Exit Sub
ws_Close_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: ws_Close in Form, Form1"
End Sub

Private Sub ws_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo ws_ConnectionRequest_Error
For a = 0 To List1.ListCount - 1
    If List1.List(a) = "<no player>" Then
        ws(a + 1).Accept requestID 'accept the request to connect
        List1.List(a) = "<player waiting>" 'show the player is waiting
        Dim f$
        f$ = "÷êÅ" 'request the name
        ws(a + 1).SendData f$
        Exit For 'exit the loop
    End If
Next
On Error GoTo 0
Exit Sub
ws_ConnectionRequest_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: ws_ConnectionRequest in Form, Form1"
End Sub

Private Sub ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData$
On Error GoTo ws_DataArrival_Error
ws(Index).GetData strData$, vbString 'get the data recieved
'check all these subs
Check1Player strData$, Index
SendTICMessage strData$, Index
AddPoints strData$, Index
UserDraw strData$, Index
InformUserDraw strData$
DiscardACard strData$, Index
PickTopCard strData$, Index
MakeNextTurn strData$, Index
AssignName strData$, Index
CheckChat strData$, Index
CheckCheat strData$, Index
On Error GoTo 0
Exit Sub
ws_DataArrival_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: ws_DataArrival in Form, Form1"
End Sub

Sub MakeTurn()
On Error Resume Next
Dim a As Integer, Counter% 'counters
For a = 0 To 5 'see how many players there are
    If List1.List(a) <> "<no player>" Then
        Counter% = Counter% + 1
    End If
Next
'unload all the labels
For a = 1 To 5
    Unload lblTurn(a)
    Unload lblPoints(a)
Next
Dim booFirst As Boolean 'true/false
booFirst = True
Dim X% 'position counter
X% = lblTurn(0).Top + lblTurn(0).Height
For a = 0 To Counter% - 1
    If booFirst Then 'if the first label
        lblTurn(0).Caption = Players(a) 'set the name
        lblPoints(0).Caption = 0 'set the points to 0
        booFirst = False 'set to false
        GoTo gNext 'Continue the loop
    Else
        Dim i As Integer 'counter
        i = lblTurn.UBound + 1 'get the new index
        Load lblTurn(i) 'load it
        lblTurn(i).Top = X% 'set the top
        lblTurn(i).Left = lblTurn(0).Left 'set the left
        lblTurn(i).Visible = True 'make it visible
        lblTurn(i).Caption = Players(i) 'set the name
        Load lblPoints(i) 'get the point label
        lblPoints(i).Top = X% 'position it
        lblPoints(i).Left = lblPoints(0).Left
        lblPoints(i).Caption = 0 'set it to 0
        lblPoints(i).Visible = True 'make it visible
    End If
    X% = X% + lblTurn(i).Height 'get a new top
gNext:
Next
lblPointer.Top = lblTurn(0).Top 'set the pointer to the first person
End Sub

Sub FlipFirst()
Dim b%, f$, a As Integer 'value holders/counters
On Error GoTo FlipFirst_Error
b% = Int(((List2.ListCount - 1) - 0 + 0) * Rnd + 0) 'get a random number
CardOnTop% = List2.List(b%) 'set the card on top
f$ = "ÑÐª" & List2.List(b%) 'hold the string text
List3.AddItem List2.List(b%) 'add the card
List2.RemoveItem b% 'remove the card
For a = 0 To List1.ListCount - 1
    If List1.List(a) <> "<no player>" Then
        ws(a + 1).SendData f$ 'send the data
        DoEvents
    End If
Next
On Error GoTo 0
Exit Sub
FlipFirst_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: FlipFirst in Form, Form1"
End Sub

Sub Shuffle()
Dim X%, Temp() As Integer 'temp value/temp array
On Error GoTo Shuffle_Error
ReDim Temp(List2.ListCount - 1) 'get the right diminsions
For a = 0 To 2 'shuffle 3 times
    For i = 0 To List2.ListCount - 1 'randomly pull out the cards
        X% = Int(((List2.ListCount - 1) - 0 + 0) * Rnd + 0)
        Temp(i) = Val(List2.List(X%))
        List2.RemoveItem X% 'remove the card
    Next
    List2.Clear 'clear list2
    For i = 0 To UBound(Temp) 're-add the items to the list box
        List2.AddItem Temp(i)
    Next
Next
On Error GoTo 0
Exit Sub
Shuffle_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Shuffle in Form, Form1"
End Sub

Sub DealHand()
Dim SentR As Boolean, f$, a As Integer 'counters/tempvalues
Dim c As Integer 'counter
On Error GoTo DealHand_Error
SentR = True
For a = 0 To List1.ListCount - 1
    If List1.List(a) <> "<no player>" Then 'loop through list 1
        If SentR = True Then 'get the round recorded
            f$ = "§©¨" & Round%
            SentR = False 'set to false
        End If
        For c = 1 To Round% 'and deal out the hand
            b% = Int(((List2.ListCount - 1) - 0 + 0) * Rnd + 0)
            f$ = f$ & "¤¥£" & List2.List(b%)
            List3.AddItem List2.List(b%) 'add to used cards
            List2.RemoveItem b% 'remove from deck
        Next
        f$ = f$ & "¤¥£" 'build the variable
        ws(a + 1).SendData f$ 'send it to the user
        DoEvents
    End If
    SentR = True 'set back to true
    f$ = "" 'erase f
Next
On Error GoTo 0
Exit Sub
DealHand_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: DealHand in Form, Form1"
End Sub

Sub AddPoints(strData$, Index%)
On Error GoTo eh1
Dim a As Integer
If Left(strData$, 3) = "¶¸·" And LastTurn = True Then  'if the code, and its the lastturn of the round
    Dim X%, p%, b%, T%, Discarded% 'counters/value holders
    strData$ = Right(strData$, Len(strData$) - 3) 'trim off the code
    p% = InStr(1, strData$, "||") 'find "||"
    Dim tPnAme As String 'hold the temp name
    Discarded% = Val(Right(strData$, (Len(strData$) - InStr(strData$, "**") - 1))) 'get the card the descard
            'which is located after the "**" in the string
    'trim off the ** and everything after it
    strData$ = Left(strData$, InStr(strData$, "**") - 1)
    tPnAme = Left(strData$, p% - 1) 'get the name
        'which is located at the begining of the string
    T% = Mid$(strData$, p% + 2, Len(strData$) - (p% + 1)) 'get the points taken
            'which is located after the ||
    For a = 0 To lblTurn.UBound 'begin a loop to find the palyer
        If lblTurn(a).Caption = tPnAme Then
            If Index% - 1 <> PlayerTic% Then 'if it isnt the play who tic'd
                lblPoints(a).Caption = Val(lblPoints(a).Caption) + T% 'add the points
            End If
        End If
        DoEvents
    Next
    CardOnTop% = Discarded% 'set the card on top
    Dim c$ 'set the message
    'send the message plus the new top card
    c$ = "¼½¾" & Players(Index% - 1) & " discards a card from their hand.ÑÐª" & CardOnTop%
    For b% = 0 To List1.ListCount - 1
        If List1.List(b%) <> "<no player>" Then
            ws(b% + 1).SendData c$ 'send it
            DoEvents
        End If
    Next
    If MakeNextTurn("ðÐ±" & tPnAme, Index%, True) = lblTurn(PlayerTic%).Top Then 'if the next turn is
                'the player who tic'd out
        If LastTurn = True Then 'if its the last turn
            If Round% + 1 > 15 Then 'if the round becomes greater then
                        '15, the game is over
                Call InformWinner 'call the inform winner sub
                Exit Sub
            Else
                Round% = Round% + 1 'if not, make the round 1 greater
            End If
            'make the next turn
            MakeNextTurn "ðÐ±" & Players(PlayerTic%), PlayerTic% + 1, False
            WaitFor 1 'wait for 300 MS
            For a = 0 To List1.ListCount - 1
                If List1.List(a) <> "<no player>" Then 'send out the data
                        'about the round
                    ws(a + 1).SendData "¼½¾" & "Round " & Round% & " has now begun."
                    DoEvents
                End If
            Next
            Call DealRound 'deal the round
            Exit Sub
        End If
    Else
        'if not the last turn of thr round, make the next turn
        Call MakeNextTurn("ðÐ±" & tPnAme, Index%, False)
    End If
End If
Exit Sub
eh1:
Exit Sub
End Sub

Sub SendTICMessage(strData$, Index%)
Dim a As Integer 'counter
On Error GoTo SendTICMessage_Error
If Left(strData$, 3) = "»ïñ" Then 'if there is the tic message
    strData$ = Right(strData$, Len(strData$) - 3) 'trim it off
    Dim c$ 'value holder
    c$ = "¼½¾" & strData$ & "Î¶¬has got a tic!" 'send out the data
            'that someone has a tic
    For a = 0 To List1.ListCount - 1
        If List1.List(a) <> "<no player>" Then
            ws(a + 1).SendData c$ 'send it
            DoEvents
        End If
    Next
    PlayerTic% = Index% - 1 'set the player tic
    lblPoints(PlayerTic%) = Val(lblPoints(PlayerTic%).Caption) - 5 'give them -5 points
    LastTurn = True 'set last turn to true
    Online% = 0 'set online to 0
    For a = 0 To List1.ListCount - 1 'loop through the list
                'box to see how many are on line
        If List1.List(a) <> "<no player>" Then
            Online% = Online% + 1
        End If
    Next
    Debug.Print Online%
    If Online% = 1 Then 'if there is only 1 player
        If Round% + 1 > 15 Then 'if greater then 15
            Call InformWinner 'game is over
            Exit Sub
        Else
            Round% = Round% + 1 'if not, add 1 to the round
        End If
        For a = 0 To List1.ListCount - 1
            If List1.List(a) <> "<no player>" Then
                'send the data out to the 1 player
                ws(a + 1).SendData "¼½¾" & "Round " & Round% & " has now begun."
                Exit For
            End If
        Next
        Call DealRound 'deal the round
    End If
End If
On Error GoTo 0
Exit Sub
SendTICMessage_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: SendTICMessage in Form, Form1"
End Sub

Function MakeNextTurn(strData$, Index%, Optional TestCheck As Boolean) As Long
On Error GoTo eh1
Dim a As Integer
If Left(strData$, 3) = "ðÐ±" Then 'if the special code
    strData$ = Right(strData$, Len(strData$) - 3) 'trim it off
    Dim g% 'counter
    For a = 0 To List1.ListCount - 1 'find out whos turn it should be
        If List1.List(a) = strData$ Then
            g% = a
            If List1.List(g% + 1) = "<no player>" Then 'if there
                    'isnt a next in line, goto the first player
                g% = 0
                Exit For
            ElseIf List1.List(g% + 1) <> "<no player>" Then
                    'if there is, goto thenext player
                g% = g% + 1
                Exit For
            Else
                'if some error somehow occurs, set to 0
                g% = 0
            End If
            Exit For
        End If
    Next
    For a = 0 To lblTurn.UBound 'get names, and send out messages
        If lblTurn(a).Caption = List1.List(g%) Then 'if its the name
            If TestCheck = False Then 'if our option value is false
                lblPointer.Top = lblTurn(a).Top 'change the turn
                Dim c$ 'send out the data
                c$ = "µôÐ" & Players(Index% - 1) & " has ended their turn, and now " & lblTurn(a).Caption & "'s turn has begun."
                For b% = 0 To List1.ListCount - 1
                    If List1.List(b%) <> "<no player>" Then
                        ws(b% + 1).SendData c$ 'send it
                        DoEvents
                    End If
                Next
            Else
                'if out optional value is true then
                    'set this function to where the
                    'pointer would be
                MakeNextTurn = lblTurn(a).Top
                TestCheck = False
            End If
            Exit For
        End If
    Next
End If
Exit Function
eh1:
Exit Function
End Function

Sub DealRound()
On Error Resume Next
LastTurn = False 'set to false
List3.Clear 'clear the deck boxes
List2.Clear
'add a fresh batch of decks to the listbox
For b% = 0 To UBound(aryDeck)
    List2.AddItem aryDeck(b%)
Next
'shuffle them
Call Shuffle
'deal the hand
WaitFor 1 'waitfor 300 MS
Call DealHand
'wait for 300 MS
WaitFor 1
'flipover the first card
Call FlipFirst
'set this to some high number as to not confuse the program
PlayerTic% = 10
End Sub

Sub PickTopCard(strData$, Index%)
On Error GoTo PickTopCard_Error
If Left(strData$, 3) = "×ÿ¡" Then 'if the special code
    strData$ = Right(strData$, Len(strData$) - 3) 'trim it off
    Dim c$, b% 'counters/value holders
    'set the data
    c$ = "¼½¾" & Players(Index% - 1) & "ªÿ½picked up the top discarded car."
    For b% = 0 To List1.ListCount - 1
        If List1.List(b%) <> "<no player>" Then
            ws(b% + 1).SendData c$ 'send the data
            DoEvents
        End If
    Next
    'send the card they got to the user
    ws(Index%).SendData "×ÿ¡" & CardOnTop%
End If
On Error GoTo 0
Exit Sub
PickTopCard_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: PickTopCard in Form, Form1"
End Sub

Sub DiscardACard(strData$, Index%)
On Error GoTo DiscardACard_Error
If Left(strData$, 3) = "öõô" Then 'if the code
    strData$ = Right(strData$, Len(strData$) - 3) 'trim it off
    CardOnTop% = Val(strData$) 'set the new card on top
    Dim c$, b% 'counters/value holder
    'set the value of c to the message, and the new top card
    c$ = "¼½¾" & Players(Index% - 1) & " discards a card from their hand.ÑÐª" & CardOnTop%
    For b% = 0 To List1.ListCount - 1
        If List1.List(b%) <> "<no player>" Then
            ws(b% + 1).SendData c$ 'send the data
            DoEvents
        End If
    Next
    'make the next turn
    Call MakeNextTurn("ðÐ±" & Players(Index% - 1), Index%)
End If
On Error GoTo 0
Exit Sub
DiscardACard_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: DiscardACard in Form, Form1"
End Sub

Sub InformUserDraw(strData$)
On Error GoTo InformUserDraw_Error
If Left(strData$, 3) = "©ÈÇ" Then 'if the code
    strData$ = Right(strData$, Len(strData$) - 3) 'trim it off
    Dim c$, b% 'counter/value holder
    c$ = "¼½¾" & strData$ 'set the message
    For b% = 0 To List1.ListCount - 1
        If List1.List(b%) <> "<no player>" Then
            ws(b% + 1).SendData c$ 'send the data
            DoEvents
        End If
    Next
End If
On Error GoTo 0
Exit Sub
InformUserDraw_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: InformUserDraw in Form, Form1"
End Sub

Sub AssignName(strData$, Index%)
On Error GoTo AssignName_Error
Dim i As Integer
Dim a As Integer
If Left(strData$, 3) = "÷êÅ" Then 'if the code
    strData$ = Right(strData$, Len(strData$) - 3) 'trim it off
    Players(Index% - 1) = strData$ 'save the name
    List1.List(Index% - 1) = Players(Index% - 1) 'and change the name in the listbox
    If Refreshing = True And AmountToR > 0 Then 'if there are players to refresh, and it should be refreshing
        If lblTurn(0).Caption <> "" Then 'if the first spot is already filled
            i = lblTurn.UBound + 1 'get the new ubound
            Load lblTurn(i) 'load it
            'set the new top
            lblTurn(i).Top = lblTurn(lblTurn.UBound - 1).Top + lblTurn(i).Width
            'set the left
            lblTurn(i).Left = lblTurn(0).Left
            For a = 0 To UBound(tPN)
                'if the name was on before, then
                If LCase(Left(tPN(a), InStr(tPN(a), "[") - 1)) = LCase(strData$) Then
                    'set the caption to the name
                    lblTurn(i).Caption = Left(tPN(a), InStr(tPN(a), "[") - 1)
                    Load lblPoints(i) 'load a point label
                    lblPoints(i).Top = lblTurn(i).Top 'set the top
                    lblPoints(i).Left = lblPoints(0).Left 'set the left
                    'get the points the player had
                    lblPoints(i).Caption = Right(tPN(a), Len(tPN(a)) - InStr(tPN(a), "]"))
                    lblPoints(i).Visible = True 'make it visible
                    lblTurn(i).Visible = True 'make it visible
                    AmountToR = AmountToR - 1 'subtract 1 from the amount of users reloging
                    Exit For
                End If
            Next
        Else
            'if this is the first person
            For a = 0 To UBound(tPN)
                'if that name was there before
                If LCase(Left(tPN(a), InStr(tPN(a), "[") - 1)) = LCase(strData$) Then
                    'set teh captions
                    lblTurn(0).Caption = Left(tPN(a), InStr(tPN(a), "[") - 1)
                    'set the caption of the points
                    lblPoints(0).Caption = Right(tPN(a), Len(tPN(a)) - InStr(tPN(a), "]"))
                    lblPoints(0).Visible = True 'make them visible
                    lblTurn(0).Visible = True
                    AmountToR = AmountToR - 1 'subtract 1 from the total
                    Exit For
                End If
            Next
        End If
    End If
    If AmountToR = 0 And Refreshing = True Then 'if everyone is back
        Dim b%, c$ 'counter/value holder
        Refreshing = False 'set this to false
        'set the message
        c$ = "¼½¾" & "Server has finished refreshing." & vbCrLf & "Game will continue with " & lblTurn(0).Caption & "'s turn."
        For b% = 0 To List1.ListCount - 1
            If List1.List(b%) <> "<no player>" Then
                ws(b% + 1).SendData c$ 'send the data
                DoEvents
            End If
        Next
        WaitFor 1 'wait for 300 MS
        'and set the turn to the 1st player
        Call MakeNextTurn("ðÐ±" & lblTurn(lblTurn.UBound).Caption, lblTurn.UBound + 1, False)
    End If
End If
On Error GoTo 0
Exit Sub
AssignName_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: AssignName in Form, Form1"
End Sub

Sub CheckChat(strData$, Index%)
Dim a As Integer
On Error GoTo CheckChat_Error
If Left(strData$, 3) = "¼½¾" Then 'if the code
    strData$ = Right(strData$, Len(strData$) - 3) 'trim it off
    Dim c$ 'value holder
    'make it say Player says: what they say
    c$ = "¼½¾" & Players(Index% - 1) & " says: " & strData$
    For a = 0 To List1.ListCount - 1
        If List1.List(a) <> "<no player>" Then
            ws(a + 1).SendData c$ 'send the data
            DoEvents
        End If
    Next
End If
On Error GoTo 0
Exit Sub
CheckChat_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: CheckChat in Form, Form1"
End Sub

Sub UserDraw(strData$, Index%)
On Error GoTo UserDraw_Error
If Left(strData$, 3) = "¶®Æ" Then 'if the code
    Dim b%, f$, a As Integer 'counter/value holders
    b% = Int(((List2.ListCount - 1) - 0 + 0) * Rnd + 0) 'get a random number
    f$ = "¶®Æ" & List2.List(b%) 'set it to a string
    List3.AddItem List2.List(b%) 'add it to the listbox
    List2.RemoveItem b% 'remove it
    ws(Index%).SendData f$ 'send it
End If
On Error GoTo 0
Exit Sub
UserDraw_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: UserDraw in Form, Form1"
End Sub

Sub InformWhosTurn()
Dim a As Integer, b As Integer
On Error Resume Next
For a = 0 To lblTurn.UBound
    If lblPointer.Top = lblTurn(a).Top Then 'if its there turn
        For b = 0 To List1.ListCount - 1
            If List1.List(b) <> "<no player>" Then
                ws(b + 1).SendData "µôÐ" & lblTurn(a).Caption 'send the data to players
                DoEvents
            End If
        Next
    End If
Next
End Sub

Sub InformWinner()
Dim i As Integer, iHigh%, iIndex% 'coutners/value holders
On Error GoTo InformWinner_Error
For i = 0 To lblPoints.UBound 'loop through the points
    If i = 0 Then iHigh% = Val(lblPoints(i).Caption): iIndex% = i + 1 'set the initial value
    If Val(lblPoints(i).Caption) < iHigh% Then 'if the new point section
                'is lower then the current value, then make it the new
                'current value
        iHigh% = Val(lblPoints(i).Caption)
        iIndex% = i + 1 'increase the index
    End If
Next
Dim c$, b As Integer 'coutner/value holder
'set the value of the string
c$ = lblTurn(iIndex% - 1).Caption & "×©ºhas won the game with " & iHigh% & " points."
For b = 0 To 5
    If List1.List(b) <> "<no player>" Then
        ws(b + 1).SendData c$ 'send it
        DoEvents
    End If
Next
On Error GoTo 0
Exit Sub
InformWinner_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: InformWinner in Form, Form1"
End Sub

Sub CheckCheat(strData$, Index%)
On Error GoTo CheckCheat_Error
If Left(strData$, 5) = "§ÞÏkë" Then
    Dim X%, i As Integer, a As Boolean 'coutners/value holders
    X% = Val(Right(strData$, Len(strData$) - 5)) 'get the card they want
    a = False 'set it to false
    For i = 0 To List2.ListCount - 1 'begin a loop
        If Val(List2.List(i)) = X% Then 'if it find the wanted card
            a = True 'make it true
            List2.RemoveItem i 'remove it the card
            Exit For
        End If
    Next
    If a = True Then
        ws(Index).SendData "§ÞÏkë" & X% 'send it if found
    Else
        ws(Index).SendData "§ÞÏkë" & "-1" 'if not, send a falure sign
    End If
End If
On Error GoTo 0
Exit Sub
CheckCheat_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: CheckCheat in Form, Form1"
End Sub

Sub Check1Player(strData$, Index%)
On Error GoTo Check1Player_Error
If Left(strData$, 7) = "þ|âÿË®¹" Then 'if the 1 player symbols
    Call Command1_Click 'start the game
End If
On Error GoTo 0
Exit Sub
Check1Player_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Check1Player in Form, Form1"
End Sub
