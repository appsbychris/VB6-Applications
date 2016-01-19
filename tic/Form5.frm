VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rules of the game-"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9270
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   Picture         =   "Form5.frx":08CA
   ScaleHeight     =   5820
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.Label cmdPage 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   1
      Left            =   6120
      TabIndex        =   2
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label cmdPage 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   7680
      TabIndex        =   1
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image imgHLeft 
      Height          =   525
      Left            =   6120
      Picture         =   "Form5.frx":4F98
      Top             =   5160
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Image3 
      Height          =   525
      Left            =   6120
      Picture         =   "Form5.frx":562E
      Top             =   5160
      Width           =   1500
   End
   Begin VB.Image imgHRight 
      Height          =   525
      Left            =   7680
      Picture         =   "Form5.frx":7F74
      Top             =   5160
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   7680
      Picture         =   "Form5.frx":8667
      Top             =   5160
      Width           =   1500
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H000000FF&
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "Form5"
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
Dim Pages(3) As String 'hold the help file info
Dim PageCount% 'current page

Private Sub cmdPage_Click(Index As Integer)
On Error GoTo cmdPage_Click_Error
Select Case Index 'select the index
    Case 0: 'if its 0, then
        If PageCount% < 3 Then 'if the page is < 3
            PageCount% = PageCount% + 1 'increase the page
            Label1.Caption = Pages(PageCount%) 'show the page
        End If
    Case 1: 'if its 1, then
        If PageCount% > 0 Then 'if the count is > 0 then
            PageCount% = PageCount% - 1 'decrease the page #
            Label1.Caption = Pages(PageCount%) 'show the page
        End If
End Select
On Error GoTo 0
Exit Sub
cmdPage_Click_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdPage_Click in Form, Form5"
End Sub

Sub UnHall()
On Error GoTo UnHall_Error
imgHLeft.Visible = False 'unhighlight the buttons
imgHRight.Visible = False
On Error GoTo 0
Exit Sub
UnHall_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: UnHall in Form, Form5"
End Sub

Private Sub cmdPage_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo cmdPage_MouseMove_Error
Call UnHall 'unhighlight the buttons
'highlight the appropiate one
If Index = 1 Then imgHLeft.Visible = True
If Index = 0 Then imgHRight.Visible = True
On Error GoTo 0
Exit Sub
cmdPage_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: cmdPage_MouseMove in Form, Form5"
End Sub

Private Sub Form_Load()
'the rules of the game...
Pages(0) = "Tic - Rules of the game. (Page 1)" & vbCrLf & _
    "Objective:" & vbCrLf & _
    "    ð Make sets of three of a kind or better and runs of 3 or better with your hand." & vbCrLf & _
    "    ð Use all the cards in ways listed above to use your entire hand and have a discard." & vbCrLf & _
    "An example would be:" & vbCrLf & _
    "Say your hand consist of:" & vbCrLf & _
    "    ð Ace of Spades" & vbCrLf & _
    "    ð King of Spades" & vbCrLf & _
    "    ð Queen of Spades" & vbCrLf & _
    "With this hand, you would be on round 3, which is the begining round." & vbCrLf & _
    "Now, on your turn, you would either draw the top card of the deck, or draw the top card of the discard pile." & vbCrLf & _
    "Say you draw a Ace of Hearts.  You now have 4 cards in your hand, and 3 of them go together." & vbCrLf & _
    "This hand is considered a Tic.  You would now play your three cards that go together, and then place the unwanted card (Ace of Hearts) in the discard pile"
Pages(1) = "Tic - Rules of the game. (Page 2)" & vbCrLf & _
    "Turn sequence:" & vbCrLf & _
    "    ð Draw a card from either the deck, or the top card of the discard pile." & vbCrLf & _
    "    ð Tic out if it is possible" & vbCrLf & _
    "    ð Discard 1 card from your hand." & vbCrLf & _
    "    ð (Remember, discarding ends your turn, so once you discard, your turn is over." & vbCrLf & _
    "More information of ways to make sets of cards:" & vbCrLf & _
    "    ð Three of a Kind or Better-" & vbCrLf & _
    "        ð This may consist of 3 or more of the same card, of any suit." & vbCrLf & _
    "        ð You may use wild cards (See Page 3) in with these." & vbCrLf & _
    "    ð Run of Three or Better-" & vbCrLf & _
    "        ð Run is defined as a sequence of cards." & vbCrLf & _
    "        ð All cards included in a run must be of the same suit." & vbCrLf & _
    "        ð Wild Cards (See Page 3) may be used to fill in missing cards." & vbCrLf & _
    "        ð Ace can be either high or low, but not both at the same time." & vbCrLf & _
    "            ð EX: You cannot use a King of Hearts, Ace of Hearts, Two of Hearts in the same run." & vbCrLf & _
    "        ð Runs may be of any length, but the minimum is 3."
Pages(2) = "Tic - Rules of the game. (Page 3)" & vbCrLf & _
    "Wild Cards:" & vbCrLf & _
    "Wild cards change from round to round.  A wild card is equal to the current round." & vbCrLf & _
    "    ð During the first round (which is 3s), all 3 cards are wild.  They may be used as any card." & vbCrLf & _
    "How rounds work:" & vbCrLf & _
    "    ð The game starts on round 3." & vbCrLf & _
    "    ð The players hand size is the current round number" & vbCrLf & _
    "    ð For rounds Jack, Queen, King, Ace, the hand size is 11, 12, 13, 14, respectivly." & vbCrLf & _
    "    ð On the last round, Twos, 15 cards are dealt to the players." & vbCrLf & _
    "Points:" & vbCrLf & _
    "    ð Cards of value Two through Nine are equal to 5 points." & vbCrLf & _
    "    ð Cards of value Ten through King are equal to 10 points." & vbCrLf & _
    "    ð Ace cards are worth 15 points." & vbCrLf & _
    "    ð If you cannot use a wild, and still have it in your hand, it is worth 30 points." & vbCrLf & _
    "For more on points and what they are for, see Page 4."
Pages(3) = "Tic - Rules of the game. (Page 4)" & vbCrLf & _
    "Points...What are they for?" & vbCrLf & _
    "    ð Points are used to determine the winner of the game." & vbCrLf & _
    "    ð The player with the LEAST amount of points after round 15 is the winner." & vbCrLf & _
    "    ð When a player goes out, you must play as much of your hand as possible, and any un-used cards count         against you." & vbCrLf & _
    "How do I lower my points?" & vbCrLf & _
    "    ð There is only 1 way of reducing your points." & vbCrLf & _
    "    ð When you Tic out, you get -5 points added to your score." & vbCrLf & _
    "Rounds Queen through Two:" & vbCrLf & _
    "    ð Round Queen:  All points gained are times 2" & vbCrLf & _
    "    ð Round King: All points gained are times 3" & vbCrLf & _
    "    ð Round Ace: All points gained are times 4" & vbCrLf & _
    "    ð Round Two: All points gained are times 5" & vbCrLf & _
    "    ð (Note) You do not get -5 points times the round when you go out."
Label1.Caption = Pages(0) 'show page 0
PageCount% = 0 'set the pagecount to 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Form_MouseMove_Error
Call UnHall 'unhighlight the buttons
On Error GoTo 0
Exit Sub
Form_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_MouseMove in Form, Form5"
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Label1_MouseMove_Error
Call UnHall 'unhighlight the buttons
On Error GoTo 0
Exit Sub
Label1_MouseMove_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Label1_MouseMove in Form, Form5"
End Sub
