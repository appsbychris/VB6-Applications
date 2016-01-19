VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mix Names"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   Icon            =   "Form3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":08CA
   ScaleHeight     =   5445
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List3 
      Height          =   840
      Left            =   280
      TabIndex        =   20
      Top             =   3960
      Width           =   2895
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   3280
      TabIndex        =   19
      Top             =   3960
      Width           =   3135
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear All Names"
      Height          =   255
      Left            =   3640
      TabIndex        =   18
      Top             =   5040
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4720
      TabIndex        =   17
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2485
      TabIndex        =   16
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3880
      TabIndex        =   15
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00008000&
      Caption         =   "Check3"
      Height          =   195
      Left            =   280
      TabIndex        =   14
      Top             =   3260
      Width           =   200
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00008000&
      Height          =   195
      Left            =   1840
      TabIndex        =   13
      Top             =   3000
      Width           =   200
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00008000&
      Height          =   195
      Left            =   280
      TabIndex        =   12
      Top             =   3020
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cl&ear"
      Height          =   255
      Left            =   4840
      TabIndex        =   11
      Top             =   2410
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Copy"
      Height          =   255
      Left            =   4840
      TabIndex        =   10
      Top             =   2080
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4120
      TabIndex        =   8
      Top             =   2160
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   2920
      TabIndex        =   7
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Save these Names"
      Height          =   255
      Left            =   160
      TabIndex        =   6
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Names 
      Height          =   285
      Index           =   4
      Left            =   160
      TabIndex        =   5
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Names 
      Height          =   285
      Index           =   3
      Left            =   160
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox Names 
      Height          =   285
      Index           =   2
      Left            =   160
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Names 
      Height          =   285
      Index           =   1
      Left            =   160
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Names 
      Height          =   285
      Index           =   0
      Left            =   160
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3640
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Mix to length:         "
      Height          =   255
      Left            =   2920
      TabIndex        =   9
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Check1_Click()
'Enables/Disables the Choose length options
On Error Resume Next
If Check1.Value = 1 Then
    Text2.Enabled = True
    Text2.BackColor = vbWhite
    Text3.Enabled = True
    Text3.BackColor = vbWhite
End If
If Check1.Value = 0 Then
    Text2.Enabled = False
    Text2.BackColor = &H80000004
    Text3.Enabled = False
    Text3.BackColor = &H80000004
End If
End Sub

Private Sub Command1_Click()
'Sub for combineing the names
On Error Resume Next
MousePointer = vbHourglass
'set the defaults to 1
If Text2.Text = "" Then Text2.Text = "1"
If Text3.Text = "" Then Text3.Text = "1"
Dim X%, Y%, z%, a$, p%, q%, r%, s%
If Check1.Value = 0 Then
    'if they want a random length, make a random number
    q% = Len(Names(0))
    For i = 1 To Val(Text1.Text) - 1
        If q% > Len(Names(i)) Then
            q% = Len(Names(i))
        End If
    Next
    r% = Len(Names(0))
    For i = 1 To Val(Text1.Text) - 1
        If r% < Len(Names(i)) Then
            r% = Len(Names(i))
        End If
    Next
    s% = Int(((r% + q%) - q% + q%) * Rnd + q%)
    p% = s%
Else
    'If they want a specific number, use it
    p% = Val(Text2.Text)
End If
'A loop of when the length variable a$
'is less then the desired length, keep adding
'onto it.
While Len(a$) < p%
    'Call the MakeName Function
    a$ = a$ & MakeName(a$)
Wend
'Makes sure the name will be the desired length
If Len(a$) > p% Then
    X% = Int(((Len(a$) - p%) - 1 + 1) * Rnd + 1)
    a$ = Mid(a$, X%, p%)
End If
a$ = LCase(a$)
Dim d$
'Capitalize the first letter
d$ = Left(a$, 1)
d$ = UCase(d$)
a$ = Mid(a$, 2)
a$ = d$ & a$
'Last name creation
If Check2.Value = 1 Then
    If Check1.Value = 0 Then
        'random length
        q% = Len(Names(0))
        For i = 1 To Val(Text1.Text) - 1
            If q% > Len(Names(i)) Then
                q% = Len(Names(i))
            End If
        Next
        r% = Len(Names(0))
        For i = 1 To Val(Text1.Text) - 1
            If r% < Len(Names(i)) Then
                r% = Len(Names(i))
            End If
        Next
        s% = Int(((r% + q%) - q% + q%) * Rnd + q%)
        p% = s%
    Else
        'Fixed length
        p% = Val(Text3.Text)
    End If
    Dim b$
    'Make the name
    While Len(b$) < p%
        b$ = b$ & MakeName(b$)
    Wend
    'Trim it to a specific size
    If Len(b$) > p% Then
        X% = Int(((Len(b$) - p%) - 1 + 1) * Rnd + 1)
        b$ = Mid(b$, X%, p%)
    End If
    b$ = LCase(b$)
    Dim e$
    'Cap the first letter
    e$ = Left(b$, 1)
    e$ = UCase(e$)
    b$ = Mid(b$, 2)
    b$ = e$ & b$
    If Check3.Value = 1 And Text5.Text <> "" And Check2.Value = 1 Then
        'Check if they want a specific letter
        b$ = Mid(b$, 2)
        b$ = UCase(Text5.Text) & b$
    End If
    'Put the names together
    a$ = a$ & " " & b$
End If
If Check3.Value = 1 And Text4.Text <> "" Then
    'Check if they want a specific letter for the begining
    'of the first name
    'Cut off first letter
    a$ = Mid(a$, 2)
    'Add the chosen letter
    a$ = UCase(Text4.Text) & a$
End If
'Add it to the list
List1.AddItem a$
'Scroll the list
List1.Selected(List1.ListCount - 1) = True
MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
On Error Resume Next
'Copy the name to the clipboard
Clipboard.Clear
Clipboard.SetText List1.Text
End Sub

Private Sub Command3_Click()
On Error Resume Next
'Clear the list box
List1.Clear
End Sub

Private Sub Command4_Click()
On Error Resume Next
'Save a name
Close #1
Dim c%, v$
Dim d As Boolean
Open App.Path & "\saved1.txt" For Append As #1
    
    For i = 0 To 4
        d = False
        'Check for blanks/duplicates
        'If no duplicates, then save it
        If Names(i) <> "" Then
            For c% = 0 To List3.ListCount - 1
                If LCase(List3.List(c%)) = LCase(Names(i).Text) Then
                    d = True
                    c% = List3.ListCount - 1
                End If
            Next
            If d = False Then
                Print #1, Names(i).Text
                List3.AddItem Names(i).Text
            Else
                v$ = v$ & Names(i).Text & ", "
            End If
        End If
    Next
Close #1
Dim g$
List3.Clear
'Reload the file
Open App.Path & "\saved1.txt" For Input As #1
    While Not EOF(1)
        Line Input #1, g$
        List3.AddItem g$
    Wend
Close #1
'If there were duplicates, tell the user there were
If v$ <> "" Then
    MsgBox "Because they already exsist in the saved names file, the following names were not saved: " & vbCrLf & v$, vbCritical, "Duplicates"
End If
End Sub

Private Sub Command5_Click()
'Clear all the saved names
If MsgBox("Are you sure you wish to delete all the names? There is no undo.", vbOKCancel + vbQuestion, "Delete All") = vbOK Then
    Open App.Path & "\saved.txt" For Output As #1
        Print #1, ""
    Close #1
    List2.Clear
    Open App.Path & "\saved1.txt" For Output As #1
        Print #1, ""
    Close #1
    List3.Clear
End If
End Sub

Private Sub Form_Load()
On Error GoTo eh1
'Close any files that are open just in case
Close #1
'Default value
Text1.Text = 2
'Open the saved names
'And load them into the list boxes
Open App.Path & "\saved.txt" For Input As #1
    While Not EOF(1)
        Line Input #1, c$
        List2.AddItem c$
    Wend
Close #1
Open App.Path & "\saved1.txt" For Input As #1
    While Not EOF(1)
        Line Input #1, c$
        List3.AddItem c$
    Wend
Close #1
Exit Sub
eh1:
Exit Sub
End Sub

Private Sub List1_DblClick()
On Error Resume Next
'Close #1 just in case
Close #1
Dim c%, v$
Dim d As Boolean
'Save the selected name to the names file
Open App.Path & "\saved.txt" For Append As #1
    'Check for duplicates and save what
    'there isnt duplicates for
    For c% = 0 To List2.ListCount - 1
        d = False
        If LCase(List2.List(c%)) = LCase(List1.Text) Then
            d = True
            c% = List2.ListCount - 1
        End If
    Next
    If d = False Then
        Print #1, List1.Text
    Else
        v$ = v$ & List1.Text & "."
    End If
Close #1
'If there are duplicates, then inform the user
If v$ <> "" Then
    MsgBox "The following name(s) could not be save because they already exsist in the saved names file:" & vbCrLf & v$, vbCritical, "Duplicate"
End If
Dim f$
'Reload the new file into the list box
List2.Clear
Open App.Path & "\saved.txt" For Input As #1
    While Not EOF(1)
        Line Input #1, f$
        List2.AddItem f$
    Wend
Close #1
End Sub

Private Sub List2_DblClick()
On Error Resume Next
'Load a name into an empty, visable text box for
'Combineing that name
Close #1
For i = 0 To Val(Text1.Text) - 1
    If Names(i).Text = "" Then
        Names(i).Text = List2.Text
        Exit Sub
    End If
Next
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
'If they right click on an item, ask em
'If they want to delete the name
If Button = 2 Then
    If MsgBox("Are you sure you wish to delete the name: " & List2.Text & ", from your saved names file?", vbOKCancel + vbQuestion) = vbOK Then
        Close #1
        'Remove the name
        List2.RemoveItem List2.ListIndex
        'And save whats in that box.
        Open App.Path & "\saved.txt" For Output As #1
            For i = 0 To List2.ListCount - 1
                Print #1, List2.List(i)
            Next
        Close #1
    End If
    'Reload that file
    List2.Clear
    Dim c$
    Open App.Path & "\saved.txt" For Input As #1
        While Not EOF(1)
            Line Input #1, c$
            List2.AddItem c$
        Wend
    Close #1
End If
End Sub

Private Sub List3_DblClick()
On Error Resume Next
Close #1
'Insert the name clicked on into an
'empty visible text box
For i = 0 To Val(Text1.Text) - 1
    If Names(i).Text = "" Then
        Names(i).Text = List3.Text
        Exit Sub
    End If
Next

End Sub

Private Sub List3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
'If right click
'Delete that name and resave it
    If MsgBox("Are you sure you wish to delete the name: " & List2.Text & ", from your saved names file?", vbOKCancel + vbQuestion) = vbOK Then
        Close #1
        List3.RemoveItem List3.ListIndex
        Open App.Path & "\saved1.txt" For Output As #1
            For i = 0 To List3.ListCount - 1
                Print #1, List3.List(i)
            Next
        Close #1
    End If
    'Reload that file
    List3.Clear
    Dim c$
    Open App.Path & "\saved1.txt" For Input As #1
        While Not EOF(1)
            Line Input #1, c$
            List3.AddItem c$
        Wend
    Close #1
End If

End Sub

Private Sub Names_GotFocus(Index As Integer)
On Error Resume Next
'Auto select everying in the text box
'Whenever it gets focus
Names(Index).SelStart = 0
Names(Index).SelLength = Len(Names(Index))

End Sub

Private Sub Names_KeyPress(Index As Integer, KeyAscii As Integer)
'If they press enter, create the name
If KeyAscii = 13 Then Call Command1_Click
End Sub

Private Sub Text1_Change()
On Error Resume Next
'Make sure the number is >=2 or <=5
If Text1.Text < 2 Then
    Text1.Text = 2
    Text1.SelStart = 0
    Text1.SelLength = 1
End If
If Text1.Text > 5 Then
    Text1.Text = 5
    Text1.SelStart = 0
    Text1.SelLength = 1
End If
'Make all the boxes invisible
For i = 0 To 4
    Names(i).Visible = False
Next
'Make the desired amount visable
For i = 0 To Val(Text1.Text) - 1
    Names(i).Visible = True
Next
'Relocate the save names button under the lowest textbox
If Val(Text1.Text) <> 5 Then
    Command4.Top = Names(Val(Text1.Text)).Top
Else
    Command4.Top = 2160
End If
End Sub

Function MakeName(c$) As String
On Error Resume Next
'THe make the name function
Dim a$, X%, Y%, z%
X% = Val(Text1.Text)
For i = 0 To X% - 1
    'It will either take 1 or 2 letters
    'from each name, from a random starting location
    'and add it to the name
    Y% = Int((Len(Names(i).Text) - 1 + 1) * Rnd + 1)
    z% = Int((2 - 1 + 1) * Rnd + 1)
    a$ = a$ & Mid(Names(i).Text, Y%, z%)
Next
'Set the function to what it made
MakeName = a$
End Function

Private Sub Text1_GotFocus()
On Error Resume Next
'Select everying in text 1
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
'If they press enter, call command1's click function
If KeyAscii = 13 Then Call Command1_Click
End Sub
