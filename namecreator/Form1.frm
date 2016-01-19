VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Name Creator"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6315
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   4410
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   200
      TabIndex        =   9
      Top             =   730
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Copy Name"
      Height          =   255
      Left            =   180
      TabIndex        =   15
      Top             =   4040
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cl&ear Box"
      Height          =   255
      Left            =   180
      TabIndex        =   14
      Top             =   3680
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Generate Name"
      Height          =   255
      Left            =   180
      TabIndex        =   13
      Top             =   3320
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FF0000&
      Caption         =   "Start w/&out Vowel"
      Height          =   195
      Left            =   200
      TabIndex        =   12
      Top             =   1280
      Width           =   200
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   780
      TabIndex        =   11
      Top             =   1500
      Width           =   270
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FF0000&
      Caption         =   "Check3"
      Height          =   195
      Left            =   200
      TabIndex        =   10
      Top             =   1600
      Width           =   200
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1040
      TabIndex        =   8
      Top             =   2680
      Width           =   290
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF0000&
      Height          =   195
      Left            =   200
      TabIndex        =   7
      Top             =   2620
      Width           =   200
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF0000&
      Caption         =   "Start with &Vowel"
      Height          =   240
      Left            =   200
      TabIndex        =   6
      Top             =   990
      Width           =   200
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   780
      TabIndex        =   5
      Top             =   1840
      Width           =   270
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00FF0000&
      Height          =   195
      Left            =   200
      TabIndex        =   4
      Top             =   1960
      Width           =   200
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1500
      TabIndex        =   3
      Top             =   180
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   2560
      TabIndex        =   2
      Top             =   320
      Width           =   3615
   End
   Begin VB.ListBox Plugs 
      Height          =   1035
      Left            =   2600
      TabIndex        =   1
      Top             =   3320
      Width           =   3615
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   6000
      TabIndex        =   0
      Top             =   4560
      Width           =   375
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "&Options"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnucombine 
         Caption         =   "&Combine Names"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnujkl2138ds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is the "pool" of letters to choose from
Const letConst = "bcdfnnghjkycclmrrnpllmmqrstvwxddyz"
Const letVowel = "aaaeieeiiouy"
Dim booStart As Boolean, i As Integer
Dim tempName As String

Private Sub Check2_Click()
On Error Resume Next
'Enables Choose length buttons, and text boxes and what not
If Check2.Value = 1 Then
    Text1.Enabled = True
    Text1.BackColor = vbWhite
    Text2.Enabled = True
    Text2.BackColor = vbWhite
End If
If Check2.Value = 0 Then
    Text1.Enabled = False
    Text1.BackColor = &H80000004
    Text2.Enabled = False
    Text2.BackColor = &H80000004
End If

End Sub

Private Sub Command1_Click()
On Error Resume Next
'Makes the name...
MousePointer = vbHourglass
'Disables this button so the user doesnt conjest the system up
Command1.Enabled = False
Dim tempConst As String, tempVowel As String
Dim X As Integer
Dim lC%, lV%, p%
'Sees if the choose length is selected
If Check2.Value = 0 Then
    'if it is, get a random number for the length
    p% = Int((10 - 3 + 3) * Rnd + 3)
Else
    'if not, just use what then input
    p% = Val(Text1.Text)
End If
'Get the pool of letters
tempConst = letConst
tempVowel = letVowel
'get the lengths of the pool
lC% = Len(tempConst)
lV% = Len(tempVowel)
'Clear the name variable
tempName = ""
'If they want a vowel in the begining...
If booStart = True Then
    'Get a random letter from the 'pool'
    X = Int((lV% - 1 + 1) * Rnd + 1)
    'And put it into the Name Variable
    tempName = tempName & Mid(tempVowel, X, 1)
    'Now loop through and finish the name
    For i = 1 To p% - 1
        'Call the function to see if we use a vowel or constenent next
        If Alternate = True Then
            'If we alternate to a const, choose one at random
            X = Int((lC% - 1 + 1) * Rnd + 1)
            tempName = tempName & Mid(tempConst, X, 1)
            'If any of the letter combos from the plugin are true, we need to
            'increase 1, since we added a letter
            If ComboVowel = True Then i = i + 1
        Else
            'If we need a vowel next, choose one at random
            X = Int((lV% - 1 + 1) * Rnd + 1)
            tempName = tempName & Mid(tempVowel, X, 1)
            'If any of the letter combos from the plugin are true, we need to
            'increase 1, since we added a letter
            If ComboConst = True Then i = i + 1
        End If
    Next
Else
    'If we want a constenent first, then
    'Get a random constenent
    X = Int((lC% - 1 + 1) * Rnd + 1)
    tempName = tempName & Mid(tempConst, X, 1)
    For i = 1 To p% - 1
        '(See above for more explenation)
        If Alternate = True Then
            'Get a vowel
            X = Int((lV% - 1 + 1) * Rnd + 1)
            tempName = tempName & Mid(tempVowel, X, 1)
            If ComboVowel = True Then i = i + 1
        Else
            'get a constenent
            X = Int((lC% - 1 + 1) * Rnd + 1)
            tempName = tempName & Mid(tempConst, X, 1)
            If ComboConst = True Then i = i + 1
        End If
    Next
End If
'Dim the Last name variable
Dim tempName1 As String
'If they want a last name
If Check1.Value = 1 Then
    'If they want to choose length, or random length
    If Check2.Value = 0 Then
        p% = Int((10 - 3 + 3) * Rnd + 3)
    Else
        p% = Val(Text2.Text)
    End If
    'if vowel first
    If booStart = True Then
        X = Int((lV% - 1 + 1) * Rnd + 1)
        tempName1 = tempName1 & Mid(tempVowel, X, 1)
        For i = 1 To p% - 1
            If Alternate = True Then
                X = Int((lC% - 1 + 1) * Rnd + 1)
                tempName1 = tempName1 & Mid(tempConst, X, 1)
                If ComboVowel = True Then i = i + 1
            Else
                X = Int((lV% - 1 + 1) * Rnd + 1)
                tempName1 = tempName1 & Mid(tempVowel, X, 1)
                If ComboConst = True Then i = i + 1
            End If
        Next
    Else
        'if const first
        X = Int((lC% - 1 + 1) * Rnd + 1)
        tempName1 = tempName1 & Mid(tempConst, X, 1)
        For i = 1 To p% - 1
            If Alternate = True Then
                X = Int((lV% - 1 + 1) * Rnd + 1)
                tempName1 = tempName1 & Mid(tempVowel, X, 1)
                If ComboVowel = True Then i = i + 1
            Else
                X = Int((lC% - 1 + 1) * Rnd + 1)
                tempName1 = tempName1 & Mid(tempConst, X, 1)
                If ComboConst = True Then i = i + 1
            End If
        Next
    End If
    Dim d$
    'Capitalize the first letter
    d$ = Left(tempName1, 1)
    d$ = UCase(d$)
    tempName1 = Mid(tempName1, 2)
    tempName1 = d$ & tempName1
    'if they want it to start with a specific letter
    If Check4.Value = 1 And Text4.Text <> "" Then
        tempName1 = Mid(tempName1, 2)
        tempName1 = UCase(Text4.Text) & tempName1
    End If
    'Put the names together
    tempName = tempName & " " & tempName1
End If
Dim f$
'Capitalize the first name
f$ = Left(tempName, 1)
f$ = UCase(f$)
tempName = Mid(tempName, 2)
tempName = f$ & tempName
'Check if they want a specific letter
If Check3.Value = 1 And Text3.Text <> "" Then
    tempName = Mid$(tempName, 2)
    tempName = UCase(Text3.Text) & tempName
End If
'add it to the list
List1.AddItem tempName
'Reendable this button
Command1.Enabled = True
'Scroll the list box
List1.Selected(List1.ListCount - 1) = True
'set focus on the button
Command1.SetFocus
'Restore the mouse to default
MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
On Error Resume Next
'Clear the list
List1.Clear
End Sub

Private Sub Command3_Click()
On Error Resume Next
'Copy to the clipboard
Clipboard.Clear
Clipboard.SetText List1.Text
End Sub

Private Sub Form_Load()
On Error Resume Next
'Ensure a random number is chosen
Randomize
On Error GoTo eh1
'Make a default vaule
Text1.Text = 5
Dim c$
'Open any saved names...
Open App.Path & "\saved.txt" For Input As #1
    While Not EOF(1)
        Line Input #1, c$
        List2.AddItem c$
    Wend
Close #1
Exit Sub
eh1:
Exit Sub

End Sub

Private Sub List1_DblClick()
On Error Resume Next
'Close #1 if it is open
Close #1
Dim c%, v$
Dim d As Boolean
'Check for duplicates, and then right to file
Open App.Path & "\saved.txt" For Append As #1
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
'If there are duplicates...
If v$ <> "" Then
    MsgBox "The following name(s) could not be save because they already exsist in the saved names file:" & vbCrLf & v$, vbCritical, "Duplicate"
End If
Dim xx$
'Reload the file into the list box
Open App.Path & "\saved.txt" For Input As #1
    While Not EOF(1)
        Line Input #1, xx$
        List2.AddItem xx$
    Wend
Close #1

End Sub

Private Sub mnucombine_Click()
On Error Resume Next
'Load the combine names form.
Load Form3
Form3.Show 1

End Sub

Private Sub mnuexit_Click()
'Quit the program
Unload Me
End Sub

Private Sub Option1_Click()
On Error Resume Next
'Set they want a vowel first
booStart = True
End Sub

Private Sub Option2_Click()
On Error Resume Next
'Set they want a const first
booStart = False
End Sub

Function Alternate() As Boolean
On Error Resume Next
Dim a As Integer
'Make a random chance of true or false, weighing more on a "true" answer
a = Int((2 - 1 + 1) * Rnd + 1)
If a = 1 Then
    a = Int((2 - 1 + 1) * Rnd + 1)
    If a = 1 Then
        a = Int((2 - 1 + 1) * Rnd + 1)
        If a = 1 Then
            Alternate = False
        End If
    End If
Else
    Alternate = True
End If
End Function

Function ComboVowel() As Boolean
On Error Resume Next
Dim a As Integer, b As String, c As String
Dim d As Integer
'Check the plun in's for any letter combonations
For a = 0 To Plugs.ListCount - 1
    'See if it the vowel plugin
    If InStr(1, LCase(Plugs.List(a)), "vowel") Then
        'if so, open it
        Open Plugs.List(a) For Input As #1
            'loop through the file
            While Not EOF(1)
                Line Input #1, b
                'split it up so i can check it
                c = b
                d = InStr(1, c, "+")
                b = LCase(Left(c, d))
                c = LCase(Right(c, 1))
                If Right(tempName, Len(b)) = b And b <> "" Then
                    'randomly add the selected letter
                    If Alternate = True Then
                        tempName = tempName & c
                        'Set this value to true, so the name will add 1 to it
                        ComboVowel = True
                        'exit the function
                        Exit Function
                    End If
                End If
            Wend
        Close #1
    End If
Next
End Function

Function ComboConst() As Boolean
On Error Resume Next
Dim a As Integer, b As String, c As String
Dim d As Integer
'see above for comments, does basically the same thing
For a = 0 To Plugs.ListCount - 1
    If InStr(1, LCase(Plugs.List(a)), "const") Then
        Open Plugs.List(a) For Input As #1
            While Not EOF(1)
                Line Input #1, b
                c = b
                d = InStr(1, c, "+")
                b = LCase(Left(c, d))
                c = LCase(Right(c, 1))
                If Right(tempName, Len(b)) = b And b <> "" Then
                    If Alternate = True Then
                        tempName = tempName & c
                        ComboConst = True
                        Exit Function
                    End If
                End If
            Wend
        Close #1
    End If
Next
End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
'If they press enter in text1, generate the name.
If KeyAscii = 13 Then Call Command1_Click
End Sub
