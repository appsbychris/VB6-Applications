VERSION 5.00
Begin VB.Form frmEmotions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Emotions"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   7815
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   240
      TabIndex        =   24
      Top             =   240
      Width           =   2415
   End
   Begin ServerEditor.UltraBox lstEmotes 
      Height          =   3735
      Left            =   240
      TabIndex        =   23
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   6588
      Style           =   3
      Color           =   0
      Fill            =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "(save)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "(new)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "< Previous"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   4680
      Width           =   1215
   End
   Begin ServerEditor.Raise Raise3 
      Height          =   495
      Left            =   2880
      TabIndex        =   21
      Top             =   4560
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   873
      Style           =   2
      Color           =   0
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   3120
      ScaleHeight     =   4095
      ScaleWidth      =   4215
      TabIndex        =   13
      Top             =   240
      Width           =   4215
      Begin VB.TextBox txtYouSee 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtOthers 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox txtSomeoneToYou 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   3120
         Width           =   3735
      End
      Begin VB.TextBox txtOthersSomeoneElse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtYouSomeoneElse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   3735
      End
      Begin VB.TextBox txtID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtSyntax 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "What you see (no target):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1905
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "What others see (no target):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   2100
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "What you see if someone targets you with an Emotion:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   2880
         Width           =   3960
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "What others see you do to someone else:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "What you see doing to someone else:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   3480
         Width           =   2715
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   225
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Syntax:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   570
      End
   End
   Begin ServerEditor.Raise Raise2 
      Height          =   4335
      Left            =   2880
      TabIndex        =   12
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   7646
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise1 
      Height          =   4335
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   7646
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise4 
      Height          =   5175
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9128
      Style           =   4
      Color           =   0
   End
End
Attribute VB_Name = "frmEmotions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem*************************************************************************************
Rem*************************************************************************************
Rem***************       Code create by Chris Van Hooser          **********************
Rem***************                  (c)2002                       **********************
Rem*************** You may use this code and freely distribute it **********************
Rem***************   If you have any questions, please email me   **********************
Rem***************          at theendorbunker@attbi.com.          **********************
Rem***************       Thanks for downloading my project        **********************
Rem***************        and i hope you can use it well.         **********************
Rem***************                frmEmotions                     **********************
Rem***************                ServerEditor                    **********************
Rem***************                Editor.vbp                      **********************
Rem*************************************************************************************
Rem*************************************************************************************

Dim lcID As Long
Dim bIs As Boolean
Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF4
        cmdSave_Click
    Case vbKeyF3
        cmdNext_Click
    Case vbKeyF2
        cmdPrevious_Click
End Select
End Sub

Private Sub txtFind_Change()
lstEmotes.SetSelected lstEmotes.FindInStr(txtFind.Text), True, True
End Sub

Private Sub cmdAddNew_Click()
Dim x As Long
Dim i As Long
Dim t As Boolean
MousePointer = vbHourglass
ReDim Preserve dbEmotions(1 To UBound(dbEmotions) + 1)
x = dbEmotions(UBound(dbEmotions) - 1).iID
x = x + 1
Do Until t = True
    t = True
    i = GetEmotionID(, x)
    If i <> 0 Then
        t = False
        x = x + 1
    End If
Loop
With dbEmotions(UBound(dbEmotions))
    .iID = x
    .sSyntax = "New Emotion"
    .sPhraseYou = "You xxx"
    .sPhraseOthers = "<player> xxx."
    .sPhraseToYou = "<player> xxx at you!"
    .sPhraseOthers2 = "<player> xxx at <victim>!"
    .sPhraseYouToOther = "You xxx at <victim>!"
End With
lcID = UBound(dbEmotions)
FillEmotes lcID, True
MousePointer = vbDefault
End Sub

Private Sub cmdNext_Click()
On Error GoTo cmdNext_Click_Error
SaveEmotes
lcID = lcID + 1
If lcID > UBound(dbEmotions) Then lcID = 1
FillEmotes lcID
On Error GoTo 0
Exit Sub
cmdNext_Click_Error:

End Sub

Private Sub cmdPrevious_Click()
On Error GoTo cmdPrevious_Click_Error
SaveEmotes
lcID = lcID - 1
If lcID < LBound(dbEmotions) Then lcID = UBound(dbEmotions)
FillEmotes lcID
On Error GoTo 0
Exit Sub
cmdPrevious_Click_Error:
End Sub

Private Sub cmdSave_Click()
SaveEmotes
End Sub

Private Sub Form_Load()
FillEmotes FillList:=True
End Sub

Sub SaveEmotes()
MousePointer = vbHourglass
With dbEmotions(lcID)
    .sSyntax = txtSyntax.Text
    .sPhraseYou = txtYouSee.Text
    .sPhraseOthers = txtOthers.Text
    .sPhraseOthers2 = txtOthersSomeoneElse
    .sPhraseToYou = txtSomeoneToYou.Text
    .sPhraseYouToOther = txtYouSomeoneElse.Text
End With
modUpdateDatabase.SaveMemoryToDatabase Emotions
FillEmotes lcID, True
MousePointer = vbDefault
End Sub

Sub FillEmotes(Optional Arg As Long = -1, Optional FillList As Boolean = False)
Dim i As Long, j As Long
Dim m As Long
Dim Arr() As String
MousePointer = vbHourglass
bIs = True
If Arg = -1 Then Arg = LBound(dbEmotions)
If FillList Then lstEmotes.Clear

For i = LBound(dbEmotions) To UBound(dbEmotions)
    With dbEmotions(i)
        If FillList Then lstEmotes.AddItem CStr(.iID & " " & .sSyntax)
        If i = Arg Then
            lcID = i
            txtID.Text = .iID
            txtSyntax.Text = .sSyntax
            txtYouSee.Text = .sPhraseYou
            txtOthers.Text = .sPhraseOthers
            txtOthersSomeoneElse.Text = .sPhraseOthers2
            txtSomeoneToYou.Text = .sPhraseToYou
            txtYouSomeoneElse.Text = .sPhraseYouToOther
            If Not FillList Then Exit For
        End If
    End With
Next
modMain.SetLstSelected lstEmotes, txtID.Text & " " & txtSyntax.Text
bIs = False
MousePointer = vbDefault
End Sub

Private Sub lstEmotes_Click()
If bIs Then Exit Sub
Dim i As Long
MousePointer = vbHourglass
For i = LBound(dbEmotions) To UBound(dbEmotions)
    With dbEmotions(i)
        If .iID & " " & .sSyntax = lstEmotes.ItemText Then
            FillEmotes i
            Exit For
        End If
    End With
Next
MousePointer = vbDefault
End Sub
