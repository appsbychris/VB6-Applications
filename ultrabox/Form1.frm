VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ultrabox Demo"
   ClientHeight    =   10650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin Project1.UltraBox lL 
      Height          =   2415
      Left            =   120
      TabIndex        =   10
      Top             =   6480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4260
      Style           =   1
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
   Begin VB.CommandButton Command8 
      Caption         =   "Refill"
      Height          =   255
      Left            =   7560
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Clear"
      Height          =   255
      Left            =   6600
      TabIndex        =   8
      Top             =   2760
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   55
      Left            =   4440
      Top             =   2280
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Get Hamburgers"
      Height          =   615
      Left            =   4440
      TabIndex        =   7
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "GO"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Selected Items"
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Deselect"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Selected"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove Item"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin Project1.UltraBox l 
      Height          =   1935
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3413
      Style           =   2
      Color           =   16777215
      Fill            =   0
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
   Begin Project1.UltraBox lubD 
      Height          =   2415
      Left            =   6000
      TabIndex        =   12
      Top             =   6480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4260
      Style           =   4
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
   Begin Project1.UltraBox ub1 
      Height          =   2415
      Left            =   2760
      TabIndex        =   13
      Top             =   6480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4260
      Style           =   3
      Color           =   16777215
      Fill            =   0
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
   Begin Project1.UltraBox l1 
      Height          =   2775
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4895
      Style           =   0
      Color           =   16777215
      Fill            =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0000
      Height          =   1575
      Left            =   120
      TabIndex        =   15
      Top             =   9000
      Width           =   8895
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   9000
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8880
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Order"
      Height          =   2775
      Left            =   5520
      TabIndex        =   6
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   315
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================================================================
'=======         FONT COLOR CODES         ==========================
'===================================================================
'=======   }1 -BLACK                      ==========================
'=======   }2 -WHITE                      ==========================
'=======   }3 -RED                        ==========================
'=======   }4 -BLUE                       ==========================
'=======   }5 -GREEN                      ==========================
'=======   }6 -YELLOW                     ==========================
'=======   }7 -GRAY                       ==========================
'=======   }8 -ORANGE                     ==========================
'=======   }9 -PURPLE                     ==========================
'=======   }0 -LIGHTBLUE                  ==========================
'===================================================================

Private Sub Command1_Click()
   On Error GoTo Command1_Click_Error

l.RemoveItem CLng(InputBox("Remove Item Index"))

   On Error GoTo 0
   Exit Sub

Command1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command1_Click of Form Form1"
End Sub

Private Sub Command2_Click()
l.SetSelected CLng(InputBox("Select Item Index")), True, True
End Sub

Private Sub Command3_Click()
l.SetSelected CLng(InputBox("Deselect Item Index")), False, True
End Sub

Private Sub Command4_Click()
Dim i As Long
Dim s As String
For i = 1 To l.ListCount
    If l.IsSelected(i) = True Then
        s = s & l.List(i) & vbCrLf
    End If
Next
MsgBox s
End Sub

Private Sub Command5_Click()
Label1.Caption = "You have chosen:" & vbCrLf
Dim i As Long
Dim s As String
For i = 1 To l1.ListCount
    If l1.IsSelected(i, True) = True Then
        s = s & l1.List(i) & vbCrLf
    End If
Next
Label1.Caption = Label1.Caption & s
End Sub

Private Sub Command6_Click()
l1.SetItemText 3, "1. Hamburger"
l1.SetEnabled 3, True
End Sub

Private Sub Command7_Click()
l.Paint = False
l.Clear
l.Paint = True
End Sub

Private Sub Command8_Click()
Dim i As Long

With l
    .Paint = False
    .AddItem "Items in box is 30, but this makes 31, and im just testing this to see how a really really really really long string will work in the list box to see if the horizontal scorll with work"
    .AddItem "Items in box is 30, but this makes 32, and im just testing this to see how a really really really really long string will work in the list box to see if the horizontal scorll with work"
    .AddItem "}3It}6}em 1 }9of }530", UseCheckBox:=True
    .AddItemProgressBar "0", vbCenter, FCOLOR:=&HC0C000, ProgressBarMax:=300, ProgressBarValue:=24, ProgressBarProgressColor:=&HFF8080
    For i = 3 To 10
        .AddItem "Item " & CStr(i) & " of 30", UseOptionBox:=True
    Next
    .AddItem "Item 11 of 30", FCOLOR:=vbGreen, HCOLOR:=vbYellow, HTEXT:=vbBlack
    .AddItem "Item 12 of 30", FCOLOR:=vbBlue, BCOLOR:=vbRed, HCOLOR:=vbBlack
    For i = 13 To 20
         .AddItem "Item " & CStr(i) & " of 30", UseCheckBox:=True
    Next
    .AddItem "Item 21 of 30", HCOLOR:=vbGreen, HTEXT:=vbYellow
    .AddItem "Item 22 of 30", FCOLOR:=vbYellow, BCOLOR:=vbBlack, HCOLOR:=vbGreen, HTEXT:=vbBlack
    For i = 23 To 30
        .AddItem "Item " & CStr(i) & " of 30"
    Next
    .Paint = True
End With
End Sub

Private Sub Form_Load()
Dim i As Long
l.AddItem "Items in box is 30, but this makes 31, and im just testing this to see how a really really really really long string will work in the list box to see if the horizontal scorll with work"
With l
    .AddItem "Items in box is 30, but this makes 32, and im just testing this to see how a really really really really long string will work in the list box to see if the horizontal scorll with work"
    .AddItem "}3It}6}em 1 }9of }530", UseCheckBox:=True
    .AddItemProgressBar "0", vbCenter, FCOLOR:=&HC0C000, ProgressBarMax:=300, ProgressBarValue:=24, ProgressBarProgressColor:=&HFF8080
    For i = 3 To 10
        .AddItem "Item " & CStr(i) & " of 30", UseOptionBox:=True
    Next
    .AddItem "Item 11 of 30", FCOLOR:=vbGreen, HCOLOR:=vbYellow, HTEXT:=vbBlack
    .AddItem "Item 12 of 30", FCOLOR:=vbBlue, BCOLOR:=vbRed, HCOLOR:=vbBlack
    For i = 13 To 20
         .AddItem "Item " & CStr(i) & " of 30", UseCheckBox:=True
    Next
    .AddItem "Item 21 of 30", HCOLOR:=vbGreen, HTEXT:=vbYellow
    .AddItem "Item 22 of 30", FCOLOR:=vbYellow, BCOLOR:=vbBlack, HCOLOR:=vbGreen, HTEXT:=vbBlack
    For i = 23 To 30
        .AddItem "Item " & CStr(i) & " of 30"
    Next
End With
With l1
    .AddItem "}3Please choose }1as many}3 items as you want-", Enabled:=False, HCOLOR:=vbBlack
    .AddItem "}6----------------------------------------", Enabled:=False
    .AddItem "1. Hamburger }3SOLD OUT", Enabled:=False, UseCheckBox:=True
    .AddItem "2. Cheese Burger", UseCheckBox:=True
    .AddItem "3. Fries", UseCheckBox:=True
    .AddItem "4. Chicken Nuggets", UseCheckBox:=True
    .AddItem "}6----------------------------------------", Enabled:=False
    .AddItem "", Enabled:=False
    .AddItem "}3Please choose }11}3 of the following-", Enabled:=False, HCOLOR:=vbBlack
    .AddItem "}6----------------------------------------", Enabled:=False
    .AddItem "1. Diet Coke", UseOptionBox:=True, OptionGroup:=0
    .AddItem "2. Mountain Dew", UseOptionBox:=True, OptionGroup:=0
    .AddItem "3. Dr. Pepper", UseOptionBox:=True, OptionGroup:=0
    .AddItem "}6----------------------------------------", Enabled:=False
    .AddItem "", Enabled:=False
    .AddItem "}3Please choose }11}3 of the following-", Enabled:=False, HCOLOR:=vbBlack
    .AddItem "}6----------------------------------------", Enabled:=False
    .AddItem "1. Cash", UseOptionBox:=True, OptionGroup:=1
    .AddItem "2. Check", UseOptionBox:=True, OptionGroup:=1
    .AddItem "3. Debit", UseOptionBox:=True, OptionGroup:=1
End With
Dim lCol As Long

For i = 1 To 24
    With ub1
        .AddItem "Item #" & i ', FCOLOR:=vbBlack, BCOLOR:=&HF4F4F4
            
    End With
Next
For i = 1 To 24
    With lL
        .AddItem i & ".  This is a longer item to see how the scrollbar works, this is item #" & i ', FCOLOR:=vbBlack, BCOLOR:=&HF4F4F4
            
    End With
Next
For i = 1 To 24
    With lubD
        .AddItem "Item (#" & i & ")"
            
    End With
Next
End Sub

Private Sub l_Change()
lblLabel.Caption = l.ItemText
l.SetSelected l.ListIndex, True
End Sub

Private Sub l_Click()
lblLabel.Caption = l.ItemText
End Sub

Private Sub Timer1_Timer()
Dim s As Single
Dim v As String
   On Error GoTo Timer1_Timer_Error

l.SetProgressValue 4, l.GetProgressValue(4) + 1
If l.GetProgressValue(4) >= l.GetProgressMax(4) Then l.SetProgressValue 4, 0
s = l.GetProgressValue(4) / l.GetProgressMax(4)
s = Round(s, 2)
s = s * 100
v = CStr(s) & "%"
l.SetItemText 4, v

   On Error GoTo 0
   Exit Sub

Timer1_Timer_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer1_Timer of Form Form1"
End Sub
