VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ultrabox Demo"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPre 
      Caption         =   "Previous"
      Height          =   375
      Left            =   8160
      TabIndex        =   46
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   9480
      TabIndex        =   45
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame fraDemo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Demo 4"
      Height          =   3855
      Index           =   3
      Left            =   120
      TabIndex        =   38
      Top             =   120
      Visible         =   0   'False
      Width           =   10455
      Begin Project1.UltraBox LB 
         Height          =   1575
         Left            =   7440
         TabIndex        =   44
         Top             =   2160
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2778
         Style           =   2
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
         Mult            =   0   'False
         Sort            =   0   'False
      End
      Begin Project1.UltraBox lN 
         Height          =   1575
         Left            =   3600
         TabIndex        =   43
         Top             =   2160
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2778
         Style           =   5
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
         Mult            =   0   'False
         Sort            =   0   'False
      End
      Begin Project1.UltraBox lR 
         Height          =   1575
         Left            =   120
         TabIndex        =   42
         Top             =   2160
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2778
         Style           =   0
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
         Mult            =   0   'False
         Sort            =   0   'False
      End
      Begin Project1.UltraBox lL 
         Height          =   1815
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   3201
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
         Mult            =   0   'False
         Sort            =   0   'False
      End
      Begin Project1.UltraBox lubD 
         Height          =   1815
         Left            =   7440
         TabIndex        =   40
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3201
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
         Mult            =   0   'False
         Sort            =   0   'False
      End
      Begin Project1.UltraBox ub1 
         Height          =   1815
         Left            =   3600
         TabIndex        =   41
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3201
         Style           =   3
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
         Mult            =   0   'False
         Sort            =   0   'False
      End
   End
   Begin VB.Frame fraDemo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Demo 3"
      Height          =   3855
      Index           =   2
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton Command10 
         Caption         =   "Fill With Random Characters"
         Height          =   735
         Left            =   5640
         TabIndex        =   36
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Sort List"
         Height          =   735
         Left            =   5640
         TabIndex        =   35
         Top             =   1080
         Width           =   1695
      End
      Begin Project1.UltraBox sort 
         Height          =   3495
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   6165
         Style           =   5
         Color           =   0
         Fill            =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mult            =   0   'False
         Sort            =   0   'False
      End
   End
   Begin VB.PictureBox picDebit 
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   2520
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   24
      Top             =   4080
      Visible         =   0   'False
      Width           =   513
   End
   Begin VB.PictureBox picCheck 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   3000
      Picture         =   "Form1.frx":0772
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   23
      Top             =   4080
      Visible         =   0   'False
      Width           =   513
   End
   Begin VB.PictureBox picCash 
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   3480
      Picture         =   "Form1.frx":0D74
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   22
      Top             =   4080
      Visible         =   0   'False
      Width           =   513
   End
   Begin VB.PictureBox picDrPepper 
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   2040
      Picture         =   "Form1.frx":14E6
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   21
      Top             =   4080
      Visible         =   0   'False
      Width           =   513
   End
   Begin VB.PictureBox picMountainDew 
      AutoSize        =   -1  'True
      Height          =   510
      Left            =   1080
      Picture         =   "Form1.frx":2668
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   513
   End
   Begin VB.PictureBox picdietCoke 
      AutoSize        =   -1  'True
      Height          =   510
      Left            =   600
      Picture         =   "Form1.frx":3172
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   19
      Top             =   4080
      Visible         =   0   'False
      Width           =   513
   End
   Begin VB.PictureBox picNuggets 
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   1560
      Picture         =   "Form1.frx":3C7C
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   513
   End
   Begin VB.PictureBox picFries 
      AutoSize        =   -1  'True
      Height          =   480
      Left            =   120
      Picture         =   "Form1.frx":472A
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   17
      Top             =   4080
      Visible         =   0   'False
      Width           =   513
   End
   Begin VB.PictureBox picCheeseburger 
      AutoSize        =   -1  'True
      Height          =   390
      Left            =   3960
      Picture         =   "Form1.frx":517C
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   16
      Top             =   4080
      Visible         =   0   'False
      Width           =   513
   End
   Begin VB.PictureBox picHamburger 
      AutoSize        =   -1  'True
      Height          =   450
      Left            =   4440
      Picture         =   "Form1.frx":59A6
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   513
   End
   Begin VB.Frame fraDemo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Demo 2"
      Height          =   3855
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton Command5 
         Caption         =   "GO"
         Height          =   615
         Left            =   6960
         TabIndex        =   12
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Get Hamburgers"
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   1335
      End
      Begin Project1.UltraBox l1 
         Height          =   2775
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   4895
         Style           =   0
         Fill            =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mult            =   -1  'True
         Sort            =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Order"
         Height          =   3495
         Left            =   8760
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraDemo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Demo 1"
      Height          =   3855
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10455
      Begin VB.CommandButton Command16 
         Caption         =   "Fill = Normal"
         Height          =   255
         Left            =   4680
         TabIndex        =   37
         Top             =   3240
         Width           =   1455
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "FindEqual"
         Height          =   255
         Left            =   4320
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "FindInStr"
         Height          =   255
         Left            =   3000
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   30
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Enable Item"
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Disable Item"
         Height          =   255
         Left            =   3120
         TabIndex        =   27
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Sorted = F"
         Height          =   255
         Left            =   7800
         TabIndex        =   26
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Add Item"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Remove Item"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Set Selected"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Deselect"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Selected Items"
         Height          =   255
         Left            =   4680
         TabIndex        =   5
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   10080
         Top             =   3480
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Clear"
         Height          =   255
         Left            =   6240
         TabIndex        =   4
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Refill"
         Height          =   255
         Left            =   6240
         TabIndex        =   3
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Enabled = T"
         Height          =   255
         Left            =   7800
         TabIndex        =   2
         Top             =   2880
         Width           =   1455
      End
      Begin Project1.UltraBox l 
         Height          =   2055
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   3625
         Style           =   2
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
         Mult            =   -1  'True
         Sort            =   0   'False
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text"
         Height          =   195
         Left            =   6000
         TabIndex        =   32
         Top             =   360
         Width           =   4275
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   5040
      Picture         =   "Form1.frx":6340
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   300
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
'=======   }i -ITALICE                    ==========================
'=======   }b -BOLD                       ==========================
'=======   }u -UNDERLINE                  ==========================
'=======   }n -NORMAL                     ==========================
'===================================================================
Dim lngIn As Long

Private Sub cmdNext_Click()
Dim i As Long
lngIn = lngIn + 1
If lngIn > fraDemo.uBound Then
    lngIn = fraDemo.lBound
End If
For i = fraDemo.lBound To fraDemo.uBound
    fraDemo(i).Visible = False
Next
fraDemo(lngIn).Visible = True
End Sub

Private Sub cmdPre_Click()
Dim i As Long
lngIn = lngIn - 1
If lngIn < fraDemo.lBound Then
    lngIn = fraDemo.uBound
End If
For i = fraDemo.lBound To fraDemo.uBound
    fraDemo(i).Visible = False
Next
fraDemo(lngIn).Visible = True
End Sub

Private Sub Command1_Click()
   On Error GoTo Command1_Click_Error

l.RemoveItem CLng(InputBox("Remove Item Index"))

   On Error GoTo 0
   Exit Sub

Command1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command1_Click of Form Form1"
End Sub

Private Sub Command10_Click()
Dim j&, i&, s$
sort.Clear
sort.Sorted = False
sort.Paint = False
Do Until j = 100
    s = s & Chr$(RndNumber(65, 118))
    If Len(s) > RndNumber(4, 17) Then
        sort.AddItem s, FCOLOR:=RGB(Val(RndNumber(0, 255)), Val(RndNumber(0, 255)), Val(RndNumber(0, 255))), BCOLOR:=vbWhite
        s = ""
        j = j + 1
    End If
    DoEvents
Loop
sort.Paint = True
End Sub

Public Function RndNumber(Min As Long, Max As Long) As Long
'gets a random number from the min to the max
RndNumber = Val((Rnd * (Max - Min)) + Min)
End Function

Private Sub Command11_Click()
sort.Sorted = True
End Sub

Private Sub Command12_Click()
l.AddItem InputBox$("Item's Text:"), BCOLOR:=vbWhite
End Sub

Private Sub Command13_Click()
If Command13.Caption = "Sorted = F" Then
    Timer1.Enabled = False
    l.Sorted = True
    Command13.Caption = "Sorted = T"
Else
    l.Sorted = False
    Command7_Click
    Command8_Click
    Command13.Caption = "Sorted = F"
    Timer1.Enabled = True
End If
End Sub

Private Sub Command14_Click()
   On Error GoTo Command14_Click_Error

l.SetEnabled CLng(InputBox("Disable Item Index")), False

   On Error GoTo 0
   Exit Sub

Command14_Click_Error:
End Sub

Private Sub Command15_Click()
On Error GoTo Command14_Click_Error

l.SetEnabled CLng(InputBox("Enable Item Index")), True

   On Error GoTo 0
   Exit Sub

Command14_Click_Error:
End Sub

Private Sub Command16_Click()
If Command16.Caption = "Fill = Normal" Then
    l.FillView = Lined
    Command16.Caption = "Fill = Lined"
Else
    l.FillView = NoStyle
    Command16.Caption = "Fill = Normal"
End If
End Sub

Private Sub Command2_Click()
   On Error GoTo Command2_Click_Error

l.SetSelected CLng(InputBox("Select Item Index")), True, True

   On Error GoTo 0
   Exit Sub

Command2_Click_Error:
End Sub

Private Sub Command3_Click()
   On Error GoTo Command3_Click_Error

l.SetSelected CLng(InputBox("Deselect Item Index")), False, True

   On Error GoTo 0
   Exit Sub

Command3_Click_Error:
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
If Command6.Caption = "Get Hamburgers" Then
    l1.SetItemText 3, "1. Hamburger"
    l1.SetEnabled 3, True
    Command6.Caption = "Sell Hamburgers"
Else
    l1.SetItemText 3, "1. Hamburger }3}bSOLD OUT"
    l1.SetEnabled 3, False
    Command6.Caption = "Get Hamburgers"
End If
End Sub

Private Sub Command7_Click()
l.Paint = False
l.Clear
l.Paint = True
End Sub

Private Sub Command8_Click()
Dim i As Long
l.AddItem "Items in box is 30, but this makes 31, and im just testing this to see how a really really really really long string will work in the list box to see if the horizontal scorll with work"
With l
    .Paint = False
    .AddItem "Items in box is 30, but this makes 32, and im just testing this to see how a really really really really long string will work in the list box to see if the horizontal scorll with work", , Picture1.Picture, RGB(128, 0, 0)
    .AddItem "}3It}6}em 1 }9of }530", UseCheckBox:=True
    .AddItemProgressBar "0", vbCenter, FCOLOR:=&HC0C000, ProgressBarMax:=300, ProgressBarValue:=24, ProgressBarProgressColor:=&HFF8080
    For i = 3 To 10
        .AddItem "Item " & CStr(i) & " of 30", , Picture1.Picture, RGB(128, 0, 0), UseOptionBox:=True
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

Private Sub Command9_Click()
l.Enabled = Not l.Enabled
Select Case l.Enabled
    Case True
        Command9.Caption = "Enabled = T"
        Timer1.Enabled = True
    Case Else
        Command9.Caption = "Enabled = F"
        Timer1.Enabled = False
End Select
End Sub

Private Sub Form_Load()
Dim i As Long
Randomize Timer
With l
    .Paint = False
    .AddItem "}bBOLD }n}iITALIC }n}uUNDERLINE }n}b}i}uALL THREE!"
    .AddItem "Items in box is 33, and im just testing this to see how a really really really really long string will work in the list box to see if the horizontal scorll with work"
    .AddItem "}3It}6}em 1 }9of }530", UseCheckBox:=True
    .AddItemProgressBar "0", vbCenter, FCOLOR:=&HC0C000, ProgressBarMax:=300, ProgressBarValue:=24, ProgressBarProgressColor:=&HFF8080
    .AddItem "Items in box is 33, and im just testing this to see how a really really really really long string will work in the list box to see if the horizontal scorll with work", , Picture1.Picture, RGB(128, 0, 0)
    For i = 3 To 10
        .AddItem "Item " & CStr(i) & " of 30", , Picture1.Picture, RGB(128, 0, 0), UseOptionBox:=True
    Next
    .AddItem "Item 11 of 33", FCOLOR:=vbGreen, HCOLOR:=vbYellow, HTEXT:=vbBlack
    .AddItem "Item 12 of 33", FCOLOR:=vbBlue, BCOLOR:=vbRed, HCOLOR:=vbBlack
    For i = 13 To 20
         .AddItem "Item " & CStr(i) & " of 30", UseCheckBox:=True
    Next
    .AddItem "Item 21 of 33", HCOLOR:=vbGreen, HTEXT:=vbYellow
    .AddItem "Item 22 of 33", FCOLOR:=vbYellow, BCOLOR:=vbBlack, HCOLOR:=vbGreen, HTEXT:=vbBlack
    For i = 23 To 30
        .AddItem "Item " & CStr(i) & " of 33"
    Next
    .Paint = True
End With
With l1
    .Paint = False
    .AddItem "}3}uPlease choose }1}ias many}3}n}u items as you want-", Enabled:=False, HCOLOR:=vbBlack
    .AddItem "}6----------------------------------------", Enabled:=False
    .AddItem "1. Hamburger }3}bSOLD OUT", pPicture:=picHamburger.Picture, TRANSColor:=RGB(0, 255, 0), Enabled:=False, UseCheckBox:=True
    .AddItem "2. Cheese Burger", pPicture:=picCheeseburger.Picture, TRANSColor:=RGB(0, 255, 0), UseCheckBox:=True
    .AddItem "3. Fries", pPicture:=picFries.Picture, TRANSColor:=RGB(0, 255, 0), UseCheckBox:=True
    .AddItem "4. Chicken Nuggets", pPicture:=picNuggets.Picture, TRANSColor:=RGB(0, 255, 0), UseCheckBox:=True
    .AddItem "}6----------------------------------------", Enabled:=False
    .AddItem "", Enabled:=False
    .AddItem "}3}uPlease choose }1}i}b1}3}n}u of the following-", Enabled:=False, HCOLOR:=vbBlack
    .AddItem "}6----------------------------------------", Enabled:=False
    .AddItem "1. Diet Coke", pPicture:=picdietCoke.Picture, TRANSColor:=RGB(0, 255, 0), UseOptionBox:=True, OptionGroup:=0
    .AddItem "2. Mountain Dew", pPicture:=picMountainDew.Picture, TRANSColor:=RGB(0, 255, 0), UseOptionBox:=True, OptionGroup:=0
    .AddItem "3. Dr. Pepper", pPicture:=picDrPepper.Picture, TRANSColor:=RGB(0, 255, 0), UseOptionBox:=True, OptionGroup:=0
    .AddItem "}6----------------------------------------", Enabled:=False
    .AddItem "", Enabled:=False
    .AddItem "}3}uPlease choose }1}i}b1}3}n}u of the following-", Enabled:=False, HCOLOR:=vbBlack
    .AddItem "}6----------------------------------------", Enabled:=False
    .AddItem "1. Cash", pPicture:=picCash.Picture, TRANSColor:=RGB(0, 255, 0), UseOptionBox:=True, OptionGroup:=1
    .AddItem "2. Check", pPicture:=picCheck.Picture, UseOptionBox:=True, OptionGroup:=1
    .AddItem "3. Debit", pPicture:=picDebit.Picture, UseOptionBox:=True, OptionGroup:=1
    .Paint = True
End With

For i = 1 To 24
    With ub1
        .Paint = False
        .AddItem "Item #" & i ', FCOLOR:=vbBlack, BCOLOR:=&HF4F4F4
        .Paint = True
    End With
Next
For i = 1 To 24
    With lL
        .Paint = False
        .AddItem i & ".  This is a longer item to see how the scrollbar works, this is item #" & i ', FCOLOR:=vbBlack, BCOLOR:=&HF4F4F4
        .Paint = True
    End With
Next
For i = 1 To 24
    With lubD
        .Paint = False
        .AddItem "Item (#" & i & ")"
        .Paint = True
    End With
    With lN
        .AddItem "Item (#" & i & ")"
    End With
    With LB
        .AddItem "Item (#" & i & ")"
    End With
    With lR
        .AddItem "Item (#" & i & ")"
    End With
Next
End Sub

Private Sub l_Click()
lblLabel.Caption = l.ItemText(True)
End Sub

Private Sub Timer1_Timer()
Dim s As Single
Static PbMax As Long
Dim lPBVal As Long
Dim v As String
   On Error GoTo Timer1_Timer_Error
l.Paint = False
If PbMax = 0 Then PbMax = l.GetProgressMax(4)
lPBVal = l.GetProgressValue(4)
If lPBVal + 1 >= PbMax Then l.SetProgressValue 4, 0: lPBVal = 0 Else l.SetProgressValue 4, lPBVal + 1: lPBVal = lPBVal + 1
s = lPBVal / PbMax
s = Round(s, 2)
s = s * 100
v = CStr(s) & "%"
l.SetItemText 4, v
l.Paint = True
   On Error GoTo 0
   Exit Sub

Timer1_Timer_Error:
l.Paint = True
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer1_Timer of Form Form1"
End Sub

Private Sub txtFind_Change()
If opt1.Value = True Then
    l.SetSelected l.FindInStr(txtFind.Text), True, True
Else
    l.SetSelected l.Find(txtFind.Text), True, True
End If
End Sub
