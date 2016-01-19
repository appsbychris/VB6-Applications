VERSION 5.00
Begin VB.Form frmRace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Race"
   ClientHeight    =   9735
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9735
   ScaleWidth      =   11535
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   240
      TabIndex        =   48
      Top             =   240
      Width           =   3255
   End
   Begin ServerEditor.UltraBox lstRaces 
      Height          =   8295
      Left            =   240
      TabIndex        =   47
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   14631
      Style           =   3
      Color           =   16777215
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
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   8655
      Left            =   3840
      ScaleHeight     =   8655
      ScaleWidth      =   7455
      TabIndex        =   20
      Top             =   240
      Width           =   7455
      Begin VB.ComboBox cboVision 
         Height          =   315
         ItemData        =   "frmRace.frx":0000
         Left            =   240
         List            =   "frmRace.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   8160
         Width           =   3615
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   240
         ScaleHeight     =   1935
         ScaleWidth      =   6855
         TabIndex        =   38
         Top             =   5400
         Width           =   6855
         Begin ServerEditor.NumOnlyText txtMinAgeS 
            Height          =   255
            Left            =   1560
            TabIndex        =   10
            Top             =   960
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   0
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtMaxAgeS 
            Height          =   255
            Left            =   1560
            TabIndex        =   11
            Top             =   1320
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   0
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtMaxAge 
            Height          =   255
            Left            =   1560
            TabIndex        =   12
            Top             =   1680
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   0
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Max Age:"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   42
            Top             =   1680
            Width           =   690
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Min Age Start:"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   41
            Top             =   960
            Width           =   1035
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Max Age Start:"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   40
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   $"frmRace.frx":00EC
            Height          =   855
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   6735
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   240
         ScaleHeight     =   975
         ScaleWidth      =   6855
         TabIndex        =   33
         Top             =   4080
         Width           =   6855
         Begin ServerEditor.NumOnlyText txtMinHP 
            Height          =   255
            Left            =   840
            TabIndex        =   8
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   0
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtMaxHP 
            Height          =   255
            Left            =   840
            TabIndex        =   9
            Top             =   600
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   0
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "HP: The starting characters Hit Points will be a random number in these boundries, inclusive."
            Height          =   195
            Index           =   11
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   6615
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Max HP:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   35
            Top             =   600
            Width           =   600
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Min HP:"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   5055
         TabIndex        =   26
         Top             =   1440
         Width           =   5055
         Begin VB.Label lblLabel 
            Caption         =   "Starting Stats: These are the stats a character will start with. Good starting numbers are mostly between 1 and 9."
            Height          =   435
            Index           =   10
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   4980
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   240
         ScaleHeight     =   1815
         ScaleWidth      =   5055
         TabIndex        =   24
         Top             =   1920
         Width           =   5055
         Begin ServerEditor.NumOnlyText txtStr 
            Height          =   255
            Left            =   960
            TabIndex        =   3
            Top             =   120
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   0
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtDex 
            Height          =   255
            Left            =   960
            TabIndex        =   5
            Top             =   840
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   0
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtInt 
            Height          =   255
            Left            =   960
            TabIndex        =   4
            Top             =   480
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   0
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtAgil 
            Height          =   255
            Left            =   960
            TabIndex        =   6
            Top             =   1200
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   0
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin ServerEditor.NumOnlyText txtCha 
            Height          =   255
            Left            =   960
            TabIndex        =   7
            Top             =   1560
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            AllowNeg        =   0   'False
            Align           =   0
            MaxLength       =   0
            Enabled         =   -1  'True
            Backcolor       =   -2147483643
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Strength:"
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   32
            Top             =   120
            Width           =   690
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Agility:"
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   31
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Intellect:"
            Height          =   195
            Index           =   4
            Left            =   0
            TabIndex        =   30
            Top             =   480
            Width           =   645
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Charm:"
            Height          =   195
            Index           =   5
            Left            =   0
            TabIndex        =   29
            Top             =   1560
            Width           =   525
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Dexterity:"
            Height          =   195
            Index           =   6
            Left            =   0
            TabIndex        =   28
            Top             =   840
            Width           =   735
         End
      End
      Begin ServerEditor.NumOnlyText txtEXP 
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         AllowNeg        =   0   'False
         Align           =   0
         MaxLength       =   0
         Enabled         =   -1  'True
         Backcolor       =   -2147483643
      End
      Begin VB.TextBox txtID 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Text            =   " "
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   1
         Text            =   " "
         Top             =   480
         Width           =   2055
      End
      Begin ServerEditor.Raise Raise3 
         Height          =   2535
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4471
         Style           =   2
         Color           =   0
      End
      Begin ServerEditor.Raise Raise4 
         Height          =   1215
         Left            =   120
         TabIndex        =   37
         Top             =   3960
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   2143
         Style           =   2
         Color           =   0
      End
      Begin ServerEditor.Raise Raise5 
         Height          =   2175
         Left            =   120
         TabIndex        =   43
         Top             =   5280
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3836
         Style           =   2
         Color           =   0
      End
      Begin VB.Label Label2 
         Caption         =   $"frmRace.frx":0206
         Height          =   615
         Left            =   120
         TabIndex        =   44
         Top             =   7560
         Width           =   7095
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "EXP Cost:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   705
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   225
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Race Name:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   255
      Left            =   10200
      TabIndex        =   15
      Top             =   9240
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "< Previous"
      Height          =   255
      Left            =   9000
      TabIndex        =   16
      Top             =   9240
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "(new)"
      Height          =   255
      Left            =   7560
      TabIndex        =   17
      Top             =   9240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "(save)"
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   9240
      Width           =   1095
   End
   Begin ServerEditor.Raise Raise1 
      Height          =   8895
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   15690
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise2 
      Height          =   8895
      Left            =   3720
      TabIndex        =   19
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   15690
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise6 
      Height          =   495
      Left            =   6240
      TabIndex        =   45
      Top             =   9120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   873
      Style           =   2
      Color           =   0
   End
   Begin ServerEditor.Raise Raise7 
      Height          =   9735
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   17171
      Style           =   4
      Color           =   0
   End
End
Attribute VB_Name = "frmRace"
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
Rem***************                frmRace                         **********************
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
lstRaces.SetSelected lstRaces.FindInStr(txtFind.Text), True, True
End Sub

Private Sub cmdNew_Click()
Dim x As Long
Dim i As Long
Dim t As Boolean
MousePointer = vbHourglass
ReDim Preserve dbRaces(1 To UBound(dbRaces) + 1)
x = dbRaces(UBound(dbRaces) - 1).iID
x = x + 1
Do Until t = True
    t = True
    i = GetRaceID(, x)
    If i <> 0 Then
        t = False
        x = x + 1
    End If
Loop
With dbRaces(UBound(dbRaces))
    .iID = x
    .sName = "New Race"
    .dEXP = 0
    .iVision = 0
    .lMaxAge = 100
    .lStartAgeMin = 16
    .lStartAgeMax = 24
    .sHP = "10:30"
    .sStats = "4:4:4:4:4"
End With
lcID = UBound(dbRaces)
FillRace lcID, True
MousePointer = vbDefault
End Sub

Private Sub cmdNext_Click()
On Error GoTo cmdNext_Click_Error
SaveRace
lcID = lcID + 1
If lcID > UBound(dbRaces) Then lcID = 1
FillRace lcID
On Error GoTo 0
Exit Sub
cmdNext_Click_Error:
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo cmdPrevious_Click_Error
SaveRace
lcID = lcID - 1
If lcID < LBound(dbRaces) Then lcID = UBound(dbRaces)
FillRace lcID
On Error GoTo 0
Exit Sub
cmdPrevious_Click_Error:
End Sub

Private Sub cmdSave_Click()
SaveRace
End Sub

Private Sub Form_Load()
FillRace FillList:=True
End Sub

Private Sub lstRaces_Click()
If bIs Then Exit Sub
Dim i As Long
MousePointer = vbHourglass
For i = LBound(dbRaces) To UBound(dbRaces)
    With dbRaces(i)
        If .iID & " " & .sName = lstRaces.ItemText Then
            FillRace i
            Exit For
        End If
    End With
Next
MousePointer = vbDefault
End Sub

Sub SaveRace()
MousePointer = vbHourglass
modUpdateDatabase.ReverseEffects lcID, Race
With dbRaces(lcID)
    .dEXP = CDbl(txtEXP.Text)
    .iVision = CLng(Left$(cboVision.list(cboVision.ListIndex), InStr(1, cboVision.list(cboVision.ListIndex), " ")))
    .lMaxAge = CLng(txtMaxAge.Text)
    .lStartAgeMax = CLng(txtMaxAgeS.Text)
    .lStartAgeMin = CLng(txtMinAgeS.Text)
    .sHP = txtMinHP.Text & ":" & txtMaxHP.Text
    .sName = txtName.Text
    .sStats = txtStr.Text & ":" & txtAgil.Text & ":" & txtInt.Text & ":" & txtCha.Text & ":" & txtDex.Text
End With
modUpdateDatabase.DoEffects lcID, Race
modUpdateDatabase.SaveMemoryToDatabase Race
FillRace lcID, True
MousePointer = vbDefault
End Sub

Sub FillRace(Optional Arg As Long = -1, Optional FillList As Boolean = False)
Dim i As Long, j As Long
Dim m As Long
Dim Arr() As String
Dim lCol As Long
MousePointer = vbHourglass
bIs = True
If Arg = -1 Then Arg = LBound(dbRaces)
If FillList Then lstRaces.Clear
For i = LBound(dbRaces) To UBound(dbRaces)
    With dbRaces(i)
        If FillList Then lstRaces.AddItem CStr(.iID & " " & .sName)
        If i = Arg Then
            lcID = i
            txtID.Text = .iID
            txtName.Text = .sName
            txtEXP.Text = .dEXP
            Arr = Split(.sStats, ":")
            txtStr.Text = Arr(0)
            txtInt.Text = Arr(2)
            txtDex.Text = Arr(4)
            txtAgil.Text = Arr(1)
            txtCha.Text = Arr(3)
            Erase Arr
            Arr = Split(.sHP, ":")
            txtMinHP.Text = Arr(0)
            txtMaxHP.Text = Arr(1)
            txtMinAgeS.Text = .lStartAgeMin
            txtMaxAgeS.Text = .lStartAgeMax
            txtMaxAge.Text = .lMaxAge
            modMain.SetCBOlstIndex cboVision, .iVision, [Vision Level]
            If Not FillList Then Exit For
        End If
    End With
Next
modMain.SetLstSelected lstRaces, txtID.Text & " " & txtName.Text
bIs = False
MousePointer = vbDefault
End Sub


