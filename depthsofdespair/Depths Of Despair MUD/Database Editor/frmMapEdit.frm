VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMapEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Large Area Edit"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12075
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   584
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   Begin TabDlg.SSTab ssTab 
      Height          =   8175
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   14420
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "General Settings"
      TabPicture(0)   =   "frmMapEdit.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblMode"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtLight"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtDesc"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtBasic"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "optIndoor(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "optIndoor(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "optIndoor(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chkOVR"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkAuto"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtDeath"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtMax"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtMob"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdLoadArea"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Enhance Descriptions"
      TabPicture(1)   =   "frmMapEdit.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5(0)"
      Tab(1).Control(1)=   "Label5(1)"
      Tab(1).Control(2)=   "txtT(9)"
      Tab(1).Control(3)=   "txtT(8)"
      Tab(1).Control(4)=   "txtT(7)"
      Tab(1).Control(5)=   "txtT(6)"
      Tab(1).Control(6)=   "txtT(5)"
      Tab(1).Control(7)=   "txtT(4)"
      Tab(1).Control(8)=   "txtT(3)"
      Tab(1).Control(9)=   "txtT(2)"
      Tab(1).Control(10)=   "txtT(1)"
      Tab(1).Control(11)=   "txtT(0)"
      Tab(1).Control(12)=   "txtD(2)"
      Tab(1).Control(13)=   "txtD(1)"
      Tab(1).Control(14)=   "txtD(0)"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton cmdLoadArea 
         Caption         =   "Load Area"
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   7560
         Width           =   3495
      End
      Begin VB.TextBox txtD 
         Height          =   285
         Index           =   0
         Left            =   -74880
         TabIndex        =   35
         Text            =   "There is a small path to the"
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txtD 
         Height          =   285
         Index           =   1
         Left            =   -74880
         TabIndex        =   34
         Text            =   "There is an exit to the"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtD 
         Height          =   285
         Index           =   2
         Left            =   -74880
         TabIndex        =   33
         Text            =   "You can see a walkway to the"
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txtT 
         Height          =   285
         Index           =   0
         Left            =   -74880
         TabIndex        =   32
         Text            =   "Pathway"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txtT 
         Height          =   285
         Index           =   1
         Left            =   -74880
         TabIndex        =   31
         Text            =   "Small Path"
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox txtT 
         Height          =   285
         Index           =   2
         Left            =   -74880
         TabIndex        =   30
         Text            =   "3 Way Intersection"
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox txtT 
         Height          =   285
         Index           =   3
         Left            =   -74880
         TabIndex        =   29
         Text            =   "Fork"
         Top             =   3480
         Width           =   2655
      End
      Begin VB.TextBox txtT 
         Height          =   285
         Index           =   4
         Left            =   -74880
         TabIndex        =   28
         Text            =   "4 Way Intersection"
         Top             =   3840
         Width           =   2655
      End
      Begin VB.TextBox txtT 
         Height          =   285
         Index           =   5
         Left            =   -74880
         TabIndex        =   27
         Text            =   "Crossroads"
         Top             =   4200
         Width           =   2655
      End
      Begin VB.TextBox txtT 
         Height          =   285
         Index           =   6
         Left            =   -74880
         TabIndex        =   26
         Text            =   "5 Way Intersection"
         Top             =   4560
         Width           =   2655
      End
      Begin VB.TextBox txtT 
         Height          =   285
         Index           =   7
         Left            =   -74880
         TabIndex        =   25
         Text            =   "6 Way Intersection"
         Top             =   4920
         Width           =   2655
      End
      Begin VB.TextBox txtT 
         Height          =   285
         Index           =   8
         Left            =   -74880
         TabIndex        =   24
         Text            =   "7 Way Intersection"
         Top             =   5280
         Width           =   2655
      End
      Begin VB.TextBox txtT 
         Height          =   285
         Index           =   9
         Left            =   -74880
         TabIndex        =   23
         Text            =   "8 Way Intersection"
         Top             =   5640
         Width           =   2655
      End
      Begin ServerEditor.NumOnlyText txtMob 
         Height          =   375
         Left            =   1200
         TabIndex        =   21
         Top             =   5280
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
      Begin MSComctlLib.Slider txtMax 
         Height          =   375
         Left            =   1200
         TabIndex        =   20
         Top             =   5760
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Max             =   6
      End
      Begin ServerEditor.ucCombo txtDeath 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   6960
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   661
         SPACESAVER      =   1
      End
      Begin VB.CheckBox chkAuto 
         Caption         =   "&Auto Exits"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkOVR 
         Caption         =   "O&verwrite Previous"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton optIndoor 
         Caption         =   "&Outdoor"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   960
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optIndoor 
         Caption         =   "&Indoor"
         Height          =   375
         Index           =   1
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optIndoor 
         Caption         =   "&Underground"
         Height          =   375
         Index           =   2
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtBasic 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "Large Forest"
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   3360
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Mode:"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin MSComctlLib.Slider txtLight 
         Height          =   495
         Left            =   1200
         TabIndex        =   22
         Top             =   6240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Min             =   -200
         Max             =   200
         TickFrequency   =   25
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "For Title:"
         Height          =   195
         Index           =   1
         Left            =   -74880
         TabIndex        =   37
         Top             =   2160
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "For Exits:"
         Height          =   195
         Index           =   0
         Left            =   -74760
         TabIndex        =   36
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Basic Title:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Basic Description:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   3120
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Light:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   6240
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Max Regen:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   5760
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mob Group:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   5280
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Death Room:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   6720
         Width           =   945
      End
      Begin VB.Label lblMode 
         AutoSize        =   -1  'True
         Caption         =   "Exit Creation"
         Height          =   195
         Left            =   960
         TabIndex        =   10
         Top             =   600
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   10680
      TabIndex        =   2
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdRandom 
      Caption         =   "Make Random Area"
      Height          =   255
      Left            =   9720
      TabIndex        =   1
      Top             =   9840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   8130
      Left            =   3840
      ScaleHeight     =   540
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   540
      TabIndex        =   0
      Top             =   120
      Width           =   8130
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuMode 
         Caption         =   "&Mode"
         Begin VB.Menu mnuExit 
            Caption         =   "&Exit Creation"
            Checked         =   -1  'True
            Shortcut        =   ^E
         End
         Begin VB.Menu mnuDelete 
            Caption         =   "Exit &Deletion"
            Shortcut        =   ^D
         End
      End
   End
   Begin VB.Menu mnuRC 
      Caption         =   "RightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuiJoin 
         Caption         =   "&Join To Existing Room"
      End
   End
End
Attribute VB_Name = "frmMapEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long

Private Const DT_LEFT = &H0

Public prevSel As Long
Private prevHover As Long
Private lW As Long
Private lH As Long

Public Function RndNumber(Min As Long, Max As Long) As Long
'gets a random number from the min to the max
Randomize Timer
RndNumber = (Rnd * (Max - Min)) + Min
End Function

Private Sub init()
Dim i As Long
Dim k As Long
lW = 18
lH = 18
With udtMapArea(0)
    .Xl = 0
    .Yl = 0
    .lIndoor = 1
End With
For i = LBound(udtMapArea) + 1 To UBound(udtMapArea)
    With udtMapArea(i)
        .Xl = udtMapArea(i - 1).Xl + lW
        .Yl = udtMapArea(i - 1).Yl
        .lIndoor = 1
        If .Xl + lW > picMain.ScaleWidth Then
            k = k + 1
            .Xl = 0
            .Yl = lH * k
        End If
    End With
Next
prevSel = -1
Load frmMapDef
frmMapDef.Show
frmMapDef.Left = Me.Left + Me.Width
End Sub

Private Sub cmdLoadArea_Click()
Dim i As Long
Dim j As Long
Dim k As Long
Dim h As Long
Dim u As Long
Dim b As Boolean
Dim c As Boolean
Dim s As String
Dim m As Long
Dim St As Long
m = 1
k = Val(InputBox("Room number to center around.", "Room Number", "1"))
If k < 1 Then k = 1
If GetMapIndex(k) = 0 Then k = 1
St = k
j = 434
Erase udtMapArea
init
picMain.Cls
Do Until b
    If k = 0 Then Exit Do
    c = False
    For i = LBound(udtMapArea) To UBound(udtMapArea)
        If k = udtMapArea(i).lRealID Then
            c = True
            Exit For
        End If
    Next
    h = GetMapIndex(k)
    If Not c Then
        With udtMapArea(j)
            .sIsRoom = True
            .lAlreadyExist = 1
            .sTitle = dbMap(h).sRoomTitle
            .sDesc = dbMap(h).sRoomDesc
            .lLight = dbMap(h).lLight
            .lMob = dbMap(h).iMobGroup
            .lDeath = dbMap(h).lDeathRoom
            .lIndoor = dbMap(h).iInDoor
            .lMaxRegen = dbMap(h).iMaxRegen
            .lRealID = k
            .lAuto = 1
            'If h - 30 >= 0 Then
            .dN = dbMap(h).lNorth
            'If h + 30 <= UBound(udtMapArea) Then
            .dS = dbMap(h).lSouth
            'If h + 1 <= UBound(udtMapArea) And CheckBound(h, h + 1) = True Then
            .ddE = dbMap(h).lEast
            'If h - 1 >= 0 And CheckBound(h, h - 1) = True Then
            .dW = dbMap(h).lWest
            'If h - 29 >= 0 And CheckBound(h, h - 29) = True Then
            .dNE = dbMap(h).lNorthEast
            'If h - 31 >= 0 And CheckBound(h, h - 31) = True Then
            .dNW = dbMap(h).lNorthWest
            'If h + 29 <= UBound(udtMapArea) And CheckBound(h, h + 29) = True Then
            .dSW = dbMap(h).lSouthWest
            'If h + 31 <= UBound(udtMapArea) And CheckBound(h, h + 31) = True Then
            .dSE = dbMap(h).lSouthEast
        End With
    End If
    If m = 1 And j - 30 >= 0 Then
        c = False
        With udtMapArea(j)
            If .dN <> 0 And .cN = 0 Then
                k = .dN
                .cN = 1
                u = j - 30
                c = True
            End If
        End With
        If Not c Then
            For i = LBound(udtMapArea) To UBound(udtMapArea)
                With udtMapArea(i)
                    If .dN <> 0 And .cN = 0 And i - 30 >= 0 Then
                        k = .dN
                        .cN = 1
                        u = i - 30
                        c = True
                        Exit For
                    End If
                End With
            Next
            If Not c Then m = 2
        End If
    ElseIf m = 1 Then
        m = 2
    End If
    If m = 2 And j + 30 <= UBound(udtMapArea) Then
        c = False
        With udtMapArea(j)
            If .dS <> 0 And .cS = 0 Then
                k = .dS
                .cS = 1
                u = j + 30
                c = True
            End If
        End With
        If Not c Then
            For i = LBound(udtMapArea) To UBound(udtMapArea)
                With udtMapArea(i)
                    If .dS <> 0 And .cS = 0 And i + 30 <= UBound(udtMapArea) Then
                        k = .dS
                        .cS = 1
                        u = i + 30
                        c = True
                        Exit For
                    End If
                End With
            Next
            If Not c Then m = 3
        End If
    ElseIf m = 2 Then
        m = 3
    End If
    If m = 3 And j + 1 <= UBound(udtMapArea) And CheckBound(j, j + 1) = True Then
        c = False
        With udtMapArea(j)
            If .ddE <> 0 And .cE = 0 Then
                k = .ddE
                .cE = 1
                u = j + 1
                c = True
            End If
        End With
        If Not c Then
            For i = LBound(udtMapArea) To UBound(udtMapArea)
                With udtMapArea(i)
                    If .ddE <> 0 And .cE = 0 And i + 1 <= UBound(udtMapArea) And CheckBound(i, i + 1) = True Then
                        k = .ddE
                        .cE = 1
                        u = i + 1
                        c = True
                        Exit For
                    End If
                End With
            Next
            If Not c Then m = 4
        End If
    ElseIf m = 3 Then
        m = 4
    End If
    If m = 4 And j - 1 >= 0 And CheckBound(j, j - 1) = True Then
        c = False
        With udtMapArea(j)
            If .dW <> 0 And .cW = 0 Then
                k = .dW
                .cW = 1
                u = j - 1
                c = True
            End If
        End With
        If Not c Then
            For i = LBound(udtMapArea) To UBound(udtMapArea)
                With udtMapArea(i)
                    If .dW <> 0 And .cW = 0 And i - 1 >= 0 And CheckBound(i, i - 1) = True Then
                        k = .dW
                        .cW = 1
                        u = i - 1
                        c = True
                        Exit For
                    End If
                End With
            Next
            If Not c Then m = 5
        End If
    ElseIf m = 4 Then
        m = 5
    End If
    If m = 5 And j - 29 >= 0 And CheckBound(j, j - 29) = True Then
        c = False
        With udtMapArea(j)
            If .dNE <> 0 And .cNE = 0 Then
                k = .dNE
                .cNE = 1
                u = j - 29
                c = True
            End If
        End With
        If Not c Then
            For i = LBound(udtMapArea) To UBound(udtMapArea)
                With udtMapArea(i)
                    If .dNE <> 0 And .cNE = 0 And i - 29 >= 0 And CheckBound(i, i - 29) = True Then
                        k = .dNE
                        .cNE = 1
                        u = i - 29
                        c = True
                        Exit For
                    End If
                End With
            Next
            If Not c Then m = 6
        End If
    ElseIf m = 5 Then
        m = 6
    End If
    If m = 6 And j - 31 >= 0 And CheckBound(j, j - 31) = True Then
        c = False
        With udtMapArea(j)
            If .dNW <> 0 And .cNW = 0 Then
                k = .dNW
                .cNW = 1
                u = j - 31
                c = True
            End If
        End With
        If Not c Then
            For i = LBound(udtMapArea) To UBound(udtMapArea)
                With udtMapArea(i)
                    If .dNW <> 0 And .cNW = 0 And i - 31 >= 0 And CheckBound(i, i - 31) = True Then
                        k = .dNW
                        .cNW = 1
                        u = i - 31
                        c = True
                        Exit For
                    End If
                End With
            Next
            If Not c Then m = 7
        End If
    ElseIf m = 6 Then
        m = 7
    End If
    If m = 7 And j + 31 <= UBound(udtMapArea) And CheckBound(j, j + 31) = True Then
        c = False
        With udtMapArea(j)
            If .dSE <> 0 And .cSE = 0 Then
                k = .dSE
                .cSE = 1
                u = j + 31
                c = True
            End If
        End With
        If Not c Then
            For i = LBound(udtMapArea) To UBound(udtMapArea)
                With udtMapArea(i)
                    If .dSE <> 0 And .cSE = 0 And i + 31 <= UBound(udtMapArea) And CheckBound(i, i + 31) = True Then
                        k = .dSE
                        .cSE = 1
                        u = i + 31
                        c = True
                        Exit For
                    End If
                End With
            Next
            If Not c Then m = 8
        End If
    ElseIf m = 7 Then
        m = 8
    End If
    If m = 8 And j + 29 <= UBound(udtMapArea) And CheckBound(j, j + 29) = True Then
        c = False
        With udtMapArea(j)
            If .dSW <> 0 And .cSW = 0 Then
                k = .dSW
                .cSW = 1
                u = j + 29
                c = True
            End If
        End With
        If Not c Then
            For i = LBound(udtMapArea) To UBound(udtMapArea)
                With udtMapArea(i)
                    If .dSW <> 0 And .cSW = 0 And i + 29 <= UBound(udtMapArea) And CheckBound(i, i + 29) = True Then
                        k = .dSW
                        .cSW = 1
                        u = i + 29
                        c = True
                        Exit For
                    End If
                End With
            Next
            If Not c Then
                For i = LBound(udtMapArea) To UBound(udtMapArea)
                    With udtMapArea(i)
                        If .dN <> 0 And .cN = 0 And i - 30 >= 0 Then
                            k = .dN
                            .cN = 1
                            u = i - 30
                            c = True
                            m = 1
                            Exit For
                        End If
                        If .dS <> 0 And .cS = 0 And i + 30 <= UBound(udtMapArea) Then
                            k = .dS
                            .cS = 1
                            u = i + 30
                            c = True
                            m = 2
                            Exit For
                        End If
                        If .ddE <> 0 And .cE = 0 And i + 1 <= UBound(udtMapArea) And CheckBound(i, i + 1) = True Then
                            k = .ddE
                            .cE = 1
                            u = i + 1
                            c = True
                            m = 3
                            Exit For
                        End If
                        If .dW <> 0 And .cW = 0 And i - 1 >= 0 And CheckBound(i, i - 1) = True Then
                            k = .dW
                            .cW = 1
                            u = i - 1
                            c = True
                            m = 4
                            Exit For
                        End If
                        If .dNE <> 0 And .cNE = 0 And i - 29 >= 0 And CheckBound(i, i - 29) = True Then
                            k = .dNE
                            .cNE = 1
                            u = i - 29
                            c = True
                            m = 5
                            Exit For
                        End If
                        If .dNW <> 0 And .cNW = 0 And i - 31 >= 0 And CheckBound(i, i - 31) = True Then
                            k = .dNW
                            .cNW = 1
                            u = i - 31
                            c = True
                            m = 6
                            Exit For
                        End If
                        If .dSE <> 0 And .cSE = 0 And i + 31 <= UBound(udtMapArea) And CheckBound(i, i + 31) = True Then
                            k = .dSE
                            .cSE = 1
                            u = i + 31
                            c = True
                            m = 7
                            Exit For
                        End If
                        If .dSW <> 0 And .cSW = 0 And i + 29 <= UBound(udtMapArea) And CheckBound(i, i + 29) = True Then
                            k = .dSW
                            .cSW = 1
                            u = i + 29
                            c = True
                            m = 8
                            Exit For
                        End If
                    End With
                Next
                If Not c Then b = True
            End If
        End If
    ElseIf m = 8 Then
        c = False
        For i = LBound(udtMapArea) To UBound(udtMapArea)
            With udtMapArea(i)
                If .dN <> 0 And .cN = 0 And i - 30 >= 0 Then
                    k = .dN
                    .cN = 1
                    u = i - 30
                    c = True
                    m = 1
                    Exit For
                End If
                If .dS <> 0 And .cS = 0 And i + 30 <= UBound(udtMapArea) Then
                    k = .dS
                    .cS = 1
                    u = i + 30
                    c = True
                    m = 2
                    Exit For
                End If
                If .ddE <> 0 And .cE = 0 And i + 1 <= UBound(udtMapArea) And CheckBound(i, i + 1) = True Then
                    k = .ddE
                    .cE = 1
                    u = i + 1
                    c = True
                    m = 3
                    Exit For
                End If
                If .dW <> 0 And .cW = 0 And i - 1 >= 0 And CheckBound(i, i - 1) = True Then
                    k = .dW
                    .cW = 1
                    u = i - 1
                    c = True
                    m = 4
                    Exit For
                End If
                If .dNE <> 0 And .cNE = 0 And i - 29 >= 0 And CheckBound(i, i - 29) = True Then
                    k = .dNE
                    .cNE = 1
                    u = i - 29
                    c = True
                    m = 5
                    Exit For
                End If
                If .dNW <> 0 And .cNW = 0 And i - 31 >= 0 And CheckBound(i, i - 31) = True Then
                    k = .dNW
                    .cNW = 1
                    u = i - 31
                    c = True
                    m = 6
                    Exit For
                End If
                If .dSE <> 0 And .cSE = 0 And i + 31 <= UBound(udtMapArea) And CheckBound(i, i + 31) = True Then
                    k = .dSE
                    .cSE = 1
                    u = i + 31
                    c = True
                    m = 7
                    Exit For
                End If
                If .dSW <> 0 And .cSW = 0 And i + 29 <= UBound(udtMapArea) And CheckBound(i, i + 29) = True Then
                    k = .dSW
                    .cSW = 1
                    u = i + 29
                    c = True
                    m = 8
                    Exit For
                End If
            End With
        Next
        If Not c Then b = True
    End If
    j = u
Loop
For i = LBound(udtMapArea) To UBound(udtMapArea)
    With udtMapArea(i)
        If .dN <> 0 Then .sExits = .sExits & ":n;"
        If .dS <> 0 Then .sExits = .sExits & ":s;"
        If .ddE <> 0 Then .sExits = .sExits & ":e;"
        If .dW <> 0 Then .sExits = .sExits & ":w;"
        If .dNE <> 0 Then .sExits = .sExits & ":ne;"
        If .dNW <> 0 Then .sExits = .sExits & ":nw;"
        If .dSE <> 0 Then .sExits = .sExits & ":se;"
        If .dSW <> 0 Then .sExits = .sExits & ":sw;"
    End With
Next
ReDrawMap
End Sub

Private Sub cmdOK_Click()
Dim i As Long
Dim j As Long
Dim x As Long
Dim t As Boolean
Dim m As Long
Dim Arr() As String
Dim s As String
SaveMemoryToDatabase Map
x = dbMap(UBound(dbMap)).lRoomID
For i = LBound(udtMapArea) To UBound(udtMapArea)
    If udtMapArea(i).sIsRoom And udtMapArea(i).lAlreadyExist = 0 Then
        x = x + 1
        Do Until t = True
            t = True
            m = GetMapIndex(x)
            If m <> 0 Then
                t = False
                x = x + 1
            End If
        Loop
        udtMapArea(i).lRealID = x
    End If
Next
For i = LBound(udtMapArea) To UBound(udtMapArea)
    If udtMapArea(i).sIsRoom And udtMapArea(i).lAlreadyExist = 0 Then
        ReDim Preserve dbMap(1 To (UBound(dbMap) + 1))
        x = UBound(dbMap)
        SplitFast udtMapArea(i).sExits, Arr, ";"
        For j = LBound(Arr) To UBound(Arr)
            With dbMap(x)
                '+- 30
            '+- 1
            '-31nw
            '+31se
            '-29ne
            '+29sw
                Select Case Arr(j)
                    Case ":n"
                        .lNorth = udtMapArea(i - 30).lRealID
                    Case ":s"
                        .lSouth = udtMapArea(i + 30).lRealID
                    Case ":e"
                        .lEast = udtMapArea(i + 1).lRealID
                    Case ":w"
                        .lWest = udtMapArea(i - 1).lRealID
                    Case ":ne"
                        .lNorthEast = udtMapArea(i - 29).lRealID
                    Case ":se"
                        .lSouthEast = udtMapArea(i + 31).lRealID
                    Case ":nw"
                        .lNorthWest = udtMapArea(i - 31).lRealID
                    Case ":sw"
                        .lSouthWest = udtMapArea(i + 29).lRealID
                End Select
            End With
        Next
        If udtMapArea(i).lJoinRoom <> 0 Then
            With dbMap(x)
                Select Case udtMapArea(i).sJoinExit
                    Case "n"
                        .lNorth = udtMapArea(i).lJoinRoom
                    Case "s"
                        .lSouth = udtMapArea(i).lJoinRoom
                    Case "e"
                        .lEast = udtMapArea(i).lJoinRoom
                    Case "w"
                        .lWest = udtMapArea(i).lJoinRoom
                    Case "ne"
                        .lNorthEast = udtMapArea(i).lJoinRoom
                    Case "se"
                        .lSouthEast = udtMapArea(i).lJoinRoom
                    Case "nw"
                        .lNorthWest = udtMapArea(i).lJoinRoom
                    Case "sw"
                        .lSouthWest = udtMapArea(i).lJoinRoom
                End Select
            End With
            With dbMap(udtMapArea(i).lJoinRoom)
                Select Case udtMapArea(i).sJoinExit
                    Case "s"
                        .lNorth = x
                    Case "n"
                        .lSouth = x
                    Case "w"
                        .lEast = x
                    Case "e"
                        .lWest = x
                    Case "nw"
                        .lNorthWest = x
                    Case "sw"
                        .lSouthWest = x
                    Case "ne"
                        .lNorthEast = x
                    Case "se"
                        .lSouthEast = x
                End Select
            End With
        End If
        With dbMap(x)
            .sRoomDesc = udtMapArea(i).sDesc
            .sRoomTitle = udtMapArea(i).sTitle
            .iMaxRegen = udtMapArea(i).lMaxRegen
            .iMobGroup = udtMapArea(i).lMob
            .lRoomID = udtMapArea(i).lRealID
            .lLight = udtMapArea(i).lLight
            .sHidden = "0"
            .sItems = "0"
            .sMonsters = "0"
            .sScript = "0"
            .sShopItems = "0"
            .lDeathRoom = udtMapArea(i).lDeath
            .iInDoor = udtMapArea(i).lIndoor
        End With
    ElseIf udtMapArea(i).lAlreadyExist <> 0 Then
        SplitFast udtMapArea(i).sExits, Arr, ";"
        With dbMap(GetMapIndex(udtMapArea(i).lRealID))
            .lNorth = 0
            .lSouth = 0
            .lEast = 0
            .lWest = 0
            .lNorthEast = 0
            .lNorthWest = 0
            .lSouthWest = 0
            .lSouthEast = 0
            For j = LBound(Arr) To UBound(Arr)
                Select Case Arr(j)
                    Case ":n"
                        .lNorth = udtMapArea(i - 30).lRealID
                    Case ":s"
                        .lSouth = udtMapArea(i + 30).lRealID
                    Case ":e"
                        .lEast = udtMapArea(i + 1).lRealID
                    Case ":w"
                        .lWest = udtMapArea(i - 1).lRealID
                    Case ":ne"
                        .lNorthEast = udtMapArea(i - 29).lRealID
                    Case ":se"
                        .lSouthEast = udtMapArea(i + 31).lRealID
                    Case ":nw"
                        .lNorthWest = udtMapArea(i - 31).lRealID
                    Case ":sw"
                        .lSouthWest = udtMapArea(i + 29).lRealID
                End Select
                .sRoomDesc = udtMapArea(i).sDesc
                .sRoomTitle = udtMapArea(i).sTitle
                .iMaxRegen = udtMapArea(i).lMaxRegen
                .iMobGroup = udtMapArea(i).lMob
                .lRoomID = udtMapArea(i).lRealID
                .lLight = udtMapArea(i).lLight
                .sHidden = "0"
                .sItems = "0"
                .sMonsters = "0"
                .sScript = "0"
                .sShopItems = "0"
                .lDeathRoom = udtMapArea(i).lDeath
                .iInDoor = udtMapArea(i).lIndoor
            Next
        End With
    End If
Next
SaveMemoryToDatabase Map
MsgBox "Done."
End Sub

Private Sub cmdRandom_Click()
Dim lMaxRooms As Long
lMaxRooms = 350
Dim lCR As Long
Do Until lCR >= lMaxRooms
    lCR = lCR + 1
Loop
End Sub

Private Sub Command1_Click()
mnuExit_Click
End Sub

Private Sub Form_Load()
Me.Show
init
ReDrawMap
picMain.SetFocus
End Sub

Private Sub mnuDelete_Click()
Select Case mnuDelete.Checked
    Case True
        mnuExit.Checked = True
        mnuDelete.Checked = False
    Case Else
        mnuExit.Checked = False
        mnuDelete.Checked = True
End Select
ReDrawMap
End Sub

Private Sub mnuExit_Click()
Select Case mnuExit.Checked
    Case True
        mnuDelete.Checked = True
        mnuExit.Checked = False
    Case Else
        mnuDelete.Checked = False
        mnuExit.Checked = True
End Select
If mnuExit.Checked Then lblMode.Caption = "Exit Creation" Else lblMode.Caption = "Exit Deletion"
ReDrawMap
End Sub

Private Sub mnuiJoin_Click()
Load frmRoomJoin
frmRoomJoin.Show 1
End Sub

Private Sub picMain_KeyDown(KeyCode As Integer, Shift As Integer)
Dim R As RECT
Dim j As Long
Dim pt As POINTAPI
Dim f As Long
Select Case KeyCode
    Case vbKeyUp, vbKeyNumpad8
        If prevSel - 30 >= LBound(udtMapArea) Then MakeRm prevSel - 30, j, R, pt
    Case vbKeyDown, vbKeyNumpad2
        If prevSel + 30 <= UBound(udtMapArea) Then MakeRm prevSel + 30, j, R, pt
    Case vbKeyLeft, vbKeyNumpad4
        If prevSel - 1 >= LBound(udtMapArea) Then MakeRm prevSel - 1, j, R, pt
    Case vbKeyRight, vbKeyNumpad6
        If prevSel + 1 <= UBound(udtMapArea) Then MakeRm prevSel + 1, j, R, pt
    Case vbKeyNumpad9
        If prevSel - 29 >= LBound(udtMapArea) Then MakeRm prevSel - 29, j, R, pt
    Case vbKeyNumpad7
        If prevSel - 31 >= LBound(udtMapArea) Then MakeRm prevSel - 31, j, R, pt
    Case vbKeyNumpad1
        If prevSel + 29 <= UBound(udtMapArea) Then MakeRm prevSel + 29, j, R, pt
    Case vbKeyNumpad3
        If prevSel + 31 <= UBound(udtMapArea) Then MakeRm prevSel + 31, j, R, pt
End Select
End Sub

Private Sub EnhanceTitle(f As Long)
Dim Arr() As String
Dim s As String
Dim i As Long
Dim n As Long
Dim j As String
With udtMapArea(f)
    If .lAuto = 0 Then
        .lMaxRegen = txtMax.Value
        .lMob = Val(txtMob.Text)
        .lLight = txtLight.Value
        .lDeath = txtDeath.Number
        s = ReplaceFast(.sExits, ":", "")
        If s <> "" Then
            SplitFast s, Arr, ";"
            n = UBound(Arr) - 1
            .sTitle = txtBasic.Text & ", " & GetTitle(n)
            If .sTitle = txtBasic.Text & ", " Then .sTitle = txtBasic.Text
            .sDesc = txtDesc.Text & " " & GetDesc(Arr)
        End If
    End If
End With
If prevSel <> -1 Then
    With udtMapArea(prevSel)
        If .lAuto = 0 Then
            s = ReplaceFast(.sExits, ":", "")
            If s <> "" Then
                SplitFast s, Arr, ";"
                n = UBound(Arr) - 1
                .sTitle = txtBasic.Text & ", " & GetTitle(n)
                If .sTitle = txtBasic.Text & ", " Then .sTitle = txtBasic.Text
                .sDesc = txtDesc.Text & " " & GetDesc(Arr)
            End If
        End If
    End With
End If
End Sub

Private Function GetDesc(SE() As String) As String
Dim s As String
Dim i As Long
Select Case RndNumber(1, 3)
    Case 1
        s = txtD(0).Text
    Case 2
        s = txtD(1).Text
    Case 3
        s = txtD(2).Text  '"You can see a walkway to the"
End Select
For i = LBound(SE) To UBound(SE)
    Select Case SE(i)
        Case "n"
            s = s & " north,"
        Case "s"
            s = s & " south,"
        Case "e"
            s = s & " east,"
        Case "w"
            s = s & " west,"
        Case "ne"
            s = s & " northeast,"
        Case "nw"
            s = s & " northwest,"
        Case "se"
            s = s & " southeast,"
        Case "sw"
            s = s & " southwest,"
    End Select
Next
s = Left$(s, Len(s) - 1) & "."
GetDesc = s
End Function

Private Function GetTitle(n As Long) As String
Select Case n
    Case 0
        GetTitle = "Dead End"
    Case 1
        Select Case RndNumber(1, 3)
            Case 1
                GetTitle = ""
            Case 2
                GetTitle = txtT(0).Text
            Case 3
                GetTitle = txtT(1).Text
        End Select
    Case 2
        Select Case RndNumber(1, 4)
            Case 1
                GetTitle = ""
            Case 2
                GetTitle = txtT(2).Text
            Case 3
                GetTitle = "Intersection"
            Case 4
                GetTitle = txtT(3).Text
        End Select
    Case 3
        Select Case RndNumber(1, 5)
            Case 1
                GetTitle = ""
            Case 2
                GetTitle = txtT(4).Text
            Case 3
                GetTitle = "Intersection"
            Case 4
                GetTitle = txtT(5).Text
            Case 5
                GetTitle = "Crossway"
        End Select
    Case 4
        Select Case RndNumber(1, 5)
            Case 1
                GetTitle = ""
            Case 2
                GetTitle = txtT(6).Text
            Case 3
                GetTitle = "Intersection"
            Case 4
                GetTitle = txtT(5).Text
            Case 5
                GetTitle = "Crossway"
        End Select
    Case 5
        Select Case RndNumber(1, 5)
            Case 1
                GetTitle = ""
            Case 2
                GetTitle = txtT(7).Text
            Case 3
                GetTitle = "Intersection"
            Case 4
                GetTitle = txtT(5).Text
            Case 5
                GetTitle = "Crossway"
        End Select
    Case 6
        Select Case RndNumber(1, 5)
            Case 1
                GetTitle = ""
            Case 2
                GetTitle = txtT(8).Text
            Case 3
                GetTitle = "Intersection"
            Case 4
                GetTitle = txtT(5).Text
            Case 5
                GetTitle = "Crossway"
        End Select
    Case 7
        Select Case RndNumber(1, 5)
            Case 1
                GetTitle = ""
            Case 2
                GetTitle = txtT(9).Text
            Case 3
                GetTitle = "Intersection"
            Case 4
                GetTitle = txtT(5).Text
            Case 5
                GetTitle = "Crossway"
        End Select
End Select
End Function

Private Sub picMain_LostFocus()
'picMain.SetFocus
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim R As RECT
Dim i As Long
Dim j As Long
Dim pt As POINTAPI
Dim f As Long
Dim Arr() As String
Dim s As String
For i = LBound(udtMapArea) To UBound(udtMapArea)
    With udtMapArea(i)
        If x >= .Xl And x <= .Xl + lW Then
            If y >= .Yl And y <= .Yl + lH Then
                If mnuExit.Checked Then
                    If .sIsRoom = False Or chkOVR.Value = 1 Then
                        If optIndoor(0).Value Then .lIndoor = 0
                        If optIndoor(1).Value Then .lIndoor = 1
                        If optIndoor(2).Value Then .lIndoor = 2
                    Else
                        prevSel = i
                        frmMapDef.txtTitle.Text = .sTitle
                        frmMapDef.txtDesc.Text = .sDesc
                        frmMapDef.txtLight.Text = .lLight
                        frmMapDef.txtMob.Text = .lMob
                        frmMapDef.txtMax.Text = .lMaxRegen
                        frmMapDef.txtDeath.Text = .lDeath
                        frmMapDef.Check1.Value = .lAuto
                        frmMapDef.optIndoor(.lIndoor).Value = True
                    End If
                    .sIsRoom = True
                End If
                If .sIsRoom Then f = i Else f = -1
                
                Exit For
            End If
        End If
    End With
Next
If Button = vbRightButton Then
    With udtMapArea(f)
        .sIsRoom = False
        .lIndoor = 1
        If .sExits <> "" Then
            s = ReplaceFast(.sExits, ":", "")
            SplitFast s, Arr, ";"
            '+- 30
            '+- 1
            '-31nw
            '+31se
            '-29ne
            '+29sw
            For i = LBound(Arr) To UBound(Arr)
                Select Case Arr(i)
                    Case "n"
                        With udtMapArea(f - 30)
                            .sExits = ReplaceFast(.sExits, ":s;", "")
                        End With
                        prevSel = f - 30
                        EnhanceTitle f
                    Case "s"
                        With udtMapArea(f + 30)
                            .sExits = ReplaceFast(.sExits, ":n;", "")
                        End With
                        prevSel = f + 30
                        EnhanceTitle f
                    Case "e"
                        With udtMapArea(f + 1)
                            .sExits = ReplaceFast(.sExits, ":w;", "")
                        End With
                        prevSel = f + 1
                        EnhanceTitle f
                    Case "w"
                        With udtMapArea(f - 1)
                            .sExits = ReplaceFast(.sExits, ":e;", "")
                        End With
                        prevSel = f - 1
                        EnhanceTitle f
                    Case "nw"
                        With udtMapArea(f - 31)
                            .sExits = ReplaceFast(.sExits, ":se;", "")
                        End With
                        prevSel = f - 31
                        EnhanceTitle f
                    Case "ne"
                        With udtMapArea(f - 29)
                            .sExits = ReplaceFast(.sExits, ":sw;", "")
                        End With
                        prevSel = f - 29
                        EnhanceTitle f
                    Case "sw"
                        With udtMapArea(f + 29)
                            .sExits = ReplaceFast(.sExits, ":ne;", "")
                        End With
                        prevSel = f + 29
                        EnhanceTitle f
                    Case "se"
                        With udtMapArea(f + 31)
                            .sExits = ReplaceFast(.sExits, ":nw;", "")
                        End With
                        prevSel = f + 31
                        EnhanceTitle f
                End Select
            Next
        End If
        .sExits = ""
        ReDrawMap
    End With
Else
    If mnuExit.Checked Then
        MakeRm f, j, R, pt
    Else
        If prevSel <> -1 And f <> -1 Then
            DrawRm prevSel
            If prevSel + 30 = f And udtMapArea(f).Yl + lH <= picMain.ScaleHeight Then
                'south
                With udtMapArea(prevSel)
                    .sExits = ReplaceFast(.sExits, ":s;", "")
                End With
                With udtMapArea(f)
                    .sExits = ReplaceFast(.sExits, ":n;", "")
                End With
            ElseIf prevSel - 30 = f And udtMapArea(f).Yl >= 0 Then
                'north
                With udtMapArea(prevSel)
                    .sExits = ReplaceFast(.sExits, ":n;", "")
                End With
                With udtMapArea(f)
                    .sExits = ReplaceFast(.sExits, ":s;", "")
                End With
            ElseIf prevSel - 1 = f And udtMapArea(f).Xl >= 0 Then
                'west
                With udtMapArea(prevSel)
                    .sExits = ReplaceFast(.sExits, ":w;", "")
                End With
                With udtMapArea(f)
                    .sExits = ReplaceFast(.sExits, ":e;", "")
                End With
            ElseIf prevSel + 1 = f And udtMapArea(f).Xl + lW <= picMain.ScaleWidth Then
                'east
                With udtMapArea(prevSel)
                    .sExits = ReplaceFast(.sExits, ":e;", "")
                End With
                With udtMapArea(f)
                    .sExits = ReplaceFast(.sExits, ":w;", "")
                End With
            ElseIf prevSel - 31 = f And udtMapArea(f).Xl >= 0 Then
                'nw
                With udtMapArea(prevSel)
                    .sExits = ReplaceFast(.sExits, ":nw;", "")
                End With
                With udtMapArea(f)
                    .sExits = ReplaceFast(.sExits, ":se;", "")
                End With
            ElseIf prevSel + 31 = f And udtMapArea(f).Yl <= picMain.ScaleHeight Then
                'se
                With udtMapArea(prevSel)
                    .sExits = ReplaceFast(.sExits, ":se;", "")
                End With
                With udtMapArea(f)
                    .sExits = ReplaceFast(.sExits, ":nw;", "")
                End With
            ElseIf prevSel - 29 = f And udtMapArea(f).Yl > -1 Then
                'ne
                With udtMapArea(prevSel)
                    .sExits = ReplaceFast(.sExits, ":ne;", "")
                End With
                With udtMapArea(f)
                    .sExits = ReplaceFast(.sExits, ":sw;", "")
                End With
            ElseIf prevSel + 29 = f And udtMapArea(f).Xl > -1 Then
                'sw
                With udtMapArea(prevSel)
                    .sExits = ReplaceFast(.sExits, ":sw;", "")
                End With
                With udtMapArea(f)
                    .sExits = ReplaceFast(.sExits, ":ne;", "")
                End With
            End If
        End If
        EnhanceTitle f
        If f <> -1 Then prevSel = f
        ReDrawMap
    End If
End If
End Sub

Public Sub DrawRm(i As Long)
Dim R As RECT
Dim pt As POINTAPI
Dim j As Long
If i = -1 Then Exit Sub
With udtMapArea(i)
    R.Left = .Xl + 3
    R.Top = .Yl + 3
    R.Right = .Xl + lW - 3
    R.Bottom = .Yl + lH - 3
    If i <> prevSel Then
        If .lAlreadyExist = 0 Then
            j = CreateSolidBrush(vbBlack)
        Else
            j = CreateSolidBrush(vbRed)
        End If
    Else
        j = CreateSolidBrush(vbGreen)
    End If
    FillRect picMain.hdc, R, j
    DeleteObject j
    'If .lAlreadyExist = 1 Then Exit Sub
    If .lIndoor = 1 Then
        R.Left = .Xl + 6
        R.Top = .Yl + 6
        R.Right = .Xl + lW - 6
        R.Bottom = .Yl + lH - 6
        j = CreateSolidBrush(vbWhite)
        FillRect picMain.hdc, R, j
        DeleteObject j
        R.Left = .Xl + 8
        R.Top = .Yl + 8
        R.Right = .Xl + lW - 8
        R.Bottom = .Yl + lH - 8
        j = CreateSolidBrush(vbBlack)
        FillRect picMain.hdc, R, j
        DeleteObject j
    ElseIf .lIndoor = 2 Then
        picMain.ForeColor = vbWhite
        MoveToEx picMain.hdc, .Xl + 3, .Yl + 9, pt
        LineTo picMain.hdc, .Xl + 6, .Yl + 6
        MoveToEx picMain.hdc, .Xl + 6, .Yl + 6, pt
        LineTo picMain.hdc, .Xl + 15, .Yl + 12
        picMain.ForeColor = vbBlack
    End If
'    If .lJoinRoom <> 0 Then
'        r.Left = .Xl + 3
'        r.Top = .Yl + 3
'        r.Right = .Xl + lW - 3
'        r.Bottom = .Yl + lH - 3
'        j = CreateSolidBrush(vbRed)
'        FillRect picMain.hdc, r, j
'        DeleteObject j
'    End If
End With
End Sub

Private Function CheckBound(curID As Long, MoveID As Long) As Boolean
Dim i As Long
CheckBound = True
i = curID
Do Until i <= 0
    i = i - 30
Loop
If i = -1 Then
    If MoveID - 1 = curID Then
        CheckBound = False
    ElseIf MoveID + 29 = curID Then
        CheckBound = False
    ElseIf MoveID - 31 = curID Then
        CheckBound = False
    End If
ElseIf i = 0 Then
    If MoveID + 1 = curID Then
        CheckBound = False
    ElseIf MoveID - 29 = curID Then
        CheckBound = False
    ElseIf MoveID + 31 = curID Then
        CheckBound = False
    End If
'Else
'    CheckBound = False
End If

End Function

Private Sub MakeRm(f As Long, j As Long, R As RECT, pt As POINTAPI)
Dim i As Long
Dim m As Long
If prevSel <> -1 Then
    i = prevSel
    Do Until i <= 0
        i = i - 30
    Loop
    If i = -1 Then
        If f - 1 = prevSel Then
            Exit Sub
        ElseIf f + 29 = prevSel Then
            Exit Sub
        ElseIf f - 31 = prevSel Then
            Exit Sub
        End If
    End If
    If i = 0 Then
        If f + 1 = prevSel Then
            Exit Sub
        ElseIf f - 29 = prevSel Then
            Exit Sub
        ElseIf f + 31 = prevSel Then
            Exit Sub
        End If
    End If
    DrawRm prevSel
End If
SetRectEmpty R
With udtMapArea(f)
    'If .lAlreadyExist = 1 Then Exit Sub
    R.Top = .Yl + 3
    R.Left = .Xl + 3
    R.Right = .Xl + lW - 3
    R.Bottom = .Yl + lH - 3
    If .sIsRoom = False Or chkOVR.Value = 1 Then
        If optIndoor(0).Value Then .lIndoor = 0
        If optIndoor(1).Value Then .lIndoor = 1
        If optIndoor(2).Value Then .lIndoor = 2
    End If
    .sIsRoom = True
End With
'+- 30
'+- 1
'-31nw
'+31se
'-29ne
'+29sw
DrawRm f
With udtMapArea(f)
    R.Left = .Xl + 3
    R.Top = .Yl + 3
    R.Right = .Xl + lW - 3
    R.Bottom = .Yl + lH - 3
End With
j = CreateSolidBrush(vbGreen)
FillRect picMain.hdc, R, j
DeleteObject j
If chkAuto.Value Then
    SetRectEmpty R
    If prevSel = -1 Then GoTo Done
    If prevSel + 30 = f And udtMapArea(f).Yl + lH <= picMain.ScaleHeight Then
        'south
        With udtMapArea(prevSel)
            .sExits = ReplaceFast(.sExits, ":s;", "")
            .sExits = .sExits & ":s;"
            R.Left = .Xl + 8
            R.Top = .Yl + 15
            R.Right = R.Left + 2
        End With
        With udtMapArea(f)
            .sExits = ReplaceFast(.sExits, ":n;", "")
            .sExits = .sExits & ":n;"
            R.Bottom = .Yl + 3
        End With
        j = CreateSolidBrush(vbBlack)
        FillRect picMain.hdc, R, j
        DeleteObject j
    ElseIf prevSel - 30 = f And udtMapArea(f).Yl >= 0 Then
        'north
        With udtMapArea(prevSel)
            .sExits = ReplaceFast(.sExits, ":n;", "")
            .sExits = .sExits & ":n;"
            R.Bottom = .Yl + 3
        End With
        With udtMapArea(f)
            .sExits = ReplaceFast(.sExits, ":s;", "")
            .sExits = .sExits & ":s;"
            R.Left = .Xl + 8
            R.Top = .Yl + 15
            R.Right = R.Left + 2
        End With
        j = CreateSolidBrush(vbBlack)
        FillRect picMain.hdc, R, j
        DeleteObject j
    ElseIf prevSel - 1 = f And udtMapArea(f).Xl >= 0 Then
        'west
        With udtMapArea(prevSel)
            .sExits = ReplaceFast(.sExits, ":w;", "")
            .sExits = .sExits & ":w;"
            R.Right = .Xl + 3
        End With
        With udtMapArea(f)
            .sExits = ReplaceFast(.sExits, ":e;", "")
            .sExits = .sExits & ":e;"
            R.Left = .Xl + 15
            R.Top = .Yl + 8
            R.Bottom = R.Top + 2
        End With
        j = CreateSolidBrush(vbBlack)
        FillRect picMain.hdc, R, j
        DeleteObject j
    ElseIf prevSel + 1 = f And udtMapArea(f).Xl + lW <= picMain.ScaleWidth Then
        'east
        With udtMapArea(prevSel)
            .sExits = ReplaceFast(.sExits, ":e;", "")
            .sExits = .sExits & ":e;"
            R.Left = .Xl + 15
            R.Top = .Yl + 8
            R.Bottom = R.Top + 2
        End With
        With udtMapArea(f)
            .sExits = ReplaceFast(.sExits, ":w;", "")
            .sExits = .sExits & ":w;"
            R.Right = .Xl + 3
        End With
        j = CreateSolidBrush(vbBlack)
        FillRect picMain.hdc, R, j
        DeleteObject j
    ElseIf prevSel - 31 = f And udtMapArea(f).Xl >= 0 Then
        'nw
        With udtMapArea(prevSel)
            .sExits = ReplaceFast(.sExits, ":nw;", "")
            .sExits = .sExits & ":nw;"
            MoveToEx picMain.hdc, .Xl + 3, .Yl + 3, pt
        End With
        With udtMapArea(f)
            .sExits = ReplaceFast(.sExits, ":se;", "")
            .sExits = .sExits & ":se;"
            LineTo picMain.hdc, .Xl + 14, .Yl + 14
        End With
    ElseIf prevSel + 31 = f And udtMapArea(f).Yl <= picMain.ScaleHeight Then
        'se
        With udtMapArea(prevSel)
            .sExits = ReplaceFast(.sExits, ":se;", "")
            .sExits = .sExits & ":se;"
            MoveToEx picMain.hdc, .Xl + 14, .Yl + 14, pt
        End With
        With udtMapArea(f)
            .sExits = ReplaceFast(.sExits, ":nw;", "")
            .sExits = .sExits & ":nw;"
            LineTo picMain.hdc, .Xl + 3, .Yl + 3
        End With
    ElseIf prevSel - 29 = f And udtMapArea(f).Yl > -1 Then
        'ne
        With udtMapArea(prevSel)
            .sExits = ReplaceFast(.sExits, ":ne;", "")
            .sExits = .sExits & ":ne;"
            MoveToEx picMain.hdc, .Xl + 14, .Yl + 3, pt
        End With
        With udtMapArea(f)
            .sExits = ReplaceFast(.sExits, ":sw;", "")
            .sExits = .sExits & ":sw;"
            LineTo picMain.hdc, .Xl + 3, .Yl + 14
        End With
    ElseIf prevSel + 29 = f And udtMapArea(f).Xl > -1 Then
        'sw
        With udtMapArea(prevSel)
            .sExits = ReplaceFast(.sExits, ":sw;", "")
            .sExits = .sExits & ":sw;"
            MoveToEx picMain.hdc, .Xl + 3, .Yl + 14, pt
        End With
        With udtMapArea(f)
            .sExits = ReplaceFast(.sExits, ":ne;", "")
            .sExits = .sExits & ":ne;"
            LineTo picMain.hdc, .Xl + 14, .Yl + 3
        End With
    End If
End If
Done:
EnhanceTitle f
m = prevSel
prevSel = f
DrawRm m
frmMapDef.txtTitle.Text = udtMapArea(f).sTitle
frmMapDef.txtDesc.Text = udtMapArea(f).sDesc
frmMapDef.txtLight.Text = udtMapArea(f).lLight
frmMapDef.txtMob.Text = udtMapArea(f).lMob
frmMapDef.txtMax.Text = udtMapArea(f).lMaxRegen
frmMapDef.txtDeath.Text = udtMapArea(f).lDeath
frmMapDef.Check1.Value = udtMapArea(f).lAuto
frmMapDef.optIndoor(udtMapArea(f).lIndoor).Value = True
picMain.Refresh
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
Dim R As RECT
Dim f As Long
Dim j As Long
Dim pt As POINTAPI
For i = LBound(udtMapArea) To UBound(udtMapArea)
    With udtMapArea(i)
        If x >= .Xl And x <= .Xl + lW Then
            If y >= .Yl And y <= .Yl + lH Then
                f = i
                Exit For
            End If
        End If
    End With
    DoEvents
Next
With udtMapArea(prevHover)
    R.Top = .Yl + 3
    R.Left = .Xl + 3
    R.Right = .Xl + lW - 3
    R.Bottom = .Yl + lH - 3
    If prevHover = prevSel Then
        If mnuExit.Checked Then
            j = CreateSolidBrush(vbGreen)
            FillRect picMain.hdc, R, j
            DeleteObject j
        Else
            j = CreateSolidBrush(vbBlue)
            FillRect picMain.hdc, R, j
            DeleteObject j
        End If
    ElseIf .sIsRoom Then
'        If .lAlreadyExist = 0 Then
'            j = CreateSolidBrush(vbBlack)
'        Else
'            j = CreateSolidBrush(vbRed)
'        End If
        DrawRm prevHover
'        FillRect picMain.hdc, R, j
'        DeleteObject j
    Else
        j = CreateSolidBrush(vbWhite)
        FillRect picMain.hdc, R, j
        DeleteObject j
    End If
End With
With udtMapArea(f)
    If .sIsRoom Then
        DrawRm f
    End If
    R.Top = .Yl + 3
    R.Left = .Xl + 3
    R.Right = .Xl + lW - 3
    R.Bottom = .Yl + lH - 3
    If udtMapArea(prevHover).sIsRoom Then
        DrawRm prevHover
    End If
    If .sIsRoom Then
        j = CreateSolidBrush(vbYellow)
        FillRect picMain.hdc, R, j
        DeleteObject j
    Else
        j = CreateSolidBrush(&HC000C0)
        FillRect picMain.hdc, R, j
        DeleteObject j
    End If
End With
prevHover = f
picMain.Refresh
End Sub

Private Sub ReDrawMap()
Dim i As Long
Dim j As Long
Dim pt As POINTAPI
Dim R As RECT
Dim Arr() As String
Dim s As String
Dim n As Long
picMain.Cls
For i = LBound(udtMapArea) To UBound(udtMapArea)
    With udtMapArea(i)
        
        If .sIsRoom Then
            If .sExits <> "" Then
                s = ReplaceFast(.sExits, ":", "")
                SplitFast s, Arr, ";"
                '+- 30
                '+- 1
                '-31nw
                '+31se
                '-29ne
                '+29sw
                If .lAlreadyExist = 1 Or .lJoinRoom <> 0 Then picMain.ForeColor = vbRed Else picMain.ForeColor = vbBlack
                For n = LBound(Arr) To UBound(Arr)
                    SetRectEmpty R
                    Select Case Arr(n)
                        Case "n"
                            If i - 30 >= 0 Then
                                With udtMapArea(i - 30)
                                    R.Bottom = .Yl + 3
                                End With
                                R.Left = .Xl + 8
                                R.Top = .Yl + 15
                                R.Right = R.Left + 2
                            End If
                        Case "s"
                            If i + 30 <= UBound(udtMapArea) Then
                                With udtMapArea(i + 30)
                                    R.Left = .Xl + 8
                                    R.Top = .Yl + 15
                                    R.Right = R.Left + 2
                                End With
                                R.Bottom = .Yl + 3
                            End If
                        Case "e"
                            If i + 1 <= UBound(udtMapArea) Then
                                With udtMapArea(i + 1)
                                    R.Left = .Xl + 15
                                    R.Top = .Yl + 8
                                    R.Bottom = R.Top + 2
                                End With
                                R.Right = .Xl + 3
                            End If
                        Case "w"
                            If i - 1 >= 0 Then
                                With udtMapArea(i - 1)
                                    R.Right = .Xl + 3
                                End With
                                R.Left = .Xl + 15
                                R.Top = .Yl + 8
                                R.Bottom = R.Top + 2
                            End If
                        Case "nw"
                            If i - 31 >= 0 Then
                                With udtMapArea(i - 31)
                                    MoveToEx picMain.hdc, .Xl + 3, .Yl + 3, pt
                                End With
                                LineTo picMain.hdc, .Xl + 14, .Yl + 14
                            End If
                        Case "ne"
                            If i - 29 >= 0 Then
                                With udtMapArea(i - 29)
                                    MoveToEx picMain.hdc, .Xl + 14, .Yl + 3, pt
                                End With
                                LineTo picMain.hdc, .Xl + 3, .Yl + 14
                            End If
                        Case "sw"
                            If i + 29 <= UBound(udtMapArea) Then
                                With udtMapArea(i + 29)
                                    MoveToEx picMain.hdc, .Xl + 3, .Yl + 14, pt
                                End With
                                LineTo picMain.hdc, .Xl + 14, .Yl + 3
                            End If
                        Case "se"
                            If i + 29 <= UBound(udtMapArea) Then
                                With udtMapArea(i + 31)
                                    MoveToEx picMain.hdc, .Xl + 14, .Yl + 14, pt
                                End With
                                LineTo picMain.hdc, .Xl + 3, .Yl + 3
                            End If
                    End Select
                    If .lAlreadyExist = 1 Then
                        j = CreateSolidBrush(vbRed)
                    Else
                        j = CreateSolidBrush(vbBlack)
                    End If
                    FillRect picMain.hdc, R, j
                    DeleteObject j
                Next
            End If
        End If
    End With
Next
For i = LBound(udtMapArea) To UBound(udtMapArea)
    With udtMapArea(i)
        If .sIsRoom Then
            SetRectEmpty R
            R.Left = .Xl + 3
            R.Top = .Yl + 3
            R.Right = .Xl + lW - 3
            R.Bottom = .Yl + lH - 3
            If prevSel = i Then
                If mnuExit.Checked Then
                    j = CreateSolidBrush(vbGreen)
                    FillRect picMain.hdc, R, j
                    DeleteObject j
                Else
                    j = CreateSolidBrush(vbRed)
                    FillRect picMain.hdc, R, j
                    DeleteObject j
                End If
            Else
                If .lAlreadyExist = 0 Then
                    j = CreateSolidBrush(vbBlack)
                Else
                    j = CreateSolidBrush(vbRed)
                End If
                FillRect picMain.hdc, R, j
                DeleteObject j
            End If
            'If .lAlreadyExist = 1 Then GoTo nNext
            If .lIndoor = 1 Then
                R.Left = .Xl + 6
                R.Top = .Yl + 6
                R.Right = .Xl + lW - 6
                R.Bottom = .Yl + lH - 6
                j = CreateSolidBrush(vbWhite)
                FillRect picMain.hdc, R, j
                DeleteObject j
                R.Left = .Xl + 8
                R.Top = .Yl + 8
                R.Right = .Xl + lW - 8
                R.Bottom = .Yl + lH - 8
                j = CreateSolidBrush(vbBlack)
                FillRect picMain.hdc, R, j
                DeleteObject j
            ElseIf .lIndoor = 2 Then
                picMain.ForeColor = vbWhite
                MoveToEx picMain.hdc, .Xl + 3, .Yl + 9, pt
                LineTo picMain.hdc, .Xl + 6, .Yl + 6
                MoveToEx picMain.hdc, .Xl + 6, .Yl + 6, pt
                LineTo picMain.hdc, .Xl + 15, .Yl + 12
                picMain.ForeColor = vbBlack
            End If
        End If
    End With
nNext:
Next
picMain.Refresh
End Sub

Public Sub DrawExit(lNewSel As Long, sDir As String)
Dim R As RECT
Dim pt As POINTAPI
Dim j As Long
Dim i As Long
i = lNewSel
picMain.ForeColor = vbRed
With udtMapArea(prevSel)
    Select Case sDir
         
        Case "n"
            With udtMapArea(i)
                R.Bottom = .Yl + 3
            End With
            R.Left = .Xl + 8
            R.Top = .Yl + 15
            R.Right = R.Left + 2
        Case "s"
            With udtMapArea(i)
                R.Left = .Xl + 8
                R.Top = .Yl + 15
                R.Right = R.Left + 2
            End With
            R.Bottom = .Yl + 3
        Case "e"
            With udtMapArea(i)
                R.Left = .Xl + 15
                R.Top = .Yl + 8
                R.Bottom = R.Top + 2
            End With
            R.Right = .Xl + 3
        Case "w"
            With udtMapArea(i)
                R.Right = .Xl + 3
            End With
            R.Left = .Xl + 15
            R.Top = .Yl + 8
            R.Bottom = R.Top + 2
        Case "nw"
            With udtMapArea(i)
                MoveToEx picMain.hdc, .Xl + 3, .Yl + 3, pt
            End With
            LineTo picMain.hdc, .Xl + 14, .Yl + 14
        Case "ne"
            With udtMapArea(i)
                MoveToEx picMain.hdc, .Xl + 14, .Yl + 3, pt
            End With
            LineTo picMain.hdc, .Xl + 3, .Yl + 14
        Case "sw"
            With udtMapArea(i)
                MoveToEx picMain.hdc, .Xl + 3, .Yl + 14, pt
            End With
            LineTo picMain.hdc, .Xl + 14, .Yl + 3
        Case "se"
            With udtMapArea(i)
                MoveToEx picMain.hdc, .Xl + 14, .Yl + 14, pt
            End With
            LineTo picMain.hdc, .Xl + 3, .Yl + 3
    End Select
End With
j = CreateSolidBrush(vbRed)
FillRect picMain.hdc, R, j
DeleteObject j
picMain.ForeColor = vbBlack
DrawRm prevSel
End Sub
