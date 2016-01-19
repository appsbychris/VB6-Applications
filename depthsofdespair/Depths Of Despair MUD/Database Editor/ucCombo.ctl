VERSION 5.00
Begin VB.UserControl ucCombo 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3525
   ScaleHeight     =   375
   ScaleWidth      =   3525
   Begin VB.ComboBox cboMain 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin ServerEditor.NumOnlyText txtMain 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
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
      MaxLength       =   5
      Enabled         =   -1  'True
      Backcolor       =   -2147483643
   End
End
Attribute VB_Name = "ucCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum ComboSetting
    Rooms = 0
    Items = 1
    Classes = 2
    Monsters = 3
    Shops = 4
    Environment = 5
    RoomType = 6
    Keys = 7
    Alignment = 8
End Enum
Public Enum SizeMode
    Normal = 0
    Small = 1
End Enum
Public lS As ComboSetting
Public Event Change()
Dim b As SizeMode


Private Sub cboMain_Change()
cboMain_Click
End Sub

Private Sub cboMain_Click()
s = cboMain.list(cboMain.ListIndex)
txtMain.Text = Left$(s, InStr(1, s, ")"))
RaiseEvent Change
End Sub

Private Sub txtMain_Change()
Dim i As Long
Dim s As String
Dim j As Long
For i = 0 To cboMain.ListCount - 1
    With cboMain
        s = .list(i)
        s = Mid$(s, 2, InStr(1, s, ")"))
        j = CLng(Val(s))
        If CLng(Val(txtMain.Text)) = j Then
            .ListIndex = i
            Exit For
        End If
    End With
    DoEvents
Next
RaiseEvent Change
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
b = PropBag.ReadProperty("SPACESAVER")
End Sub

Private Sub UserControl_Resize()
If b = Normal Then
    UserControl.Width = 4560
Else
    UserControl.Width = 3525
End If
UserControl.Height = 375
End Sub

Public Property Get ListSettings() As ComboSetting
ListSettings = lS
End Property

Public Property Let ListSettings(ByVal v As ComboSetting)
lS = v
ChangeList
End Property

Private Sub ChangeList()
Dim i As Long
cboMain.Clear
cboMain.Width = 3735
txtMain.Visible = True
cboMain.Left = 840
txtMain.AllowNegative = False
Select Case lS
    Case Rooms
        cboMain.AddItem "(0) None"
        For i = 1 To UBound(dbMap)
            With dbMap(i)
                cboMain.AddItem "(" & .lRoomID & ") " & .sRoomTitle
            End With
            DoEvents
        Next
    Case Items
        cboMain.AddItem "(0) None"
        For i = 1 To UBound(dbItems)
            With dbItems(i)
                cboMain.AddItem "(" & .iID & ") " & .sItemName
            End With
            DoEvents
        Next
    Case Classes
        cboMain.AddItem "(0) None"
        For i = 1 To UBound(dbClass)
            With dbClass(i)
                cboMain.AddItem "(" & .iID & ") " & .sName
            End With
            DoEvents
        Next
    Case Monsters
        cboMain.AddItem "(0) None"
        For i = 1 To UBound(dbMonsters)
            With dbMonsters(i)
                cboMain.AddItem "(" & .lID & ") " & .sMonsterName
            End With
            DoEvents
        Next
    Case Shops
        cboMain.AddItem "(0) None"
        For i = 1 To UBound(dbShops)
            With dbShops(i)
                cboMain.AddItem "(" & .iID & ") " & .sShopName
            End With
            DoEvents
        Next
    Case Environment
        With cboMain
            .AddItem "(0) Outdoor"
            .AddItem "(1) Indoor"
            .AddItem "(2) Underground"
        End With
    Case RoomType
        With cboMain
            .AddItem "(0) Normal"
            .AddItem "(1) Shop"
            .AddItem "(2) Level Trainer"
            .AddItem "(3) Arena"
            .AddItem "(4) Boss"
            .AddItem "(5) Bank"
            .AddItem "(6) Class Trainer"
        End With
    Case Keys
        cboMain.AddItem "(0) None"
        For i = 1 To UBound(dbItems)
            With dbItems(i)
                If .sWorn = "key" Then
                    cboMain.AddItem "(" & .iID & ") " & .sItemName
                End If
            End With
            DoEvents
        Next
    Case Alignment
        cboMain.Left = 0
        cboMain.Width = 2895
        txtMain.Visible = False
        txtMain.AllowNegative = True
        With cboMain
            .AddItem "(120) Choas Lord"
            .AddItem "(99) Vile Outlaw"
            .AddItem "(69) Outlawed Scum"
            .AddItem "(39) Petty Thief"
            .AddItem "(0) Citizen"
            .AddItem "(-40) Law-abiding Citizen"
            .AddItem "(-70) Peace Keeper"
            .AddItem "(-99) Law Enforcer"
            .AddItem "(-175) Peace Lord"
        End With
        
End Select
cboMain.ListIndex = 0
End Sub

Public Property Get Number() As Long
Number = Val(txtMain.Text)
End Property

Public Property Let Number(ByVal v As Long)
txtMain.Text = v
End Property

Public Property Get Enabled() As Boolean
Enabled = cboMain.Enabled
End Property

Public Property Let Enabled(ByVal b As Boolean)
cboMain.Enabled = b
txtMain.Enabled = b
End Property

Public Property Get SpaceSaver() As SizeMode
SpaceSaver = b
End Property

Public Property Let SpaceSaver(ByVal b1 As SizeMode)
b = b1
If b = Small Then
    cboMain.Width = 2655
    UserControl.Width = 3525
Else
    cboMain.Width = 3735
    UserControl.Width = 4560
End If
PropertyChanged SpaceSaver
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "SPACESAVER", b
End Sub
