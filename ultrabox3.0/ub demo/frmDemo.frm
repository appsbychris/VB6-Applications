VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "Demo"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   10935
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
   ScaleHeight     =   8190
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBackPic 
      Caption         =   "Toggle Back Picture"
      Height          =   375
      Left            =   3120
      TabIndex        =   59
      Top             =   7800
      Width           =   2775
   End
   Begin VB.CommandButton cmdSelectStyle 
      Caption         =   "Toggle Select Style (Normal/Fade)"
      Height          =   375
      Left            =   3120
      TabIndex        =   58
      Top             =   7320
      Width           =   2775
   End
   Begin VB.CommandButton cmdListAllBut 
      Caption         =   "List All Selected Check And Option"
      Height          =   375
      Left            =   120
      TabIndex        =   57
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "List All Selected Items"
      Height          =   375
      Left            =   120
      TabIndex        =   56
      Top             =   3960
      Width           =   2775
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   6000
      TabIndex        =   55
      Top             =   250
      Width           =   4935
   End
   Begin VB.OptionButton optInstr 
      Caption         =   "Find InStr"
      Height          =   255
      Left            =   7800
      TabIndex        =   54
      Top             =   0
      Width           =   1335
   End
   Begin VB.OptionButton optExact 
      Caption         =   "Find Exact"
      Height          =   255
      Left            =   6000
      TabIndex        =   53
      Top             =   0
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdHighlight 
      Caption         =   "Randomly Change Items Highlight"
      Height          =   375
      Left            =   120
      TabIndex        =   48
      Top             =   7320
      Width           =   2775
   End
   Begin VB.CommandButton cmdHighText 
      Caption         =   "Randomly Change Items Sel.Txt.Col"
      Height          =   375
      Left            =   120
      TabIndex        =   47
      Top             =   7800
      Width           =   2775
   End
   Begin VB.CommandButton cmdMulti 
      Caption         =   "Toggle Multiselect (Currently: False)"
      Height          =   375
      Left            =   3120
      TabIndex        =   46
      Top             =   5400
      Width           =   2775
   End
   Begin VB.CommandButton cmdRandomFore 
      Caption         =   "Randomly Change Items Forecolor"
      Height          =   375
      Left            =   120
      TabIndex        =   45
      Top             =   6840
      Width           =   2775
   End
   Begin VB.CommandButton cmdItemBack 
      Caption         =   "Randomly Change Items Backcolor"
      Height          =   375
      Left            =   120
      TabIndex        =   44
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton cmdBackColor 
      Caption         =   "Randomly Change Backcolor"
      Height          =   375
      Left            =   120
      TabIndex        =   43
      Top             =   5880
      Width           =   2775
   End
   Begin VB.CommandButton FillView 
      Caption         =   "Toggle Fill View"
      Height          =   375
      Left            =   3120
      TabIndex        =   42
      Top             =   6840
      Width           =   2775
   End
   Begin VB.CommandButton cmdBorder 
      Caption         =   "Toggle Border Style"
      Height          =   375
      Left            =   3120
      TabIndex        =   41
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Timer timPB 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1000
      Left            =   10080
      Top             =   8400
   End
   Begin VB.CommandButton cmdOptChk 
      Caption         =   "Toggle Option/Check/Normal"
      Height          =   375
      Left            =   3120
      TabIndex        =   40
      Top             =   5880
      Width           =   2775
   End
   Begin VB.CommandButton cmdProgress 
      Caption         =   "Add A Progress Bar"
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   4920
      Width           =   2775
   End
   Begin VB.CommandButton cmdSetPic 
      Caption         =   "Change Selected Items Picture"
      Height          =   375
      Left            =   3120
      TabIndex        =   38
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Item  ..."
      Height          =   375
      Left            =   3120
      TabIndex        =   37
      Top             =   3960
      Width           =   2775
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Set Sort (Currently: False)"
      Height          =   375
      Left            =   3120
      TabIndex        =   36
      Top             =   4920
      Width           =   2775
   End
   Begin VB.CommandButton cmdEnableMain 
      Caption         =   "Enable/Disable Ultrabox"
      Height          =   375
      Left            =   3120
      TabIndex        =   35
      Top             =   3480
      Width           =   2775
   End
   Begin VB.CommandButton cmdEnable 
      Caption         =   "Enable/Disable Item ..."
      Height          =   375
      Left            =   3120
      TabIndex        =   34
      Top             =   3000
      Width           =   2775
   End
   Begin VB.PictureBox picOver 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   9120
      Picture         =   "frmDemo.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   33
      Top             =   8520
      Width           =   300
   End
   Begin VB.PictureBox picCheck 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   8520
      Picture         =   "frmDemo.frx":0102
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   32
      Top             =   8520
      Width           =   513
   End
   Begin VB.PictureBox picDebit 
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   7920
      Picture         =   "frmDemo.frx":0704
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   31
      Top             =   8520
      Width           =   513
   End
   Begin VB.CommandButton cmdPics 
      Caption         =   "Fill With Pictures And Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   3480
      Width           =   2775
   End
   Begin VB.PictureBox picHamburger 
      AutoSize        =   -1  'True
      Height          =   450
      Left            =   8880
      Picture         =   "frmDemo.frx":0E76
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   29
      Top             =   8880
      Width           =   513
   End
   Begin VB.PictureBox picCheeseburger 
      AutoSize        =   -1  'True
      Height          =   390
      Left            =   8400
      Picture         =   "frmDemo.frx":1810
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   28
      Top             =   8880
      Width           =   513
   End
   Begin VB.PictureBox picFries 
      AutoSize        =   -1  'True
      Height          =   480
      Left            =   5400
      Picture         =   "frmDemo.frx":203A
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   27
      Top             =   8880
      Width           =   513
   End
   Begin VB.PictureBox picNuggets 
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   6840
      Picture         =   "frmDemo.frx":2A8C
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   26
      Top             =   8880
      Width           =   513
   End
   Begin VB.PictureBox picdietCoke 
      AutoSize        =   -1  'True
      Height          =   510
      Left            =   5880
      Picture         =   "frmDemo.frx":353A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   25
      Top             =   8880
      Width           =   513
   End
   Begin VB.PictureBox picMountainDew 
      AutoSize        =   -1  'True
      Height          =   510
      Left            =   6360
      Picture         =   "frmDemo.frx":4044
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   24
      Top             =   8880
      Width           =   513
   End
   Begin VB.PictureBox picDrPepper 
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   7320
      Picture         =   "frmDemo.frx":4B4E
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   23
      Top             =   8880
      Width           =   513
   End
   Begin VB.PictureBox picCash 
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   7920
      Picture         =   "frmDemo.frx":5CD0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   22
      Top             =   8880
      Width           =   513
   End
   Begin VB.CommandButton cmdRandom 
      Caption         =   "Fill With Random Characters"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4920
      Top             =   8760
   End
   Begin UltraboxDemo.UltraBox ubMain 
      Height          =   7575
      Left            =   6000
      TabIndex        =   0
      Top             =   600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   13361
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
      SELECTSTYLE     =   0
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "_ItemAdded:"
      Height          =   195
      Index           =   11
      Left            =   120
      TabIndex        =   52
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label lblItemAdded 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   195
      Left            =   1440
      TabIndex        =   51
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblListCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   1440
      TabIndex        =   50
      Top             =   2760
      Width           =   90
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "List Count:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   49
      Top             =   2760
      Width           =   780
   End
   Begin VB.Label lblDoubleClick 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   195
      Left            =   1440
      TabIndex        =   20
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "_DoubleClick::"
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   1020
   End
   Begin VB.Label lblVerticalScroll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   195
      Left            =   1440
      TabIndex        =   18
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblMouseUp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   195
      Left            =   1440
      TabIndex        =   17
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblMouseMove 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   195
      Left            =   1440
      TabIndex        =   16
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblMouseDown 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   195
      Left            =   1440
      TabIndex        =   15
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblItemClicked 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   195
      Left            =   1440
      TabIndex        =   14
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblItemChecked 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   195
      Left            =   1440
      TabIndex        =   13
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblHorizontalScroll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   195
      Left            =   1440
      TabIndex        =   12
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblCtrl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   195
      Left            =   1440
      TabIndex        =   11
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblClick 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   195
      Left            =   1440
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "_VerticalScroll:"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1050
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "_MouseUp:"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   810
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "_MouseMove:"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1005
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "_MouseDown:"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1020
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "_ItemClicked:"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "_ItemChecked:"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "_HorizontalScroll:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "_Ctrl:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   405
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "_Click:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBackColor_Click()
ubMain.Color = RGB(RndNumber(0, 255), RndNumber(0, 255), RndNumber(0, 255))
End Sub

Private Sub cmdBackPic_Click()
If ubMain.BackPic Is Nothing Then
    Set ubMain.BackPic = LoadPicture(App.Path & "\sky.jpg")
Else
    Set ubMain.BackPic = Nothing
End If
ubMain.Refresh
End Sub

Private Sub cmdBorder_Click()
'Public Enum View
'    RaisedEdge = 0
'    SunkenEdge = 1
'    BumpedEdge = 2
'    EtchedEdge = 3
'    LineEdge = 4
'    None = 5
'End Enum
If ubMain.Style + 1 < 6 Then
    ubMain.Style = ubMain.Style + 1
Else
    ubMain.Style = 0
End If
End Sub

Private Sub cmdDelete_Click()
Dim s As String
Dim l As Long
s = InputBox$("Enter an index number from 1 to " & ubMain.ListCount & ".", "Index", "1")
If s <> vbNullString Then
    l = Val(s)
    ubMain.RemoveItem l
End If
End Sub

Private Sub cmdEnable_Click()
Dim s As String
Dim l As Long
s = InputBox$("Enter an index number from 1 to " & ubMain.ListCount & ".", "Index", "1")
If s <> vbNullString Then
    l = Val(s)
    ubMain.SetEnabled l, Not ubMain.ItemEnabled(l)
End If
End Sub

Private Sub cmdEnableMain_Click()
ubMain.Enabled = Not ubMain.Enabled
End Sub

Private Sub cmdHighlight_Click()
If ubMain.ListIndex < 1 Then Exit Sub
ubMain.SetItemColors ubMain.ListIndex, [Highlight Color], RGB(RndNumber(0, 255), RndNumber(0, 255), RndNumber(0, 255))
End Sub

Private Sub cmdHighText_Click()
If ubMain.ListIndex < 1 Then Exit Sub
ubMain.SetItemColors ubMain.ListIndex, [Highlight Text], RGB(RndNumber(0, 255), RndNumber(0, 255), RndNumber(0, 255))
End Sub

Private Sub cmdItemBack_Click()
If ubMain.ListIndex < 1 Then Exit Sub
ubMain.SetItemColors ubMain.ListIndex, [Back Color], RGB(RndNumber(0, 255), RndNumber(0, 255), RndNumber(0, 255))
End Sub

Private Sub cmdList_Click()
Dim i As Long
Dim s As String
For i = 1 To ubMain.ListCount
    If ubMain.IsSelected(i, , True) Then
        s = s & "SELECTED INDEX: " & i & ": " & ubMain.List(i) & vbCrLf
    End If
    If ubMain.IsSelected(i, True) Then
        If ubMain.ItemTypeX(i) = [Check Box] Then
            s = s & "SELECT CHECK: " & i & ": " & ubMain.List(i) & vbCrLf
        ElseIf ubMain.ItemTypeX(i) = [Option Box] Then
            s = s & "SELECT OPTION: " & i & " (OptGrp:" & ubMain.GetItemOptGrp(i) & "): " & ubMain.List(i) & vbCrLf
        End If
    End If
Next
MsgBox s
End Sub

Private Sub cmdListAllBut_Click()
Dim i As Long
Dim s As String
For i = 1 To ubMain.ListCount
    If ubMain.IsSelected(i, True) Then
        If ubMain.ItemTypeX(i) = [Check Box] Then
            s = s & "SELECT CHECK: " & i & ": " & ubMain.List(i) & vbCrLf
        ElseIf ubMain.ItemTypeX(i) = [Option Box] Then
            s = s & "SELECT OPTION: " & i & " (OptGrp:" & ubMain.GetItemOptGrp(i) & "): " & ubMain.List(i) & vbCrLf
        End If
    End If
Next
MsgBox s

End Sub

Private Sub cmdMulti_Click()
Select Case cmdMulti.Caption
    Case "Toggle Multiselect (Currently: False)"
        cmdMulti.Caption = "Toggle Multiselect (Currently: True)"
        ubMain.MultiSelect = True
    Case "Toggle Multiselect (Currently: True)"
        cmdMulti.Caption = "Toggle Multiselect (Currently: False)"
        ubMain.MultiSelect = False
End Select
End Sub

Private Sub cmdOptChk_Click()
Static ii As Long
Dim j As Long
Dim s As String
If ubMain.ListIndex < 1 Then Exit Sub
Select Case ubMain.ItemTypeX(ubMain.ListIndex)
    Case 0
        ubMain.MakeItemX ubMain.ListIndex, [Check Box]
    Case 1
        ubMain.MakeItemX ubMain.ListIndex, Normal
    Case 3
        s = InputBox$("What option group would you like to add it to?", "Option Group", "0")
        If s <> vbNullString Then
            If Len(s) > 6 Then s = Left$(s, 6)
            j = Val(s)
            ubMain.MakeItemX ubMain.ListIndex, [Option Box], j
        End If
End Select
End Sub

Private Sub cmdPics_Click()
With ubMain
    .Clear
    cmdSort.Caption = "Set Sort (Currently: False)"
    .Sorted = False
    .Paint = False
    .Font.Size = 14
    .Refresh
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
End Sub

Private Function RndNumber(Min As Long, Max As Long) As Long
'gets a random number from the min to the max
RndNumber = Val((Rnd * (Max - Min)) + Min)
End Function

Private Sub cmdProgress_Click()
cmdSort.Caption = "Set Sort (Currently: False)"
ubMain.Sorted = False
ubMain.AddItemProgressBar "0%", vbCenter, FCOLOR:=vbBlack
Load timPB(timPB.UBound + 1)
With timPB(timPB.UBound)
    .Tag = ubMain.ListCount
    .Enabled = True
End With
End Sub

Private Sub cmdRandom_Click()
Dim j As Long
Dim i As Long
Dim s As String
ubMain.Clear
ubMain.Color = vbWhite
cmdSort.Caption = "Set Sort (Currently: False)"
ubMain.Sorted = False
ubMain.Paint = False
ubMain.Font.Size = 8
ubMain.Refresh
Do Until j = 100
    s = s & Chr$(RndNumber(65, 118))
    If Len(s) > RndNumber(4, 17) Then
        ubMain.AddItem s, FCOLOR:=RGB(Val(RndNumber(0, 255)), Val(RndNumber(0, 255)), Val(RndNumber(0, 255))), BCOLOR:=vbWhite
        s = ""
        j = j + 1
    End If
    DoEvents
Loop
ubMain.Paint = True
End Sub

Private Sub cmdRandomFore_Click()
If ubMain.ListIndex < 1 Then Exit Sub
ubMain.SetItemColors ubMain.ListIndex, [Fore Color], RGB(RndNumber(0, 255), RndNumber(0, 255), RndNumber(0, 255))
End Sub

Private Sub cmdSelectStyle_Click()
ubMain.SelectStyle = IIf(ubMain.SelectStyle = Default, Faded, Default)
End Sub

Private Sub cmdSetPic_Click()
If ubMain.ListIndex > 0 Then
    ubMain.Paint = False
    Select Case RndNumber(1, 11)
        Case 1
            ubMain.SetItemPicture ubMain.ListIndex, picFries.Picture, RGB(0, 255, 0)
        Case 2
            ubMain.SetItemPicture ubMain.ListIndex, picdietCoke.Picture, RGB(0, 255, 0)
        Case 3
            ubMain.SetItemPicture ubMain.ListIndex, picMountainDew.Picture, RGB(0, 255, 0)
        Case 4
            ubMain.SetItemPicture ubMain.ListIndex, picNuggets.Picture, RGB(0, 255, 0)
        Case 5
            ubMain.SetItemPicture ubMain.ListIndex, picDrPepper.Picture, RGB(0, 255, 0)
        Case 6
            ubMain.SetItemPicture ubMain.ListIndex, picCash.Picture, RGB(0, 255, 0)
        Case 7
            ubMain.SetItemPicture ubMain.ListIndex, picCheeseburger.Picture, RGB(0, 255, 0)
        Case 8
            ubMain.SetItemPicture ubMain.ListIndex, picHamburger.Picture, RGB(0, 255, 0)
        Case 9
            ubMain.SetItemPicture ubMain.ListIndex, picDebit.Picture
        Case 10
            ubMain.SetItemPicture ubMain.ListIndex, picCheck.Picture
        Case 11
            ubMain.SetItemPicture ubMain.ListIndex, picOver.Picture
    End Select
    ubMain.Paint = True
End If
End Sub

Private Sub cmdSort_Click()
ubMain.Paint = False
Select Case cmdSort.Caption
    Case "Set Sort (Currently: False)"
        cmdSort.Caption = "Set Sort (Currently: True)"
        ubMain.Sorted = True
    Case "Set Sort (Currently: True)"
        cmdSort.Caption = "Set Sort (Currently: False)"
        ubMain.Sorted = False
End Select
ubMain.Paint = True
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub FillView_Click()
If ubMain.FillView = NoStyle Then
    ubMain.FillView = Lined
Else
    ubMain.FillView = NoStyle
End If
End Sub

Private Sub Form_Load()
ubMain.Paint = False
For i = 0 To 100
    ubMain.AddItem "test" & i
Next
ubMain.Paint = True
End Sub

Private Sub Timer1_Timer()
Me.Caption = Now
End Sub

Private Sub timPB_Timer(Index As Integer)
   On Error GoTo timPB_Timer_Error
If ubMain.GetProgressValue(Val(timPB(Index).Tag)) + 1 > ubMain.GetProgressMax(Val(timPB(Index).Tag)) Then ubMain.SetProgressValue Val(timPB(Index).Tag), 0
ubMain.SetProgressValue Val(timPB(Index).Tag), ubMain.GetProgressValue(Val(timPB(Index).Tag)) + 1
ubMain.SetItemText Val(timPB(Index).Tag), Round((ubMain.GetProgressValue(Val(timPB(Index).Tag)) / ubMain.GetProgressMax(Val(timPB(Index).Tag))) * 100, 0) & "%"

   On Error GoTo 0
   Exit Sub

timPB_Timer_Error:
timPB(Index).Enabled = False
Unload timPB(Index)
End Sub

Private Sub txtFind_Change()
If optExact.Value Then ubMain.SetSelected ubMain.Find(txtFind.Text), True, True, True
If optInstr.Value Then ubMain.SetSelected ubMain.FindInStr(txtFind.Text), True, True, True
End Sub

Private Sub ubMain_Click()
lblClick.Caption = Now
End Sub

Private Sub ubMain_Ctrl(Value As Long)
lblCtrl.Caption = "(Value=" & CStr(Value) & ") " & Now
End Sub

Private Sub ubMain_DoubleClick()
lblDoubleClick.Caption = Now
End Sub

Private Sub ubMain_HorizontalScroll(lValue As Long)
lblHorizontalScroll.Caption = "(lValue=" & CStr(lValue) & ") " & Now
End Sub

Private Sub ubMain_ItemAdded()
lblItemAdded.Caption = Now
lblListCount.Caption = ubMain.ListCount
End Sub

Private Sub ubMain_ItemChecked(Index As Long)
lblItemChecked.Caption = "(Index=" & CStr(Index) & ") " & Now
End Sub

Private Sub ubMain_ItemClicked(Index As Long)
lblItemClicked.Caption = "(Index=" & CStr(Index) & ") " & Now
End Sub

Private Sub ubMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMouseDown.Caption = "(Button=" & CStr(Button) & ",Shift=" & CStr(Shift) & ",X=" & CStr(X) & ",Y=" & CStr(Y) & ") " & Now
End Sub

Private Sub ubMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMouseMove.Caption = "(Button=" & CStr(Button) & ",Shift=" & CStr(Shift) & ",X=" & CStr(X) & ",Y=" & CStr(Y) & ") " & Now
End Sub

Private Sub ubMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMouseUp.Caption = "(Button=" & CStr(Button) & ",Shift=" & CStr(Shift) & ",X=" & CStr(X) & ",Y=" & CStr(Y) & ") " & Now
End Sub

Private Sub ubMain_VerticalScroll(lValue As Long)
lblVerticalScroll.Caption = "(lValue=" & CStr(lValue) & ") " & Now
End Sub
