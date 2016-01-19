VERSION 5.00
Begin VB.UserControl IntelliSense 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   595
   ToolboxBitmap   =   "IntelliSense.ctx":0000
   Begin VB.PictureBox picLstMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2160
      Left            =   5760
      ScaleHeight     =   144
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   1995
      Begin VB.ListBox lstMain 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         IntegralHeight  =   0   'False
         ItemData        =   "IntelliSense.ctx":0312
         Left            =   45
         List            =   "IntelliSense.ctx":0325
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   45
         Width           =   1695
      End
      Begin ServerEditor.Raise rsMain 
         Height          =   1830
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   3228
         Style           =   0
         Color           =   0
      End
   End
   Begin VB.PictureBox picSide 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   0
      ScaleHeight     =   231
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
   End
   Begin VB.TextBox txtMain 
      Appearance      =   0  'Flat
      Height          =   3735
      HideSelection   =   0   'False
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "IntelliSense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Private Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long


Private Const WM_USER As Long = &H400
Private Const EM_GETTEXTRANGE As Long = (WM_USER + 75)
Private Const TB_GETTEXTROWS As Long = (WM_USER + 61)
Private Const EM_GETLINE As Long = &HC4
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_LINEFROMCHAR As Long = &HC9
Private Const EM_LINELENGTH = &HC1
Private Const EM_GETFIRSTVISIBLELINE As Long = &HCE

Private Const LB_SELECTSTRING = &H18C
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const HWND_BOTTOM As Long = 1
Private Const HWND_TOP As Long = 0
Private Const SW_HIDE As Long = 0
Private Const WM_PASTE As Long = &H302
Private Const DT_ACCEPT_DBCS As Long = (&H20)
Private Const DT_AGENT As Long = (&H3)
Private Const DT_BOTTOM As Long = &H8
Private Const DT_CALCRECT As Long = &H400
Private Const DT_CENTER As Long = &H1
Private Const DT_CHARSTREAM As Long = 4
Private Const DT_DISPFILE As Long = 6
Private Const DT_DISTLIST As Long = (&H1)
Private Const DT_EDITABLE As Long = (&H2)
Private Const DT_EDITCONTROL As Long = &H2000
Private Const DT_END_ELLIPSIS As Long = &H8000
Private Const DT_EXPANDTABS As Long = &H40
Private Const DT_EXTERNALLEADING As Long = &H200
Private Const DT_FOLDER As Long = (&H1000000)
Private Const DT_FOLDER_LINK As Long = (&H2000000)
Private Const DT_FOLDER_SPECIAL As Long = (&H4000000)
Private Const DT_FORUM As Long = (&H2)
Private Const DT_GLOBAL As Long = (&H20000)
Private Const DT_HIDEPREFIX As Long = &H100000
Private Const DT_INTERNAL As Long = &H1000
Private Const DT_LEFT As Long = &H0
Private Const DT_LOCAL As Long = (&H30000)
Private Const DT_MAILUSER As Long = (&H0)
Private Const DT_METAFILE As Long = 5
Private Const DT_MODIFIABLE As Long = (&H10000)
Private Const DT_MODIFYSTRING As Long = &H10000
Private Const DT_MULTILINE As Long = (&H1)
Private Const DT_NOCLIP As Long = &H100
Private Const DT_NOFULLWIDTHCHARBREAK As Long = &H80000
Private Const DT_NOPREFIX As Long = &H800
Private Const DT_NOT_SPECIFIC As Long = (&H50000)
Private Const DT_ORGANIZATION As Long = (&H4)
Private Const DT_PASSWORD_EDIT As Long = (&H10)
Private Const DT_PATH_ELLIPSIS As Long = &H4000
Private Const DT_PLOTTER As Long = 0
Private Const DT_PREFIXONLY As Long = &H200000
Private Const DT_PRIVATE_DISTLIST As Long = (&H5)
Private Const DT_RASCAMERA As Long = 3
Private Const DT_RASDISPLAY As Long = 1
Private Const DT_RASPRINTER As Long = 2
Private Const DT_REMOTE_MAILUSER As Long = (&H6)
Private Const DT_REQUIRED As Long = (&H4)
Private Const DT_RIGHT As Long = &H2
Private Const DT_RTLREADING As Long = &H20000
Private Const DT_SET_IMMEDIATE As Long = (&H8)
Private Const DT_SET_SELECTION As Long = (&H40)
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_TABSTOP As Long = &H80
Private Const DT_TOP As Long = &H0
Private Const DT_VCENTER As Long = &H4
Private Const DT_WAN As Long = (&H40000)
Private Const DT_WORD_ELLIPSIS As Long = &H40000
Private Const DT_WORDBREAK As Long = &H10

Private Const BDR_RAISEDOUTER As Long = &H1
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BDR_SUNKENINNER As Long = &H8
Private Const BDR_RAISEDINNER = &H4
Private Const EDGE_BUMP As Long = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED As Long = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED As Long = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN As Long = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const BF_BOTTOM As Long = &H8
Private Const BF_LEFT As Long = &H1
Private Const BF_RIGHT As Long = &H4
Private Const BF_TOP As Long = &H2
Private Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    lLeft As Long
    lTop As Long
    lRight As Long
    lBottom As Long
End Type
Private Type Words
    sKeyWord As String
    sMethods As String
    sTrigger As String
    sEndLine As String
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private sString As String
Private IntelliSenseIsShown As Boolean
Private IntelliSenseHasAWord As Boolean
Private TipIsShown As Boolean
Private lH As Long
Private udtKeyWords() As Words
Private TipMode As Boolean
Private CMethod As String
Public Event MethodHasChanged()


Private Sub lstMain_DblClick()
HideIntelliSense True, lH
End Sub

Private Sub lstMain_KeyPress(KeyAscii As Integer)
txtMain_KeyPress KeyAscii
End Sub

Private Sub txtMain_Change()
PaintNumbers
End Sub

Public Sub PaintNumbers()
Dim lBegin As Long
Dim lHeight As Long
Dim lLines As Long
Dim R As RECT
Dim j As Long
lBegin = SendMessage(txtMain.hWnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
lHeight = UserControl.TextHeight("X")
lLines = txtMain.Height \ lHeight
picSide.Cls
'MsgBox (lBegin + 1) & " " & ((lLines + 1) + (lBegin + 1))
For i = (lBegin + 1) To (lLines + 1) + (lBegin + 1)
    With R
        .lTop = (j * lHeight) + 1
        .lBottom = lHeight + .lTop
        .lLeft = 1
        .lRight = (picSide.ScaleWidth)
        If .lTop > picSide.ScaleHeight - lHeight Then Exit For
    End With
    DrawText picSide.hdc, CStr(i), Len(CStr(i)), R, BF_RECT
    j = j + 1
Next
picSide.Refresh
End Sub

Private Sub txtMain_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyDown
        If IntelliSenseIsShown Then
            IntelliSenseMoveDown
            For i = LBound(udtKeyWords) To UBound(udtKeyWords)
                With udtKeyWords(i)
                    If .sKeyWord = lstMain.list(lstMain.ListIndex) Then
                        CMethod = udtKeyWords(i).sMethods
                        RaiseEvent MethodHasChanged
                        Exit For
                    End If
                End With
                DoEvents
            Next
            KeyCode = 0
        End If
        txtMain_Change
    Case vbKeyUp
        If IntelliSenseIsShown Then
            IntelliSenseMoveUp
            For i = LBound(udtKeyWords) To UBound(udtKeyWords)
                With udtKeyWords(i)
                    If .sKeyWord = lstMain.list(lstMain.ListIndex) Then
                        CMethod = udtKeyWords(i).sMethods
                        RaiseEvent MethodHasChanged
                        Exit For
                    End If
                End With
                DoEvents
            Next
            KeyCode = 0
        End If
        txtMain_Change
    Case vbKeyTab
        txtMain.SelText = "    "
        KeyCode = 0
End Select
End Sub

Private Sub txtMain_KeyPress(KeyAscii As Integer)
Dim p As POINTAPI
Dim i As Long
Dim s As String
Dim t As String
GetCaretPos p

Select Case KeyAscii
    Case Asc(".")
        On Error GoTo BadMove
        If IntelliSenseIsShown Then HideIntelliSense True, txtMain.hWnd
        For i = LBound(udtKeyWords) To UBound(udtKeyWords)
            If txtMain.SelStart - Len(udtKeyWords(i).sTrigger) >= 0 Then
                s = LCase$(Mid$(txtMain.Text, (txtMain.SelStart + 1) - Len(udtKeyWords(i).sTrigger), Len(udtKeyWords(i).sTrigger)))
                If s = LCase$(udtKeyWords(i).sTrigger) Then
                    If udtKeyWords(i).sKeyWord <> "" Then
                        If Not IntelliSenseIsShown Then
                            ShowIntelliSense p.x, p.y, txtMain.hWnd, udtKeyWords(i).sTrigger
                            CMethod = udtKeyWords(i).sMethods
                            RaiseEvent MethodHasChanged
                            Exit For
                        End If
                    End If
                End If
            End If
            DoEvents
        Next
    Case Asc(" "), Asc(vbTab), Asc(","), Asc(")"), Asc("(")
        If IntelliSenseIsShown Then
            HideIntelliSense True, txtMain.hWnd
            If KeyAscii = Asc(" ") Or KeyAscii = Asc(vbTab) Then KeyAscii = 0
        Else
            If KeyAscii = Asc(vbTab) Then KeyAscii = 0
        End If
    Case vbKeyBack
        If IntelliSenseHasAWord Then
            If IntelliSenseIsShown Then
                IntelliSenseRemoveLastLetterOfWord
                IntelliSenseOffsetLeft p.x - (Screen.TwipsPerPixelX \ 2)
            Else
                HideIntelliSense
            End If
        Else
            HideIntelliSense
        End If
    Case 13
        If IntelliSenseIsShown Then HideIntelliSense True, txtMain.hWnd
        
    Case Else
        If IntelliSenseIsShown Then
            IntelliSenseAppendToCurrentWord Chr$(KeyAscii)
            IntelliSenseOffsetLeft p.x
        End If
End Select
Exit Sub
BadMove:
MsgBox "error " & Err.Description & " " & Err.Number
End Sub

Private Sub UserControl_Initialize()
With UserControl
    .Width = txtMain.Width + picSide.Width
    .Height = txtMain.Height
End With
txtMain.Left = picSide.ScaleWidth
ReDim udtKeyWords(0)
txtMain_Change
End Sub

Private Sub UserControl_Resize()
With txtMain
    .Width = UserControl.ScaleWidth - picSide.ScaleWidth
    .Height = UserControl.ScaleHeight
    .Left = picSide.ScaleWidth
End With
picSide.Height = UserControl.ScaleHeight
End Sub

Private Sub AlwaysOnTop(hWnd As Long, SetOnTop As Boolean)
    Dim lFlag As Long
    Dim R As RECT
    GetWindowRect hWnd, R
    If SetOnTop Then lFlag = HWND_TOPMOST Else lFlag = HWND_NOTOPMOST
    SetWindowPos hWnd, lFlag, _
        R.lLeft, _
        R.lTop, _
        R.lRight - R.lLeft, _
        R.lBottom - R.lTop, _
        SWP_NOACTIVATE Or SWP_SHOWWINDOW
    SetRectEmpty R
End Sub

Private Sub ShowIntelliSense(x As Long, y As Long, ControlhWnd As Long, sTrigger As String)
    Dim R As RECT
    Dim lBDWP As Long
    Dim i As Long
    lstMain.Clear
    For i = LBound(udtKeyWords) To UBound(udtKeyWords)
        If udtKeyWords(i).sKeyWord <> "" Then
            If LCase$(udtKeyWords(i).sTrigger) = LCase$(sTrigger) Then
                lstMain.AddItem udtKeyWords(i).sKeyWord
            End If
        End If
        DoEvents
    Next
    IntelliSenseAutoSizeIntelliSense
    'If TipIsShown Then HideTip
    GetWindowRect picLstMain.hWnd, R
    With R
        .lLeft = x + (Screen.TwipsPerPixelX \ 2) + picSide.ScaleWidth
        .lTop = y + (Screen.TwipsPerPixelY)
        .lRight = picLstMain.ScaleWidth + .lLeft
        .lBottom = picLstMain.ScaleHeight + .lTop
    End With
    If R.lTop > txtMain.Height - picLstMain.ScaleHeight Then
        R.lTop = R.lTop - picLstMain.ScaleHeight - (UserControl.TextHeight("X") + 1)
        R.lBottom = R.lBottom - picLstMain.ScaleHeight - (UserControl.TextHeight("X") + 1)
    End If
    If R.lLeft > txtMain.Width - picLstMain.ScaleWidth Then
        R.lLeft = txtMain.Width - picLstMain.ScaleWidth + picSide.ScaleWidth
        R.lRight = picLstMain.ScaleWidth + R.lLeft
    End If
    lBDWP = BeginDeferWindowPos(1)
    DeferWindowPos lBDWP, picLstMain.hWnd, HWND_TOP, R.lLeft, R.lTop, R.lRight - R.lLeft, R.lBottom - R.lTop, SWP_SHOWWINDOW
    EndDeferWindowPos lBDWP
    GetClientRect picLstMain.hWnd, R
    OffsetRect R, 3, 3
    FillRect picLstMain.hdc, R, CreateSolidBrush(&H808080)
    picLstMain.Refresh
    AlwaysOnTop picLstMain.hWnd, True
    If picLstMain.Top < 0 Then
        With picLstMain
            .Height = picLstMain.Height + picLstMain.Top
            .Top = 6
        End With
        With lstMain
            .Height = picLstMain.Height - 6
        End With
        With rsMain
            .Height = lstMain.Height + 6
        End With
    End If
    If lstMain.Visible = False Then lstMain.Visible = True
    IntelliSenseIsShown = True
    SetRectEmpty R
End Sub

Private Sub HideIntelliSense(Optional AddHighlightedWord As Boolean = False, Optional txtBoxhWnd As Long)
    Dim s As String, t As String
    AlwaysOnTop picLstMain.hWnd, False
    ShowWindow picLstMain.hWnd, SW_HIDE
    IntelliSenseIsShown = False
    If AddHighlightedWord Then
        s = Clipboard.GetText
        Clipboard.Clear
        t = lstMain.list(lstMain.ListIndex)
        IntelliSenseRemoveCurrentTypeWord txtMain.SelStart, txtMain.hWnd
        Clipboard.SetText t
        SendMessageByString txtBoxhWnd, WM_PASTE, Len(t), t
        Clipboard.Clear
        Clipboard.SetText s
        For i = LBound(udtKeyWords) To UBound(udtKeyWords)
            With udtKeyWords(i)
                If .sKeyWord = t Then
                    CMethod = udtKeyWords(i).sMethods
                    RaiseEvent MethodHasChanged
                    Exit For
                End If
            End With
            DoEvents
        Next
    End If
    sString = ""
    lstMain.ListIndex = -1
    txtMain.SetFocus
    IntelliSenseHasAWord = False
End Sub

Private Sub IntelliSenseMoveDown()
    If lstMain.ListIndex + 1 < lstMain.ListCount - 1 Then
        lstMain.ListIndex = lstMain.ListIndex + 1
    Else
        lstMain.ListIndex = lstMain.ListCount - 1
    End If
End Sub

Private Sub IntelliSenseMoveUp()
    If lstMain.ListIndex - 1 < lstMain.ListCount - 1 Then
        lstMain.ListIndex = lstMain.ListIndex - 1
    Else
        lstMain.ListIndex = 0
    End If
End Sub

Private Sub IntelliSenseAppendToCurrentWord(sNew As String)
Dim RetVal As Long
Dim i As Long
sString = sString & sNew
RetVal = SendMessage(UCase$(lstMain.hWnd), LB_SELECTSTRING, ByVal -1, ByVal UCase$(sString))
If RetVal > -1 Then lstMain.Selected(RetVal) = True
For i = LBound(udtKeyWords) To UBound(udtKeyWords)
    If lstMain.list(lstMain.ListIndex) = udtKeyWords(i).sKeyWord Then
        CMethod = udtKeyWords(i).sMethods
        RaiseEvent MethodHasChanged
    End If
    DoEvents
Next
If sString <> "" Then IntelliSenseHasAWord = True
End Sub

Private Sub IntelliSenseOffsetLeft(x As Long)
picLstMain.Left = x + (Screen.TwipsPerPixelX \ 2) + picSide.ScaleWidth
If picLstMain.Left > txtMain.Width - picLstMain.Width Then
    picLstMain.Left = txtMain.Width - picLstMain.Width + picSide.ScaleWidth
End If
End Sub

Private Sub IntelliSenseRemoveLastLetterOfWord()
Dim RetVal As Long
If Len(sString) > 0 Then sString = Left$(sString, Len(sString) - 1)
RetVal = SendMessage(UCase$(lstMain.hWnd), LB_SELECTSTRING, ByVal -1, ByVal UCase$(sString))
If RetVal > -1 Then lstMain.Selected(RetVal) = True
If sString <> "" Then IntelliSenseHasAWord = True
End Sub

Public Sub IntelliSenseAddWords(sKeyWord As String, sMethods As String, sTrigger As String, Optional sEndLine As String = "")
ReDim Preserve udtKeyWords(UBound(udtKeyWords) + 1)
With udtKeyWords(UBound(udtKeyWords))
    .sKeyWord = sKeyWord
    .sMethods = sMethods
    .sTrigger = sTrigger
    .sEndLine = sEndLine
End With
End Sub

Public Sub IntelliSenseAddWordsFile(sFile As String)
Dim s As String
Dim i As Long
Dim m As Long
Dim n As Long
Dim Arr() As String
Dim Arr2() As String
Dim sKeyWord As String
Dim sMethods As String
Dim sTrigger As String
Open sFile For Binary As #1
    s = Input$(LOF(1), 1)
Close #1
Do While Len(s) > 0
    m = InStr(1, s, "[Start Method List=")
    If m = 0 Then Exit Do
    n = InStr(m, s, "=")
    m = InStr(n, s, "]")
    sTrigger = Mid$(s, n + 1, m - n - 1)
    n = InStr(m, s, "[End Method List=" & sTrigger & "]")
    Arr = Split(Mid$(s, m + 1, n - m - 1), vbCrLf)
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) <> "" Then
            Arr(i) = Replace$(Arr(i), "[", "")
            Arr(i) = Replace$(Arr(i), "]", "")
            Arr(i) = Replace$(Arr(i), vbCrLf, "")
            Arr2 = Split(Arr(i), ",", 2)
            IntelliSenseAddWords Arr2(0), Arr2(1), sTrigger
        End If
        DoEvents
    Next
    s = Mid$(s, n + Len("[End Method List=" & sTrigger & "]"))
    DoEvents
Loop
End Sub

Private Sub IntelliSenseRemoveCurrentTypeWord(lSelStart As Long, ControlhWnd As Long)
Dim n As Long
Dim s As String
Dim i As Long, j As Long
On Error GoTo BadLine
j = InStrRev(txtMain.Text, ".", txtMain.SelStart)
txtMain.SelStart = j
j = txtMain.SelStart
On Error GoTo EndLine
Do Until i <> 0
    j = j + 1
    If j > Len(txtMain) Then GoTo EndLine
    If Mid$(txtMain.Text, j, 1) = vbCr Or Mid$(txtMain.Text, j, 1) = vbLf Or Mid$(txtMain.Text, j, 1) = " " Then
        i = 1
        'j = j + 1
        Exit Do
    End If
    DoEvents
Loop
EndLine:
j = j - 1
txtMain.SelLength = j - txtMain.SelStart
BadLine:
End Sub

'Public Function TipIsStillInTipZone(lSelStart As Long, ControlhWnd As Long) As Boolean
'Dim n As Long
'Dim s As String
'Dim i As Long
'Dim m As Long
'Dim CurCom As String
''n = SendMessage(ControlhWnd, EM_LINEFROMCHAR, ByVal lSelStart, ByVal 0&)
''s = Space$(400)
''SendMessage ControlhWnd, EM_GETLINE, n, ByVal s
''s = Trim$(s)
''i = InStrRev(s, " ")
''If i > lSelStart Then
''    Do Until i < lSelStart
''        i = InStrRev(s, " ", i - 1)
''        DoEvents
''    Loop
''End If
''CurCom = Mid$(s, i + 1, lSelStart - i)
''If CurCom <> "" Then
'CurCom = IntelliSenseGetLastKeyWord(txtMain.hWnd, txtMain.SelStart)
'For i = LBound(udtKeyWords) To UBound(udtKeyWords)
'    With udtKeyWords(i)
'        If LCase$(.sTrigger) = LCase$(CurCom) Then
'            n = DCount(.sMethods, ",")
'            Exit For
'        End If
'    End With
'    DoEvents
'Next
'If n > 0 Then
'    Do Until i > 0
'        i = InStrRev(txtMain.Text, ".", txtMain.SelStart)
'        If LCase$(Mid$(txtMain.Text, i - Len(CurCom), Len(CurCom))) = LCase$(CurCom) Then
'            Exit Do
'        Else
'            i = 0
'        End If
'        DoEvents
'    Loop
'    Do Until m > 0
'        m = InStr(i, txtMain.Text, ",")
'        If m > lSelStart Then
'            Exit Do
'        Else
'            m = 0
'        End If
'        DoEvents
'    Loop
'End If
'End Function

Private Function DCount(sInString As String, sCountString As String) As Long
'finds out how many a certain string appears in another string
Dim iNextOccur As Long
iNextOccur = 0
'If InStr(iNextOccur, sInString, sCountString) = 0 Then Exit Function Else DCount = DCount + 1
Do
    iNextOccur = InStr(iNextOccur + 1, sInString, sCountString)
    If iNextOccur > 0 Then DCount = DCount + 1 Else Exit Do
    DoEvents
Loop
End Function

Private Function IntelliSenseGetLastKeyWord(ControlhWnd As Long, lSelStart As Long) As String
Dim n As Long
Dim s As String
Dim i As Long
Dim CurCom As String
Dim bFound As Boolean
On Error GoTo BadLine
n = SendMessage(ControlhWnd, EM_LINEFROMCHAR, ByVal lSelStart, ByVal 0&)
s = Space$(400)
SendMessage ControlhWnd, EM_GETLINE, n, ByVal s
s = Trim$(s)
i = Len(s)
TryAgain:
    Debug.Print i
    i = InStrRev(s, ".", i)
    If i = 0 Then GoTo BadLine
    n = InStrRev(s, " ", i)
    If n = 0 Then
        CurCom = Left$(s, i - 1)
    Else
        CurCom = Mid$(s, n + 1, i - n - 1)
    End If
    For n = LBound(udtKeyWords) To UBound(udtKeyWords)
        With udtKeyWords(n)
            If LCase$(CurCom) = LCase$(.sTrigger) Then
                If .sTrigger <> "" Then
                    If .sEndLine = "" Then
                        i = i - 1
                        GoTo TryAgain
                    End If
                    bFound = True
                    Exit For
                End If
            End If
        End With
        DoEvents
    Next
    If Not bFound Then
        i = i - 1
        GoTo TryAgain
    End If
    'Debug.Print CurCom
    IntelliSenseGetLastKeyWord = CurCom
BadLine:
End Function

Private Sub IntelliSenseAutoSizeIntelliSense()
Dim i As Long
Dim n As Long
Dim m As Long
Dim o As Long
Dim l As POINTAPI
For i = 0 To lstMain.ListCount - 1
    n = UserControl.TextWidth(lstMain.list(i))
    If n > m Then m = n
    DoEvents
Next
m = m + 23
If lstMain.ListCount < 10 Then
    o = (lstMain.ListCount * UserControl.TextHeight("X")) + 8
Else
    o = (10 * UserControl.TextHeight("X")) + 8
End If
With lstMain
    .Width = m
    .Height = o
End With
With picLstMain
    .Width = lstMain.Width + 6
    .Height = lstMain.Height + 6
End With
With rsMain
    .Width = lstMain.Width + 6
    .Height = lstMain.Height + 6
End With
End Sub

'Public Sub TipDetermineTip(lSelStart As Long, x As Long, y As Long, ControlhWnd As Long)
'Dim i As Long
'For i = LBound(udtKeyWords) To UBound(udtKeyWords)
'    If txtMain.SelStart - Len(udtKeyWords(i).sKeyWord) > 0 Then
'        If LCase$(Mid$(txtMain.Text, _
'            txtMain.SelStart - Len(udtKeyWords(i).sKeyWord) + 1, _
'            Len(udtKeyWords(i).sKeyWord))) = LCase$(udtKeyWords(i).sKeyWord) Then
'                If udtKeyWords(i).sKeyWord <> "" Then
'                    ShowTip udtKeyWords(i).sKeyWord, x, y, ControlhWnd, udtKeyWords(i).sTrigger
'                    Exit For
'                End If
'        End If
'    End If
'    DoEvents
'Next
'End Sub
'
'Public Sub ShowTip(sKeyWord As String, x As Long, y As Long, ControlhWnd As Long, sTrigger As String)
'    Dim i As Long
'    Dim r As RECT
'    Dim lBDWP As Long
'    For i = LBound(udtKeyWords) To UBound(udtKeyWords)
'        If LCase$(sKeyWord) = LCase$(udtKeyWords(i).sKeyWord) Then
'            If LCase$(sTrigger) = LCase$(udtKeyWords(i).sTrigger) Then
'                If udtKeyWords(i).sKeyWord <> "" Then
'                    'draw tip
'                    TipMode = True
'
'                    lstMain.Visible = False
'                    picMain.FontBold = True
'
'                        picMain.Height = picMain.TextHeight("X") + (Screen.TwipsPerPixelY \ 2)
'                        picMain.Width = picMain.TextWidth(udtKeyWords(i).sKeyWord & udtKeyWords(i).sMethods)
'
'                    picMain.FontBold = False
'
'                    picMain.Cls
'                    GetWindowRect picMain.hWnd, r
'                    With r
'                        .lLeft = x + (Screen.TwipsPerPixelX \ 2) + picSide.ScaleWidth
'                        .lTop = y + (Screen.TwipsPerPixelY)
'                        .lRight = picMain.ScaleWidth + .lLeft
'                        .lBottom = picMain.ScaleHeight + .lTop
'                    End With
'                    If r.lLeft > txtMain.Width - picMain.ScaleWidth Then
'                        r.lLeft = txtMain.Width - picMain.ScaleWidth
'                        r.lRight = r.lLeft + picMain.ScaleWidth
'                    End If
'                    If r.lTop > txtMain.Height - picMain.ScaleHeight Then
'                        r.lTop = txtMain.Height - picMain.ScaleHeight - ((UserControl.TextHeight("X") + 1) * 2)
'                        r.lBottom = r.lTop + picMain.ScaleHeight
'                    End If
'                    lBDWP = BeginDeferWindowPos(1)
'                    DeferWindowPos lBDWP, picMain.hWnd, HWND_TOP, r.lLeft, r.lTop, r.lRight - r.lLeft, r.lBottom - r.lTop, SWP_SHOWWINDOW
'                    EndDeferWindowPos lBDWP
'                    GetClientRect picMain.hWnd, r
'                    With picMain
'
'                        .FontBold = True
'                            r.lLeft = r.lLeft + 1
'                            DrawText .hdc, udtKeyWords(i).sKeyWord, Len(udtKeyWords(i).sKeyWord), r, DT_LEFT
'                            r.lLeft = r.lLeft + .TextWidth(udtKeyWords(i).sKeyWord)
'                        .FontBold = False
'
'                        Select Case Right$(udtKeyWords(i).sKeyWord, 1)
'                            Case "(", "="
'                                DrawText .hdc, udtKeyWords(i).sMethods, Len(udtKeyWords(i).sMethods), r, DT_LEFT
'                            Case Else
'                                DrawText .hdc, " " & udtKeyWords(i).sMethods, Len(" " & udtKeyWords(i).sMethods), r, DT_LEFT
'                        End Select
'
'                        r.lLeft = r.lLeft + .TextWidth(udtKeyWords(i).sMethods)
'
'                        .FontBold = True
'                            Select Case Right$(udtKeyWords(i).sKeyWord, 1)
'                                Case "("
'                                    DrawText .hdc, ")", Len(")"), r, DT_LEFT
'                            End Select
'                        .FontBold = False
'
'                    End With
'                    picMain.Refresh
'                    AlwaysOnTop picMain.hWnd, True
'                    TipIsShown = True
'                    Exit For
'                End If
'            End If
'        End If
'        DoEvents
'    Next
'    SetRectEmpty r
'End Sub
'
'Public Sub HideTip()
'    AlwaysOnTop picMain.hWnd, False
'    ShowWindow picMain.hWnd, SW_HIDE
'    TipIsShown = False
'    TipMode = False
'End Sub


Public Property Get Params() As String
Params = CMethod
End Property


Public Property Get Text() As String
Text = txtMain.Text
End Property

Public Property Let Text(ByVal s As String)
txtMain.Text = s
End Property

Public Sub IntelliSenseStartSubclassing()
modSub.SubclassTextbox txtMain
End Sub

Public Property Get Enabled() As Boolean
Enabled = txtMain.Enabled
End Property

Public Property Let Enabled(ByVal b As Boolean)
txtMain.Enabled = b
End Property

Public Sub DebugPrintKeywords()
Dim i As Long
For i = LBound(udtKeyWords) To UBound(udtKeyWords)
    Debug.Print udtKeyWords(i).sTrigger
Next
End Sub
