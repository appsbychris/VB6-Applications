VERSION 5.00
Begin VB.UserControl eButton 
   AccessKeys      =   "F"
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2910
   KeyPreview      =   -1  'True
   ScaleHeight     =   1590
   ScaleWidth      =   2910
End
Attribute VB_Name = "eButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : eButton
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Const BDR_RAISEDOUTER As Long = &H1
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BDR_SUNKENINNER As Long = &H8
Private Const BDR_RAISEDINNER = &H4
Private Const EDGE_BUMP As Long = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED As Long = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED As Long = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN As Long = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const BF_ADJUST As Long = &H2000
Private Const BF_BOTTOM As Long = &H8

Private Const BF_DIAGONAL As Long = &H10
Private Const BF_FLAT As Long = &H4000
Private Const BF_LEFT As Long = &H1
Private Const BF_MIDDLE As Long = &H800
Private Const BF_MONO As Long = &H8000
Private Const BF_RIGHT As Long = &H4
Private Const BF_SOFT As Long = &H1000
Private Const BF_TOP As Long = &H2
Private Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_TOPLEFT As Long = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT As Long = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT As Long = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT As Long = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_DIAGONAL_ENDBOTTOMLEFT As Long = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMRIGHT As Long = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Private Const BF_DIAGONAL_ENDTOPLEFT As Long = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Private Const BF_DIAGONAL_ENDTOPRIGHT As Long = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Private Const DFC_BUTTON As Long = 4
Private Const DFC_CAPTION As Long = 1
Private Const DFC_MENU As Long = 2
Private Const DFC_POPUPMENU As Long = 5
Private Const DFC_SCROLL As Long = 3
Private Const DFCS_BUTTON3STATE As Long = &H8
Private Const DFCS_BUTTONCHECK As Long = &H0
Private Const DFCS_BUTTONPUSH As Long = &H10
Private Const DFCS_BUTTONRADIO As Long = &H4
Private Const DFCS_BUTTONRADIOIMAGE As Long = &H1
Private Const DFCS_BUTTONRADIOMASK As Long = &H2
Private Const DFCS_CAPTIONCLOSE As Long = &H0
Private Const DFCS_CAPTIONHELP As Long = &H4
Private Const DFCS_CAPTIONMAX As Long = &H2
Private Const DFCS_CAPTIONMIN As Long = &H1
Private Const DFCS_CAPTIONRESTORE As Long = &H3
Private Const DFCS_MENUARROW As Long = &H0
Private Const DFCS_MENUBULLET As Long = &H2
Private Const DFCS_MENUCHECK As Long = &H1
Private Const DFCS_SCROLLCOMBOBOX As Long = &H5
Private Const DFCS_SCROLLDOWN As Long = &H1
Private Const DFCS_SCROLLLEFT As Long = &H2
Private Const DFCS_SCROLLRIGHT As Long = &H3
Private Const DFCS_SCROLLSIZEGRIP As Long = &H8
Private Const DFCS_SCROLLUP As Long = &H0
Private Const DFCS_ADJUSTRECT As Long = &H2000
Private Const DFCS_CHECKED As Long = &H400
Private Const DFCS_FLAT As Long = &H4000
Private Const DFCS_INACTIVE As Long = &H100
Private Const DFCS_MONO As Long = &H8000
Private Const DFCS_PUSHED As Long = &H200

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



Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private R As RECT
Private bDown As Boolean
Private bHover As Boolean
Private bHasDrawn As Boolean
Private oColor As OLE_COLOR
Private BCOLOR As OLE_COLOR
Private FCOLOR As OLE_COLOR
Public Enum eStyle
    Normal = 0
    Flat = 1
    InverseFlat = 2
    FlatNormal = 3
    SinkHover = 4
    RaiseHover = 5
End Enum

Public Enum CaptionAlignment
    AlignLeft = 0
    AlignRight = 1
    AlignCenter = 2
End Enum
    
Private CS As eStyle
Private CA As CaptionAlignment
Private sCap As String
Private bHasFiredEnter As Boolean
Private bHasFiredLeave As Boolean
Private bHasFocus As Boolean
Private fFont As StdFont
Private bBoldOnHighlight As Boolean
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnter()
Public Event MouseLeave()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)


Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
UserControl_Click
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_GotFocus()
bHasFocus = True
DrawButton
End Sub

Private Sub UserControl_Initialize()
DrawButton
SetAccessKey
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
If KeyAscii = 13 Then RaiseEvent Click
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
bHasFocus = False
DrawButton
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
bDown = True
DrawButton
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With UserControl
    If GetCapture <> .hwnd Then SetCapture .hwnd
    If X < 0 Or Y < 0 Or X > .ScaleWidth Or Y > .ScaleHeight Then    ' we're over the control
      ReleaseCapture
        bHasFiredEnter = False
        bHover = False
        bDown = False
        bHasDrawn = False
        DrawButton
        If Not bHasFiredLeave Then RaiseEvent MouseLeave
        bHasFiredLeave = True
    Else
        bHover = True
        If Not bHasDrawn Then DrawButton
        bHasDrawn = True
        If Not bHasFiredEnter Then RaiseEvent MouseEnter
        bHasFiredEnter = True
        bHasFiredLeave = False
    End If
End With

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bDown = False
DrawButton
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
DrawButton
End Sub

Private Sub UserControl_Resize()
'cmdB.Height = UserControl.Height
'cmdB.Width = UserControl.Width
DrawButton
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
With PropBag
    CS = .ReadProperty("Style", 0)
    sCap = .ReadProperty("Cap", "eButton")
    Set fFont = .ReadProperty("Font", UserControl.Font)
    oColor = .ReadProperty("hCol", vbBlue)
    BCOLOR = .ReadProperty("bCol", &H8000000F)
    FCOLOR = .ReadProperty("fCol", vbBlack)
    bBoldOnHighlight = .ReadProperty("BoldOn", True)
    CA = .ReadProperty("CA", 2)
End With
With UserControl
    .BackColor = BCOLOR
    .ForeColor = FCOLOR
    Set .Font = fFont
    SetAccessKey
End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Style", CS, 0
    .WriteProperty "Cap", sCap, "eButton"
    .WriteProperty "Font", fFont ', 'UserControl.Font
    .WriteProperty "hCol", oColor, vbBlue
    .WriteProperty "bCol", BCOLOR, &H8000000F
    .WriteProperty "fCol", FCOLOR, vbBlack
    .WriteProperty "BoldOn", bBoldOnHighlight, True
    .WriteProperty "CA", CA
End With
End Sub

Private Sub DrawButton()
Select Case CS
    Case 0
        'SetWindowLong cmdB.hwnd, GWL_STYLE, WS_CHILD Or BS_FLAT
        DrawAB EDGE_RAISED, EDGE_ETCHED
    Case 1
        DrawAB EDGE_ETCHED, EDGE_BUMP
        'SetWindowLong cmdB.hwnd, GWL_STYLE, WS_CHILD Or BS_FLAT
    Case 2
        DrawAB EDGE_BUMP, EDGE_ETCHED
    Case 3
        DrawAB EDGE_ETCHED, EDGE_RAISED
    Case 4
        DrawAB EDGE_RAISED, EDGE_SUNKEN
    Case 5
        DrawAB EDGE_SUNKEN, EDGE_RAISED
End Select
End Sub

Private Sub DrawAB(sSt1 As Long, sSt2 As Long)
Dim h As Long
Dim Al As Long
UserControl.Cls
h = UserControl.hdc
GetClientRect UserControl.hwnd, R
If bHover Then FillRect h, R, CreateSolidBrush(oColor)
DrawEdge h, R, sSt1, BF_RECT
If bHover Then
    InflateRect R, -1, -1
    DrawEdge h, R, sSt2, BF_RECT
    InflateRect R, 1, 1
    If bBoldOnHighlight Then If UserControl.FontBold = False Then UserControl.FontBold = True
Else
    If UserControl.FontBold = True Then UserControl.FontBold = False
End If
If bHasFocus Then
    InflateRect R, -4, -4
    DrawFocusRect h, R
    InflateRect R, 4, 4
End If
Select Case CA
    Case 0
        Al = DT_LEFT
    Case 1
        Al = DT_RIGHT
    Case 2
        Al = DT_CENTER
End Select
InflateRect R, -6, -6
If Not bDown Then
    DrawText h, sCap, Len(sCap), R, DT_VCENTER Or Al Or DT_SINGLELINE
Else
    OffsetRect R, 2, 2
    DrawText h, sCap, Len(sCap), R, DT_VCENTER Or Al Or DT_SINGLELINE
End If
End Sub

Private Sub SetAccessKey()
Dim i As Long
Dim b As Boolean
i = InStr(1, sCap, "&")
If i Then
    If Mid$(sCap, i + 1, 1) <> "&" Then
        UserControl.AccessKeys = Mid$(sCap, i + 1, 1)
        b = True
    Else
        Do Until i = 0
            i = InStr(i + 2, sCap, "&")
            If Mid$(sCap, i + 1, 1) <> "&" Then
                UserControl.AccessKeys = Mid$(sCap, i + 1, 1)
                b = True
                i = 0
            End If
        Loop
    End If
End If
If Not b Then UserControl.AccessKeys = ""
End Sub

Public Property Get Style() As eStyle
Style = CS
End Property

Public Property Let Style(ByVal eS As eStyle)
CS = eS
DrawButton
PropertyChanged "Style"
End Property

Public Property Get Caption() As String
Caption = sCap
End Property

Public Property Let Caption(ByVal s As String)
sCap = s
SetAccessKey
DrawButton
PropertyChanged "Cap"
End Property

Public Property Get Font() As StdFont
If fFont Is Nothing Then Set fFont = UserControl.Font
Set Font = fFont
End Property

Public Property Set Font(ByVal f As StdFont)
Set fFont = f
Set UserControl.Font = fFont
DrawButton
PropertyChanged "Font"
End Property

Public Property Get HighlightColor() As OLE_COLOR
HighlightColor = oColor
End Property

Public Property Let HighlightColor(ByVal o As OLE_COLOR)
oColor = o
PropertyChanged "hCol"
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = BCOLOR
End Property

Public Property Let BackColor(ByVal o As OLE_COLOR)
BCOLOR = o
UserControl.BackColor = BCOLOR
DrawButton
PropertyChanged "bCol"
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = FCOLOR
End Property

Public Property Let ForeColor(ByVal f As OLE_COLOR)
FCOLOR = f
UserControl.ForeColor = FCOLOR
DrawButton
PropertyChanged "fCol"
End Property

Public Property Get BoldOnHighlight() As Boolean
BoldOnHighlight = bBoldOnHighlight
End Property

Public Property Let BoldOnHighlight(ByVal b As Boolean)
bBoldOnHighlight = b
PropertyChanged "BoldOn"
End Property

Public Property Get CaptionAlign() As CaptionAlignment
CaptionAlign = CA
End Property

Public Property Let CaptionAlign(ByVal s As CaptionAlignment)
CA = s
DrawButton
PropertyChanged "CA"
End Property
