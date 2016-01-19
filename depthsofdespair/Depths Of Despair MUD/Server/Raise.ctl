VERSION 5.00
Begin VB.UserControl Raise 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "Raise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : Raise
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
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

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private R As RECT
Public Enum Styles
    RaisedEdge = 0
    SunkenEdge = 1
    BumpedEdge = 2
    EtchedEdge = 3
    MonoButton = 4
    CheckedMonoButton = 5
    CheckedSunken = 6
End Enum
Public Event Click()
Dim CurStyle As Styles
Dim CurColor As OLE_COLOR
Private Rgn As Long
Private LB As Long

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
GetClientRect UserControl.hwnd, R
UserControl.Cls
If Color Then
    Rgn = CreateRectRgn(R.Left, R.Top, R.Right, R.Bottom)
    LB = CreateSolidBrush(CurColor)
    FillRgn UserControl.hdc, Rgn, LB
End If
Select Case CurStyle
    Case 0
        DrawEdge UserControl.hdc, R, EDGE_RAISED, BF_RECT
    Case 1
        DrawEdge UserControl.hdc, R, EDGE_SUNKEN, BF_RECT
    Case 2
        DrawEdge UserControl.hdc, R, EDGE_BUMP, BF_RECT
    Case 3
        DrawEdge UserControl.hdc, R, EDGE_ETCHED, BF_RECT
    Case 4
        DrawFrameControl UserControl.hdc, R, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_MONO
    Case 5
        DrawFrameControl UserControl.hdc, R, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_MONO Or DFCS_CHECKED
    Case 6
        DrawFrameControl UserControl.hdc, R, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_CHECKED
End Select
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
With PropBag
    CurStyle = .ReadProperty("Style")
    ', 0)
    CurColor = .ReadProperty("Color") ', &H8000000F)
End With
UserControl_Initialize
End Sub

Private Sub UserControl_Resize()
UserControl_Initialize
End Sub

Public Property Get Style() As Styles
Style = CurStyle
End Property

Public Property Let Style(ByVal NewStyle As Styles)
CurStyle = NewStyle
UserControl_Initialize
PropertyChanged "Style"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Style", CurStyle ', 0
    .WriteProperty "Color", CurColor ', &H8000000F
End With
End Sub

Public Property Get Color() As OLE_COLOR
Color = CurColor
End Property

Public Property Let Color(ByVal NewColor As OLE_COLOR)
CurColor = NewColor
UserControl_Initialize
End Property
