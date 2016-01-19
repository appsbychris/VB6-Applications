VERSION 5.00
Begin VB.UserControl UltraBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2925
   ScaleWidth      =   5805
   ToolboxBitmap   =   "ucListBox.ctx":0000
   Begin VB.HScrollBar HS 
      Height          =   255
      LargeChange     =   10
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.VScrollBar VS 
      Height          =   2055
      LargeChange     =   2
      Left            =   3360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1815
      ScaleWidth      =   3375
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "UltraBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'---------------------------------------------------------------------------------------
' Module    : UltraBox
' DateTime  : 120403 16:15
' Author    : Chris Van Hooser
' Copyright : 2002, Spike Technologies
' Purpose   :
'---------------------------------------------------------------------------------------
'
Option Explicit


'APIs, taken from API Guide
'Draw the edge
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'Fill rectangles
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'Create solid brushes
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Get the user controls rectangle
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'draw various windows controls
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
'draw focus RECTS
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
'draw text
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'Make a rect larger or smaller
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
'Clean up resources
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Offset a rect by a set amount
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
'Set the text color of a hdc
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'Get text dimensions
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
'Get system Colors
Private Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'Constants from API Viewer and API Guide
Private Const COLOR_SCROLLBAR = 0 'The Scrollbar colour
Private Const COLOR_BACKGROUND = 1 'Colour of the background with no wallpaper
Private Const COLOR_ACTIVECAPTION = 2 'Caption of Active Window
Private Const COLOR_INACTIVECAPTION = 3 'Caption of Inactive window
Private Const COLOR_MENU = 4 'Menu
Private Const COLOR_WINDOW = 5 'Windows background
Private Const COLOR_WINDOWFRAME = 6 'Window frame
Private Const COLOR_MENUTEXT = 7 'Window Text
Private Const COLOR_WINDOWTEXT = 8 '3D dark shadow (Win95)
Private Const COLOR_CAPTIONTEXT = 9 'Text in window caption
Private Const COLOR_ACTIVEBORDER = 10 'Border of active window
Private Const COLOR_INACTIVEBORDER = 11 'Border of inactive window
Private Const COLOR_APPWORKSPACE = 12 'Background of MDI desktop
Private Const COLOR_HIGHLIGHT = 13 'Selected item background
Private Const COLOR_HIGHLIGHTTEXT = 14 'Selected menu item
Private Const COLOR_BTNFACE = 15 'Button
Private Const COLOR_BTNSHADOW = 16 '3D shading of button
Private Const COLOR_GRAYTEXT = 17 'Grey text, of zero if dithering is used.
Private Const COLOR_BTNTEXT = 18 'Button text
Private Const COLOR_INACTIVECAPTIONTEXT = 19 'Text of inactive window
Private Const COLOR_BTNHIGHLIGHT = 20 '3D highlight of button
Private Const COLOR_2NDACTIVECAPTION = 27 'Win98 only: 2nd active window color
Private Const COLOR_2NDINACTIVECAPTION = 28 'Win98 only: 2nd inactive window color


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

Private Const DFC_BUTTON As Long = 4

Private Const DFCS_BUTTONCHECK As Long = &H0
Private Const DFCS_BUTTONRADIO = &H4

Private Const DFCS_CHECKED As Long = &H400

Private Const DT_CENTER As Long = &H1
Private Const DT_LEFT As Long = &H0
Private Const DT_RIGHT As Long = &H2

Private Type SIZE
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private R As RECT 'Client rectangle

'5 styles for the looks of the listbox
Public Enum View
    RaisedEdge = 0
    SunkenEdge = 1
    BumpedEdge = 2
    EtchedEdge = 3
    LineEdge = 4
    None = 5
End Enum
Public Enum FillStyle
    Normal = 0
    Lined = 1
End Enum
Public Enum SysMet
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
    SM_CXVSCROLL = 2
    SM_CYHSCROLL = 3
    SM_CYCAPTION = 4
    SM_CXBORDER = 5
    SM_CYBORDER = 6
    SM_CXDLGFRAME = 7
    SM_CYDLGFRAME = 8
    SM_CYHTHUMB = 9
    SM_CXHTHUMB = 10
    SM_CXICON = 11
    SM_CYICON = 12
    SM_CXCURSOR = 13
    SM_CYCURSOR = 14
    SM_CYMENU = 15
    SM_CXFULLSCREEN = 16
    SM_CYFULLSCREEN = 17
    SM_CYKANJIWINDOW = 18
    SM_MOUSEPRESENT = 19
    SM_CYVSCROLL = 20
    SM_CXHSCROLL = 21
    SM_DEBUG = 22
    SM_SWAPBUTTON = 23
    SM_CXMIN = 24
    SM_CYMIN = 25
    SM_CXSIZE = 26
    SM_CYSIZE = 27
    SM_CXMINTRACK = 28
    SM_CYMINTRACK = 29
    SM_CXDOUBLECLK = 30
    SM_CYDOUBLECLK = 31
    SM_CXICONSPACING = 32
    SM_CYICONSPACING = 33
    SM_MENUDROPALIGNMENT = 34
    SM_PENWINDOWS = 35
    SM_DBCSENABLED = 36
    SM_CMOUSEBUTTONS = 37
    SM_CMETRICS = 38
    SM_CLEANBOOT = 39
    SM_CXMAXIMIZED = 40
    SM_CXMAXTRACK = 41
    SM_CXMENUCHECK = 42
    SM_CXMENUSIZE = 43
    SM_CXMINIMIZED = 44
    SM_CYMAXIMIZED = 45
    SM_CYMAXTRACK = 46
    SM_CYMENUCHECK = 47
    SM_CYMENUSIZE = 48
    SM_CYMINIMIZED = 49
    SM_CYSMCAPTION = 50
    SM_MIDEASTENABLED = 51
    SM_NETWORK = 52
    SM_SLOWMACHINE = 53
End Enum
'Each item in the listbox has these properties
Private Type ListStyle
    sCaption As String
    bSelected As Boolean
    lHighlightColor As OLE_COLOR
    lForeColor As OLE_COLOR
    lBackColor As OLE_COLOR
    lHightlightText As OLE_COLOR
    bUseCheckBox As Boolean
    bUseOptionBox As Boolean
    iCheck As Integer
    iOpt As Integer
    lOptionGroup As Long
    bEnabled As Boolean
    bUseProgress As Boolean
    lProgressMax As Long
    lProgressValue As Long
    lProgressBarColor As OLE_COLOR
    lAlignment As Long
    lDrawFR As Long
End Type

Private Type TextObj
    lWidTWIPS As Long
    lWidPIXELS As Long
    lHeiTWIPS As Long
    lHeiPIXELS As Long
End Type
Private SystemTextColor As Long
Private SystemHighlightTextColor As Long
Private CurStyle As View  'styles
Private CurColor As OLE_COLOR 'BG colors
Private lB As Long 'Color
Private aList() As ListStyle 'List items
Private lTop As Long 'Top item index
Private lSelected As Long 'Which item is selected
'Private bHasSB As Boolean 'Have the little thingy on the scrollbar
Private MaxLen As Long
Private udtT As TextObj
'Basic events
Public Event Click()
Public Event DoubleClick()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event VerticalScroll(lValue As Long)
Public Event HorizontalScroll(lValue As Long)
Public Event Change()
Private bPaint As Boolean
Private fFill As FillStyle
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

Private Sub HS_Change()
DrawInit
RaiseEvent HorizontalScroll(HS.Value)
End Sub

Private Sub HS_Scroll()
DrawInit
RaiseEvent HorizontalScroll(HS.Value)
End Sub

Private Sub VS_Change()
DrawInit
RaiseEvent VerticalScroll(VS.Value)
End Sub

Private Sub VS_Scroll()
DrawInit
RaiseEvent VerticalScroll(VS.Value)
End Sub

Private Sub picMain_Click()
RaiseEvent Click
End Sub

Private Sub picMain_DblClick()
RaiseEvent DoubleClick
End Sub

Private Sub picMain_GotFocus()
   On Error GoTo picMain_GotFocus_Error

If lSelected = -1 Then Exit Sub
aList(lSelected).lDrawFR = 0
DrawInit

   On Error GoTo 0
   Exit Sub

picMain_GotFocus_Error:
    lSelected = -1
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picMain_GotFocus of User Control UltraBox"
End Sub

Private Sub picMain_LostFocus()
If lSelected = -1 Then Exit Sub
aList(lSelected).lDrawFR = 1
DrawInit
End Sub

Private Sub UserControl_Initialize()
With udtT
    .lHeiPIXELS = (ScaleY(picMain.TextHeight("X" & vbCrLf & "W" & vbCrLf & "y" & vbCrLf & "g"), 1, 3) \ 4) + 2
    .lHeiTWIPS = (picMain.TextHeight("X" & vbCrLf & "W" & vbCrLf & "y" & vbCrLf & "g") \ 4) + ScaleY(2, 3, 1)
    .lWidPIXELS = ScaleX(picMain.TextWidth("XW"), 1, 3) \ 2
    .lWidTWIPS = picMain.TextWidth("XW") \ 2
End With
SystemTextColor = GetSysColor(COLOR_WINDOWTEXT)
SystemHighlightTextColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
lSelected = -1
bPaint = True
DrawInit
End Sub

Private Sub picMain_KeyDown(KeyCode As Integer, Shift As Integer)
'For moving the selection up and down.
Dim i&
Dim n&
   On Error GoTo picMain_KeyDown_Error

For i& = 1 To UBound(aList)
    With aList(i)
        'Make sure something is selected
        If .bSelected = True Then Exit For
        If i& = UBound(aList) Then Exit Sub
    End With
    DoEvents
Next
Select Case KeyCode
    Case vbKeyDown
        'Find the next item, going down the list
        For i& = 1 To UBound(aList)
            With aList(i)
                If .bSelected = True Then
                    If i + 1 <= UBound(aList) Then
                        .bSelected = False
                        n& = i
                        Do Until lSelected <> i Or n& > UBound(aList)
                            If aList(n& + 1).bEnabled = True Then
                                aList(n& + 1).bSelected = True
                                If UBound(aList) - DeterVisible + 1 < n& + 1 Then
                                    lTop = UBound(aList) - DeterVisible + 1
                                Else
                                    lTop = n& + 1
                                    VS.Value = n& '+ 1
                                End If
                                lSelected = n& + 1
                            End If
                            n& = n& + 1
                            DoEvents
                        Loop
                    End If
                    Exit For
                End If
            End With
            DoEvents
        Next
    Case vbKeyUp
        'Find thenext item going up the list
        For i& = 1 To UBound(aList)
            With aList(i)
                If .bSelected = True Then
                    If i - 1 >= LBound(aList) + 1 Then
                        .bSelected = False
                        n& = i&
                        Do Until lSelected <> i Or n& < 1
                            If aList(n& - 1).bEnabled = True Then
                                aList(n& - 1).bSelected = True
                                If UBound(aList) - DeterVisible + 1 < n& - 1 Then
                                    lTop = UBound(aList) - DeterVisible + 1
                                Else
                                    lTop = n& - 1
                                    VS.Value = n& - 2
                                End If
                                lSelected = n& - 1
                            End If
                            n& = n& - 1
                            DoEvents
                        Loop
                    End If
                    Exit For
                End If
            End With
            DoEvents
        Next
End Select
DrawInit 'Refresh
RaiseEvent Change
RaiseEvent Click
   On Error GoTo 0
   Exit Sub

picMain_KeyDown_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picMain_KeyDown of User Control UltraBox"
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'To find out where/what they clicked
Dim l%, i&
l% = GetItemClick(x, y)
If l% <> 0 Then
    For i& = 1 To UBound(aList)
        With aList(i)
            If i& = l% Then
                .bSelected = True 'Find the item they clicked and selected it
                lSelected = i&
            Else
                .bSelected = False 'and make all the others not selected
            End If
        End With
        DoEvents
    Next
End If
DrawInit
RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
DrawInit
RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Resize()
With picMain
    .Top = 0
    .Left = 0
    .Width = UserControl.ScaleWidth
    .Height = UserControl.ScaleHeight
End With
DrawInit 'refresh
End Sub

Private Sub UserControl_Terminate()

   On Error GoTo UserControl_Terminate_Error

Erase aList 'Clear up memory

   On Error GoTo 0
   Exit Sub

UserControl_Terminate_Error:
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
'Propertys to retrieve
With PropBag
    CurStyle = .ReadProperty("Style")
    CurColor = .ReadProperty("Color")
    fFill = .ReadProperty("Fill")
    If Not .ReadProperty("Font") Is Nothing Then Set picMain.Font = .ReadProperty("Font")
End With
With udtT
    .lHeiPIXELS = ScaleY(picMain.TextHeight("X" & vbCrLf & "W" & vbCrLf & "y"), 1, 3) \ 3
    .lHeiTWIPS = picMain.TextHeight("X" & vbCrLf & "W" & vbCrLf & "y") \ 3
    .lWidPIXELS = ScaleX(picMain.TextWidth("XW"), 1, 3) \ 2
    .lWidTWIPS = picMain.TextWidth("XW") \ 2
End With
DrawInit
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'properties to save
With PropBag
    .WriteProperty "Style", CurStyle
    .WriteProperty "Color", CurColor
    .WriteProperty "Fill", fFill
    .WriteProperty "Font", picMain.Font
End With
End Sub

'================
'PROPERTIES

'Backcolor
Public Property Get Color() As OLE_COLOR
Color = CurColor
End Property

Public Property Let Color(ByVal NewColor As OLE_COLOR)
CurColor = NewColor
DrawInit
End Property

'Current look
Public Property Get Style() As View
Style = CurStyle
End Property

Public Property Let Style(ByVal NewStyle As View)
CurStyle = NewStyle
DrawInit
PropertyChanged "Style"
End Property

'Font of the items
Public Property Get Font() As StdFont
Set Font = picMain.Font
End Property

Public Property Set Font(ByVal f As StdFont)
Set picMain.Font = f
PropertyChanged "Font"
With udtT
    .lHeiPIXELS = ScaleY(picMain.TextHeight("X" & vbCrLf & "W" & vbCrLf & "y"), 1, 3) \ 3
    .lHeiTWIPS = picMain.TextHeight("X" & vbCrLf & "W" & vbCrLf & "y") \ 3
    .lWidPIXELS = ScaleX(picMain.TextWidth("XW"), 1, 3) \ 2
    .lWidTWIPS = picMain.TextWidth("XW") \ 2
End With
End Property
'================


'================
'PRIVATE FUNCTIONS/SUBS

Private Function GetItemClick(x As Single, y As Single) As Integer
'To determine what thing they click in the listbox
Dim lH          As Long     'Letter height
Dim i           As Long     'Counter
Dim b           As RECT     'Rect for positioning of things
Dim lCounter    As Long     'Counter
Dim lN          As Long     'Text height in PIXELS
Dim j           As Long     'Counter
Dim bBNeedSB    As Boolean  'If they need the scrollbar
Dim bNeedHSB    As Boolean
Dim lOFF        As Long
   On Error GoTo GetItemClick_Error

bBNeedSB = DoINeedaVSB 'Check if they need it
bNeedHSB = DoINeedaHSB
If bNeedHSB Then lOFF = -HS.Value
If bBNeedSB Then lTop = VS.Value + 1
lH = udtT.lHeiTWIPS 'picMain.TextHeight("X") 'Get the height in TWIPS
lN = udtT.lHeiPIXELS  'Height in PIXELS
lCounter = 0 'Set to 0
For i = lTop To UBound(aList) 'Loop from where the list begins
    With aList(i)
        If .bEnabled = True Then 'Make sure it is enabled
            'Check if they are using option/check boxes
            'on the item, so you can make the RECT structure
            'the correct size
            If (Not .bUseCheckBox) And (Not .bUseOptionBox) Then
                b.Left = 3
                If bNeedHSB Then OffsetRect b, lOFF, 0
            Else
                'If they are using the Check/Option 's,
                'Determine which one, and adjust and
                'check the rect area.
                If .bUseCheckBox Then
                    'Checkbox dimensions
                    With b
                        .Left = 3
                        .Top = 3 + ((lCounter) * ScaleY(lH, 1, 3))
                        .Right = b.Left + 12
                        .Bottom = b.Top + ScaleY(lH, 1, 3) + 3
                        If bNeedHSB Then OffsetRect b, lOFF, 0
                        .Left = ScaleX(.Left, 3, 1) 'TWIPS
                        .Top = ScaleY(.Top, 3, 1) 'TWIPS
                        .Right = ScaleX(.Right, 3, 1) 'TWIPS
                        .Bottom = ScaleY(.Bottom, 3, 1) 'TWIPS
                    End With
                    'Check if they clicked in the Checkbox
                    If (x > 1 And x < b.Right) And (y > b.Top And y < b.Bottom) Then
                        Select Case .iCheck
                            Case 0
                                .iCheck = 1 'Check it
                            Case 1
                                .iCheck = 0 'UnCheck it
                        End Select
                        Exit For
                    End If
                    'If they didn't click in the Checkbox, adjust the RECT
                    'To be able to check for an item/ other things
                    b.Left = b.Right + 1
                ElseIf .bUseOptionBox Then
                    'If an option box
                    'Option box dimensions
                    With b
                        .Left = 3
                        .Top = 3 + ((lCounter) * ScaleY(lH, 1, 3))
                        .Right = b.Left + 12
                        .Bottom = b.Top + ScaleY(lH, 1, 3) + 3
                        If bNeedHSB Then OffsetRect b, lOFF, 0
                        'TWIPS
                        .Left = ScaleX(.Left, 3, 1)
                        .Top = ScaleY(.Top, 3, 1)
                        .Right = ScaleX(.Right, 3, 1)
                        .Bottom = ScaleY(.Bottom, 3, 1)
                    End With
                    'If they clicked on it
                    If (x > 1 And x < b.Right) And (y > b.Top And y < b.Bottom) Then
                        Select Case .iOpt
                            Case 0 'If it isn't selected,
                                'We have to loop through all the items
                                'and take out whatever is in the option
                                'group, and unselected them, since an
                                'option button allows only 1 choice.
                                For j = 1 To UBound(aList)
                                    With aList(j)
                                        If .bUseOptionBox Then
                                            If .lOptionGroup = aList(i).lOptionGroup Then
                                                .iOpt = 0
                                            End If
                                        End If
                                    End With
                                    DoEvents
                                Next
                                .iOpt = 1 'Select this one
        
                        End Select
                        Exit For
                    End If
                    'set up the rect if the option box wasn't click
                    'for checking of others items
                    b.Left = b.Right + 1
                End If
            End If
        End If
        'Adjust the RECT accordinly, depending on if they
        'have a scroll bar showing.
        If bBNeedSB Then
            b.Right = picMain.Width - ScaleY(GetVSBWidth, 3, 1) - ScaleX(2, 3, 1)
        Else
            b.Right = picMain.Width - ScaleX(3, 3, 1)
        End If
        'Set up the RECT for the 1 item, dpeneding on
        'how far down the list we are.
        b.Top = 1 + ((lCounter) * lH)
        b.Bottom = b.Top + lH
        If (x > 1 And x < b.Right) And (y > b.Top And y < b.Bottom) And .bEnabled Then
            GetItemClick = i 'If they clicked the item, set the
            'value to the function, and get out of the loop
            Exit For
        End If
        lCounter = lCounter + 1 'Increase the counter
        'We don't want to check for invisible items, so if
        'we have check more then are visible, get out of
        'the loop
        If lCounter >= DeterVisible Then Exit For
    End With
Next

   On Error GoTo 0
   Exit Function

GetItemClick_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetItemClick of User Control UltraBox"
End Function



Private Sub DrawList()
Dim i           As Long     'Counter
Dim b           As RECT     'Item's RECT structure
Dim lH          As Long     'Text height
Dim lC          As Long     'Color
Dim bBHasSB     As Boolean  'Scroll bar flag
Dim lCounter    As Long     'Counter
Dim bBNeedSB    As Boolean  'Need the scroll bar flag
Dim v           As Long     'Color
Dim t           As Long     'For Instr
Dim tArr()      As String   'Temp array for dif colored strings
Dim bNHSB       As Boolean
Dim j           As Long
Dim pb          As RECT
Dim lVal        As Single
Dim lOFF        As Long
Dim lCol        As Long
On Error GoTo DrawList_Error

'Check the limitations
If UBound(aList) = 0 Then HS.Visible = False: VS.Visible = False: Exit Sub
If lTop = 0 Then lTop = 1
If lTop > UBound(aList) Then lTop = UBound(aList)
'Now, see if they need a scroll bar
bBNeedSB = DoINeedaVSB
bNHSB = DoINeedaHSB
If bNHSB Then lOFF = -HS.Value: DrawHSB Else HS.Visible = False
If bBNeedSB Then lTop = VS.Value + 1: DrawSB Else VS.Visible = False
lCounter = 0 'Set the counter to 0
'Get the height of a letter in PIXELS
lH = udtT.lHeiPIXELS
For i = lTop To UBound(aList) 'Loop from the first visible one
    With aList(i)
        'If they aren't using a check or option box,
        'adjust the rect accordenly
        If (Not .bUseCheckBox) And (Not .bUseOptionBox) Then
            b.Left = 3
            If bNHSB Then OffsetRect b, lOFF, 0
        Else
            'Since they are using a check/option box...
            If .bUseCheckBox Then
                'Get the checkbox dimensions
                With b
                    .Left = 3
                    .Top = 3 + ((lCounter) * lH)
                    .Right = .Left + 12
                    .Bottom = .Top + lH + 3
                End With
                If bNHSB Then OffsetRect b, lOFF, 0
                Select Case .iCheck
                    Case 0 'Unchecked
                        DrawFrameControl picMain.hdc, b, DFC_BUTTON, DFCS_BUTTONCHECK
                    Case 1 'Checked
                        DrawFrameControl picMain.hdc, b, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_CHECKED
                End Select
                b.Left = b.Right + 1 'Prepare for text with 1 pixel spacing
            ElseIf .bUseOptionBox Then
                'Option box
                With b
                    .Left = 3
                    .Top = 3 + ((lCounter) * lH)
                    .Right = .Left + 12
                    .Bottom = .Top + lH + 3
                End With
                If bNHSB Then OffsetRect b, lOFF, 0
                Select Case .iOpt
                    Case 0 'Unselected
                        DrawFrameControl picMain.hdc, b, DFC_BUTTON, DFCS_BUTTONRADIO
                    Case 1 'Selected
                        DrawFrameControl picMain.hdc, b, DFC_BUTTON, DFCS_BUTTONRADIO Or DFCS_CHECKED
                End Select
                'Prepare for more text
                b.Left = b.Right + 1
            End If
        End If
        'If the need a scrollbar
        'adjust the Rect 18 pixels smaller
        With b
            If bBNeedSB Then
                .Right = ScaleX(picMain.Width - 1, 1, 3) - GetVSBWidth - 2
            Else
                .Right = ScaleX(picMain.Width - 1, 1, 3) - 3
            End If
            .Top = 3 + ((lCounter) * lH)
            .Bottom = .Top + lH
        End With
        If .bSelected = False Then 'If the items isn't selected
            'Create a solid brush of the backcolor they chose for the item
            If fFill = Normal Then
                lC = CreateSolidBrush(.lBackColor)
            Else
                If lCol = 0 Then
                    lC = CreateSolidBrush(vbWhite)
                    lCol = 1
                Else
                    lC = CreateSolidBrush(&HEFEFEF)
                    lCol = 0
                End If
            End If
            'Fill the rect with that color
            FillRect picMain.hdc, b, lC
            'Clean up resources
            DeleteObject lC
            '///////////////Progressbar//////////////////
            If .bUseProgress Then
                lC = CreateSolidBrush(.lProgressBarColor)
                CopyRect pb, b
                If bNHSB Then OffsetRect pb, lOFF, 0
                DrawEdge picMain.hdc, pb, EDGE_ETCHED, BF_RECT
                InflateRect pb, -2, -2
                lVal = .lProgressValue / .lProgressMax
                lVal = Round(lVal, 2)
                pb.Right = pb.Right * lVal
                FillRect picMain.hdc, pb, lC
                DeleteObject lC
                b.Left = b.Left + 2
            End If
            
            'If it is enabled, set the
            'forecolor accordenly
            If .bEnabled Then
                SetTextColor picMain.hdc, .lForeColor
            Else
                'Get the default GRAYed out color of the system
                SetTextColor picMain.hdc, GetSysColor(COLOR_GRAYTEXT)
            End If
            'Check if they are using custom colors...
            t = InStr(1, .sCaption, "}")
            If t = 0 Then 'If not, just draw out the text
                DrawText picMain.hdc, aList(i).sCaption, Len(aList(i).sCaption), b, .lAlignment
            Else
                'If they are, split it into an array
                tArr = Split(.sCaption, "}")
                'Loop through the array
                For j = LBound(tArr) To UBound(tArr)
                    If tArr(j) <> "" Then 'Make sure there is text in it
                                            'before doing anything
                        'Now, call my function to get the LONG
                        'value of the color from the string
                        'This is kind of how a string will split.
                        '0-
                        '1-}1This }2Is a }1T}2e}3s}4test
                        '2-1 This
                        '3-2Is a
                        '4-1 t
                        '5-2e
                        '6-3 s
                        '7-4 test
                        v = GetTextColorFromNum(Left$(tArr(j), 1))
                        If v = -1 Then '-1 means no custom color, and
                                    'the } found was suppose to be at
                                    'the front, so put it back
                            tArr(j) = "}" & tArr(j)
                        ElseIf j <> 0 Then 'If it is not the FIRST item
                                           'In the array, set the custom
                                           'color, and trim off the color
                                           'number
                            SetTextColor picMain.hdc, v
                            tArr(j) = Mid$(tArr(j), 2)
                        End If
                        'Draw the text of that color
                        DrawText picMain.hdc, tArr(j), Len(tArr(j)), b, .lAlignment 'DT_LEFT
                        'Offset the rect by tthe amount of the text printed out.
                        OffsetRect b, ScaleX(picMain.TextWidth(tArr(j)), 1, 3), 0
                    End If
                Next
            End If
        Else
            'If the item is selected
            'Get the highlight color and make a brush
            lC = CreateSolidBrush(.lHighlightColor)
            If lCol = 1 Then lCol = 0 Else lCol = 1
            'Fill the items RECT with the color
            FillRect picMain.hdc, b, lC
            'Clean up resources
            DeleteObject lC
            '///////////////Progressbar//////////////////
            If .bUseProgress Then
                lC = CreateSolidBrush(.lProgressBarColor)
                CopyRect pb, b
                DrawEdge picMain.hdc, pb, EDGE_ETCHED, BF_RECT
                InflateRect pb, -2, -2
                lVal = .lProgressValue / .lProgressMax
                lVal = Round(lVal, 2)
                pb.Right = pb.Right * lVal
                FillRect picMain.hdc, pb, lC
                DeleteObject lC
                InvertRect picMain.hdc, pb
            End If
            'Draw the focus rect around the item
            If .lDrawFR = 0 Then
                SetTextColor picMain.hdc, vbBlack
                DrawFocusRect picMain.hdc, b
            End If
            If .bUseProgress Then b.Left = b.Left + 2
            'And change the text color to the HIGHLIGHT text color
            SetTextColor picMain.hdc, .lHightlightText
            t = InStr(1, .sCaption, "}") 'See if there is custom colors
            If t = 0 Then 'If not, just put out the text
                DrawText picMain.hdc, aList(i).sCaption, Len(aList(i).sCaption), b, .lAlignment
            Else
                'Else, use custom colors
                'See above for explanation
                tArr = Split(.sCaption, "}")
                For j = LBound(tArr) To UBound(tArr)
                    If tArr(j) <> "" Then
                        v = GetTextColorFromNum(Left$(tArr(j), 1))
                        If v = -1 Then
                            tArr(j) = "}" & tArr(j)
                        Else
                            SetTextColor picMain.hdc, v
                            tArr(j) = Mid$(tArr(j), 2)
                        End If
                        DrawText picMain.hdc, tArr(j), Len(tArr(j)), b, .lAlignment 'DT_LEFT
                        OffsetRect b, ScaleX(picMain.TextWidth(tArr(j)), 1, 3), 0
                    End If
                Next
            End If
        End If
        lCounter = lCounter + 1 'Increase the counter
        'We don't want to go over the amount of
        'visible items, so check that
        If lCounter >= DeterVisible Then
            'if we are over, draw the scroll bar if needed
            'and exit the loop.
            If bNHSB Then DrawHSB
            If Not bBHasSB And bBNeedSB Then DrawSB: bBHasSB = True
            Exit For
        End If
    End With
Next
   On Error GoTo 0
   Exit Sub

DrawList_Error:
    
End Sub

Private Function GetVSBWidth() As Long
GetVSBWidth = GetSysMetrics(SM_CXVSCROLL)
End Function

Private Function GetHSBHeight() As Long
GetHSBHeight = GetSysMetrics(SM_CYHSCROLL)
End Function

Private Function DoINeedaVSB() As Boolean
'This is to check if we need a scroll bar
Dim lH As Long
Dim lR As Long
Dim p As SIZE
   On Error GoTo DoINeedaVSB_Error

lR = picMain.Height
'Get the height of a letter
GetTextExtentPoint32 picMain.hdc, "X", 1, p
lH = p.y * UBound(aList) 'and take that height, and multiply it
                        'by how many items we have
lH = ScaleY(lH, 3, 1) 'make it in TWIPS
If lH >= lR Then DoINeedaVSB = True 'if it is more then the height of the
                                    'user control, we need one.

   On Error GoTo 0
   Exit Function

DoINeedaVSB_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DoINeedaVSB of User Control UltraBox"
End Function

Private Function DoINeedaHSB() As Boolean
Dim s As Long
Dim l As Long
s = MaxLen
If DoINeedaVSB Then l = GetVSBWidth + 2 Else l = 0
If s > picMain.Width - ScaleX(l, 3, 1) Then DoINeedaHSB = True
End Function

Private Function SetSysMetIndex(SysMetVal As SysMet) As Long
Select Case SysMetVal
    Case 0 To 23
        SetSysMetIndex = SysMetVal
    Case 24:
        SetSysMetIndex = 28
    Case 25:
        SetSysMetIndex = 29
    Case 26:
        SetSysMetIndex = 30
    Case 27:
        SetSysMetIndex = 31
    Case 28:
        SetSysMetIndex = 34
    Case 29:
        SetSysMetIndex = 35
    Case 30:
        SetSysMetIndex = 36
    Case 31:
        SetSysMetIndex = 37
    Case 32:
        SetSysMetIndex = 38
    Case 33:
        SetSysMetIndex = 39
    Case 34:
        SetSysMetIndex = 40
    Case 35:
        SetSysMetIndex = 41
    Case 36:
        SetSysMetIndex = 42
    Case 37:
        SetSysMetIndex = 43
    Case 38:
        SetSysMetIndex = 44
    Case 39:
        SetSysMetIndex = 67
    Case 40:
        SetSysMetIndex = 61
    Case 41:
        SetSysMetIndex = 59
    Case 42:
        SetSysMetIndex = 71
    Case 43:
        SetSysMetIndex = 54
    Case 44:
        SetSysMetIndex = 57
    Case 45:
        SetSysMetIndex = 62
    Case 46:
        SetSysMetIndex = 60
    Case 47:
        SetSysMetIndex = 72
    Case 48:
        SetSysMetIndex = 55
    Case 49:
        SetSysMetIndex = 58
    Case 50:
        SetSysMetIndex = 51
    Case 51:
        SetSysMetIndex = 74
    Case 52:
        SetSysMetIndex = 63
    Case 53:
        SetSysMetIndex = 73
End Select
End Function

Private Function GetSysMetrics(GetWhat As SysMet) As Long
GetSysMetrics = GetSystemMetrics(SetSysMetIndex(GetWhat))
End Function


Private Function GetTextColorFromNum(lNum As String) As Long
'Function to take a custom color, and make it into its LONG value
Select Case lNum
    Case "0"
        GetTextColorFromNum = &H808000 'LIGHTBLUE
    Case "1"
        GetTextColorFromNum = vbBlack
    Case "2"
        GetTextColorFromNum = vbWhite
    Case "3"
        GetTextColorFromNum = vbRed
    Case "4"
        GetTextColorFromNum = vbBlue
    Case "5"
        GetTextColorFromNum = vbGreen
    Case "6"
        GetTextColorFromNum = vbYellow
    Case "7"
        GetTextColorFromNum = &HC0C0C0 'GRAY
    Case "8"
        GetTextColorFromNum = &H80FF& 'ORAnGE
    Case "9"
        GetTextColorFromNum = &H800080 'PURPLE
    Case Else
        GetTextColorFromNum = -1 'Not a custom color
End Select
End Function

Private Function ReplaceColors(ByVal s As String) As String
'Replaces all the custom colros so the
'user can just get the text value of
'the item
s = Replace$(s, "}0", "")
s = Replace$(s, "}1", "")
s = Replace$(s, "}2", "")
s = Replace$(s, "}3", "")
s = Replace$(s, "}4", "")
s = Replace$(s, "}5", "")
s = Replace$(s, "}6", "")
s = Replace$(s, "}7", "")
s = Replace$(s, "}8", "")
s = Replace$(s, "}9", "")
ReplaceColors = s
End Function

Private Sub DrawInit()
'Begining of the drawing to the user control
Dim b As Long
If bPaint = False Then Exit Sub
GetClientRect picMain.hWnd, R 'Get the user controls RECT
picMain.Cls 'Clear the screen
If Color Then 'If they have a custom color
    'create a brush
    lB = CreateSolidBrush(CurColor)
    'And fill the user control with it
    FillRect picMain.hdc, R, lB
    DeleteObject lB 'Clean up resources
End If
'Draw the list
DrawList
'Draw a border depending on the style selected
Select Case CurStyle
    Case 0
        DrawEdge picMain.hdc, R, EDGE_RAISED, BF_RECT
    Case 1
        DrawEdge picMain.hdc, R, EDGE_SUNKEN, BF_RECT
    Case 2
        DrawEdge picMain.hdc, R, EDGE_BUMP, BF_RECT
    Case 3
        DrawEdge picMain.hdc, R, EDGE_ETCHED, BF_RECT
    Case 4
        b = CreateSolidBrush(vbBlack)
        FrameRect picMain.hdc, R, b
        DeleteObject b
End Select
picMain.Refresh 'Refresht he user control
End Sub

Private Sub DrawSB()
'Drawing the scroll bar
VS.Visible = True
VS.Top = ScaleY(3, 3, 1)
VS.Height = UserControl.Height - ScaleY(6, 3, 1)
VS.Width = ScaleX(GetVSBWidth, 3, 1)
VS.Left = picMain.Width - VS.Width - ScaleX(2, 3, 1)
VS.Max = UBound(aList) - DeterVisible
End Sub

Private Sub DrawHSB()
Dim HH As Long
Dim vW As Long
HH = GetHSBHeight
vW = GetVSBWidth
HS.Visible = True
HS.Width = picMain.Width - ScaleX(6, 3, 1)
If DoINeedaVSB Then HS.Width = HS.Width - ScaleX(vW + 4, 3, 1)
HS.Height = ScaleY(HH, 3, 1)
HS.Top = picMain.Height - HS.Height - ScaleY(3, 3, 1)
HS.Left = ScaleX(3, 3, 1)
HS.Max = ScaleX(MaxLen, 1, 3) - ScaleX(picMain.Width, 1, 3) + (vW + 8)
End Sub

Private Function DeterVisible() As Long
Dim lH As Long
Dim lY As Long
'Determines how many items are visible at 1 time
lH = udtT.lHeiTWIPS
If DoINeedaHSB Then lY = picMain.Height - HS.Height - ScaleY(6, 3, 1) Else lY = picMain.Height - ScaleY(6, 3, 1)
lH = lY \ lH
DeterVisible = lH
End Function
'===================

'===================
'PUBLIC FUNCTIONS

Public Function ListIndex() As Long
ListIndex = lSelected
End Function

Public Function ItemText() As String
'Gets the currently selected item's text
If lSelected = -1 Then Exit Function
ItemText = ReplaceColors(aList(lSelected).sCaption)
End Function

Public Function ListCount() As Long
On Error GoTo ListCount_Error
'Gets the listcount
ListCount = UBound(aList)

On Error GoTo 0
Exit Function
ListCount_Error:
    ReDim aList(0)
    ListCount = 0
End Function

Public Function list(Index As Long) As String
On Error GoTo ItemTextFromIndex_Error
'Get a specific item's text
list = ReplaceColors(aList(Index).sCaption)

On Error GoTo 0
Exit Function

ItemTextFromIndex_Error:
    
End Function

Public Sub RemoveItem(Index As Long)
On Error GoTo RemoveItem_Error
'Removes an item
Dim i As Long
If Index <> 0 Then
    For i = Index To UBound(aList) - 1
        aList(i) = aList(i + 1)
    Next
    ReDim Preserve aList(UBound(aList) - 1)
    DrawInit 'refresh the screen
End If

On Error GoTo 0
Exit Sub
RemoveItem_Error:
    
End Sub

Public Function IsSelected(Index As Long, Optional CheckAndOptionOnly As Boolean = False) As Boolean
On Error GoTo IsSelected_Error
'Will check if a certain item is selected
If Not CheckAndOptionOnly Then
    'This will flag selected if it is CHECKED, SELECTED, or the option box is CLICKED
    If aList(Index).bSelected = True Or aList(Index).iCheck = 1 Or aList(Index).iOpt = 1 Then
        IsSelected = True
    End If
Else
    'This will only check option boxes and check boxes
    If aList(Index).iCheck = 1 Or aList(Index).iOpt = 1 Then
        IsSelected = True
    End If
End If
   On Error GoTo 0
   Exit Function

IsSelected_Error:
End Function

Public Sub SetProgressValue(Index As Long, NewValue As Long)

   On Error GoTo SetProgressValue_Error

    With aList(Index)
        If .bUseProgress Then
            .lProgressValue = NewValue
            If .lProgressValue > .lProgressMax Then .lProgressValue = .lProgressMax
        End If
    End With
    DrawInit
   On Error GoTo 0
   Exit Sub

SetProgressValue_Error:

End Sub

Public Function GetProgressValue(Index As Long) As Long
   On Error GoTo GetProgressValue_Error

    With aList(Index)
        If .bUseProgress Then
            GetProgressValue = .lProgressValue
        End If
    End With
    DrawInit
   On Error GoTo 0
   Exit Function

GetProgressValue_Error:

End Function

Public Function GetProgressMax(Index As Long) As Long
   On Error GoTo GetProgressValue_Error

    With aList(Index)
        If .bUseProgress Then
            GetProgressMax = .lProgressMax
        End If
    End With
    DrawInit
   On Error GoTo 0
   Exit Function

GetProgressValue_Error:

End Function

Public Sub SetSelected(Index As Long, ByVal Selected As Boolean, Optional MoveSelectionToTop As Boolean = False)
Dim i As Long
On Error GoTo SetSelected_Error
'Sets a certain item selected
With aList(Index)
    If .bUseCheckBox = True Then .iCheck = IIf(Selected = True, 1, 0)
    If .bUseOptionBox = True Then .iOpt = IIf(Selected = True, 1, 0)
    .bSelected = Selected
    lSelected = Index
End With
'This will uncheck all other option boxs in the same group
For i = 1 To UBound(aList)
    If i <> Index Then
        With aList(i)
            If Selected = True Then
                If .lOptionGroup = aList(i).lOptionGroup Then
                    .iOpt = 0
                End If
                .bSelected = False
                            
            End If
        End With
    End If
Next
If Index > lTop Then
    If Index - 1 > VS.Max Then VS.Value = VS.Max Else If MoveSelectionToTop Then VS.Value = Index - 1
Else
    VS.Value = Index - 1
End If
DrawInit 'refresh

On Error GoTo 0
Exit Sub
SetSelected_Error:

End Sub

Public Sub SetEnabled(Index As Long, ByVal Enabled As Boolean)
On Error GoTo SetEnabled_Error
'This will set an item enabled or disabled
With aList(Index)
    .bEnabled = Enabled
    If Not Enabled Then .bSelected = False
End With
DrawInit 'refresh

On Error GoTo 0
Exit Sub
SetEnabled_Error:

End Sub

Public Sub SetItemText(Index As Long, NewText As String)
On Error GoTo SetItemText_Error
'Set an items text
With aList(Index)
    .sCaption = NewText
End With
DrawInit

On Error GoTo 0
Exit Sub
SetItemText_Error:
End Sub

Public Sub Clear()
ReDim aList(0)
lSelected = -1
DrawInit
End Sub

Public Sub AddItemProgressBar(sText As String, Optional Alignment As AlignmentConstants = vbLeftJustify, Optional Enabled As Boolean = True, Optional FCOLOR As OLE_COLOR = -1, Optional BCOLOR As OLE_COLOR = -1, Optional HCOLOR As OLE_COLOR = -1, Optional HTEXT As OLE_COLOR = -1, Optional ProgressBarMax As Long = 100, Optional ProgressBarValue As Long = 1, Optional ProgressBarProgressColor As OLE_COLOR = -1)
AddItem sText, Alignment, Enabled, FCOLOR, BCOLOR, HCOLOR, HTEXT, , , , True, ProgressBarMax, ProgressBarValue, ProgressBarProgressColor
End Sub

Public Sub AddItemOption(sText As String, Optional Alignment As AlignmentConstants = vbLeftJustify, Optional Enabled As Boolean = True, Optional FCOLOR As OLE_COLOR = -1, Optional BCOLOR As OLE_COLOR = -1, Optional HCOLOR As OLE_COLOR = -1, Optional HTEXT As OLE_COLOR = -1, Optional OptionGroup As Long = 0)
AddItem sText, Alignment, Enabled, FCOLOR, BCOLOR, HCOLOR, HTEXT, False, True, OptionGroup
End Sub

Public Sub AddItemCheck(sText As String, Optional Alignment As AlignmentConstants = vbLeftJustify, Optional Enabled As Boolean = True, Optional FCOLOR As OLE_COLOR = -1, Optional BCOLOR As OLE_COLOR = -1, Optional HCOLOR As OLE_COLOR = -1, Optional HTEXT As OLE_COLOR = -1)
AddItem sText, Alignment, Enabled, FCOLOR, BCOLOR, HCOLOR, HTEXT, True
End Sub

Public Sub AddItem(sText As String, Optional Alignment As AlignmentConstants = vbLeftJustify, Optional Enabled As Boolean = True, Optional FCOLOR As OLE_COLOR = -1, Optional BCOLOR As OLE_COLOR = -1, Optional HCOLOR As OLE_COLOR = -1, Optional HTEXT As OLE_COLOR = -1, Optional UseCheckBox As Boolean = False, Optional UseOptionBox As Boolean = False, Optional OptionGroup As Long = 0, Optional UseProgressBar As Boolean = False, Optional ProgressBarMax As Long = 100, Optional ProgressBarValue As Long = 1, Optional ProgressBarProgressColor As OLE_COLOR = -1)
On Error GoTo AddItem_Error
'Function to add items
'Add 1 more item
ReDim Preserve aList(UBound(aList) + 1)
With aList(UBound(aList))
    'Set all the properties
    .sCaption = Replace$(sText, vbCrLf, "") 'No enters
    If TextWidth(.sCaption) > MaxLen Then MaxLen = picMain.TextWidth(.sCaption)
    If BCOLOR = -1 Then
        .lBackColor = CurColor
    Else
        .lBackColor = BCOLOR
    End If
    If FCOLOR = -1 Then
        .lForeColor = SystemTextColor
    Else
        .lForeColor = FCOLOR
    End If
    If HTEXT = -1 Then
        .lHightlightText = SystemHighlightTextColor
    Else
        .lHightlightText = HTEXT
    End If
    If HCOLOR = -1 Then
        .lHighlightColor = GetSysColor(COLOR_HIGHLIGHT)
    Else
        .lHighlightColor = HCOLOR
    End If
    .bUseCheckBox = UseCheckBox
    .bUseOptionBox = UseOptionBox
    .bUseProgress = UseProgressBar
    .lProgressMax = ProgressBarMax
    .lProgressValue = ProgressBarValue
    If ProgressBarProgressColor = -1 Then
        .lProgressBarColor = .lHighlightColor
    Else
        .lProgressBarColor = ProgressBarProgressColor
    End If
    Select Case Alignment
        Case vbLeftJustify
            .lAlignment = DT_LEFT
        Case vbRightJustify
            .lAlignment = DT_RIGHT
        Case vbCenter
            .lAlignment = DT_CENTER
    End Select
    .lOptionGroup = OptionGroup
    .bEnabled = Enabled
    If .bUseCheckBox And .bUseOptionBox Then .bUseOptionBox = False
End With

If bPaint Then DrawInit

On Error GoTo 0
Exit Sub
'If there is an error, it is because aList() isn't
'dimensioned yet, so Redim it to 0 items
'And then continue
AddItem_Error:
    ReDim aList(0)
    ReDim Preserve aList(UBound(aList) + 1)
    With aList(UBound(aList))
        .sCaption = Replace$(sText, vbCrLf, "")
        If TextWidth(.sCaption) > MaxLen Then MaxLen = picMain.TextWidth(.sCaption)
        If BCOLOR = -1 Then
            .lBackColor = CurColor
        Else
            .lBackColor = BCOLOR
        End If
        If FCOLOR = -1 Then
            .lForeColor = SystemTextColor
        Else
            .lForeColor = FCOLOR
        End If
        If HTEXT = -1 Then
            .lHightlightText = SystemHighlightTextColor
        Else
            .lHightlightText = HTEXT
        End If
        If HCOLOR = -1 Then
            .lHighlightColor = GetSysColor(COLOR_HIGHLIGHT)
        Else
            .lHighlightColor = HCOLOR
        End If
        .bUseCheckBox = UseCheckBox
        .bUseOptionBox = UseOptionBox
        .lOptionGroup = OptionGroup
        .lProgressMax = ProgressBarMax
        .lProgressValue = ProgressBarValue
        If ProgressBarProgressColor = -1 Then
            .lProgressBarColor = .lHighlightColor
        Else
            .lProgressBarColor = ProgressBarProgressColor
        End If
        .lOptionGroup = OptionGroup
        .bEnabled = Enabled
        If .bUseCheckBox And .bUseOptionBox Then .bUseOptionBox = False
    End With
    DrawInit
End Sub

Public Property Get Paint() As Boolean
Paint = bPaint
End Property

Public Property Let Paint(ByVal b As Boolean)
bPaint = b
If bPaint = True Then DrawInit
End Property

Public Function Find(ByVal s As String) As Long
Dim i As Long
   On Error GoTo Find_Error

For i = LBound(aList) To UBound(aList)
    With aList(i)
        If .bEnabled = True Then
            If LCase$(Left$(.sCaption, Len(s))) = LCase$(s) Then
                Find = i
                Exit Function
            End If
        End If
    End With
Next
Find = 0

   On Error GoTo 0
   Exit Function

Find_Error:

    Find = 0
End Function

Public Function FindInStr(ByVal s As String) As Long
Dim i As Long
   On Error GoTo FindInStr_Error

For i = LBound(aList) To UBound(aList)
    With aList(i)
        If .bEnabled = True Then
            If InStr(1, LCase$(.sCaption), s) Then
                FindInStr = i
                Exit Function
            End If
        End If
    End With
Next
FindInStr = 0

   On Error GoTo 0
   Exit Function

FindInStr_Error:

    FindInStr = 0
End Function

Public Property Get FillView() As FillStyle
FillView = fFill
End Property

Public Property Let FillView(ByVal f As FillStyle)
fFill = f
PropertyChanged "Fill"
End Property


