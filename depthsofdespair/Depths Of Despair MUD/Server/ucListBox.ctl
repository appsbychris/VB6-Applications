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
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : UltraBox
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
'UltraBox 2.0
'-----------------------------------------------------------------------------------
'
Option Explicit


'APIs, taken from API Guide
'Draw the edge
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'Fill rectangles
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
'Create solid brushes
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
'Get the user controls rectangle
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'draw various windows controls
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
'draw focus RECTS
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
'draw text
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'Make a rect larger or smaller
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
'Clean up resources
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Offset a rect by a set amount
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
'Set the text color of a hdc
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'Get system Colors
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Private Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long
'Constants from API Viewer and API Guide
'Private Const COLOR_SCROLLBAR = 0 'The Scrollbar colour
'Private Const COLOR_BACKGROUND = 1 'Colour of the background with no wallpaper
'Private Const COLOR_ACTIVECAPTION = 2 'Caption of Active Window
'Private Const COLOR_INACTIVECAPTION = 3 'Caption of Inactive window
'Private Const COLOR_MENU = 4 'Menu
'Private Const COLOR_WINDOW = 5 'Windows background
'Private Const COLOR_WINDOWFRAME = 6 'Window frame
'Private Const COLOR_MENUTEXT = 7 'Window Text
Private Const COLOR_WINDOWTEXT = 8 '3D dark shadow (Win95)
'Private Const COLOR_CAPTIONTEXT = 9 'Text in window caption
'Private Const COLOR_ACTIVEBORDER = 10 'Border of active window
'Private Const COLOR_INACTIVEBORDER = 11 'Border of inactive window
'Private Const COLOR_APPWORKSPACE = 12 'Background of MDI desktop
Private Const COLOR_HIGHLIGHT = 13 'Selected item background
Private Const COLOR_HIGHLIGHTTEXT = 14 'Selected menu item
'Private Const COLOR_BTNFACE = 15 'Button
'Private Const COLOR_BTNSHADOW = 16 '3D shading of button
Private Const COLOR_GRAYTEXT = 17 'Grey text, of zero if dithering is used.
'Private Const COLOR_BTNTEXT = 18 'Button text
'Private Const COLOR_INACTIVECAPTIONTEXT = 19 'Text of inactive window
'Private Const COLOR_BTNHIGHLIGHT = 20 '3D highlight of button'
'Private Const COLOR_2NDACTIVECAPTION = 27 'Win98 only: 2nd active window color
'Private Const COLOR_2NDINACTIVECAPTION = 28 'Win98 only: 2nd inactive window color

Private Const HS_CROSS As Long = 4

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
Private Const DFCS_INACTIVE = &H100

Private Const DT_CENTER As Long = &H1
Private Const DT_LEFT As Long = &H0
Private Const DT_RIGHT As Long = &H2

'Private Type SIZE
'    X As Long
'    Y As Long
'End Type
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
    NoStyle = 0
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

Public Enum sColorss
    [Highlight Color] = 0
    [Fore Color] = 1
    [Back Color] = 2
    [Highlight Text] = 3
End Enum

Public Enum SelSty
    [Default] = 0
    [Faded] = 1
End Enum

Public Enum coE
    [Option Box] = 0
    [Check Box] = 1
    [Progress Bar] = 2
    [Normal] = 3
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
    pIcon As StdPicture
    lTrans As OLE_COLOR
    TransMap As String
    Mapped As Boolean
    bPrev As Boolean
End Type

Private Type TextObj
    lWidTWIPS As Long
    lWidPIXELS As Long
    lHeiTWIPS As Long
    lHeiPIXELS As Long
End Type

Public BackPic As StdPicture
Private SystemTextColor As Long
Private SystemHighlightTextColor As Long
Private CurStyle As View  'styles
Private CurColor As OLE_COLOR 'BG colors
Private LB As Long 'Color
Private aList() As ListStyle 'List items
Private lTop As Long 'Top item index
Private lSelected As Long 'Which item is selected
'Private bHasSB As Boolean 'Have the little thingy on the scrollbar
Private MaxLen As Long
Private udtT As TextObj
'Basic events
Public Event Click()
Public Event DoubleClick()
Public Event ItemClicked(Index As Long)
Public Event ItemChecked(Index As Long)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event VerticalScroll(lValue As Long)
Public Event HorizontalScroll(lValue As Long)
Public Event ItemAdded()
Public Event Ctrl(Value As Long)
Private bPaint As Boolean
Private bMult As Boolean
Private fFill As FillStyle
Private sSty As SelSty
Private lCtrlD As Long
Private VSBW As Long
Private HSBH As Long
Private bSort As Boolean
Public CtrlValue As Long


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
'=======   }r -RESET COLOR                ==========================
'=======   }i -ITALIC                     ==========================
'=======   }b -BOLD                       ==========================
'=======   }u -UNDERLINE                  ==========================
'=======   }n -NORMAL                     ==========================
'===================================================================

Private Function DeterVisible(NeedVSB As Boolean) As Long
Dim lH As Long
Dim lY As Long
'Determines how many items are visible at 1 time
lH = udtT.lHeiTWIPS
If DoINeedaHSB(NeedVSB, VSBW) Then lY = picMain.Height - HS.Height - ScaleY(6, 3, 1) Else lY = picMain.Height - ScaleY(6, 3, 1)
lH = lY \ lH
'If DoINeedaHSB(NeedVSB) Then lH = lH - 1
DeterVisible = lH
End Function
'===================

Private Function DoINeedaHSB(NeedVSB As Boolean, Optional VSBWidth As Long) As Boolean
Dim s As Long
Dim l As Long
s = MaxLen
If NeedVSB Then l = VSBWidth + 2 Else l = 0
If s > picMain.Width - ScaleX(l, 3, 1) Then DoINeedaHSB = True
End Function

Private Function DoINeedaVSB(Optional NeedHSB As Boolean = False) As Boolean
'This is to check if we need a scroll bar
Dim lH As Long
Dim lR As Long
   On Error GoTo DoINeedaVSB_Error

lR = picMain.Height
'Get the height of a letter

lH = udtT.lHeiPIXELS * UBound(aList)  'and take that height, and multiply it
                        'by how many items we have
If NeedHSB Then lH = lH + HSBH + 2
lH = ScaleY(lH, 3, 1) 'make it in TWIPS
If lH >= lR Then DoINeedaVSB = True 'if it is more then the height of the
                                    'user control, we need one.

   On Error GoTo 0
   Exit Function

DoINeedaVSB_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DoINeedaVSB of User Control UltraBox"
End Function

Private Function GetHSBHeight() As Long
GetHSBHeight = GetSysMetrics(SM_CYHSCROLL)
End Function

Private Function GetItemClick(X As Single, Y As Single) As Long
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
If bBNeedSB Then
    bNeedHSB = DoINeedaHSB(bBNeedSB, VSBW)
    lTop = VS.Value + 1
Else
    bNeedHSB = DoINeedaHSB(bBNeedSB)
End If
If bNeedHSB Then lOFF = -HS.Value
lH = udtT.lHeiTWIPS  'Get the height in TWIPS
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
                If .bUseCheckBox Then
                    'Checkbox dimensions
                    'Check if they clicked in the Checkbox
                    If (X > 1 And X < b.Right) And (Y > b.Top And Y < b.Bottom) Then
                        Select Case .iCheck
                            Case 0
                                .iCheck = 1 'Check it
                                RaiseEvent ItemChecked(i)
                            Case 1
                                .iCheck = 0 'UnCheck it
                        End Select
                        GetItemClick = i
                        SetRectEmpty b
                        Exit For
                    End If
                    'If they didn't click in the Checkbox, adjust the RECT
                    'To be able to check for an item/ other things
                ElseIf .bUseOptionBox Then
                    'If they clicked on it
                    If (X > 1 And X < b.Right) And (Y > b.Top And Y < b.Bottom) Then
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
                                                'Debug.Print j & " OPT = 0"
                                            End If
                                        End If
                                    End With
                                    DoEvents
                                Next
                                .iOpt = 1 'Select this one
                                RaiseEvent ItemChecked(i)
                                GetItemClick = i
                        End Select
                        SetRectEmpty b
                        Exit For
                    End If
                End If
                b.Left = b.Right + 1
            End If
        End If
        'Adjust the RECT accordinly, depending on if they
        'have a scroll bar showing.
        If bBNeedSB Then
            b.Right = picMain.Width - ScaleY(VSBW, 3, 1) - ScaleX(2, 3, 1)
        Else
            b.Right = picMain.Width - ScaleX(3, 3, 1)
        End If
        'Set up the RECT for the 1 item, dpeneding on
        'how far down the list we are.
        b.Top = 1 + ((lCounter) * lH)
        b.Bottom = b.Top + lH
        If (X > 1 And X < b.Right) And (Y > b.Top And Y < b.Bottom) And .bEnabled Then
            GetItemClick = i 'If they clicked the item, set the
            'value to the function, and get out of the loop
            SetRectEmpty b
            Exit For
        End If
        lCounter = lCounter + 1 'Increase the counter
        'We don't want to check for invisible items, so if
        'we have check more then are visible, get out of
        'the loop
        If lCounter >= DeterVisible(bBNeedSB) Then Exit For
    End With
Next

   On Error GoTo 0
   Exit Function

GetItemClick_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetItemClick of User Control UltraBox"
End Function

Private Function GetSysMetrics(GetWhat As SysMet) As Long
GetSysMetrics = GetSystemMetrics(SetSysMetIndex(GetWhat))
End Function

Private Function GetTextColorFromNum(lNum As String) As Long
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
'=======   }r -RESET COLOR                ==========================
'=======   }i -ITALIC                     ==========================
'=======   }b -BOLD                       ==========================
'=======   }u -UNDERLINE                  ==========================
'=======   }n -NORMAL                     ==========================
'===================================================================
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
    Case "i"
        GetTextColorFromNum = -2
    Case "b"
        GetTextColorFromNum = -3
    Case "u"
        GetTextColorFromNum = -4
    Case "n"
        GetTextColorFromNum = -5
    Case "r"
        GetTextColorFromNum = -6
    Case Else
        GetTextColorFromNum = -1 'Not a custom color
End Select
End Function

Private Function GetVSBWidth() As Long
GetVSBWidth = GetSysMetrics(SM_CXVSCROLL)
End Function

Private Function ReplaceColors(ByVal s As String) As String
'Replaces all the custom colros so the
'user can just get the text value of
'the item
If InStr(1, s, "}") = 0 Then ReplaceColors = s: Exit Function
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
s = Replace$(s, "}i", "")
s = Replace$(s, "}b", "")
s = Replace$(s, "}u", "")
s = Replace$(s, "}n", "")
s = Replace$(s, "}r", "")
ReplaceColors = s
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

Private Sub CALCMAXLEN(s As String, Optional bPic As Boolean = False)
Dim tArr()  As String
Dim j       As Long
Dim v       As Long
Dim i       As Long
If InStr(1, s, "}") > 0 Then
    tArr = Split(s, "}")
    For j = LBound(tArr) To UBound(tArr)
        If tArr(j) <> "" Then
            v = GetTextColorFromNum(Left$(tArr(j), 1))
            If v = -1 And Left$(tArr(j), 1) = "}" Then
                tArr(j) = "}" & tArr(j)
            ElseIf j <> 0 And v > -1 Then
                tArr(j) = Mid$(tArr(j), 2)
            ElseIf j <> 0 And v < -1 Then
                Select Case v
                    Case -2 'Italic
                        UserControl.FontItalic = True
                    Case -3 'Bold
                        UserControl.FontBold = True
                    Case -4 'Underline
                        UserControl.FontUnderline = True
                    Case -5 'Normal
                        UserControl.FontItalic = False
                        UserControl.FontBold = False
                        UserControl.FontUnderline = False
                End Select
                tArr(j) = Mid$(tArr(j), 2)
            End If
            i = i + TextWidth(tArr(j))
        End If
    Next
    UserControl.FontItalic = False
    UserControl.FontBold = False
    UserControl.FontUnderline = False
Else
    i = picMain.TextWidth(s)
End If
If bPic Then
    If i + ScaleX((udtT.lHeiPIXELS * 2) - 3, 3, 1) > MaxLen Then MaxLen = i + ScaleX((udtT.lHeiPIXELS * 2) - 3, 3, 1)
Else
    If i > MaxLen Then MaxLen = i
End If
End Sub

Private Sub DRAWCUSTOMSTRING(sCap As String, lTHDC As Long, ByRef rRECT As RECT, lFOR As OLE_COLOR, lALI As Long)
Dim tArr()  As String
Dim j       As Long
Dim v       As Long
'If they are, split it into an array
tArr = Split(sCap, "}")
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
        If v = -1 And Left$(tArr(j), 1) = "}" Then '-1 means no custom color, and
                    'the } found was suppose to be at
                    'the front, so put it back
            tArr(j) = "}" & tArr(j)
        ElseIf j <> 0 And v > -1 Then 'If it is not the FIRST item
                           'In the array, set the custom
                           'color, and trim off the color
                           'number
            SetTextColor lTHDC, v
            tArr(j) = Mid$(tArr(j), 2)
        ElseIf j <> 0 And v < -1 Then
                Select Case v
                    Case -2 'Italic
                        picMain.FontItalic = True
                    Case -3 'Bold
                        picMain.FontBold = True
                    Case -4 'Underline
                        picMain.FontUnderline = True
                    Case -5 'Normal
                        picMain.FontItalic = False
                        picMain.FontBold = False
                        picMain.FontUnderline = False
                    Case -6
                        SetTextColor lTHDC, lFOR
                End Select
            tArr(j) = Mid$(tArr(j), 2)
        End If
        'Draw the text of that color
        DrawText lTHDC, tArr(j), Len(tArr(j)), rRECT, lALI
        'Offset the rect by tthe amount of the text printed out.
        OffsetRect rRECT, ScaleX(picMain.TextWidth(tArr(j)), 1, 3), 0
    End If
Next
Erase tArr
picMain.FontItalic = False
picMain.FontBold = False
picMain.FontUnderline = False
End Sub

Private Sub DrawHSB()
Dim HH          As Long
Dim vW          As Long
Dim bBNeedSB    As Boolean
Dim bNHSB       As Boolean
bBNeedSB = DoINeedaVSB
If bBNeedSB Then
    bNHSB = DoINeedaHSB(bBNeedSB, VSBW)
    lTop = VS.Value + 1
    DrawSB
Else
    bNHSB = DoINeedaHSB(bBNeedSB)
    bBNeedSB = DoINeedaVSB(bNHSB)
    If bBNeedSB Then
        bNHSB = DoINeedaHSB(bBNeedSB, VSBW)
    End If
End If
HH = GetHSBHeight
vW = GetVSBWidth
HS.Visible = True
HS.Width = picMain.Width - ScaleX(6, 3, 1)
If bBNeedSB Then HS.Width = HS.Width - ScaleX(vW + 1, 3, 1)
HS.Height = ScaleY(HH, 3, 1)
HS.Top = picMain.Height - HS.Height - ScaleY(3, 3, 1)
HS.Left = ScaleX(3, 3, 1)
HS.max = ScaleX(MaxLen, 1, 3) - ScaleX(picMain.Width, 1, 3) + (vW + 8)
End Sub

Private Sub DrawInit()
'Begining of the drawing to the user control
Dim b   As Long
Dim LBM As Long
If bPaint = False Then Exit Sub
bPaint = False
GetClientRect picMain.hwnd, R 'Get the user controls RECT
picMain.Cls 'Clear the screen
If Color And (BackPic Is Nothing Or fFill = Lined) Then 'If they have a custom color
    'create a brush
    LB = CreateSolidBrush(CurColor)
    'And fill the user control with it
    FillRect picMain.hdc, R, LB
    DeleteObject LB 'Clean up resources
ElseIf Not BackPic Is Nothing Then
    LBM = CreateCompatibleDC(GetDC(0))
    SelectObject LBM, BackPic
    StretchBlt picMain.hdc, 0, 0, ScaleX(picMain.ScaleWidth, 1, 3) * 1.8, ScaleY(picMain.ScaleHeight, 1, 3) * 1.8, LBM, 0, 0, ScaleX(BackPic.Width, 1, 3), ScaleY(BackPic.Height, 1, 3), vbSrcCopy
    DeleteDC LBM
    DeleteObject LBM
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
SetRectEmpty R
bPaint = True
picMain.Refresh 'Refresht he user control
End Sub

Private Sub DrawList() 'Optional DrawAll As Boolean = False, Optional lHDC As Long = -1, Optional bPrinter As Boolean = False)
Dim i           As Long     'Counter
Dim b           As RECT     'Item's RECT structure
Dim lH          As Long     'Text height
Dim lC          As Long     'Color
Dim bBHasSB     As Boolean  'Scroll bar flag
Dim lCounter    As Long     'Counter
Dim bBNeedSB    As Boolean  'Need the scroll bar flag
Dim t           As Long     'For Instr
Dim bNHSB       As Boolean
Dim pb          As RECT
Dim lOFF        As Long
Dim lCol        As Long
Dim lhDC        As Long
Dim lT          As Long
On Error GoTo DrawList_Error
'Check the limitations
If UBound(aList) = 0 Then HS.Visible = False: VS.Visible = False: Exit Sub
If lTop = 0 Then lTop = 1
If lTop > UBound(aList) Then lTop = UBound(aList)
'Now, see if they need a scroll bar
bBNeedSB = DoINeedaVSB
If bBNeedSB Then
    bNHSB = DoINeedaHSB(bBNeedSB, VSBW)
    lTop = VS.Value + 1
    DrawSB
Else
    bNHSB = DoINeedaHSB(bBNeedSB)
    bBNeedSB = DoINeedaVSB(bNHSB)
    If Not bBNeedSB Then
        VS.Visible = False
    Else
        bNHSB = DoINeedaHSB(bBNeedSB, VSBW)
        lTop = VS.Value + 1
        DrawSB
    End If
End If
lhDC = picMain.hdc 'If lHDC = -1 Then lHDC = picMain.hdc
If bNHSB Then lOFF = -HS.Value: DrawHSB Else HS.Visible = False
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
            With b
                .Left = 3
                .Top = 3 + ((lCounter) * lH)
                .Right = .Left + 12
                .Bottom = .Top + lH + 3
            End With
            If bNHSB Then OffsetRect b, lOFF, 0
            If .bUseCheckBox Then
                'Get the checkbox dimensions
                If .bEnabled = True Then
                    Select Case .iCheck
                        Case 0 'Unchecked
                            DrawFrameControl lhDC, b, DFC_BUTTON, DFCS_BUTTONCHECK
                        Case 1 'Checked
                            DrawFrameControl lhDC, b, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_CHECKED
                    End Select
                Else
                    Select Case .iCheck
                        Case 0 'Unchecked
                            DrawFrameControl lhDC, b, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_INACTIVE
                        Case 1 'Checked
                            DrawFrameControl lhDC, b, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_CHECKED Or DFCS_INACTIVE
                    End Select
                End If
            ElseIf .bUseOptionBox Then
                'Option box
                If .bEnabled = True Then
                    Select Case .iOpt
                        Case 0 'Unselected
                            DrawFrameControl lhDC, b, DFC_BUTTON, DFCS_BUTTONRADIO
                        Case 1 'Selected
                            DrawFrameControl lhDC, b, DFC_BUTTON, DFCS_BUTTONRADIO Or DFCS_CHECKED
                    End Select
                Else
                    Select Case .iOpt
                        Case 0 'Unselected
                            DrawFrameControl lhDC, b, DFC_BUTTON, DFCS_BUTTONRADIO Or DFCS_INACTIVE
                        Case 1 'Selected
                            DrawFrameControl lhDC, b, DFC_BUTTON, DFCS_BUTTONRADIO Or DFCS_CHECKED Or DFCS_INACTIVE
                    End Select
                End If
            End If
            b.Left = b.Right + 1 'Prepare for text with 1 pixel spacing
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
        If (.bSelected = False) Or (.bSelected = True And i <> lSelected) Then 'If the items isn't selected
            'Create a solid brush of the backcolor they chose for the item
            If fFill = NoStyle Then
                If Not BackPic Is Nothing And .bSelected = False Then
                    lC = -111
                Else
                    If (.bSelected = True And i <> lSelected) Then
                        If sSty = Faded Then
                            lC = -112
                        Else
                            lC = CreateSolidBrush(.lHighlightColor)
                        End If
                    Else
                        lC = CreateSolidBrush(.lBackColor)
                    End If
                End If
            Else
                If lCol = 0 Then
                    lC = CreateSolidBrush(vbWhite)
                    lCol = 1
                Else
                    lC = CreateSolidBrush(&HEFEFEF)
                    lCol = 0
                End If
            End If
            If Not .pIcon Is Nothing Then lT = 1 Else lT = 0
            If lC <> -111 And lC <> -112 Then
                'Fill the rect with that color
                FillRect lhDC, b, lC
                'Clean up resources
                DeleteObject lC
            ElseIf lC = -112 Then
                If lT = 1 Then OffsetRect b, (b.Bottom - b.Top) + 3, 0
                DRAWFADESELECT lhDC, b, .lHighlightColor
                If lT = 1 Then OffsetRect b, -((b.Bottom - b.Top) + 3), 0
            End If
            If lT = 1 Then
                DRAWPICON lhDC, i, lH, b, .pIcon, .lTrans, .lBackColor, lCol
                'lOFF = (lH * 2) - 3
                lT = 1
                OffsetRect b, (b.Bottom - b.Top) + 3, 0
            End If
            '///////////////Progressbar//////////////////
            If .bUseProgress Then
                DRAWPROGRESS lhDC, b, bNHSB, lOFF, .lProgressValue, .lProgressMax, .lProgressBarColor, lT
                b.Left = b.Left + 2
            End If
            'If it is enabled, set the
            'forecolor accordenly
            If .bEnabled Then
                If (.bSelected = True And i <> lSelected) Then
                    SetTextColor lhDC, .lHightlightText
                Else
                    SetTextColor lhDC, .lForeColor
                End If
            Else
                'Get the default GRAYed out color of the system
                SetTextColor lhDC, GetSysColor(COLOR_GRAYTEXT)
            End If
            If (i = lSelected) Then
                If .lDrawFR = 0 Then
                    SetTextColor lhDC, vbBlack
                    DrawFocusRect lhDC, b
                End If
            End If
            'Check if they are using custom colors...
            t = InStr(1, .sCaption, "}")
            If t = 0 Then 'If not, just draw out the text
                DrawText lhDC, aList(i).sCaption, Len(aList(i).sCaption), b, .lAlignment
            Else
                DRAWCUSTOMSTRING .sCaption, lhDC, b, .lForeColor, .lAlignment
            End If
        Else
            'If the item is selected
            lT = 0
            If Not .pIcon Is Nothing Then
                DRAWPICON lhDC, i, lH, b, .pIcon, .lTrans, .lBackColor, lCol
                'lOFF = (lH * 2) - 3
                lT = 1
                OffsetRect b, (b.Bottom - b.Top) + 3, 0
            End If
            'Get the highlight color and make a brush
            
            '///////////////Progressbar//////////////////
            If .bUseProgress Then
                DRAWPROGRESS lhDC, b, bNHSB, lOFF, .lProgressValue, .lProgressMax, .lProgressBarColor, lT
                CopyRect pb, b
                If lT = 1 Then pb.Right = pb.Right - 3 - (udtT.lHeiPIXELS * 2)
                InflateRect pb, -2, -2
                InvertRect lhDC, pb
                With b
                    .Left = .Left + 2
                    .Top = .Top + 2
                    .Bottom = .Bottom - 3
                    .Right = .Right - 3
                    If lT = 1 Then .Right = .Right - (udtT.lHeiPIXELS * 2) - 3
                End With
            Else
                'If lT = 1 Then b.Right = b.Right - (udtT.lHeiPIXELS * 2)
                If sSty = Faded And fFill <> Lined Then
                    DRAWFADESELECT lhDC, b, .lHighlightColor
                Else
                    lC = CreateSolidBrush(.lHighlightColor)
                    If lCol = 1 Then lCol = 0 Else lCol = 1
                    'Fill the items RECT with the color
                    FillRect lhDC, b, lC
                    'Clean up resources
                    DeleteObject lC
                End If
            End If
            'Draw the focus rect around the item
            If .lDrawFR = 0 Then
                SetTextColor lhDC, vbBlack
                DrawFocusRect lhDC, b
            End If
            'And change the text color to the HIGHLIGHT text color
            If .bUseProgress Then InflateRect b, 2, 2
            SetTextColor lhDC, .lHightlightText
            t = InStr(1, .sCaption, "}") 'See if there is custom colors
            If t = 0 Then 'If not, just put out the text
                DrawText lhDC, aList(i).sCaption, Len(aList(i).sCaption), b, .lAlignment
            Else
                'Else, use custom colors
                DRAWCUSTOMSTRING .sCaption, lhDC, b, .lForeColor, .lAlignment
            End If
        End If
        lCounter = lCounter + 1 'Increase the counter
        'We don't want to go over the amount of
        'visible items, so check that
        If (lCounter >= DeterVisible(bBNeedSB)) Then 'And Not DrawAll) Or i = UBound(aList) Then
            'if we are over, draw the scroll bar if needed
            'and exit the loop.
            If bNHSB Then DrawHSB
            If Not bBHasSB And bBNeedSB Then DrawSB: bBHasSB = True
            Exit For
        End If
    End With
Next
SetRectEmpty b
SetRectEmpty pb
   On Error GoTo 0
   Exit Sub

DrawList_Error:
    
End Sub

Private Sub DRAWFADESELECT(lTHDC As Long, rRECT As RECT, lCO As OLE_COLOR)
Dim i As Long
i = CreateSolidBrush(lCO)
FillRect UserControl.hdc, rRECT, i
DeleteObject i
BitBlt lTHDC, rRECT.Left, rRECT.Top, rRECT.Right - rRECT.Left, rRECT.Bottom - rRECT.Top, UserControl.hdc, rRECT.Left, rRECT.Top, vbSrcPaint
End Sub

Private Sub DRAWPICON(lTHDC As Long, iID As Long, lHEI As Long, rRECT As RECT, pICO As StdPicture, lTRA As Long, lBAC As Long, lCO)
Dim LBM     As Long
Dim tempR   As RECT
Dim j       As Long
Dim t       As Long
Dim bp      As Boolean
Dim s       As String
Dim c       As Long
Dim b       As Boolean
s = ""
LBM = CreateCompatibleDC(GetDC(0))
SelectObject LBM, pICO
BitBlt UserControl.hdc, 0, 0, ScaleX(pICO.Width, 1, 3), ScaleY(pICO.Height, 1, 3), LBM, 0, 0, vbSrcCopy
DeleteDC LBM
LBM = UserControl.hdc
If Not BackPic Is Nothing Then bp = True
If lTRA <> -1 Then
    If aList(iID).Mapped = False Then
        tempR.Right = ScaleX(pICO.Width, vbTwips, vbPixels)
        tempR.Bottom = ScaleY(pICO.Height, vbTwips, vbPixels)
        For j = tempR.Top To tempR.Bottom
            For t = tempR.Left To tempR.Right '+ (lH * 2) - 3
                If GetPixel(LBM, t, j) = lTRA Then
                    If fFill = NoStyle Then
                        If bp Then
                            c = GetPixel(lTHDC, rRECT.Left + t, rRECT.Top + j) And lBAC
                            SetPixel LBM, t, j, c
                            s = s & t & "," & j & ","
                        Else
                            SetPixel LBM, t, j, lBAC
                            s = s & t & "," & j & ","
                        End If
                    Else
                        Select Case lCO
                            Case 1
                                SetPixel LBM, t, j, vbWhite
                                s = s & t & "," & j & ","
                            Case 0
                                SetPixel LBM, t, j, &HEFEFEF
                                s = s & t & "," & j & ","
                        End Select
                    End If
                End If
            Next
        Next
        aList(iID).TransMap = s
        aList(iID).Mapped = True
    Else
        s = aList(iID).TransMap
        If s <> "" Then
            Do Until b
                c = InStr(1, s, ",")
                t = Left$(s, c)
                s = Mid$(s, c + 1)
                c = InStr(1, s, ",")
                j = Left$(s, c)
                s = Mid$(s, c + 1)
                If fFill = NoStyle Then
                    If bp Then
                        c = GetPixel(lTHDC, rRECT.Left + t, rRECT.Top + j) And lBAC
                        SetPixel LBM, t, j, c
                    Else
                        SetPixel LBM, t, j, lBAC
                    End If
                Else
                    Select Case lCO
                        Case 1
                            SetPixel LBM, t, j, vbWhite
                        Case 0
                            SetPixel LBM, t, j, &HEFEFEF
                    End Select
                End If
                If InStr(1, s, ",") = 0 Then b = True
            Loop
        End If
    End If
End If
StretchBlt UserControl.hdc, 0, 0, (lHEI * 2) - 3, (lHEI * 2) - 3, LBM, 0, 0, ScaleX(pICO.Width, 1, 3), ScaleY(pICO.Height, 1, 3), vbSrcCopy
BitBlt lTHDC, rRECT.Left, rRECT.Top, (lHEI), (lHEI), UserControl.hdc, 0, 0, vbSrcCopy
End Sub

Private Sub DRAWPROGRESS(lTHDC As Long, ByRef rRECT As RECT, bHOR As Boolean, lOS As Long, lPV As Long, lPM As Long, lCol As OLE_COLOR, lICON As Long)
Dim pb      As RECT
Dim lC      As Long
Dim dVal    As Double
lC = CreateSolidBrush(lCol)
CopyRect pb, rRECT
If bHOR Or lOS <> 0 Then OffsetRect pb, lOS, 0
If lICON = 1 Then pb.Right = pb.Right - (udtT.lHeiPIXELS * 2) - 3
DrawEdge lTHDC, pb, EDGE_ETCHED, BF_RECT
'If lICON = 1 Then pb.Left = pb.Left + (udtT.lHeiPIXELS * 2) - 3
InflateRect pb, -2, -2
dVal = lPV / lPM
dVal = Round(dVal, 2)
pb.Right = pb.Left + ((pb.Right - pb.Left) * dVal)

 'pb.Right = pb.Right + pb.Left
FillRect lTHDC, pb, lC
DeleteObject lC
SetRectEmpty pb
End Sub

Private Sub DrawSB()
'Drawing the scroll bar
VS.Visible = True
VS.Top = ScaleY(3, 3, 1)
VS.Height = UserControl.Height - ScaleY(6, 3, 1)
VS.Width = ScaleX(GetVSBWidth, 3, 1)
VS.Left = picMain.Width - VS.Width - ScaleX(2, 3, 1)
VS.max = UBound(aList) - DeterVisible(True)
End Sub

Private Sub HS_Change()
If bPaint Then DrawInit
RaiseEvent HorizontalScroll(HS.Value)
End Sub

Private Sub HS_Scroll()
If bPaint Then DrawInit
RaiseEvent HorizontalScroll(HS.Value)
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
lCtrlD = 0
If bPaint Then DrawInit

   On Error GoTo 0
   Exit Sub

picMain_GotFocus_Error:
    lSelected = -1
End Sub

Private Sub picMain_KeyDown(KeyCode As Integer, Shift As Integer)
'For moving the selection up and down.
Dim d As Long
Dim i As Long
Dim j As Long
   On Error GoTo picMain_KeyDown_Error
If bPaint = False Then Exit Sub
If (KeyCode = 17 Or KeyCode = 16) And bMult = True And lCtrlD <> 1 Then
    lCtrlD = 1
    CtrlValue = 1
    RaiseEvent Ctrl(1)
ElseIf (KeyCode = 17 Or KeyCode = 16) And bMult = True And lCtrlD = 1 Then
    Exit Sub
End If
If lSelected < 1 Then Exit Sub
d = DeterVisible(DoINeedaVSB)
Select Case KeyCode
    Case vbKeyDown
        'Find the next item, going down the list
        If lSelected + 1 <= UBound(aList) Then
            i = lSelected
            Do
                i = i + 1
                DoEvents
            Loop Until aList(i).bEnabled = True Or i >= UBound(aList)
            If i <= UBound(aList) Then
                If (lCtrlD = 0) Then aList(lSelected).bSelected = False
                If lCtrlD = 0 And bMult Then
                    For j = LBound(aList) To UBound(aList)
                        If j <> i And j <> lSelected Then
                            If aList(j).bSelected Then
                                aList(j).bSelected = False
                            End If
                        End If
                    Next
                End If
                If aList(i).bSelected = True And bMult And lCtrlD = 1 Then
                    aList(i).bSelected = False
                Else
                    aList(i).bSelected = True
                End If
                If i > lTop + d - 1 Then VS.Value = VS.Value + i - lSelected
                lSelected = i
            End If
            If bPaint Then DrawInit
        End If
    Case vbKeyUp
        'Find thenext item going up the list
        If lSelected - 1 > 0 Then
            i = lSelected
            Do
                i = i - 1
                DoEvents
            Loop Until aList(i).bEnabled = True Or i < 1
            If i > 0 Then
                If lCtrlD = 0 Then aList(lSelected).bSelected = False
                If lCtrlD = 0 And bMult Then
                    For j = LBound(aList) To UBound(aList)
                        If j <> i And j <> lSelected Then
                            If aList(j).bSelected Then
                                aList(j).bSelected = False
                            End If
                        End If
                    Next
                End If
                If aList(i).bSelected = True And bMult And lCtrlD = 1 Then
                    aList(i).bSelected = False
                Else
                    aList(i).bSelected = True
                End If
                If i < lTop Then VS.Value = VS.Value - (lSelected - i)
                lSelected = i
            End If
            If bPaint Then DrawInit
        End If
End Select
RaiseEvent Click
On Error GoTo 0
Exit Sub
picMain_KeyDown_Error:
End Sub

Private Sub picMain_KeyUp(KeyCode As Integer, Shift As Integer)
If (KeyCode = 17 Or KeyCode = 16) Then lCtrlD = 0: CtrlValue = 0: RaiseEvent Ctrl(0)
End Sub

Private Sub picMain_LostFocus()
If lSelected = -1 Then Exit Sub
aList(lSelected).lDrawFR = 1
lCtrlD = 0
If bPaint Then DrawInit
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'To find out where/what they clicked
On Error GoTo eh1
Dim l As Long
Dim i As Long
l = GetItemClick(X, Y)
If l <> 0 Then
    For i = 1 To UBound(aList)
        With aList(i)
            If i = l Then
                If lCtrlD = 1 And .bSelected = True Then
                    .bSelected = False
                Else
                    .bSelected = True 'Find the item they clicked and selected it
                End If
                lSelected = i
            ElseIf lCtrlD = 0 Then
                .bSelected = False   'and make all the others not selected
            End If
        End With
        DoEvents
    Next
End If
eh1:
DrawInit
RaiseEvent ItemClicked(lSelected)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'DrawInit
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub SortIt()
Dim tmpList As ListStyle 'temp list item
Dim i       As Long 'max check
Dim j       As Long 'counter
Dim l       As Long 'offset
Dim u       As Long 'ubound()
Dim LB      As Long 'lbound()
Dim b       As Long 'item switching
   On Error GoTo SortIt_Error
    u = UBound(aList)
    If u < 1 Then Exit Sub
    LB = LBound(aList)
    LB = LB + 1
    l = (u - LB) \ 2
    Do Until l < 1
        i = u - l
        Do
            b = 0
            For j = LB To i
                If UCase(ReplaceColors(aList(j).sCaption)) > UCase(ReplaceColors(aList(j + l).sCaption)) Then
                    tmpList = aList(j)
                    aList(j) = aList(j + l)
                    aList(j + l) = tmpList
                    b = j
                End If
            Next
            i = b - l
        Loop While b > 0
        l = l \ 2
    Loop
    If bPaint Then DrawInit
   On Error GoTo 0
   Exit Sub
SortIt_Error:
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
SetStretchBltMode picMain.hdc, vbPaletteModeNone
bPaint = True
VSBW = GetVSBWidth
HSBH = GetHSBHeight
ReDim aList(0)
DrawInit
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
'Propertys to retrieve
With PropBag
    CurStyle = .ReadProperty("Style")
    CurColor = .ReadProperty("Color", vbWhite)
    fFill = .ReadProperty("Fill")
    sSty = .ReadProperty("SELECTSTYLE")
    bMult = .ReadProperty("Mult")
    bSort = .ReadProperty("Sort")
    Set BackPic = .ReadProperty("BACKPICTURE", Nothing)
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
Dim i As Long
On Error GoTo UserControl_Terminate_Error
For i = 1 To UBound(aList)
    Set aList(i).pIcon = Nothing
Next
Erase aList 'Clear up memory
Set BackPic = Nothing
SetRectEmpty R
On Error GoTo 0
Exit Sub
UserControl_Terminate_Error:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'properties to save
With PropBag
    .WriteProperty "Style", CurStyle
    .WriteProperty "Color", CurColor, vbWhite
    .WriteProperty "Fill", fFill
    .WriteProperty "Font", picMain.Font
    .WriteProperty "Mult", bMult
    .WriteProperty "Sort", bSort
    .WriteProperty "BACKPICTURE", BackPic
    .WriteProperty "SELECTSTYLE", sSty
End With
End Sub

Private Sub VS_Change()
If bPaint Then DrawInit
RaiseEvent VerticalScroll(VS.Value)
End Sub

Private Sub VS_Scroll()
If bPaint Then DrawInit
RaiseEvent VerticalScroll(VS.Value)
End Sub

Public Function Find(ByVal s As String, Optional bRemoveTags As Boolean = False) As Long
Dim i As Long
Dim a As String
Dim b As String
For i = LBound(aList) To UBound(aList)
    With aList(i)
        If .bEnabled = True Then
            If bRemoveTags Then
                a = LCase$(ReplaceColors(Left$(.sCaption, Len(s))))
                b = LCase$(ReplaceColors(s))
            Else
                a = LCase$(Left$(.sCaption, Len(s)))
                b = LCase$(s)
            End If
            If a = b Then
                Find = i
                Exit Function
            End If
        End If
    End With
Next
Find = 1
End Function

Public Function FindInStr(ByVal s As String) As Long
Dim i As Long
For i = LBound(aList) To UBound(aList)
    With aList(i)
        If .bEnabled = True Then
            If InStr(1, LCase$(.sCaption), LCase$(s)) Then
                FindInStr = i
                Exit Function
            End If
        End If
    End With
Next
FindInStr = 1
End Function

Public Function GetItemOptGrp(Index As Long) As Long
With aList(Index)
    GetItemOptGrp = .lOptionGroup
End With
End Function

Public Function GetProgressMax(Index As Long) As Long
On Error GoTo GetProgressValue_Error

With aList(Index)
    If .bUseProgress Then
        GetProgressMax = .lProgressMax
    End If
End With

On Error GoTo 0
Exit Function
GetProgressValue_Error:
End Function

Public Function GetProgressValue(Index As Long) As Long
On Error GoTo GetProgressValue_Error

With aList(Index)
    If .bUseProgress Then
        GetProgressValue = .lProgressValue
    End If
End With

On Error GoTo 0
Exit Function
GetProgressValue_Error:
End Function

Public Function IsSelected(Index As Long, Optional CheckAndOptionOnly As Boolean = False, Optional NoCheckAndOption As Boolean = False) As Boolean
On Error GoTo IsSelected_Error
'Will check if a certain item is selected
If Not CheckAndOptionOnly And Not NoCheckAndOption Then
    'This will flag selected if it is CHECKED, SELECTED, or the option box is CLICKED
    If aList(Index).bSelected = True Or aList(Index).iCheck = 1 Or aList(Index).iOpt = 1 Then
        IsSelected = True
    End If
ElseIf Not CheckAndOptionOnly And NoCheckAndOption Then
    If aList(Index).bSelected = True Then
        IsSelected = True
    End If
ElseIf CheckAndOptionOnly Then
    'This will only check option boxes and check boxes
    If aList(Index).iCheck = 1 Or aList(Index).iOpt = 1 Then
        IsSelected = True
    End If
End If

On Error GoTo 0
Exit Function
IsSelected_Error:
End Function

Public Function ItemEnabled(Index As Long) As Boolean
With aList(Index)
    ItemEnabled = .bEnabled
End With
End Function

Public Function ItemText(Optional KeepTags As Boolean = False) As String
'Gets the currently selected item's text
If lSelected = -1 Then Exit Function
If KeepTags Then
    ItemText = aList(lSelected).sCaption
Else
    ItemText = ReplaceColors(aList(lSelected).sCaption)
End If
End Function

Public Function ItemTypeX(Index As Long) As coE
With aList(Index)
    If .bUseCheckBox = True Then ItemTypeX = [Check Box]
    If .bUseOptionBox = True Then ItemTypeX = [Option Box]
    If .bUseProgress = True Then ItemTypeX = [Progress Bar]
    If .bUseCheckBox = False And .bUseOptionBox = False And .bUseProgress = False Then ItemTypeX = Normal
End With
End Function

Public Function List(Index As Long, Optional KeepTags As Boolean = False) As String
On Error GoTo ItemTextFromIndex_Error
'Get a specific item's text
If KeepTags Then
    List = aList(Index).sCaption
Else
    List = ReplaceColors(aList(Index).sCaption)
End If
On Error GoTo 0
Exit Function

ItemTextFromIndex_Error:
    
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

Public Function ListIndex() As Long
ListIndex = lSelected
End Function

'Backcolor
Public Property Get Color() As OLE_COLOR
Color = CurColor
End Property

Public Property Get Enabled() As Boolean
Enabled = picMain.Enabled
End Property

Public Property Get FillView() As FillStyle
FillView = fFill
End Property

'Font of the items
Public Property Get Font() As StdFont
Set Font = picMain.Font
End Property

Public Property Get MultiSelect() As Boolean
MultiSelect = bMult
End Property

Public Property Get Paint() As Boolean
Paint = bPaint
End Property

Public Property Get Sorted() As Boolean
Sorted = bSort
End Property

'Current look
Public Property Get Style() As View
Style = CurStyle
End Property

Public Property Let Color(ByVal NewColor As OLE_COLOR)
Dim i As Long
For i = LBound(aList) To UBound(aList)
    If aList(i).lBackColor = CurColor Then
        aList(i).lBackColor = NewColor
    End If
    DoEvents
Next
CurColor = NewColor
If bPaint Then DrawInit
End Property

Public Property Let Enabled(ByVal b As Boolean)
Dim i As Long
picMain.Enabled = b
If picMain.Enabled = False Then
    VS.Enabled = False
    HS.Enabled = False
    For i = LBound(aList) To UBound(aList)
        With aList(i)
            .bPrev = .bEnabled
            .bEnabled = False
            .bSelected = False
        End With
    Next
    If bPaint Then DrawInit
Else
    VS.Enabled = True
    HS.Enabled = True
    For i = LBound(aList) To UBound(aList)
        With aList(i)
            .bEnabled = .bPrev
        End With
    Next
    If bPaint Then DrawInit
End If
End Property

Public Property Let FillView(ByVal f As FillStyle)
fFill = f
If bPaint Then DrawInit
PropertyChanged "Fill"
End Property

Public Property Let MultiSelect(ByVal b As Boolean)
bMult = b
PropertyChanged "Mult"
End Property

Public Property Let Paint(ByVal b As Boolean)
bPaint = b
If bPaint Then DrawInit
End Property

Public Property Let Sorted(ByVal b As Boolean)
bSort = b
If bSort Then SortIt
PropertyChanged "Sort"
End Property

Public Property Let Style(ByVal NewStyle As View)
CurStyle = NewStyle
DrawInit
PropertyChanged "Style"
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

Public Sub AddItem(sText As String, Optional Alignment As AlignmentConstants = vbLeftJustify, Optional pPicture As StdPicture, Optional TRANSColor As OLE_COLOR = -1, Optional Enabled As Boolean = True, Optional FCOLOR As OLE_COLOR = -1, Optional BCOLOR As OLE_COLOR = -1, Optional HCOLOR As OLE_COLOR = -1, Optional HTEXT As OLE_COLOR = -1, Optional UseCheckBox As Boolean = False, Optional UseOptionBox As Boolean = False, Optional OptionGroup As Long = 0, Optional UseProgressBar As Boolean = False, Optional ProgressBarMax As Long = 100, Optional ProgressBarValue As Long = 1, Optional ProgressBarProgressColor As OLE_COLOR = -1)
'Function to add items
'Add 1 more item
Dim s As String
ReDim Preserve aList(UBound(aList) + 1)
With aList(UBound(aList))
    'Set all the properties
    .sCaption = Replace$(sText, vbCrLf, "") 'No enters
    s = .sCaption
    If pPicture Is Nothing Then
        CALCMAXLEN s
    Else
        CALCMAXLEN s, True
    End If
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
    Set .pIcon = pPicture
    .lTrans = TRANSColor
    If bSort And bPaint Then SortIt
    RaiseEvent ItemAdded
End With

If bPaint Then DrawInit
    
End Sub

Public Sub AddItemCheck(sText As String, Optional Alignment As AlignmentConstants = vbLeftJustify, Optional Enabled As Boolean = True, Optional FCOLOR As OLE_COLOR = -1, Optional BCOLOR As OLE_COLOR = -1, Optional HCOLOR As OLE_COLOR = -1, Optional HTEXT As OLE_COLOR = -1)
AddItem sText, Alignment, , , Enabled, FCOLOR, BCOLOR, HCOLOR, HTEXT, True
End Sub

Public Sub AddItemOption(sText As String, Optional Alignment As AlignmentConstants = vbLeftJustify, Optional Enabled As Boolean = True, Optional FCOLOR As OLE_COLOR = -1, Optional BCOLOR As OLE_COLOR = -1, Optional HCOLOR As OLE_COLOR = -1, Optional HTEXT As OLE_COLOR = -1, Optional OptionGroup As Long = 0)
AddItem sText, Alignment, , , Enabled, FCOLOR, BCOLOR, HCOLOR, HTEXT, False, True, OptionGroup
End Sub

Public Sub AddItemProgressBar(sText As String, Optional Alignment As AlignmentConstants = vbLeftJustify, Optional Enabled As Boolean = True, Optional FCOLOR As OLE_COLOR = -1, Optional BCOLOR As OLE_COLOR = -1, Optional HCOLOR As OLE_COLOR = -1, Optional HTEXT As OLE_COLOR = -1, Optional ProgressBarMax As Long = 100, Optional ProgressBarValue As Long = 1, Optional ProgressBarProgressColor As OLE_COLOR = -1)
AddItem sText, Alignment, , , Enabled, FCOLOR, BCOLOR, HCOLOR, HTEXT, , , , True, ProgressBarMax, ProgressBarValue, ProgressBarProgressColor
End Sub

Public Sub Clear()
ReDim aList(0)
MaxLen = 0
HS.Value = 0
VS.Value = 0
HS.Visible = False
VS.Visible = False
lSelected = -1
If bPaint Then DrawInit
End Sub

Public Sub MakeItemX(Index As Long, wch As coE, Optional OptGrp As Long, Optional lPrgBaMa As Long, Optional lPrgBaCol As OLE_COLOR)
Select Case wch
    Case [Option Box]
        aList(Index).bUseCheckBox = False
        With aList(Index)
            .bUseOptionBox = True
            .lOptionGroup = OptGrp
        End With
    Case [Check Box]
        aList(Index).bUseCheckBox = True
        aList(Index).bUseOptionBox = False
        aList(Index).iOpt = 0
    Case [Progress Bar]
        With aList(Index)
            .bUseProgress = True
            .lProgressBarColor = lPrgBaCol
            .lProgressMax = lPrgBaMa
        End With
    Case [Normal]
        aList(Index).bUseCheckBox = False
        aList(Index).bUseOptionBox = False
        aList(Index).iOpt = 0
End Select
If bPaint Then DrawInit
End Sub

Public Sub Refresh()
Dim i As Long
With udtT
    .lHeiPIXELS = ScaleY(picMain.TextHeight("X" & vbCrLf & "W" & vbCrLf & "y"), 1, 3) \ 3
    .lHeiTWIPS = picMain.TextHeight("X" & vbCrLf & "W" & vbCrLf & "y") \ 3
    .lWidPIXELS = ScaleX(picMain.TextWidth("XW"), 1, 3) \ 2
    .lWidTWIPS = picMain.TextWidth("XW") \ 2
End With
Set UserControl.Font = picMain.Font
For i = 1 To UBound(aList)
    CALCMAXLEN aList(i).sCaption, IIf(aList(i).pIcon Is Nothing, False, True)
    'aList(i).TransMap = ""
    'aList(i).Mapped = False
Next
If bSort Then SortIt
If bPaint Then DrawInit
End Sub

Public Sub RemoveItem(Index As Long)
On Error GoTo RemoveItem_Error
'Removes an item
Dim i As Long
If Index <> 0 Then
    Set aList(Index).pIcon = Nothing
    For i = Index To UBound(aList) - 1
        aList(i) = aList(i + 1)
    Next
    ReDim Preserve aList(UBound(aList) - 1)
    If bPaint Then DrawInit 'refresh the screen
End If

On Error GoTo 0
Exit Sub
RemoveItem_Error:
    
End Sub

Public Sub SetEnabled(Index As Long, ByVal Enabled As Boolean)
On Error GoTo SetEnabled_Error
'This will set an item enabled or disabled
With aList(Index)
    .bEnabled = Enabled
    If Not Enabled Then
        .bSelected = False
        .iCheck = 0
        .iOpt = 0
    End If
End With
If bPaint Then DrawInit 'refresh

On Error GoTo 0
Exit Sub
SetEnabled_Error:
End Sub

Public Sub SetItemColors(Index As Long, eWhich As sColorss, lCol As OLE_COLOR)
'    [Highlight Color] = 0
'    [Fore Color] = 1
'    [Back Color] = 2
'    [Highlight Text] = 3
Select Case eWhich
    Case 0
        aList(Index).lHighlightColor = lCol
    Case 1
        aList(Index).lForeColor = lCol
    Case 2
        aList(Index).lBackColor = lCol
    Case 3
        aList(Index).lHightlightText = lCol
End Select
If bPaint Then DrawInit
End Sub

Public Sub SetItemPicture(Index As Long, sPic As StdPicture, Optional Trans As OLE_COLOR = -1)
With aList(Index)
    Set .pIcon = Nothing
    Set .pIcon = sPic
    .lTrans = Trans
    .Mapped = False
    .TransMap = ""
End With
If bPaint Then DrawInit
End Sub

Public Sub SetItemText(Index As Long, NewText As String)
On Error GoTo SetItemText_Error
'Set an items text
With aList(Index)
    .sCaption = NewText
End With
If bPaint Then DrawInit

On Error GoTo 0
Exit Sub
SetItemText_Error:
End Sub

Public Sub SetProgressValue(Index As Long, NewValue As Long)

   On Error GoTo SetProgressValue_Error

    With aList(Index)
        If .bUseProgress Then
            .lProgressValue = NewValue
            If .lProgressValue > .lProgressMax Then .lProgressValue = .lProgressMax
        End If
    End With
    If bPaint Then DrawInit
   On Error GoTo 0
   Exit Sub

SetProgressValue_Error:

End Sub

Public Sub SetSelected(Index As Long, ByVal Selected As Boolean, Optional MoveSelectionToTop As Boolean = False, Optional justselect As Boolean = False)
Dim i As Long
On Error GoTo SetSelected_Error
'Sets a certain item selected
With aList(Index)
    If Not justselect Then
        If .bUseCheckBox = True Then .iCheck = IIf(Selected = True, 1, 0)
        If .bUseOptionBox = True Then .iOpt = IIf(Selected = True, 1, 0)
    End If
    .bSelected = Selected
    lSelected = Index
End With
'This will uncheck all other option boxs in the same group
For i = 1 To UBound(aList)
    If i <> Index Then
        With aList(i)
            If Selected = True Then
                If Not justselect Then
                    If .lOptionGroup = aList(i).lOptionGroup Then
                        .iOpt = 0
                    End If
                End If
                If lCtrlD = 0 Then .bSelected = False
            End If
        End With
    End If
    DoEvents
Next
If Index > lTop Then
    If Index - 1 > VS.max Then VS.Value = VS.max Else If MoveSelectionToTop Then VS.Value = Index - 1
Else
    VS.Value = Index - 1
End If
If bPaint Then DrawInit 'refresh

On Error GoTo 0
Exit Sub
SetSelected_Error:
End Sub

Public Property Get SelectStyle() As SelSty
SelectStyle = sSty
End Property

Public Property Let SelectStyle(ByVal sty As SelSty)
sSty = sty
If bPaint Then DrawInit
PropertyChanged "SELECTSTYLE"
End Property

