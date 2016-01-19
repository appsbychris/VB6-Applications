Attribute VB_Name = "modAPIs"
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long

Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal idHook As Long) As Long


'creating buffers / loading sprites
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long


'loading sprites
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

'cleanup
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public lMenuClick As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const COLOR_SCROLLBAR = 0 'The Scrollbar colour
Public Const COLOR_BACKGROUND = 1 'Colour of the background with no wallpaper
Public Const COLOR_ACTIVECAPTION = 2 'Caption of Active Window
Public Const COLOR_INACTIVECAPTION = 3 'Caption of Inactive window
Public Const COLOR_MENU = 4 'Menu
Public Const COLOR_WINDOW = 5 'Windows background
Public Const COLOR_WINDOWFRAME = 6 'Window frame
Public Const COLOR_MENUTEXT = 7 'Window Text
Public Const COLOR_WINDOWTEXT = 8 '3D dark shadow (Win95)
Public Const COLOR_CAPTIONTEXT = 9 'Text in window caption
Public Const COLOR_ACTIVEBORDER = 10 'Border of active window
Public Const COLOR_INACTIVEBORDER = 11 'Border of inactive window
Public Const COLOR_APPWORKSPACE = 12 'Background of MDI desktop
Public Const COLOR_HIGHLIGHT = 13 'Selected item background
Public Const COLOR_HIGHLIGHTTEXT = 14 'Selected menu item
Public Const COLOR_BTNFACE = 15 'Button
Public Const COLOR_BTNSHADOW = 16 '3D shading of button
Public Const COLOR_GRAYTEXT = 17 'Grey text, of zero if dithering is used.
Public Const COLOR_BTNTEXT = 18 'Button text
Public Const COLOR_INACTIVECAPTIONTEXT = 19 'Text of inactive window
Public Const COLOR_BTNHIGHLIGHT = 20 '3D highlight of button
Public Const COLOR_2NDACTIVECAPTION = 27 'Win98 only: 2nd active window color
Public Const COLOR_2NDINACTIVECAPTION = 28 'Win98 only: 2nd inactive window color
'Public Enum SysMet
'    SM_CXSCREEN = 0
'    SM_CYSCREEN = 1
'    SM_CXVSCROLL = 2
'    SM_CYHSCROLL = 3
'    SM_CYCAPTION = 4
'    SM_CXBORDER = 5
'    SM_CYBORDER = 6
'    SM_CXDLGFRAME = 7
'    SM_CYDLGFRAME = 8
'    SM_CYHTHUMB = 9
'    SM_CXHTHUMB = 10
'    SM_CXICON = 11
'    SM_CYICON = 12
'    SM_CXCURSOR = 13
'    SM_CYCURSOR = 14
'    SM_CYMENU = 15
'    SM_CXFULLSCREEN = 16
'    SM_CYFULLSCREEN = 17
'    SM_CYKANJIWINDOW = 18
'    SM_MOUSEPRESENT = 19
'    SM_CYVSCROLL = 20
'    SM_CXHSCROLL = 21
'    SM_DEBUG = 22
'    SM_SWAPBUTTON = 23
'    SM_CXMIN = 24
'    SM_CYMIN = 25
'    SM_CXSIZE = 26
'    SM_CYSIZE = 27
'    SM_CXMINTRACK = 28
'    SM_CYMINTRACK = 29
'    SM_CXDOUBLECLK = 30
'    SM_CYDOUBLECLK = 31
'    SM_CXICONSPACING = 32
'    SM_CYICONSPACING = 33
'    SM_MENUDROPALIGNMENT = 34
'    SM_PENWINDOWS = 35
'    SM_DBCSENABLED = 36
'    SM_CMOUSEBUTTONS = 37
'    SM_CMETRICS = 38
'    SM_CLEANBOOT = 39
'    SM_CXMAXIMIZED = 40
'    SM_CXMAXTRACK = 41
'    SM_CXMENUCHECK = 42
'    SM_CXMENUSIZE = 43
'    SM_CXMINIMIZED = 44
'    SM_CYMAXIMIZED = 45
'    SM_CYMAXTRACK = 46
'    SM_CYMENUCHECK = 47
'    SM_CYMENUSIZE = 48
'    SM_CYMINIMIZED = 49
'    SM_CYSMCAPTION = 50
'    SM_MIDEASTENABLED = 51
'    SM_NETWORK = 52
'    SM_SLOWMACHINE = 53
'End Enum


Public Const BDR_RAISEDOUTER As Long = &H1
Public Const BDR_SUNKENOUTER As Long = &H2
Public Const BDR_SUNKENINNER As Long = &H8
Public Const BDR_RAISEDINNER = &H4
Public Const EDGE_BUMP As Long = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED As Long = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_RAISED As Long = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN As Long = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const BF_ADJUST As Long = &H2000
Public Const BF_BOTTOM As Long = &H8

Public Const BF_DIAGONAL As Long = &H10
Public Const BF_FLAT As Long = &H4000
Public Const BF_LEFT As Long = &H1
Public Const BF_MIDDLE As Long = &H800
Public Const BF_MONO As Long = &H8000
Public Const BF_RIGHT As Long = &H4
Public Const BF_SOFT As Long = &H1000
Public Const BF_TOP As Long = &H2
Public Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_TOPLEFT As Long = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT As Long = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT As Long = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT As Long = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT As Long = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT As Long = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT As Long = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDTOPRIGHT As Long = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Public Const DFC_BUTTON As Long = 4
Public Const DFC_CAPTION As Long = 1
Public Const DFC_MENU As Long = 2
Public Const DFC_POPUPMENU As Long = 5
Public Const DFC_SCROLL As Long = 3
Public Const DFCS_BUTTON3STATE As Long = &H8
Public Const DFCS_BUTTONCHECK As Long = &H0
Public Const DFCS_BUTTONPUSH As Long = &H10
Public Const DFCS_BUTTONRADIO As Long = &H4
Public Const DFCS_BUTTONRADIOIMAGE As Long = &H1
Public Const DFCS_BUTTONRADIOMASK As Long = &H2
Public Const DFCS_CAPTIONCLOSE As Long = &H0
Public Const DFCS_CAPTIONHELP As Long = &H4
Public Const DFCS_CAPTIONMAX As Long = &H2
Public Const DFCS_CAPTIONMIN As Long = &H1
Public Const DFCS_CAPTIONRESTORE As Long = &H3
Public Const DFCS_MENUARROW As Long = &H0
Public Const DFCS_MENUBULLET As Long = &H2
Public Const DFCS_MENUCHECK As Long = &H1
Public Const DFCS_SCROLLCOMBOBOX As Long = &H5
Public Const DFCS_SCROLLDOWN As Long = &H1
Public Const DFCS_SCROLLLEFT As Long = &H2
Public Const DFCS_SCROLLRIGHT As Long = &H3
Public Const DFCS_SCROLLSIZEGRIP As Long = &H8
Public Const DFCS_SCROLLUP As Long = &H0
Public Const DFCS_ADJUSTRECT As Long = &H2000
Public Const DFCS_CHECKED As Long = &H400
Public Const DFCS_FLAT As Long = &H4000
Public Const DFCS_INACTIVE As Long = &H100
Public Const DFCS_MONO As Long = &H8000
Public Const DFCS_PUSHED As Long = &H200

'Public Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Enum Styles
    RaisedEdge = 0
    SunkenEdge = 1
    BumpedEdge = 2
    EtchedEdge = 3
    MonoButton = 4
    CheckedMonoButton = 5
    CheckedSunken = 6
End Enum
Public Type MenuSystem
    bIsTopLevel As Boolean
    sCaption As String
    bHasMenuBitmap As Boolean
    pMenuBitmap As StdPicture
    lForeColor As Long
    lBackColor As Long
    lSelectedColor As Long
    bHasCustomFont As Boolean
    fFont As StdFont
    lOwnerMenu As Long
    lID As Long
    bEnabled As Boolean
    bChecked As Boolean
    bSelected As Boolean
    bMenuIsShowing As Boolean
End Type

Public Enum MenuOptions
    bIsTopLevel = 0
    sCaption = 1
    pMenuBitmap = 2
    lForeColor = 3
    lBackColor = 4
    lSelectedColor = 5
    fFont = 6
    lOwnerMenu = 7
    bEnabled = 8
    bChecked = 9
End Enum
Public Type MenuSymbols
    lArrow As Long
    lCheck As Long
    lBullet As Long
    lArrowMask As Long
    lCheckMask As Long
    lBulletMask As Long
    lArrowSprite As Long
    lCheckSprite As Long
    lBulletSprite As Long
End Type
Public udtMSymbols As MenuSymbols
Public aMenuSystem() As MenuSystem
Public lMenuHeight As Long
Public lMenuOpen As Long
'
' Types for API Calls
'
Public Const FW_THIN = 100
Public Const FW_EXTRALIGHT = 200
Public Const FW_LIGHT = 300
Public Const FW_NORMAL = 400
Public Const FW_MEDIUM = 500
Public Const FW_SEMIBOLD = 600
Public Const FW_BOLD = 700
Public Const FW_EXTRABOLD = 800
Public Const FW_HEAVY = 900
Public Const SPI_GETNONCLIENTMETRICS = 41&
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64
Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(LF_FACESIZE - 1) As Byte
End Type

Public Type NONCLIENTMETRICS
        cbSize As Long
        iBorderWidth As Long
        iScrollWidth As Long
        iScrollHeight As Long
        iCaptionWidth As Long
        iCaptionHeight As Long
        lfCaptionFont As LOGFONT
        iSMCaptionWidth As Long
        iSMCaptionHeight As Long
        lfSMCaptionFont As LOGFONT
        iMenuWidth As Long
        iMenuHeight As Long
        lfMenuFont As LOGFONT
        lfStatusFont As LOGFONT
        lfMessageFont As LOGFONT
End Type
'
' Constants
'
Public Enum SystemFontTypesEnum
    sfte_CaptionFont
    sfte_SmallCaptionFont
    sfte_MenuFont
    sfte_StatusFont
    sfte_MessageFont
End Enum


'
' API Calls
'
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

'Public Function GetSystemFont(ByVal nFontType As SystemFontTypesEnum) As StdFont
'
'    Dim NCM As NONCLIENTMETRICS
'    Dim nRet As Long
'    Dim sFontname As String
'    Dim nLogPixY As Single
'    Dim fntNew As StdFont
'    Dim uFont As LOGFONT
'
'    NCM.cbSize = Len(NCM)
'    nRet = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0&, NCM, 0&)
'    If nRet = 0& Then
'        Exit Function
'    End If
'    nLogPixY = 1440 / Screen.TwipsPerPixelY
'    Select Case nFontType
'        Case sfte_CaptionFont
'            uFont = NCM.lfCaptionFont
'        Case sfte_SmallCaptionFont
'            uFont = NCM.lfSMCaptionFont
'        Case sfte_MenuFont
'            uFont = NCM.lfMenuFont
'        Case sfte_StatusFont
'            uFont = NCM.lfStatusFont
'        Case sfte_MessageFont
'            uFont = NCM.lfMessageFont
'        Case Else
'            Exit Function
'    End Select
'    Set fntNew = New StdFont
'    With uFont
'        sFontname = StrConv(.lfFaceName, vbUnicode)
'        sFontname = Left$(sFontname, InStr(1, sFontname, vbNullChar) - 1)
'    fntNew.Name = sFontname
'    fntNew.SIZE = Abs((.lfHeight * 72) / nLogPixY)
'        If .lfWeight >= FW_BOLD Then
'            fntNew.Bold = True
'        Else
'            fntNew.Bold = False
'        End If
'    fntNew.Italic = CBool(.lfItalic)
'    End With
'    Set GetSystemFont = fntNew
'End Function
'
'
'Private Function SetSysMetIndex(SysMetVal As SysMet) As Long
'Select Case SysMetVal
'    Case 0:
'        SetSysMetIndex = 0
'    Case 1:
'        SetSysMetIndex = 1
'    Case 2:
'        SetSysMetIndex = 2
'    Case 3:
'        SetSysMetIndex = 3
'    Case 4:
'        SetSysMetIndex = 4
'    Case 5:
'        SetSysMetIndex = 5
'    Case 6:
'        SetSysMetIndex = 6
'    Case 7:
'        SetSysMetIndex = 7
'    Case 8:
'        SetSysMetIndex = 8
'    Case 9:
'        SetSysMetIndex = 9
'    Case 10:
'        SetSysMetIndex = 10
'    Case 11:
'        SetSysMetIndex = 11
'    Case 12:
'        SetSysMetIndex = 12
'    Case 13:
'        SetSysMetIndex = 13
'    Case 14:
'        SetSysMetIndex = 14
'    Case 15:
'        SetSysMetIndex = 15
'    Case 16:
'        SetSysMetIndex = 16
'    Case 17:
'        SetSysMetIndex = 17
'    Case 18:
'        SetSysMetIndex = 18
'    Case 19:
'        SetSysMetIndex = 19
'    Case 20:
'        SetSysMetIndex = 20
'    Case 21:
'        SetSysMetIndex = 21
'    Case 22:
'        SetSysMetIndex = 22
'    Case 23:
'        SetSysMetIndex = 23
'    Case 24:
'        SetSysMetIndex = 28
'    Case 25:
'        SetSysMetIndex = 29
'    Case 26:
'        SetSysMetIndex = 30
'    Case 27:
'        SetSysMetIndex = 31
'    Case 28:
'        SetSysMetIndex = 34
'    Case 29:
'        SetSysMetIndex = 35
'    Case 30:
'        SetSysMetIndex = 36
'    Case 31:
'        SetSysMetIndex = 37
'    Case 32:
'        SetSysMetIndex = 38
'    Case 33:
'        SetSysMetIndex = 39
'    Case 34:
'        SetSysMetIndex = 40
'    Case 35:
'        SetSysMetIndex = 41
'    Case 36:
'        SetSysMetIndex = 42
'    Case 37:
'        SetSysMetIndex = 43
'    Case 38:
'        SetSysMetIndex = 44
'    Case 39:
'        SetSysMetIndex = 67
'    Case 40:
'        SetSysMetIndex = 61
'    Case 41:
'        SetSysMetIndex = 59
'    Case 42:
'        SetSysMetIndex = 71
'    Case 43:
'        SetSysMetIndex = 54
'    Case 44:
'        SetSysMetIndex = 57
'    Case 45:
'        SetSysMetIndex = 62
'    Case 46:
'        SetSysMetIndex = 60
'    Case 47:
'        SetSysMetIndex = 72
'    Case 48:
'        SetSysMetIndex = 55
'    Case 49:
'        SetSysMetIndex = 58
'    Case 50:
'        SetSysMetIndex = 51
'    Case 51:
'        SetSysMetIndex = 74
'    Case 52:
'        SetSysMetIndex = 63
'    Case 53:
'        SetSysMetIndex = 73
'End Select
'End Function
'
'Public Function GetSysMetrics(GetWhat As SysMet) As Long
'GetSysMetrics = GetSystemMetrics(SetSysMetIndex(GetWhat))
'End Function

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hWnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Public Function Mask(PicHdc As Long, ByRef lMaskhdc As Long, R As RECT)
Dim i As Integer
Dim j As Integer
Dim lColor As Long
For i = R.Top To R.Bottom
    For j = R.Left To R.Right
        If GetPixel(PicHdc, j, i) = vbWhite Then
            lColor = RGB(255, 255, 255)
        Else
            lColor = RGB(0, 0, 0)
        End If
        SetPixel lMaskhdc, j, i, lColor
    Next
Next
End Function

Public Function Sprite(PicHdc As Long, ByRef lSpriteHdc, R As RECT)
Dim i As Integer
Dim j As Integer
Dim lColor As Long
For i = R.Top To R.Bottom
    For j = R.Left To R.Right
        If GetPixel(PicHdc, j, i) = vbWhite Then
            lColor = vbBlack
        Else
            lColor = GetPixel(PicHdc, j, i)
        End If
        SetPixel lSpriteHdc, j, i, lColor
    Next
Next
End Function

Function Keyboard(ByVal idHook As Long, ByVal lParam As Long, ByVal wParam As Long) As Long
Dim m As Long
Dim j As Long
Dim n As String
Dim tmp As String
'Dim frm As frmMenu
For Each frm In Forms
    j = j + 1
Next
If j = 0 Then
    For j = (LBound(aMenuSystem) + 1) To UBound(aMenuSystem)
        With aMenuSystem(j)
            If .bIsTopLevel Then
                If InStr(1, .sCaption, "&") Then
                    tmp = .sCaption
                    tmp = Replace$(tmp, "&&", "")
                    m = InStr(1, .sCaption, "&")
                    If m = 0 Then Exit For
                    n = Mid$(tmp, m + 1, 1)
                    'if lparam = asc(
                End If
            End If
        End With
    Next
Else

End If
Keyboard = 0
End Function


