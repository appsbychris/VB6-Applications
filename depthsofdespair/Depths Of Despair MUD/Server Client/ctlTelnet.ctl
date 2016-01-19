VERSION 5.00
Begin VB.UserControl ctlTelnet 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   12
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   691
   Begin VB.Timer timCur 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ctlTelnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As udtSIZE) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As udtBoxPos) As Long
Private Type RECT
    lLeft As Long
    lTop As Long
    lRight As Long
    lBottom As Long
End Type
Private Type Dims
    lWidth As Long
    lHeight As Long
End Type
Private Type CurrentPos
    X As Long
    Y As Long
End Type
Private Type TextSettings
    lColor As Long
    bBack As Boolean
    lBack As Long
    lMode As Long
End Type
Private iFlash As Integer
Private CurR As RECT
Private udtDims As Dims
Private udtCP As CurrentPos
Private Const Esc As String = ""
Private Const BLACK           As String = "[0m[30m"

Private Const RED            As String = "[0m[31m"
Private Const bRED           As String = "[1m[31m"

Private Const GREEN          As String = "[0m[32m"
Private Const bGREEN         As String = "[1m[32m"

Private Const YELLOW         As String = "[0m[33m"
Private Const bYELLOW        As String = "[1m[33m"

Private Const BLUE           As String = "[0m[34m"
Private Const bBLUE          As String = "[1m[34m"


Private Const MAGNETA        As String = "[0m[35m"
Private Const bMAGNETA       As String = "[1m[35m"

Private Const LIGHTBLUE      As String = "[0m[36m"
Private Const bLIGHTBLUE     As String = "[1m[36m"

Private Const WHITE          As String = "[0m[37m"
Private Const bWHITE         As String = "[1m[37m"

Private Const BGRED          As String = "[0m[41m"
Private Const BGGREEN        As String = "[0m[42m"
Private Const BGYELLOW       As String = "[0m[43m"
Private Const BGBLUE         As String = "[0m[44m"
Private Const BGPURPLE       As String = "[0m[45m"
Private Const BGLIGHTBLUE    As String = "[0m[46m"

Private Const MOVELEFTONE     As String = "[1D"
Private Const ERASETOLEFT     As String = "[0K"
Private Const MOVELEFT23      As String = "[23D"
Private Const MOVERIGHTNUM    As String = "[#D"
Private Const ANSICLS         As String = "[2J"
Private Const MOVECURSOR      As String = "[25;25H"
Private Const UP_ARROW        As String = "[A"
Private Const DOWN_ARROW      As String = "[B"
Private Const RIGHT_ARROW     As String = "[C"
Private Const LEFT_ARROW      As String = "[D"

Private sTxt As String
Private tSet As TextSettings
Private lLines As Long

Private aTelnetC(24) As String
Private sTyped As String
Private sCText As String

Private Sub timCur_Timer()
Dim lSolidBrush As Long
With CurR
    .lTop = udtCP.Y + udtDims.lHeight - 1
    .lLeft = udtCP.X
    .lRight = .lLeft + udtDims.lWidth
    .lBottom = .lTop + 1
End With
With UserControl
    iFlash = iFlash + 1
    If iFlash > 1 Then iFlash = 0
    If iFlash = 0 Then
        lSolidBrush = CreateSolidBrush(vbBlack)
        FillRect .hdc, CurR, lSolidBrush
        DeleteObject lSolidBrush
    Else
        lSolidBrush = CreateSolidBrush(vbWhite)
        FillRect .hdc, CurR, lSolidBrush
        DeleteObject lSolidBrush
    End If
    .Refresh
End With
End Sub

Private Sub UserControl_Initialize()
Init
End Sub

Private Sub Init()
Dim i As Long
With udtDims
    .lHeight = UserControl.TextHeight("X")
    .lWidth = UserControl.TextWidth("W")
End With
With udtCP
    .X = 0
    .Y = 0
End With
With UserControl
    .Width = ScaleX(udtDims.lWidth, 3, 1) * 80
    .Height = ScaleY(udtDims.lHeight, 3, 1) * 25
End With
'For i = 0 To 24
'    aTelnetC(i) = "-1"
'Next
SetTextColor UserControl.hdc, vbWhite
End Sub

Private Sub CalibrateLines()
Dim m As Long
Dim n As Long
Dim Arr() As String
Dim i As Long
Dim bMult As String
Dim t As Boolean
Dim a As Boolean
Dim sESC As String
Dim p As String
Dim v As String
Dim w As Long
Dim g As Long
Dim h As Long
Dim Xx As Long
Dim lLenP As Long
Dim r As RECT
Dim lSolidBrush As Long
Dim lB As Long, uB As Long
Dim aCls As Boolean
Dim bDone As Boolean
Dim s As String
Dim lOldX As Long, lOldY As Long
lOldX = udtCP.X
lOldY = udtCP.Y
s = sTxt
SplitFast s, Arr, vbCrLf
lB = LBound(Arr)
uB = UBound(Arr)
'If uB > 24 Then lB = lB + 1
If uB > 50 Then
    s = ""
    For i = uB - 51 To uB
        s = s & Arr(i) & vbCrLf
    Next
    s = Left$(s, Len(s) - 2)
    sTxt = s
End If
SplitFast s, Arr, vbCrLf
lB = LBound(Arr)
uB = UBound(Arr)

lLines = 25
Do Until bDone
    With udtCP
        .X = 0
        .Y = 0
    End With
    lB = LBound(Arr)
    If uB - lLines > lB Then lB = uB - lLines
    s = ""
    For i = lB To uB
        If Arr(i) <> "" Then
            s = Arr(i)
            t = False
            Do
                m = InStr(1, s, Esc)
                If m <> 0 Then
                    If m > 1 Then
                        p = Left$(s, m - 1)
                        s = Mid$(s, m)
                        m = InStr(1, s, Esc)
                        udtCP.X = udtCP.X + (Len(p) * udtDims.lWidth)
                    End If
                    p = ""
                    n = 1
                    a = False
                    Do
                        p = p & Mid$(s, n, 1)
                        lLenP = Len(p)
                        Select Case Mid$(s, n, 1)
                            Case "m"
                                
                                a = True
                            Case "A"
                                p = Mid$(p, 3)
                                p = Left$(p, Len(p) - 1)
                                If p = "" Then p = "1"
                                If Val(p) < 1 Then p = 1
                                udtCP.Y = udtCP.Y - (udtDims.lHeight * Val(p))
                                If udtCP.Y < 0 Then udtCP.Y = 0
                                a = True
                            Case "B"
                                p = Mid$(p, 3)
                                p = Left$(p, Len(p) - 1)
                                If p = "" Then p = "1"
                                If Val(p) < 1 Then p = 1
                                udtCP.Y = udtCP.Y + (udtDims.lHeight * Val(p))
                                If udtCP.Y > UserControl.Height - udtDims.lHeight Then udtCP.Y = UserControl.Height - udtDims.lHeight
                                a = True
                            
                            Case "C"
                                p = Mid$(p, 3)
                                p = Left$(p, Len(p) - 1)
                                If p = "" Then p = "1"
                                If Val(p) < 1 Then p = 1
                                udtCP.X = udtCP.X + (udtDims.lWidth * Val(p))
                                If udtCP.X > UserControl.Width - udtDims.lWidth Then
                                    Do Until udtCP.X < UserControl.Width - udtDims.lWidth
                                        udtCP.X = udtCP.X - (udtDims.lWidth * 80)
                                        udtCP.Y = udtCP.Y + udtDims.lHeight
                                    Loop
                                    If udtCP.Y > UserControl.Height - udtDims.lHeight Then udtCP.Y = UserControl.Height - udtDims.lHeight
                                End If
                                a = True
                            Case "D"
                                p = Mid$(p, 3)
                                p = Left$(p, Len(p) - 1)
                                If p = "" Then p = "1"
                                If Val(p) < 1 Then p = 1
                                udtCP.X = udtCP.X - (udtDims.lWidth * Val(p))
                                If udtCP.X < 0 Then udtCP.X = 0
    '                                Do Until udtCP.X < UserControl.Width - udtDims.lWidth
    '                                    udtCP.X = udtCP.X + (udtDims.lWidth * 80)
    '                                    udtCP.Y = udtCP.Y - udtDims.lHeight
    '                                Loop
    '                                If udtCP.Y < 0 Then udtCP.Y = 0
                                'End If
                                a = True
                            Case "H", "f"
                                p = Mid$(p, 3)
                                p = Left$(p, Len(p) - 1)
                                If InStr(1, p, ";") = 0 Or p = "" Then
                                    udtCP.X = 0
                                    udtCP.Y = 0
                                Else
                                    Dim tA() As String
                                    SplitFast p, tA, ";"
                                    If UBound(tA) > 0 Then
                                        If Val(tA(0)) > 80 Then tA(0) = "80"
                                        If Val(tA(1)) > 25 Then tA(1) = "25"
                                        udtCP.X = Val(tA(1) - 1) * udtDims.lWidth
                                        udtCP.Y = Val(tA(0) - 1) * udtDims.lHeight
                                    End If
                                End If
                                a = True
                            Case "K"
                                a = True
                            Case "J"
                                p = Mid$(p, 3)
                                p = Left$(p, Len(p) - 1)
                                Select Case p
                                    Case "2"
                                        udtCP.X = 0
                                        udtCP.Y = 0
                                        UserControl.Cls
                                        If UBound(Arr) - i < 25 Then t = True: bDone = True: lLines = 25
                                End Select
                                a = True
                        End Select
                        n = n + 1
                        If n > 10 Then a = True
                        
                        If s = "" Then a = True
                    Loop Until a
                    s = Mid$(s, lLenP + 1)
                    If s = "" Then t = True
                Else
                    udtCP.X = udtCP.X + (Len(s) * udtDims.lWidth)
                    t = True
                End If
            Loop Until t
            If i <> UBound(Arr) Then
                udtCP.X = 0
                udtCP.Y = udtCP.Y + udtDims.lHeight
            End If
        Else
            If i <> UBound(Arr) Then
                udtCP.X = 0
                udtCP.Y = udtCP.Y + udtDims.lHeight
            End If
        End If
    Next
    If lLines > uB Then
        bDone = True
    ElseIf udtCP.Y < (udtDims.lHeight * 24) And uB > 25 Then
        lLines = lLines + 1
    Else
        bDone = True
    End If
Loop
udtCP.X = lOldX
udtCP.Y = lOldY
End Sub

Private Function ParseEsc(ByVal s As String)
Dim m As Long
Dim n As Long
Dim Arr() As String
Dim i As Long
Dim bMult As String
Dim t As Boolean
Dim a As Boolean
Dim sESC As String
Dim p As String
Dim v As String
Dim w As Long
Dim g As Long
Dim h As Long
Dim Xx As Long
Dim lLenP As Long
Dim r As RECT
Dim lSolidBrush As Long
Dim lB As Long, uB As Long
Dim aCls As Boolean
UserControl.Cls
'aCls = False
With udtCP
    .X = 0
    .Y = 0
End With
SplitFast s, Arr, vbCrLf
lB = LBound(Arr)
uB = UBound(Arr)
'If uB > 24 Then lB = lB + 1
If uB > 50 Then
    s = ""
    For i = uB - 51 To uB
        s = s & Arr(i) & vbCrLf
    Next
    s = Left$(s, Len(s) - 2)
    sTxt = s
End If
SplitFast s, Arr, vbCrLf
lB = LBound(Arr)
uB = UBound(Arr)
CalibrateLines
If uB - lLines > lB Then lB = uB - lLines
s = ""
For i = lB To uB
    If Arr(i) <> "" Then
        s = Arr(i)
        t = False
        Do
            m = InStr(1, s, Esc)
            If m <> 0 Then
                If m > 1 Then
                    p = Left$(s, m - 1)
                    DrawSegment p
                    s = Mid$(s, m)
                    m = InStr(1, s, Esc)
                    udtCP.X = udtCP.X + (Len(p) * udtDims.lWidth)
                End If
                p = ""
                n = 1
                a = False
                Do
                    p = p & Mid$(s, n, 1)
                    lLenP = Len(p)
                    Select Case Mid$(s, n, 1)
                        Case "m"
                            p = Mid$(p, 3)
                            p = Left$(p, Len(p) - 1)
                            Select Case Val(p)
                                Case 0 'Dim
                                    tSet.lMode = 0
                                    tSet.bBack = False
                                    tSet.lBack = 0
                                Case 1 'Bright
                                    tSet.lMode = 1
                                Case 30 'Black
                                    tSet.lColor = vbBlack
                                Case 31 'Red
                                    If tSet.lMode = 0 Then
                                        tSet.lColor = &H80&
                                    Else
                                        tSet.lColor = &HFF&
                                    End If
                                Case 32 'Green
                                    If tSet.lMode = 0 Then
                                        tSet.lColor = &H8000&
                                    Else
                                        tSet.lColor = &HFF00&
                                    End If
                                Case 33 'yellow
                                    If tSet.lMode = 0 Then
                                        tSet.lColor = &HC0C0&
                                    Else
                                        tSet.lColor = &HFFFF&
                                    End If
                                Case 34 'Blue
                                    If tSet.lMode = 0 Then
                                        tSet.lColor = &H800000
                                    Else
                                        tSet.lColor = &HFF0000
                                    End If
                                Case 35 'Magneta
                                    If tSet.lMode = 0 Then
                                        tSet.lColor = &H800080
                                    Else
                                        tSet.lColor = &HFF00FF
                                    End If
                                Case 36 'Cyan
                                    If tSet.lMode = 0 Then
                                        tSet.lColor = &H808000
                                    Else
                                        tSet.lColor = &HFFFF00
                                    End If
                                Case 37 'White
                                    If tSet.lMode = 0 Then
                                        tSet.lColor = &HC0C0C0
                                    Else
                                        tSet.lColor = &HFFFFFF
                                    End If
                                Case 40 'Black
                                    tSet.lColor = &HC0C0C0
                                    tSet.lBack = vbBlack
                                    tSet.bBack = True
                                Case 41 'red
                                    tSet.lColor = &HC0C0C0
                                    tSet.bBack = True
                                    If tSet.lMode = 0 Then
                                        tSet.lBack = &H80&
                                    Else
                                        tSet.lBack = &HFF&
                                    End If
                                Case 42 'green
                                    tSet.lColor = &HC0C0C0
                                    tSet.bBack = True
                                    If tSet.lMode = 0 Then
                                        tSet.lBack = &H8000&
                                    Else
                                        tSet.lBack = &HFF00&
                                    End If
                                Case 43 'yellow
                                    tSet.lColor = vbBlack
                                    tSet.bBack = True
                                    If tSet.lMode = 0 Then
                                        tSet.lBack = &HC0C0&
                                    Else
                                        tSet.lBack = &HFFFF&
                                    End If
                                Case 44 'blue
                                    tSet.lColor = &HC0C0C0
                                    tSet.bBack = True
                                    If tSet.lMode = 0 Then
                                        tSet.lBack = &H800000
                                    Else
                                        tSet.lBack = &HFF0000
                                    End If
                                Case 45 'magneta
                                    tSet.lColor = &HC0C0C0
                                    tSet.bBack = True
                                    If tSet.lMode = 0 Then
                                        tSet.lBack = &H800080
                                    Else
                                        tSet.lBack = &HFF00FF
                                    End If
                                Case 46 'cyan
                                    tSet.lColor = &HC0C0C0
                                    tSet.bBack = True
                                    If tSet.lMode = 0 Then
                                        tSet.lBack = &H808000
                                    Else
                                        tSet.lBack = &HFFFF00
                                    End If
                                Case 47 'white
                                    tSet.lColor = &HC0C0C0
                                    tSet.bBack = True
                                    If tSet.lMode = 0 Then
                                        tSet.lBack = &HC0C0C0
                                    Else
                                        tSet.lBack = &HFFFFFF
                                    End If
                            End Select
                            SetTextColor UserControl.hdc, tSet.lColor
                            a = True
                        Case "A"
                            p = Mid$(p, 3)
                            p = Left$(p, Len(p) - 1)
                            If p = "" Then p = "1"
                            If Val(p) < 1 Then p = 1
                            udtCP.Y = udtCP.Y - (udtDims.lHeight * Val(p))
                            If udtCP.Y < 0 Then udtCP.Y = 0
                            a = True
                        Case "B"
                            p = Mid$(p, 3)
                            p = Left$(p, Len(p) - 1)
                            If p = "" Then p = "1"
                            If Val(p) < 1 Then p = 1
                            udtCP.Y = udtCP.Y + (udtDims.lHeight * Val(p))
                            If udtCP.Y > UserControl.Height - udtDims.lHeight Then udtCP.Y = UserControl.Height - udtDims.lHeight
                            a = True
                        
                        Case "C"
                            p = Mid$(p, 3)
                            p = Left$(p, Len(p) - 1)
                            If p = "" Then p = "1"
                            If Val(p) < 1 Then p = 1
                            udtCP.X = udtCP.X + (udtDims.lWidth * Val(p))
                            If udtCP.X > UserControl.Width - udtDims.lWidth Then
                                Do Until udtCP.X < UserControl.Width - udtDims.lWidth
                                    udtCP.X = udtCP.X - (udtDims.lWidth * 80)
                                    udtCP.Y = udtCP.Y + udtDims.lHeight
                                Loop
                                If udtCP.Y > UserControl.Height - udtDims.lHeight Then udtCP.Y = UserControl.Height - udtDims.lHeight
                            End If
                            a = True
                        Case "D"
                            p = Mid$(p, 3)
                            p = Left$(p, Len(p) - 1)
                            If p = "" Then p = "1"
                            If Val(p) < 1 Then p = 1
                            udtCP.X = udtCP.X - (udtDims.lWidth * Val(p))
                            If udtCP.X < 0 Then udtCP.X = 0
'                                Do Until udtCP.X < UserControl.Width - udtDims.lWidth
'                                    udtCP.X = udtCP.X + (udtDims.lWidth * 80)
'                                    udtCP.Y = udtCP.Y - udtDims.lHeight
'                                Loop
'                                If udtCP.Y < 0 Then udtCP.Y = 0
                            'End If
                            a = True
                        Case "H", "f"
                            p = Mid$(p, 3)
                            p = Left$(p, Len(p) - 1)
                            If InStr(1, p, ";") = 0 Or p = "" Then
                                udtCP.X = 0
                                udtCP.Y = 0
                            Else
                                Dim tA() As String
                                SplitFast p, tA, ";"
                                If UBound(tA) > 0 Then
                                    If Val(tA(0)) > 25 Then tA(0) = "25"
                                    If Val(tA(1)) > 80 Then tA(1) = "80"
                                    udtCP.Y = (Val(tA(0)) - 1) * udtDims.lHeight
                                    udtCP.X = (Val(tA(1)) - 1) * udtDims.lWidth
                                End If
                            End If
                            a = True
                        Case "K"
                            p = Mid$(p, 3)
                            p = Left$(p, Len(p) - 1)
                            Select Case p
                                Case "", "0"
                                    w = UserControl.Width - udtCP.X
                                    w = w \ udtDims.lWidth
                                    v = Space$(w)
                                    DrawSegment v, True
                                Case "1"
                                    w = udtCP.X \ udtDims.lWidth
                                    v = Space$(w)
                                    w = udtCP.X
                                    udtCP.X = 0
                                    DrawSegment v, True
                                    udtCP.X = w
                                Case "2"
                                    w = udtCP.X
                                    udtCP.X = 0
                                    v = Space$(80)
                                    DrawSegment v, True
                                    udtCP.X = w
                            End Select
                            a = True
                        Case "J"
                            p = Mid$(p, 3)
                            p = Left$(p, Len(p) - 1)
                            Select Case p
                                Case "", "0"
                                    w = 25 - (udtCP.Y \ udtDims.lHeight)
                                    h = udtCP.Y
                                    Xx = udtCP.X
                                    udtCP.X = 0
                                    v = Space$(80)
                                    For g = 1 To w
                                        DrawSegment v, True
                                        udtCP.Y = udtCP.Y + udtDims.lHeight
                                    Next
                                    udtCP.Y = h
                                    udtCP.X = Xx
                                Case "1"
                                    w = (udtCP.Y \ udtDims.lHeight)
                                    h = udtCP.Y
                                    Xx = udtCP.X
                                    udtCP.X = 0
                                    v = Space$(80)
                                    For g = 1 To w
                                        DrawSegment v, True
                                        udtCP.Y = udtCP.Y - udtDims.lHeight
                                    Next
                                    udtCP.Y = h
                                    udtCP.X = Xx
                                Case "2"
                                    udtCP.X = 0
                                    udtCP.Y = 0
                                    UserControl.Cls
                            End Select
                            a = True
                            
                    End Select
                    n = n + 1
                    If n > 10 Then a = True
                    If s = "" Then a = True
                Loop Until a
                s = Mid$(s, lLenP + 1)
                If s = "" Then t = True
            Else
                DrawSegment s
                udtCP.X = udtCP.X + (Len(s) * udtDims.lWidth)
                t = True
            End If
        Loop Until t
        If i <> UBound(Arr) Then
            udtCP.X = 0
            udtCP.Y = udtCP.Y + udtDims.lHeight
        End If
    Else
        If i <> UBound(Arr) Then
            udtCP.X = 0
            udtCP.Y = udtCP.Y + udtDims.lHeight
        End If
    End If
Next

UserControl.Refresh
End Function

Public Sub DrawSegment(s As String, Optional ClearIT As Boolean = False)
Dim l As Long
Dim r As RECT
With r
    .lTop = udtCP.Y
    .lLeft = udtCP.X
    .lRight = .lLeft + (udtDims.lWidth * Len(s))
    .lBottom = .lTop + udtDims.lHeight
End With
If ClearIT Then
    If GetTextColor(UserControl.hdc) <> vbBlack Then SetTextColor UserControl.hdc, vbBlack
    l = CreateSolidBrush(vbBlack)
Else
    If GetTextColor(UserControl.hdc) <> tSet.lColor Then SetTextColor UserControl.hdc, tSet.lColor
    If tSet.bBack Then
        l = CreateSolidBrush(tSet.lBack)
    Else
        l = CreateSolidBrush(vbBlack)
    End If
End If
FillRect UserControl.hdc, r, l
DeleteObject l
TextOut UserControl.hdc, udtCP.X, udtCP.Y, s, Len(s)
End Sub


Private Sub DoTextOut(WithString As String)
Dim r As RECT
Dim lSolidBrush As String
With r
    .lTop = udtCP.Y
    .lLeft = udtCP.X
    .lRight = .lLeft + (udtDims.lWidth * Len(WithString))
    .lBottom = .lTop + udtDims.lHeight
End With
lSolidBrush = CreateSolidBrush(vbBlack)
FillRect UserControl.hdc, r, lSolidBrush
DeleteObject lSolidBrush
TextOut UserControl.hdc, udtCP.X, udtCP.Y, WithString, Len(WithString)
End Sub

Private Function Set_Color(ESC_Sequence As String) As Integer
Set_Color = 0
With UserControl
    Select Case ESC_Sequence
        Case BLACK
            SetTextColor .hdc, vbBlack
        Case RED
            SetTextColor .hdc, &H80&
        Case bRED
            SetTextColor .hdc, &HFF&
        Case GREEN
            SetTextColor .hdc, &H8000&
        Case bGREEN
            SetTextColor .hdc, &HFF00&
        Case YELLOW
            SetTextColor .hdc, &HC0C0&
        Case bYELLOW
            SetTextColor .hdc, &HFFFF&
        Case BLUE
            SetTextColor .hdc, &H800000
        Case bBLUE
            SetTextColor .hdc, &HFF0000
        Case MAGNETA
            SetTextColor .hdc, &H800080
        Case bMAGNETA
            SetTextColor .hdc, &HFF00FF
        Case LIGHTBLUE
            SetTextColor .hdc, &H808000
        Case bLIGHTBLUE
            SetTextColor .hdc, &HFFFF00
        Case WHITE
            SetTextColor .hdc, &HC0C0C0
        Case bWHITE
            SetTextColor .hdc, &HFFFFFF
        Case BGRED
            SetTextColor .hdc, &H80&
            Set_Color = 1
        Case BGGREEN
            SetTextColor .hdc, &H8000&
            Set_Color = 1
        Case BGYELLOW
            SetTextColor .hdc, &HC0C0&
            Set_Color = 1
        Case BGBLUE
            SetTextColor .hdc, &H800000
            Set_Color = 1
        Case BGPURPLE
            SetTextColor .hdc, &H400040
            Set_Color = 1
        Case BGLIGHTBLUE
            SetTextColor .hdc, &HFFFF00
            Set_Color = 1
        Case Else
            SetTextColor .hdc, &HC0C0C0
    End Select
End With
End Function

Private Sub DrawBackLine(UseText As String)
Dim orgX As Long
Dim orgY As Long
Dim ly2 As Long
Dim lx2 As Long
Dim uRECT As RECT
Dim lSolidBrush As Long
With uRECT
    .lTop = udtCP.Y
    .lLeft = udtCP.X
    .lBottom = .lTop + udtDims.lHeight
    .lRight = .lLeft + (udtDims.lWidth * Len(UseText))
End With
With udtCP
    orgY = .Y
    orgX = .X
End With
With UserControl
    lSolidBrush = CreateSolidBrush(GetTextColor(.hdc))
    FillRect .hdc, uRECT, lSolidBrush
    SetTextColor .hdc, vbWhite
    TextOut .hdc, orgX, orgY, UseText, Len(UseText)
    DeleteObject lSolidBrush
End With
With udtCP
    .Y = orgY
    .X = uRECT.lRight
End With
End Sub

Private Sub UserControl_Resize()
Init
End Sub

Public Sub TypedText(sChr As String)
'Dim s As String
'Dim i As Long
'sTyped = sTyped & sChr
'For i = 0 To 24
'    If aTelnetC(i) <> "-1" Then
'        If vba.right$(s, 2) <> vbCrLf Then
'            s = s & aTelnetC(i) & vbCrLf
'        Else
'            s = s & aTelnetC(i)
'        End If
'    End If
'Next
'If Len(s) > 2 Then s = vba.left$(s, Len(s) - 2)
'ParseEsc s '& sTyped
End Sub

Public Sub TypedBackspace()
'Dim s As String
'Dim i As Long
'If Len(sTyped) > 0 Then sTyped = vba.left$(sTyped, Len(sTyped) - 1)
'For i = 0 To 24
'    If aTelnetC(i) <> "-1" Then
'        If vba.right$(s, 2) <> vbCrLf Then
'            s = s & aTelnetC(i) & vbCrLf
'        Else
'            s = s & aTelnetC(i)
'        End If
'    End If
'Next
'If Len(s) > 2 Then s = vba.left$(s, Len(s) - 2)
'ParseEsc s '& sTyped
sTxt = Left$(sTxt, Len(sTxt) - 1)
ParseEsc sTxt
End Sub

Public Sub TypedEnter()
sTyped = ""
End Sub

Public Sub FeedMe(ByVal AString As String)
Dim i As Long
Dim a As Long
Dim b As Boolean
Dim Arr() As String
Dim s As String
'sCText = sCText & AString
'SplitFast AString, Arr, vbCrLf
'For a = LBound(Arr) To UBound(Arr)
'    If Arr(a) <> "" Then
'        For i = 0 To 24
'            If aTelnetC(i) = "-1" Then
'                aTelnetC(i) = Arr(a) & vbCrLf
'                b = True
'                Exit For
'            End If
'        Next
'        If Not b Then
'            For i = 0 To 23
'                aTelnetC(i) = aTelnetC(i + 1)
'            Next
'            aTelnetC(24) = Arr(a) & vbCrLf
'        End If
'    End If
'Next
'For i = 0 To 24
'    If aTelnetC(i) <> "-1" Then
'        If vba.right$(s, 2) <> vbCrLf Then
'            s = s & aTelnetC(i) & vbCrLf
'        Else
'            s = s & aTelnetC(i)
'        End If
'    End If
'Next
'If OccCount(sCText, vbCrLf) > 25 Then
'    sCText = VBA.Mid$(sCText, GetPosMidStr(sCText, vbCrLf, 25))
'End If
AString = Replace$(AString, Chr$(0), "")
AString = Replace$(AString, Chr$(vbKeyBack), "")
sTxt = sTxt & AString
ParseEsc sTxt '& sTyped
End Sub

Private Function OccCount(sInString As String, sCountString As String) As Long
'finds out how many a certain string appears in another string
Dim iNextOccur As Long
iNextOccur = 0
'If InStr(iNextOccur, sInString, sCountString) = 0 Then Exit Function Else DCount = DCount + 1
Do
    iNextOccur = InStr(iNextOccur + 1, sInString, sCountString)
    If iNextOccur > 0 Then OccCount = OccCount + 1 Else Exit Do
    DoEvents
Loop
End Function

Private Function GetPosMidStr(sInString As String, sSearch As String, lCount As Long) As Long
'finds out how many a certain string appears in another string
Dim iNextOccur As Long
Dim iInC As Long
iNextOccur = Len(sInString)
'If InStr(iNextOccur, sInString, sCountString) = 0 Then Exit Function Else DCount = DCount + 1
Do
    If iNextOccur > 1 Then iNextOccur = iNextOccur - 1 Else GetPosMidStr = 2: Exit Function
    iNextOccur = InStrRev(sInString, sSearch, iNextOccur)
    iInC = iInC + 1
    If iInC = lCount Then
        GetPosMidStr = iNextOccur
        Exit Do
    End If
    DoEvents
Loop
End Function

Public Sub SetCursorEnabled(TrueFalse As Boolean)
timCur.Enabled = TrueFalse
End Sub

Public Sub SetFontSize(intSize As Integer)
UserControl.FontSize = intSize
Init
ParseEsc sTxt
End Sub
